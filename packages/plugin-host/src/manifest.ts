/**
 * 清单校验。
 *
 * 原则：**宁可拒绝加载，也不做静默兜底**。插件清单错误如果被宽容处理，
 * 用户会得到"装了但没效果"的体验，比直接报错更难排查。
 */
import {
  SUPPORTED_CONTRIBUTION_KEYS,
  normalizeShortcut,
  type CommandContribution,
  type PluginManifest,
  type PluginPermission,
  type PluginTier,
  type TemplateContribution,
  type ValidationIssue,
  type ValidationResult,
} from '@mini-office/plugin-api';
import { macroKindOf, validateMacroSteps } from './macro';

const ID_PATTERN = /^[a-z0-9][a-z0-9-]*(\.[a-z0-9-]+)+$/;
const POINT_ID_PATTERN = /^[a-z0-9][a-z0-9-]*$/;
const VERSION_PATTERN = /^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/;
const DOCUMENT_KINDS = ['docx', 'xlsx', 'pptx'] as const;

const TIERS: PluginTier[] = ['declarative', 'sandboxed', 'trusted'];

const PERMISSIONS: PluginPermission[] = [
  'document:read',
  'document:write',
  'fs:read',
  'fs:write',
  'network',
  'clipboard',
  'print',
  'shell:exec',
];

export interface ParseManifestResult extends ValidationResult {
  manifest?: PluginManifest;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function readString(source: Record<string, unknown>, key: string, issues: ValidationIssue[], required = true): string {
  const value = source[key];
  if (typeof value === 'string' && value.trim()) return value.trim();
  if (required) issues.push({ path: key, message: '缺少必填字符串字段' });
  return '';
}

/** 解析 semver 三元组；无法解析时返回 null。 */
export function parseVersion(value: string): [number, number, number] | null {
  const match = /^(\d+)\.(\d+)\.(\d+)/.exec(value);
  if (!match) return null;
  return [Number(match[1]), Number(match[2]), Number(match[3])];
}

function compareVersion(a: [number, number, number], b: [number, number, number]): number {
  for (let index = 0; index < 3; index += 1) {
    if (a[index] !== b[index]) return a[index] < b[index] ? -1 : 1;
  }
  return 0;
}

/**
 * 比较两个版本号，`a > b` 返回正数，相等返回 0，无法解析返回 null。
 *
 * 预发布后缀（`1.0.0-beta.1`）按同一三元组处理，与 `satisfiesRange` 的宽松口径一致：
 * 插件系统的版本判断只需要回答"远端是否更新"，不做 npm 级别的精确排序。
 */
export function compareVersions(a: string, b: string): number | null {
  const left = parseVersion(a);
  const right = parseVersion(b);
  if (!left || !right) return null;
  return compareVersion(left, right);
}

/**
 * 极简 semver 范围判断，支持 `^`、`~`、`>=`、`>`、`<=`、`<`、`=` 与空格连接。
 * 不引入 semver 依赖，因为插件系统必须保持零额外体积。
 */
export function satisfiesRange(version: string, range: string): boolean {
  const target = parseVersion(version);
  if (!target) return false;
  const clauses = range.trim().split(/\s+/).filter(Boolean);
  if (clauses.length === 0) return true;
  return clauses.every((clause) => {
    const match = /^(\^|~|>=|<=|>|<|=)?(.+)$/.exec(clause);
    if (!match) return false;
    const operator = match[1] ?? '=';
    const bound = parseVersion(match[2]);
    if (!bound) return false;
    const diff = compareVersion(target, bound);
    switch (operator) {
      case '^':
        // 只允许同一 major，且不低于下界。
        return target[0] === bound[0] && diff >= 0;
      case '~':
        return target[0] === bound[0] && target[1] === bound[1] && diff >= 0;
      case '>=':
        return diff >= 0;
      case '>':
        return diff > 0;
      case '<=':
        return diff <= 0;
      case '<':
        return diff < 0;
      default:
        return diff === 0;
    }
  });
}

/**
 * 校验命令贡献点。
 *
 * 命令在 v1 只有"宏"一种形态，所以这里同时承担了三件事：
 * id 唯一、步骤合法、整条命令的类型可推断且一致。
 */
function validateCommands(
  raw: unknown[] | undefined,
  issues: ValidationIssue[],
  warnings: ValidationIssue[],
): CommandContribution[] {
  if (!raw) return [];
  const result: CommandContribution[] = [];
  const seen = new Set<string>();

  raw.forEach((item, index) => {
    const path = `contributes.commands[${index}]`;
    if (!isPlainObject(item)) {
      issues.push({ path, message: '命令必须是对象' });
      return;
    }
    const id = typeof item.id === 'string' ? item.id.trim() : '';
    if (!id) {
      issues.push({ path: `${path}.id`, message: '缺少命令 id' });
    } else if (!POINT_ID_PATTERN.test(id)) {
      issues.push({ path: `${path}.id`, message: 'id 只能包含小写字母、数字与连字符' });
    } else if (seen.has(id)) {
      issues.push({ path: `${path}.id`, message: `命令 id 重复：${id}` });
    } else {
      seen.add(id);
    }

    const title = typeof item.title === 'string' ? item.title.trim() : '';
    if (!title) issues.push({ path: `${path}.title`, message: '缺少命令标题' });

    const stepIssues = validateMacroSteps(item.steps, `${path}.steps`);
    issues.push(...stepIssues);

    if (stepIssues.length === 0) {
      const kind = macroKindOf(item.steps as never);
      if (!kind) {
        issues.push({ path: `${path}.steps`, message: '无法推断目标文档类型，或步骤混用了多种类型' });
      }
    }

    if (item.shortcut !== undefined && typeof item.shortcut !== 'string') {
      issues.push({ path: `${path}.shortcut`, message: 'shortcut 必须是字符串，如 Mod+Shift+L' });
    } else if (typeof item.shortcut === 'string' && !normalizeShortcut(item.shortcut)) {
      // 不阻断安装：快捷键写错不影响命令本身可用，但必须让用户在卡片上看到原因。
      warnings.push({ path: `${path}.shortcut`, message: '快捷键格式无效（需形如 Mod+Shift+L），该命令将只能从命令面板触发' });
    }

    if (issues.some((issue) => issue.path.startsWith(path))) return;

    const command: CommandContribution = { id, title, steps: item.steps as CommandContribution['steps'] };
    if (typeof item.description === 'string') command.description = item.description;
    if (typeof item.keywords === 'string') command.keywords = item.keywords;
    if (typeof item.shortcut === 'string') command.shortcut = item.shortcut;
    result.push(command);
  });

  return result;
}

/** 校验模板贡献点。文件是否真的在包里，由 installer 在解包后检查。 */
function validateTemplates(raw: unknown[] | undefined, issues: ValidationIssue[]): TemplateContribution[] {
  if (!raw) return [];
  const result: TemplateContribution[] = [];
  const seen = new Set<string>();

  raw.forEach((item, index) => {
    const path = `contributes.templates[${index}]`;
    if (!isPlainObject(item)) {
      issues.push({ path, message: '模板必须是对象' });
      return;
    }
    const id = typeof item.id === 'string' ? item.id.trim() : '';
    if (!id) {
      issues.push({ path: `${path}.id`, message: '缺少模板 id' });
    } else if (!POINT_ID_PATTERN.test(id)) {
      issues.push({ path: `${path}.id`, message: 'id 只能包含小写字母、数字与连字符' });
    } else if (seen.has(id)) {
      issues.push({ path: `${path}.id`, message: `模板 id 重复：${id}` });
    } else {
      seen.add(id);
    }

    const label = typeof item.label === 'string' ? item.label.trim() : '';
    if (!label) issues.push({ path: `${path}.label`, message: '缺少模板名称' });

    const kind = typeof item.kind === 'string' ? item.kind : '';
    if (!(DOCUMENT_KINDS as readonly string[]).includes(kind)) {
      issues.push({ path: `${path}.kind`, message: `kind 必须是 ${DOCUMENT_KINDS.join(' / ')}` });
    }

    const file = typeof item.file === 'string' ? item.file.trim() : '';
    if (!file) {
      issues.push({ path: `${path}.file`, message: '缺少模板文件路径' });
    } else if (file.startsWith('/') || file.includes('\\') || file.split('/').includes('..')) {
      issues.push({ path: `${path}.file`, message: '模板路径必须是包内相对路径' });
    } else if (kind && !file.toLowerCase().endsWith(`.${kind}`)) {
      issues.push({ path: `${path}.file`, message: `模板文件扩展名应与 kind（.${kind}）一致` });
    }

    if (issues.some((issue) => issue.path.startsWith(path))) return;

    const template: TemplateContribution = {
      id,
      label,
      kind: kind as TemplateContribution['kind'],
      file,
    };
    if (typeof item.description === 'string') template.description = item.description;
    result.push(template);
  });

  return result;
}

/**
 * 校验原始清单对象。
 *
 * @param raw 从 manifest.json 解析出的对象
 * @param currentAppVersion 宿主版本，用于 engines.miniOffice 检查；缺省则跳过
 */
export function validateManifest(raw: unknown, currentAppVersion?: string): ParseManifestResult {
  const issues: ValidationIssue[] = [];
  const warnings: ValidationIssue[] = [];

  if (!isPlainObject(raw)) {
    return { ok: false, issues: [{ path: '', message: 'manifest 必须是 JSON 对象' }], warnings };
  }

  const id = readString(raw, 'id', issues);
  if (id && !ID_PATTERN.test(id)) {
    issues.push({ path: 'id', message: 'id 需为反向域名格式，如 mini-office.theme-classic' });
  }

  const name = readString(raw, 'name', issues);
  const version = readString(raw, 'version', issues);
  if (version && !VERSION_PATTERN.test(version)) {
    issues.push({ path: 'version', message: 'version 需为语义化版本，如 1.0.0' });
  }
  const publisher = readString(raw, 'publisher', issues);

  const tierValue = raw.tier;
  const tier = typeof tierValue === 'string' && TIERS.includes(tierValue as PluginTier) ? (tierValue as PluginTier) : undefined;
  if (!tier) issues.push({ path: 'tier', message: 'tier 必须是 declarative / sandboxed / trusted' });

  const engines = isPlainObject(raw.engines) ? raw.engines : {};
  const miniOffice = typeof engines.miniOffice === 'string' ? engines.miniOffice.trim() : '';
  const api = typeof engines.api === 'string' ? engines.api.trim() : '';
  if (!miniOffice) issues.push({ path: 'engines.miniOffice', message: '缺少兼容的 Mini Office 版本范围' });
  if (!api) issues.push({ path: 'engines.api', message: '缺少兼容的插件 API 版本范围' });

  if (currentAppVersion && miniOffice && !satisfiesRange(currentAppVersion, miniOffice)) {
    issues.push({
      path: 'engines.miniOffice',
      message: `该插件要求 Mini Office ${miniOffice}，当前为 ${currentAppVersion}`,
    });
  }

  const permissionsRaw = Array.isArray(raw.permissions) ? raw.permissions : [];
  const permissions: PluginPermission[] = [];
  for (const item of permissionsRaw) {
    if (typeof item !== 'string' || !PERMISSIONS.includes(item as PluginPermission)) {
      issues.push({ path: 'permissions', message: `未知权限：${String(item)}` });
      continue;
    }
    permissions.push(item as PluginPermission);
  }

  // 声明式插件不执行代码，因此不允许声明任何权限。
  if (tier === 'declarative' && permissions.length > 0) {
    issues.push({ path: 'permissions', message: '声明式插件不执行代码，权限必须为空' });
  }

  const contributesRaw = isPlainObject(raw.contributes) ? raw.contributes : {};

  // 未实现的贡献点必须拒绝，而不是静默忽略：否则插件作者会以为生效了，
  // 用户会得到"装了但没效果"，排查成本远高于安装时报错。
  for (const key of Object.keys(contributesRaw)) {
    if (!(SUPPORTED_CONTRIBUTION_KEYS as readonly string[]).includes(key)) {
      issues.push({
        path: `contributes.${key}`,
        message: `当前版本不支持贡献点 ${key}，可用：${SUPPORTED_CONTRIBUTION_KEYS.join(' / ')}`,
      });
    }
  }

  const contributes = {
    themes: Array.isArray(contributesRaw.themes) ? contributesRaw.themes : undefined,
    commands: Array.isArray(contributesRaw.commands) ? contributesRaw.commands : undefined,
    templates: Array.isArray(contributesRaw.templates) ? contributesRaw.templates : undefined,
  };
  const contributionCount =
    (contributes.themes?.length ?? 0) + (contributes.commands?.length ?? 0) + (contributes.templates?.length ?? 0);
  if (contributionCount === 0) {
    issues.push({ path: 'contributes', message: '插件未声明任何贡献点，安装后不会有任何效果' });
  }

  // 声明式插件不能带权限，宏因此也不能有"读文件""访问网络"这类能力。
  const commands = validateCommands(contributes.commands, issues, warnings);
  const templates = validateTemplates(contributes.templates, issues);

  if (issues.length > 0) {
    return { ok: false, issues, warnings };
  }

  const manifest: PluginManifest = {
    id,
    name,
    version,
    publisher,
    tier: tier as PluginTier,
    engines: { miniOffice, api },
    permissions,
    contributes: {
      themes: contributes.themes as PluginManifest['contributes']['themes'],
      commands: commands.length > 0 ? commands : undefined,
      templates: templates.length > 0 ? templates : undefined,
    },
  };

  if (typeof raw.description === 'string') manifest.description = raw.description;
  if (typeof raw.homepage === 'string') manifest.homepage = raw.homepage;
  if (typeof raw.license === 'string') manifest.license = raw.license;
  if (typeof raw.icon === 'string') manifest.icon = raw.icon;

  return { ok: true, issues, warnings, manifest };
}

/** 从 manifest.json 字节解析。 */
export function parseManifestBytes(bytes: Uint8Array, currentAppVersion?: string): ParseManifestResult {
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return { ok: false, issues: [{ path: '', message: 'manifest.json 不是合法的 UTF-8 文本' }], warnings: [] };
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (error) {
    return {
      ok: false,
      issues: [{ path: '', message: `manifest.json 不是合法 JSON：${(error as Error).message}` }],
      warnings: [],
    };
  }
  return validateManifest(parsed, currentAppVersion);
}
