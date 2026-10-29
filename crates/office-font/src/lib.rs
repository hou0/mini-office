use office_core::OfficeError;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{Duration, Instant};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FontInfo {
    pub family: String,
    pub style: String,
    pub path: String,
    pub has_chinese: bool,
    /// 功能：同一字体族的其他名称（中英文/不同平台记录），用于缺失字体判断；时间：2026-10-29；作者：lq。
    #[serde(default)]
    pub aliases: Vec<String>,
}

const CJK_FONT_HINTS: &[&str] = &[
    "songti",
    "heiti",
    "kaiti",
    "fangsong",
    "pingfang",
    "hiragino",
    "simsun",
    "simhei",
    "kaiti",
    "yahei",
    "dengxian",
    "source han",
    "noto sans cjk",
    "noto serif cjk",
    "wenquanyi",
    "微软雅黑",
    "宋体",
    "黑体",
    "楷体",
    "仿宋",
    "等线",
    "思源",
];

pub fn font_directories() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = Vec::new();
    #[cfg(target_os = "macos")]
    {
        dirs.push(PathBuf::from("/System/Library/Fonts"));
        dirs.push(PathBuf::from("/System/Library/Fonts/Supplemental"));
        dirs.push(PathBuf::from("/Library/Fonts"));
        if let Some(home) = dirs::home_dir() {
            dirs.push(home.join("Library/Fonts"));
        }
    }
    #[cfg(target_os = "windows")]
    {
        if let Some(windows) = std::env::var_os("WINDIR") {
            dirs.push(PathBuf::from(windows).join("Fonts"));
        }
        if let Some(local) = dirs::data_local_dir() {
            dirs.push(local.join("Microsoft/Windows/Fonts"));
        }
    }
    #[cfg(target_os = "linux")]
    {
        dirs.push(PathBuf::from("/usr/share/fonts"));
        dirs.push(PathBuf::from("/usr/local/share/fonts"));
        if let Some(home) = dirs::home_dir() {
            dirs.push(home.join(".fonts"));
            dirs.push(home.join(".local/share/fonts"));
        }
    }
    dirs
}

fn has_chinese_name(family: &str) -> bool {
    family.chars().any(|character| {
        matches!(character, '\u{3400}'..='\u{9fff}' | '\u{f900}'..='\u{faff}')
    }) || {
        let lower = family.to_ascii_lowercase();
        CJK_FONT_HINTS
            .iter()
            .any(|hint| lower.contains(&hint.to_ascii_lowercase()))
    }
}

/// 功能：读取字体文件内全部家族名，中文名优先作为展示名，其余作为别名；时间：2026-10-29；作者：lq。
///
/// 背景：中文字体常同时带 Windows(英文/中文) 与 Macintosh 两套 name 记录。旧实现只取第一条，
/// 结果展示成 "FZXiaoBiaoSong-B05S" 这类英文名，文档里的 "方正小标宋简体" 就被误判为缺失字体。
fn face_names(data: &[u8], index: u32) -> Option<(String, String, Vec<String>)> {
    let face = ttf_parser::Face::parse(data, index).ok()?;
    let mut families: Vec<String> = Vec::new();
    let mut style: Option<String> = None;
    for name in face.names() {
        // 只认家族与排版家族两类记录，子家族（Regular/Bold）单独作为 style。
        if name.name_id == ttf_parser::name_id::FAMILY
            || name.name_id == ttf_parser::name_id::TYPOGRAPHIC_FAMILY
        {
            if let Some(value) = name.to_string() {
                let trimmed = value.trim();
                if !trimmed.is_empty() && !families.iter().any(|item| item == trimmed) {
                    families.push(trimmed.to_string());
                }
            }
        }
        if name.name_id == ttf_parser::name_id::SUBFAMILY && style.is_none() {
            style = name.to_string();
        }
    }
    if families.is_empty() {
        return None;
    }
    // 展示名优先取含中日韩字符的名称，找不到再退回第一条记录。
    let primary = families
        .iter()
        .find(|family| has_chinese_name(family))
        .cloned()
        .unwrap_or_else(|| families[0].clone());
    let aliases = families
        .into_iter()
        .filter(|family| family != &primary)
        .collect::<Vec<_>>();
    Some((primary, style.unwrap_or_else(|| "Regular".to_string()), aliases))
}

fn parse_font_file(path: &Path, output: &mut Vec<FontInfo>, seen: &mut HashSet<String>) {
    let data = match fs::read(path) {
        Ok(data) => data,
        Err(_) => return,
    };
    let count = ttf_parser::fonts_in_collection(&data).unwrap_or(1);
    for index in 0..count {
        if let Some((family, style, aliases)) = face_names(&data, index) {
            let key = format!("{}::{}", family.to_ascii_lowercase(), style.to_ascii_lowercase());
            if !seen.insert(key) {
                continue;
            }
            output.push(FontInfo {
                has_chinese: has_chinese_name(&family) || aliases.iter().any(|alias| has_chinese_name(alias)),
                family,
                style,
                path: path.to_string_lossy().to_string(),
                aliases,
            });
        }
    }
}

fn walk(directory: &Path, depth: usize, output: &mut Vec<FontInfo>, seen: &mut HashSet<String>) {
    if depth > 2 || output.len() > 4_000 {
        return;
    }
    let entries = match fs::read_dir(directory) {
        Ok(entries) => entries,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            walk(&path, depth + 1, output, seen);
            continue;
        }
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if matches!(extension.as_str(), "ttf" | "otf" | "ttc" | "otc") {
            parse_font_file(&path, output, seen);
        }
    }
}

pub fn list_fonts() -> Vec<FontInfo> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for directory in font_directories() {
        walk(&directory, 0, &mut output, &mut seen);
    }
    output.sort_by(|a, b| {
        b.has_chinese
            .cmp(&a.has_chinese)
            .then_with(|| a.family.to_lowercase().cmp(&b.family.to_lowercase()))
    });
    output
}

/// 功能：字体缓存，带过期时间与手动失效，用户新装字体后无需重启应用；时间：2026-10-29；作者：lq。
///
/// 旧实现用 `OnceLock` 永久缓存：用户在系统里安装字体后，应用仍按旧清单判断缺失字体，
/// 于是“自己安装的字体”一直显示兼容性告警。改为 30 秒 TTL + 显式失效。
const FONT_CACHE_TTL: Duration = Duration::from_secs(30);

static FONT_CACHE: Mutex<Option<(Instant, Vec<FontInfo>)>> = Mutex::new(None);

pub fn cached_fonts() -> Vec<FontInfo> {
    let mut guard = match FONT_CACHE.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    };
    if let Some((stamp, fonts)) = guard.as_ref() {
        if stamp.elapsed() < FONT_CACHE_TTL {
            return fonts.clone();
        }
    }
    let fonts = list_fonts();
    *guard = Some((Instant::now(), fonts.clone()));
    fonts
}

/// 功能：清空字体缓存，供“重新扫描字体”与字体目录变化后立即生效；时间：2026-10-29；作者：lq。
pub fn invalidate_fonts() {
    let mut guard = match FONT_CACHE.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    };
    *guard = None;
}

/// 功能：强制重新扫描并回填缓存；时间：2026-10-29；作者：lq。
pub fn refresh_fonts() -> Vec<FontInfo> {
    invalidate_fonts();
    cached_fonts()
}

pub fn font_families() -> Vec<String> {
    let mut families = Vec::new();
    let mut seen = HashSet::new();
    for font in list_fonts() {
        let key = font.family.to_ascii_lowercase();
        if seen.insert(key) {
            families.push(font.family);
        }
    }
    families
}

pub fn require_font(family: &str) -> Result<FontInfo, OfficeError> {
    let lower = family.to_ascii_lowercase();
    list_fonts()
        .into_iter()
        .find(|font| font.family.to_ascii_lowercase() == lower)
        .ok_or_else(|| OfficeError::Other(format!("缺少字体：{family}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scans_system_font_directories() {
        let fonts = list_fonts();
        assert!(!fonts.is_empty(), "system should provide at least one font");
        // Chinese coverage depends on the host's installed fonts.
    }

    #[test]
    fn directories_are_platform_specific() {
        assert!(!font_directories().is_empty());
    }
}
