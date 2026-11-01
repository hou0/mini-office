// 功能：提供 PDF、Markdown、HTML 的统一只读预览视图，采用浏览器能力与零依赖 Markdown 渲染控制体积；时间：2026-10-31；作者：lq。
import { useEffect, useMemo, useState, type ReactNode } from 'react';
import type { DocumentSession, PreviewDocumentData } from '@mini-office/document-model';

function inlineMarkdown(value: string): ReactNode[] {
  const nodes: ReactNode[] = [];
  const pattern = /(`[^`]+`|\*\*[^*]+\*\*|\*[^*]+\*|\[([^\]]+)\]\(([^)]+)\))/g;
  let cursor = 0;
  for (const match of value.matchAll(pattern)) {
    const index = match.index ?? 0;
    if (index > cursor) nodes.push(value.slice(cursor, index));
    const token = match[0];
    if (token.startsWith('`')) nodes.push(<code key={`${index}-code`} className="rounded bg-[var(--mo-surface-2)] px-1 py-0.5 text-[0.9em]">{token.slice(1, -1)}</code>);
    else if (token.startsWith('**')) nodes.push(<strong key={`${index}-strong`}>{token.slice(2, -2)}</strong>);
    else if (token.startsWith('*')) nodes.push(<em key={`${index}-em`}>{token.slice(1, -1)}</em>);
    else nodes.push(<a key={`${index}-link`} href={match[3]} target="_blank" rel="noreferrer" className="text-[var(--mo-accent)] underline">{match[2]}</a>);
    cursor = index + token.length;
  }
  if (cursor < value.length) nodes.push(value.slice(cursor));
  return nodes;
}

type DiagramNode = { id: string; label: string; depth: number; shape: 'round' | 'diamond' | 'pill' };
type DiagramEdge = { from: string; to: string; label?: string };

// 功能：解析常见 Mermaid 脑图与流程图语法，使用固定布局生成轻量 SVG；时间：2026-10-31；作者：lq。
function parseDiagram(source: string): { nodes: DiagramNode[]; edges: DiagramEdge[]; direction: 'lr' | 'tb' } | null {
  const lines = source.split(/\r?\n/).map((line) => line.trimEnd()).filter((line) => line.trim() && !line.trimStart().startsWith('%%'));
  if (!lines.length) return null;
  const first = lines[0].trim();
  const isMindmap = /^mindmap$/i.test(first);
  if (!isMindmap && !/^(flowchart|graph)\s+(TD|TB|LR|RL|BT)\b/i.test(first)) return null;
  const direction = /\b(LR|RL)\b/i.test(first) ? 'lr' : 'tb';
  const nodes = new Map<string, DiagramNode>();
  const edges: DiagramEdge[] = [];
  const addNode = (id: string, label?: string, shape?: DiagramNode['shape'], depth = 0) => {
    if (!nodes.has(id)) nodes.set(id, { id, label: label?.trim() || id, shape: shape ?? 'round', depth });
    else if (label?.trim()) nodes.get(id)!.label = label.trim();
  };
  if (isMindmap) {
    const stack: Array<{ indent: number; id: string }> = [];
    lines.slice(1).forEach((line, index) => {
      const indent = line.search(/\S|$/);
      const raw = line.trim();
      const match = raw.match(/^([\w-]+)?\s*(?:\(\((.*?)\)\)|\[([^\]]+)\]|\{([^}]+)\}|(.+))$/);
      if (!match) return;
      const id = match[1] || `mind-${index}`;
      const label = match[2] || match[3] || match[4] || match[5] || id;
      const shape = match[4] ? 'diamond' : match[3] ? 'pill' : 'round';
      while (stack.length && indent <= stack[stack.length - 1].indent) stack.pop();
      const depth = stack.length;
      addNode(id, label, shape, depth);
      if (stack.length) edges.push({ from: stack[stack.length - 1].id, to: id });
      stack.push({ indent, id });
    });
  } else {
    const body = lines.slice(1).join('\n');
    const edgePattern = /([\w-]+)(?:\s*(?:\[([^\]]+)\]|\{([^}]+)\}|\(([^)]+)\)))?\s*(?:-->|---|-.->)\s*(?:\|([^|]+)\|\s*)?([\w-]+)(?:\s*(?:\[([^\]]+)\]|\{([^}]+)\}|\(([^)]+)\)))?/g;
    for (const match of body.matchAll(edgePattern)) {
      const [, from, fromBox, fromDiamond, fromRound, edgeLabel, to, toBox, toDiamond, toRound] = match;
      addNode(from, fromBox || fromDiamond || fromRound, fromDiamond ? 'diamond' : fromBox ? 'pill' : 'round');
      addNode(to, toBox || toDiamond || toRound, toDiamond ? 'diamond' : toBox ? 'pill' : 'round');
      edges.push({ from, to, label: edgeLabel?.trim() });
    }
    const standalone = /([A-Za-z][\w-]*)(?:\s*\[([^\]]+)\]|\s*\{([^}]+)\}|\s*\(([^)]+)\))?/g;
    for (const match of body.matchAll(standalone)) {
      const [, id, box, diamond, round] = match;
      if (id === 'subgraph' || id === 'end') continue;
      addNode(id, box || diamond || round, diamond ? 'diamond' : box ? 'pill' : 'round');
    }
  }
  return nodes.size ? { nodes: [...nodes.values()], edges, direction } : null;
}

function DiagramPreview({ source }: { source: string }) {
  const diagramType = source.split(/\r?\n/).map((line) => line.trim()).find((line) => line && !line.startsWith('%')) ?? '';
  if (/^sequenceDiagram$/i.test(diagramType)) return <SequenceDiagramPreview source={source} />;
  if (/^erDiagram$/i.test(diagramType)) return <EntityDiagramPreview source={source} />;
  const diagram = parseDiagram(source);
  if (!diagram) return <pre className="my-3 overflow-auto rounded-lg bg-[var(--mo-surface-2)] p-4 text-xs">{source}</pre>;
  const positions = new Map<string, { x: number; y: number }>();
  diagram.nodes.forEach((node, index) => {
    positions.set(node.id, diagram.direction === 'lr'
      ? { x: 110 + node.depth * 220, y: 55 + index * 72 }
      : { x: 120 + (index % 4) * 220, y: 55 + node.depth * 100 + Math.floor(index / 4) * 72 });
  });
  const width = diagram.direction === 'lr' ? Math.max(700, 180 + Math.max(...diagram.nodes.map((node) => node.depth), 0) * 220) : 900;
  const height = Math.max(150, 100 + Math.max(...diagram.nodes.map((node) => positions.get(node.id)?.y ?? 0)));
  return <div className="my-4 overflow-auto rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/35 p-3">
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} role="img" aria-label="Markdown 图表预览" className="min-w-[620px]">
      <defs><marker id="mo-arrow" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--mo-accent)" /></marker></defs>
      {diagram.edges.map((edge, index) => {
        const from = positions.get(edge.from); const to = positions.get(edge.to);
        if (!from || !to) return null;
        const labelX = (from.x + to.x) / 2; const labelY = (from.y + to.y) / 2 - 6;
        return <g key={`edge-${index}`}><line x1={from.x} y1={from.y} x2={to.x} y2={to.y} stroke="var(--mo-accent)" strokeWidth="1.5" markerEnd="url(#mo-arrow)" />{edge.label ? <text x={labelX} y={labelY} textAnchor="middle" className="fill-[var(--mo-text-muted)] text-[11px]">{edge.label}</text> : null}</g>;
      })}
      {diagram.nodes.map((node) => {
        const point = positions.get(node.id)!; const wide = Math.max(92, Math.min(180, node.label.length * 13 + 28));
        return <g key={node.id} transform={`translate(${point.x},${point.y})`}>
          {node.shape === 'diamond' ? <path d={`M 0 -25 L ${wide / 2} 0 L 0 25 L ${-wide / 2} 0 Z`} fill="var(--mo-surface)" stroke="var(--mo-accent)" /> : <rect x={-wide / 2} y={-25} width={wide} height="50" rx={node.shape === 'pill' ? 25 : 8} fill="var(--mo-surface)" stroke="var(--mo-accent)" />}
          <text textAnchor="middle" dominantBaseline="central" className="fill-[var(--mo-text-primary)] text-[12px]">{node.label}</text>
        </g>;
      })}
    </svg>
  </div>;
}

// 功能：渲染常见 Mermaid 时序图，满足接口调用和业务流程文档的快速阅读；时间：2026-10-31；作者：lq。
function SequenceDiagramPreview({ source }: { source: string }) {
  const participants = new Map<string, string>();
  const messages: Array<{ from: string; to: string; text: string }> = [];
  source.split(/\r?\n/).forEach((line) => {
    const clean = line.trim();
    const participant = clean.match(/^participant\s+(\w+)(?:\s+as\s+(.+))?$/i);
    const message = clean.match(/^(\w+)\s*[-=]+>>?\s*(\w+)\s*:\s*(.+)$/);
    if (participant) participants.set(participant[1], participant[2] || participant[1]);
    else if (message) { participants.set(message[1], message[1]); participants.set(message[2], message[2]); messages.push({ from: message[1], to: message[2], text: message[3] }); }
  });
  const names = [...participants.keys()];
  const width = Math.max(680, names.length * 180 + 80);
  return <div className="my-4 overflow-auto rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/35 p-3"><svg viewBox={`0 0 ${width} ${Math.max(220, 120 + messages.length * 58)}`} width="100%" height={Math.max(220, 120 + messages.length * 58)} className="min-w-[620px]" role="img" aria-label="Markdown 时序图预览">
    <defs><marker id="mo-sequence-arrow" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--mo-accent)" /></marker></defs>
    {names.map((name, index) => { const x = 70 + index * 180; return <g key={name}><rect x={x - 56} y="12" width="112" height="34" rx="8" fill="var(--mo-surface)" stroke="var(--mo-accent)" /><text x={x} y="34" textAnchor="middle" className="fill-[var(--mo-text-primary)] text-[12px]">{participants.get(name)}</text><line x1={x} y1="46" x2={x} y2={90 + messages.length * 58} stroke="var(--mo-line-strong)" strokeDasharray="4 4" /></g>; })}
    {messages.map((message, index) => { const from = 70 + names.indexOf(message.from) * 180; const to = 70 + names.indexOf(message.to) * 180; const y = 112 + index * 58; return <g key={`${message.from}-${message.to}-${index}`}><line x1={from} y1={y} x2={to} y2={y} stroke="var(--mo-accent)" markerEnd="url(#mo-sequence-arrow)" /><text x={(from + to) / 2} y={y - 8} textAnchor="middle" className="fill-[var(--mo-text-muted)] text-[11px]">{message.text}</text></g>; })}
  </svg></div>;
}

// 功能：渲染常见 Mermaid 实体关系图，使用实体卡片和关系线保持体积可控；时间：2026-10-31；作者：lq。
function EntityDiagramPreview({ source }: { source: string }) {
  const entities = new Map<string, string>();
  const relations: Array<{ from: string; to: string; label: string }> = [];
  source.split(/\r?\n/).forEach((line) => {
    const clean = line.trim();
    const relation = clean.match(/^(\w+)\s+[^\s]+\s+(\w+)\s*:\s*(.+)$/);
    if (!relation) return;
    entities.set(relation[1], relation[1]); entities.set(relation[2], relation[2]); relations.push({ from: relation[1], to: relation[2], label: relation[3] });
  });
  const names = [...entities.keys()]; const width = Math.max(680, Math.ceil(names.length / 3) * 230); const height = Math.max(180, Math.ceil(names.length / 3) * 120);
  const point = (index: number) => ({ x: 100 + (index % 3) * 220, y: 55 + Math.floor(index / 3) * 120 });
  return <div className="my-4 overflow-auto rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/35 p-3"><svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} className="min-w-[620px]" role="img" aria-label="Markdown 实体关系图预览">
    <defs><marker id="mo-er-arrow" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--mo-accent)" /></marker></defs>
    {relations.map((relation, index) => { const from = point(names.indexOf(relation.from)); const to = point(names.indexOf(relation.to)); return <g key={`relation-${index}`}><line x1={from.x} y1={from.y} x2={to.x} y2={to.y} stroke="var(--mo-accent)" markerEnd="url(#mo-er-arrow)" /><text x={(from.x + to.x) / 2} y={(from.y + to.y) / 2 - 6} textAnchor="middle" className="fill-[var(--mo-text-muted)] text-[11px]">{relation.label}</text></g>; })}
    {names.map((name, index) => { const { x, y } = point(index); return <g key={name}><rect x={x - 72} y={y - 24} width="144" height="48" rx="8" fill="var(--mo-surface)" stroke="var(--mo-accent)" /><text x={x} y={y + 4} textAnchor="middle" className="fill-[var(--mo-text-primary)] text-[12px]">{name}</text></g>; })}
  </svg></div>;
}

function MarkdownPreview({ text }: { text: string }) {
  // 功能：隐藏 Markdown 内嵌 HTML 注释，避免文档作者备注污染阅读视图；时间：2026-10-31；作者：lq。
  const blocks = text.replaceAll('\r\n', '\n').replace(/<!--[\s\S]*?-->/g, '').split('\n');
  const elements: React.ReactNode[] = [];
  let list: string[] = [];
  const flushList = () => {
    if (!list.length) return;
    elements.push(<ul key={`list-${elements.length}`} className="my-3 list-disc space-y-1 pl-6">{list.map((item, index) => <li key={index}>{inlineMarkdown(item)}</li>)}</ul>);
    list = [];
  };
  for (let index = 0; index < blocks.length; index += 1) {
    const line = blocks[index];
    const fence = line.match(/^```\s*(mermaid|mindmap|flowchart|graph)?\s*$/i);
    if (fence) {
      flushList();
      const diagramLines: string[] = [];
      index += 1;
      while (index < blocks.length && !/^```\s*$/.test(blocks[index])) diagramLines.push(blocks[index++]);
      const diagramSource = diagramLines.join('\n');
      const inferredDiagram = diagramSource.split(/\r?\n/).map((item) => item.trim()).find((item) => item && !item.startsWith('%')) ?? '';
      if ((fence[1] && /^(mermaid|mindmap|flowchart|graph)$/i.test(fence[1])) || /^(mindmap|flowchart|graph|sequenceDiagram|erDiagram)$/i.test(inferredDiagram)) elements.push(<DiagramPreview key={`diagram-${index}`} source={diagramSource} />);
      else elements.push(<pre key={`code-${index}`} className="my-3 overflow-auto rounded-lg bg-[var(--mo-surface-2)] p-4 text-xs">{diagramLines.join('\n')}</pre>);
      continue;
    }
    if (/^\s*[-*+]\s+/.test(line)) { list.push(line.replace(/^\s*[-*+]\s+/, '')); continue; }
    flushList();
    if (!line.trim()) { elements.push(<div key={`space-${index}`} className="h-3" />); continue; }
    const heading = line.match(/^(#{1,6})\s+(.+)$/);
    if (heading) {
      const level = heading[1].length;
      const Tag = `h${level}` as keyof JSX.IntrinsicElements;
      elements.push(<Tag key={index} className={`mt-5 mb-2 font-semibold ${level === 1 ? 'text-2xl' : level === 2 ? 'text-xl' : 'text-lg'}`}>{inlineMarkdown(heading[2])}</Tag>);
      continue;
    }
    if (/^>\s?/.test(line)) { elements.push(<blockquote key={index} className="my-3 border-l-2 border-[var(--mo-accent)] pl-3 text-[var(--mo-text-muted)]">{inlineMarkdown(line.replace(/^>\s?/, ''))}</blockquote>); continue; }
    elements.push(<p key={index} className="my-2 leading-7">{inlineMarkdown(line)}</p>);
  }
  flushList();
  return <article className="mx-auto max-w-3xl px-8 py-8 text-[var(--mo-text-primary)]">{elements}</article>;
}

export function PreviewViewer({ session }: { session: DocumentSession }) {
  const data = session.data as PreviewDocumentData;
  const [pdfUrl, setPdfUrl] = useState<string>();
  useEffect(() => {
    if (data.format !== 'pdf' || !data.bytes) return;
    const url = URL.createObjectURL(new Blob([new Uint8Array(data.bytes)], { type: 'application/pdf' }));
    setPdfUrl(url);
    return () => URL.revokeObjectURL(url);
  }, [data.bytes, data.format]);
  const html = useMemo(() => data.format === 'html' ? data.text ?? '' : '', [data.format, data.text]);
  if (data.format === 'pdf') return pdfUrl ? <iframe title={session.displayName} src={pdfUrl} className="h-full w-full border-0 bg-white" /> : null;
  if (data.format === 'html') return <iframe title={session.displayName} srcDoc={html} sandbox="allow-scripts" className="h-full w-full border-0 bg-white" />;
  return <div className="h-full overflow-auto bg-[var(--mo-surface)]"><MarkdownPreview text={data.text ?? ''} /></div>;
}
