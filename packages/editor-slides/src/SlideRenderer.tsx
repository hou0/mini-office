// 功能：渲染可选择、可编辑、可拖动的 PPT 对象，同时复用演示和导出视图；时间：2026-10-24；作者：lq。
import { memo, type CSSProperties, type MouseEvent, type ReactNode } from 'react';
import type {
  SlideDeck,
  SlideElement,
  SlideFill,
  SlideGroupElement,
  SlideImageElement,
  SlideLine,
  SlideModel,
  SlideShapeElement,
  SlideTableElement,
  SlideTextElement,
  SlideTextParagraph,
  SlideTextRun,
} from '@mini-office/document-model';
import { DOC_FONT_STACK, fontStackFor } from '@mini-office/design-system';

function fillStyle(fill: SlideFill | undefined): CSSProperties {
  if (!fill || fill.type === 'none') return {};
  if (fill.type === 'solid') return { background: fill.color };
  if (fill.type === 'gradient') {
    const stops = [...fill.stops]
      .sort((a, b) => a.offset - b.offset)
      .map((stop) => `${stop.color} ${Math.round(stop.offset * 100)}%`)
      .join(', ');
    return { background: `linear-gradient(${(fill.angle ?? 0) + 90}deg, ${stops})` };
  }
  return {
    backgroundImage: `url("${fill.src}")`,
    backgroundSize: 'cover',
    backgroundPosition: 'center',
  };
}

function lineStyle(line: SlideLine | undefined): CSSProperties {
  if (!line?.color) return {};
  return {
    borderWidth: `${Math.max(0.5, line.widthPx ?? 1)}px`,
    borderStyle: line.dash === 'dot' ? 'dotted' : line.dash === 'dash' ? 'dashed' : 'solid',
    borderColor: line.color,
  };
}

function shapeCss(shape: string | undefined): CSSProperties {
  switch (shape) {
    case 'ellipse':
      return { borderRadius: '50%' };
    case 'roundRect':
      return { borderRadius: '12px' };
    case 'triangle':
      return { clipPath: 'polygon(50% 0%, 100% 100%, 0% 100%)' };
    case 'rtTriangle':
      return { clipPath: 'polygon(0% 0%, 0% 100%, 100% 100%)' };
    case 'diamond':
      return { clipPath: 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)' };
    case 'parallelogram':
      return { clipPath: 'polygon(20% 0%, 100% 0%, 80% 100%, 0% 100%)' };
    case 'trapezoid':
      return { clipPath: 'polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%)' };
    case 'pentagon':
      return { clipPath: 'polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)' };
    case 'hexagon':
      return { clipPath: 'polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%)' };
    case 'star5':
    case 'star':
      return {
        clipPath:
          'polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)',
      };
    case 'rightArrow':
      return { clipPath: 'polygon(0% 30%, 65% 30%, 65% 5%, 100% 50%, 65% 95%, 65% 70%, 0% 70%)' };
    case 'leftArrow':
      return { clipPath: 'polygon(100% 30%, 35% 30%, 35% 5%, 0% 50%, 35% 95%, 35% 70%, 100% 70%)' };
    case 'upArrow':
      return { clipPath: 'polygon(30% 100%, 30% 35%, 5% 35%, 50% 0%, 95% 35%, 70% 35%, 70% 100%)' };
    case 'downArrow':
      return { clipPath: 'polygon(30% 0%, 30% 65%, 5% 65%, 50% 100%, 95% 65%, 70% 65%, 70% 0%)' };
    case 'chevron':
      return { clipPath: 'polygon(0% 0%, 70% 0%, 100% 50%, 70% 100%, 0% 100%, 30% 50%)' };
    default:
      return {};
  }
}

function runStyle(run: SlideTextRun, colorFallback = '#1D2939'): CSSProperties {
  const style = run.style;
  return {
    color: style.color ?? colorFallback,
    fontSize: style.fontSizePt ? `${style.fontSizePt}pt` : undefined,
    fontWeight: style.bold ? 700 : undefined,
    fontStyle: style.italic ? 'italic' : undefined,
    textDecoration: [style.underline ? 'underline' : '', style.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' ') || undefined,
    backgroundColor: style.highlight,
    fontFamily: style.fontFamily
      ? `${fontStackFor(style.fontFamily)}, ${DOC_FONT_STACK}`
      : undefined,
    letterSpacing: style.letterSpacingPx ? `${style.letterSpacingPx}px` : undefined,
  };
}

function paragraphsNode(
  paragraphs: SlideTextParagraph[],
  options: { defaultSizePt?: number; color?: string } = {},
): ReactNode {
  return paragraphs.map((paragraph, index) => {
    const lineHeight = paragraph.lineSpacing ? paragraph.lineSpacing : 1.25;
    return (
      <p
        key={index}
        style={{
          margin: 0,
          marginTop: paragraph.spaceBeforePt ? `${paragraph.spaceBeforePt}pt` : undefined,
          marginBottom: paragraph.spaceAfterPt ? `${paragraph.spaceAfterPt}pt` : undefined,
          textAlign: paragraph.align ?? 'left',
          lineHeight,
          whiteSpace: 'pre-wrap',
          overflowWrap: 'break-word',
          position: 'relative',
          paddingLeft: paragraph.level ? `${paragraph.level * 22}px` : undefined,
          fontSize: options.defaultSizePt ? `${options.defaultSizePt}pt` : undefined,
          color: options.color,
        }}
      >
        {paragraph.bullet && paragraph.bulletChar !== '' ? (
          <span
            style={{
              position: 'absolute',
              left: Math.max(0, (paragraph.level ?? 0) * 22 - 14),
              color: 'inherit',
            }}
          >
            {paragraph.bulletChar ?? (paragraph.level ? '–' : '•')}
          </span>
        ) : null}
        {paragraph.runs.length === 0 ? '\u00A0' : null}
        {paragraph.runs.map((run, runIndex) =>
          run.lineBreak ? (
            <br key={runIndex} />
          ) : (
            <span key={runIndex} style={runStyle(run, options.color)}>
              {run.text}
            </span>
          ),
        )}
      </p>
    );
  });
}

function TextBody({
  element,
  editable = false,
  onTextChange,
}: {
  element: SlideTextElement | SlideShapeElement;
  editable?: boolean;
  onTextChange?: (elementId: string, value: string) => void;
}) {
  const paragraphs = element.type === 'text' ? element.paragraphs : element.paragraphs ?? [];
  const vertical = element.verticalAlign;
  const margins = element.type === 'text' ? element.margins : undefined;
  return (
    <div
      contentEditable={editable}
      suppressContentEditableWarning
      onBlur={
        editable && onTextChange
          ? (event) => onTextChange(element.id, event.currentTarget.innerText)
          : undefined
      }
      style={{
        position: 'absolute',
        inset: 0,
        display: 'flex',
        flexDirection: 'column',
        justifyContent:
          vertical === 'middle' ? 'center' : vertical === 'bottom' ? 'flex-end' : 'flex-start',
        paddingTop: margins?.top ?? 4,
        paddingRight: margins?.right ?? 8,
        paddingBottom: margins?.bottom ?? 4,
        paddingLeft: margins?.left ?? 8,
        boxSizing: 'border-box',
        overflow: 'hidden',
      }}
    >
      {paragraphsNode(paragraphs)}
    </div>
  );
}

function elementFrame(element: SlideElement, extra?: CSSProperties): CSSProperties {
  return {
    position: 'absolute',
    left: element.x,
    top: element.y,
    width: Math.max(1, element.w),
    height: Math.max(1, element.h),
    transform: `rotate(${element.rotation ?? 0}deg) scaleX(${element.flipH ? -1 : 1}) scaleY(${element.flipV ? -1 : 1})`,
    transformOrigin: 'center center',
    boxSizing: 'border-box',
    ...extra,
  };
}

function selectedStyle(selected?: boolean): CSSProperties {
  return selected
    ? { outline: '2px solid var(--mo-accent, #2E7D72)', outlineOffset: 2, cursor: 'move' }
    : {};
}

interface ElementViewProps {
  element: SlideElement;
  editable?: boolean;
  selected?: boolean;
  onSelect?: (id: string) => void;
  onTextChange?: (elementId: string, value: string) => void;
  onMove?: (elementId: string, dx: number, dy: number) => void;
}

function interactionProps({ element, editable, selected, onSelect, onMove }: Omit<ElementViewProps, 'onTextChange'>): {
  onClick?: (event: MouseEvent) => void;
  onPointerDown?: (event: React.PointerEvent) => void;
  'data-slide-element-id': string;
} {
  let start: { x: number; y: number } | null = null;
  return {
    'data-slide-element-id': element.id,
    onClick: editable && onSelect ? (event) => { event.stopPropagation(); onSelect(element.id); } : undefined,
    onPointerDown: editable && onMove && onSelect ? (event) => {
      if (event.button !== 0 || event.target instanceof HTMLElement && event.target.isContentEditable) return;
      event.stopPropagation();
      onSelect(element.id);
      start = { x: event.clientX, y: event.clientY };
      (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
      const move = (moveEvent: PointerEvent) => {
        if (!start) return;
        onMove(element.id, moveEvent.clientX - start.x, moveEvent.clientY - start.y);
        start = { x: moveEvent.clientX, y: moveEvent.clientY };
      };
      const end = () => { start = null; window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', end); };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', end, { once: true });
    } : undefined,
  };
}

function ShapeView({ element, editable, selected, onSelect, onTextChange, onMove }: Omit<ElementViewProps, 'element'> & { element: SlideShapeElement }) {
  if (element.shape === 'line' || element.shape === 'straightConnector1') {
    const color = element.line?.color ?? '#1D2939';
    const widthPx = Math.max(1, element.line?.widthPx ?? 1);
    return (
      <div {...interactionProps({ element, editable, selected, onSelect, onMove })} style={elementFrame(element, { display: 'flex', alignItems: 'center', ...selectedStyle(selected) })}>
        <div
          style={{
            width: '100%',
            height: widthPx,
            background: color,
            borderStyle: element.line?.dash === 'dot' ? 'dotted' : 'solid',
          }}
        />
      </div>
    );
  }
  return (
    <div
      {...interactionProps({ element, editable, selected, onSelect, onMove })}
      style={elementFrame(element, {
        ...fillStyle(element.fill),
        ...lineStyle(element.line),
        ...shapeCss(element.shape),
        borderColor: element.line?.color,
        borderWidth: element.line?.color ? `${Math.max(0.5, element.line.widthPx ?? 1)}px` : undefined,
        borderStyle: element.line?.color ? 'solid' : undefined,
        overflow: 'hidden',
        ...selectedStyle(selected),
      })}
    >
      {element.paragraphs?.length ? <TextBody element={element} editable={editable} onTextChange={onTextChange} /> : null}
    </div>
  );
}

function TextView({ element, editable, selected, onSelect, onTextChange, onMove }: Omit<ElementViewProps, 'element'> & { element: SlideTextElement }) {
  return (
    <div
      {...interactionProps({ element, editable, selected, onSelect, onMove })}
      style={elementFrame(element, {
        ...fillStyle(element.fill),
        ...lineStyle(element.line),
        ...(element.shape ? shapeCss(element.shape) : {}),
        overflow: 'hidden',
        ...selectedStyle(selected),
      })}
    >
      <TextBody element={element} editable={editable} onTextChange={onTextChange} />
    </div>
  );
}

function ImageView({ element, editable, selected, onSelect, onMove }: Omit<ElementViewProps, 'element'> & { element: SlideImageElement }) {
  const crop = element.crop;
  const visibleWidth = crop ? element.w / Math.max(0.01, 1 - crop.left - crop.right) : element.w;
  const visibleHeight = crop ? element.h / Math.max(0.01, 1 - crop.top - crop.bottom) : element.h;
  return (
    <div
      {...interactionProps({ element, editable, selected, onSelect, onMove })}
      style={elementFrame(element, {
        overflow: 'hidden',
        ...selectedStyle(selected),
      })}
    >
      <img
        src={element.src}
        alt={element.alt ?? element.name ?? '幻灯片图片'}
        draggable={false}
        style={{
          position: 'absolute',
          width: visibleWidth,
          height: visibleHeight,
          left: crop ? -crop.left * visibleWidth : 0,
          top: crop ? -crop.top * visibleHeight : 0,
          objectFit: 'fill',
          userSelect: 'none',
        }}
      />
    </div>
  );
}

function TableView({ element, editable, selected, onSelect, onMove }: Omit<ElementViewProps, 'element'> & { element: SlideTableElement }) {
  const columns = element.columnWidths.length || Math.max(1, ...element.rows.map((row) => row.cells.length));
  const totalColumnWidth = element.columnWidths.reduce((sum, width) => sum + width, 0) || columns;
  return (
    <div {...interactionProps({ element, editable, selected, onSelect, onMove })} style={elementFrame(element, { overflow: 'hidden', ...selectedStyle(selected) })}>
      <table
        style={{
          width: '100%',
          height: '100%',
          tableLayout: 'fixed',
          borderCollapse: 'collapse',
          background: '#FFFFFF',
        }}
      >
        <colgroup>
          {Array.from({ length: columns }, (_, index) => (
            <col
              key={index}
              style={{
                width: `${
                  ((element.columnWidths[index] ?? totalColumnWidth / columns) / totalColumnWidth) * 100
                }%`,
              }}
            />
          ))}
        </colgroup>
        <tbody>
          {element.rows.map((row, rowIndex) => (
            <tr key={rowIndex} style={{ height: element.rowHeights[rowIndex] }}>
              {row.cells.map((cell, cellIndex) => (
                <td
                  key={cellIndex}
                  colSpan={cell.colSpan}
                  rowSpan={cell.rowSpan}
                  style={{
                    ...fillStyle(cell.fill),
                    border: cell.borders?.color
                      ? `${cell.borders.widthPx ?? 1}px solid ${cell.borders.color}`
                      : '1px solid #D0D5DD',
                    verticalAlign:
                      cell.verticalAlign === 'middle'
                        ? 'middle'
                        : cell.verticalAlign === 'bottom'
                          ? 'bottom'
                          : 'top',
                    padding: '4px 8px',
                    fontSize: '14pt',
                    overflow: 'hidden',
                    boxSizing: 'border-box',
                  }}
                >
                  {paragraphsNode(cell.text, { defaultSizePt: 14 })}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function GroupView({ element, editable, selected, onSelect, onTextChange, onMove }: Omit<ElementViewProps, 'element'> & { element: SlideGroupElement }) {
  return (
    <div {...interactionProps({ element, editable, selected, onSelect, onMove })} style={elementFrame(element, selectedStyle(selected))}>
      {element.children.map((child) => (
        <ElementView key={child.id} element={child} editable={editable} selected={selected} onSelect={onSelect} onTextChange={onTextChange} onMove={onMove} />
      ))}
    </div>
  );
}

export function ElementView({ element, editable = false, selected = false, onSelect, onTextChange, onMove }: ElementViewProps) {
  switch (element.type) {
    case 'text':
      return <TextView element={element} editable={editable} selected={selected} onSelect={onSelect} onTextChange={onTextChange} onMove={onMove} />;
    case 'shape':
      return <ShapeView element={element} editable={editable} selected={selected} onSelect={onSelect} onTextChange={onTextChange} onMove={onMove} />;
    case 'image':
      return <ImageView element={element} editable={editable} selected={selected} onSelect={onSelect} onMove={onMove} />;
    case 'table':
      return <TableView element={element} editable={editable} selected={selected} onSelect={onSelect} onMove={onMove} />;
    case 'group':
      return <GroupView element={element} editable={editable} selected={selected} onSelect={onSelect} onTextChange={onTextChange} onMove={onMove} />;
    default:
      return null;
  }
}

export const SlideSurface = memo(function SlideSurface({
  deck,
  slide,
  width,
  height,
  className,
  editable = false,
  selectedId,
  onSelect,
  onTextChange,
  onMove,
}: {
  deck: SlideDeck;
  slide: SlideModel;
  width?: number;
  height?: number;
  className?: string;
  editable?: boolean;
  selectedId?: string | null;
  onSelect?: (id: string) => void;
  onTextChange?: (elementId: string, value: string) => void;
  onMove?: (elementId: string, dx: number, dy: number) => void;
}) {
  return (
    <div
      data-slide-surface
      className={className}
      style={{
        position: 'relative',
        width: width ?? deck.widthPx,
        height: height ?? deck.heightPx,
        overflow: 'hidden',
        background: '#FFFFFF',
        ...fillStyle(slide.background),
      }}
    >
      {slide.elements.map((element) => (
        <ElementView key={element.id} element={element} editable={editable} selected={selectedId === element.id} onSelect={onSelect} onTextChange={onTextChange} onMove={onMove} />
      ))}
    </div>
  );
});

export function SlideRenderer({
  deck,
  slide,
  scale = 1,
  className,
  editable = false,
  selectedId,
  onSelect,
  onTextChange,
  onMove,
}: {
  deck: SlideDeck;
  slide: SlideModel;
  scale?: number;
  className?: string;
  editable?: boolean;
  selectedId?: string | null;
  onSelect?: (id: string) => void;
  onTextChange?: (elementId: string, value: string) => void;
  onMove?: (elementId: string, dx: number, dy: number) => void;
}) {
  return (
    <div
      className={className}
      style={{
        width: deck.widthPx * scale,
        height: deck.heightPx * scale,
        position: 'relative',
        flexShrink: 0,
      }}
    >
      <div
        style={{
          width: deck.widthPx,
          height: deck.heightPx,
          transform: `scale(${scale})`,
          transformOrigin: 'top left',
          position: 'absolute',
          left: 0,
          top: 0,
        }}
      >
        <SlideSurface deck={deck} slide={slide} editable={editable} selectedId={selectedId} onSelect={onSelect} onTextChange={onTextChange} onMove={onMove} />
      </div>
    </div>
  );
}
