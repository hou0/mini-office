/** OOXML uses English Metric Units, twips and points. All helpers target CSS pixels at 96 DPI. */
export const EMU_PER_INCH = 914_400;
export const EMU_PER_PT = 12_700;
export const EMU_PER_PX = 9_525;
export const TWIPS_PER_INCH = 1_440;
export const TWIPS_PER_PT = 20;
export const TWIPS_PER_PX = 15;
export const PT_PER_PX = 0.75;
export const PX_PER_PT = 4 / 3;

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

export function emuToPt(emu: number): number {
  return emu / EMU_PER_PT;
}

export function twipsToPx(twips: number): number {
  return twips / TWIPS_PER_PX;
}

export function pxToTwips(px: number): number {
  return Math.round(px * TWIPS_PER_PX);
}

export function twipsToPt(twips: number): number {
  return twips / TWIPS_PER_PT;
}

export function ptToPx(pt: number): number {
  return pt * PX_PER_PT;
}

export function pxToPt(px: number): number {
  return px * PT_PER_PX;
}

export function halfPointsToPt(value: number): number {
  return value / 2;
}

export function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2);
}

export function mmToPx(mm: number): number {
  return (mm / 25.4) * 96;
}

export function pxToMm(px: number): number {
  return (px / 96) * 25.4;
}

export function inchesToPx(inches: number): number {
  return inches * 96;
}

export const PAGE_SIZES_PX = {
  A4: { width: mmToPx(210), height: mmToPx(297) },
  A3: { width: mmToPx(297), height: mmToPx(420) },
  A5: { width: mmToPx(148), height: mmToPx(210) },
  Letter: { width: 816, height: 1056 },
  Legal: { width: 816, height: 1344 },
  B5: { width: mmToPx(176), height: mmToPx(250) },
} as const;

export const DEFAULT_PAGE_MARGIN_PX = { top: 96, right: 90, bottom: 96, left: 90 };

/** Screen pixels per point at a given CSS zoom (for canvas rendering). */
export function scalePt(pt: number, zoom: number): number {
  return pt * PX_PER_PT * zoom;
}

/** PowerPoint slide sizes. */
export const SLIDE_SIZES = {
  '16:9': { width: 12_192_000, height: 6_858_000 },
  '4:3': { width: 9_144_000, height: 6_858_000 },
} as const;
