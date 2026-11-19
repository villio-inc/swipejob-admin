import * as XLSX from "xlsx";
import { SLOTS, SLOT_IN_SHEET, Slot, zeroSlots } from "./slots";

/**
 * 委託講師配置表の「一般充足」シートを読む。
 *
 * 採用のゴールはこのシートの**不足数が全部0になること**（2026-11-05 指示）。
 * 校舎 × 分野 × 時間帯 のコマごとに「あと何人分足りないか」が入っていて、
 * 右端に校舎ごとの判定（充足 / 空欄）と、必要な採用人数が入っている。
 *
 * 不足のコマ数（合計）と必要な採用人数（不足数）は別物で、
 * 「1名が約3日シフトに入れる」前提で人数に直されている。
 * こちらで計算し直すと元の表と食い違うので、**シートの数字をそのまま読む**。
 *
 * **黄色のセルが「まだ埋まっていないコマ」**（2026-11-19 指示）。
 * シートの式では合計が3コマ以下の校舎を「充足」として不足数0にしているので、
 * 数字が入っていても不足として数えないコマがある。
 * 数字だけを見て全部を不足として扱うと、充足している校舎まで赤く出て
 * 「全校舎が足りない」ように見えてしまうので、**印の付いたコマだけを数える**。
 */

/** 配置表が持っている分野。採用管理表の「分野」より粗い（Officeなどの枠は無い）。 */
export const PLACEMENT_FIELDS = ["IT", "WEB", "CAD"] as const;

export type PlacementField = (typeof PLACEMENT_FIELDS)[number];

export const FIELD_LABEL: Record<PlacementField, string> = {
  IT: "IT",
  WEB: "Webデザイン",
  CAD: "CAD",
};

/**
 * 「まだ埋まっていない」の印。黄色のべた塗り。
 *
 * 色の指定は FFFF00 のこともあれば FFFFFF00（先頭2桁は不透明度）のこともあり、
 * 少し違う黄色が使われることもあるので、**黄色の範囲**で見る。
 * 行の網掛け（灰）や校舎名のオレンジ（FFC000）は入らないようにしている。
 */
export function isShortageMark(rgb: string | undefined): boolean {
  const hex = String(rgb ?? "").toUpperCase().replace(/[^0-9A-F]/g, "");
  if (hex.length < 6) return false;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(hex.length - 6 + i, hex.length - 6 + i + 2), 16));
  // 赤と緑が高く、青がはっきり低いものを黄色とみなす。
  // 白(FFFFFF)は青が下がらないので入らず、校舎名のオレンジ(FFC000)は緑が足りないので入らない。
  return r >= 0xe0 && g >= 0xd0 && b <= g - 0x20;
}

function markOf(cell: XLSX.CellObject | undefined): boolean {
  const style = (cell as { s?: { patternType?: string; fgColor?: { rgb?: string } } } | undefined)?.s;
  if (!style || style.patternType !== "solid") return false;
  return isShortageMark(style.fgColor?.rgb);
}

export interface PlacementRow {
  /** 「仙台駅前」。応募データと突き合わせるときは normalizeShopKey を通す。 */
  shopShortName: string;
  manager: string;
  /** 存続校舎 / FY28判断校舎 */
  shopStatus: string;
  /** 現状と方針（シートに書かれているコメント） */
  note: string;
  /** 分野 × 時間帯 の不足コマ数（シートの数字そのまま） */
  shortage: Record<PlacementField, Record<Slot, number>>;
  /** そのコマに「まだ埋まっていない」の印（黄色）が付いているか */
  marked: Record<PlacementField, Record<Slot, boolean>>;
  /** 不足コマの合計（シートの「合計」列） */
  slots: number;
  /** 一般充足判定。「充足」なら埋まっている。 */
  judgement: string;
  /** 不足数＝これから採る人数（シートの「不足数」列） */
  needHires: number;
  /** 全分野募集停止（STOP など）。空欄なら募集中。 */
  stop: string;
}

export interface PlacementSheet {
  rows: PlacementRow[];
  /** 読んだシート名 */
  sheetName: string;
  sourceFileName: string;
  importedAt: string;
  /** 読み飛ばした行数（校舎名が空など） */
  skipped: number;
  /**
   * 黄色の印が付いたコマの数。
   * 0 なら印を使っていないファイルなので、数字の入ったコマを全部不足として数える。
   */
  markedCells: number;
}

function text(v: unknown): string {
  return String(v ?? "")
    .replace(/[\s　]+/g, " ")
    .trim();
}

function num(v: unknown): number {
  const n = Number(String(v ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

/** 見出しの「IT\n【平日】」のような書き方から、分野と時間帯を取る */
function slotColumn(header: string): { field: PlacementField; slot: Slot } | null {
  const t = header.replace(/[\s　]/g, "");
  for (const field of PLACEMENT_FIELDS) {
    if (!t.startsWith(field)) continue;
    for (const slot of SLOTS) {
      if (t.includes(`【${SLOT_IN_SHEET[slot]}】`)) return { field, slot };
    }
  }
  return null;
}

const emptyShortage = (): Record<PlacementField, Record<Slot, number>> => ({
  IT: zeroSlots(),
  WEB: zeroSlots(),
  CAD: zeroSlots(),
});

const noSlots = (): Record<Slot, boolean> => ({ 平日昼: false, 平日夜: false, 土日: false });

const emptyMarks = (): Record<PlacementField, Record<Slot, boolean>> => ({
  IT: noSlots(),
  WEB: noSlots(),
  CAD: noSlots(),
});

/**
 * シートを選ぶ。既定は「一般充足」。
 *
 * 同じ形のシートが「完全充足」「最低条件」などいくつもあるので、
 * 名前で決め打ちせず、見出しの形（校舎名 ＋ IT【平日】）が合うシートの中から
 * 「一般充足」を優先する。名前が変わっても読めるようにするため。
 */
function pickSheet(wb: XLSX.WorkBook): { name: string; header: number; cols: string[] } | null {
  const candidates: { name: string; header: number; cols: string[] }[] = [];

  for (const name of wb.SheetNames) {
    const grid = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name], { header: 1, defval: "", raw: false });
    for (let i = 0; i < Math.min(grid.length, 10); i++) {
      const cols = (grid[i] ?? []).map((c) => text(c));
      const hasShop = cols.some((c) => c.replace(/[\s　]/g, "") === "校舎名");
      const hasSlot = cols.some((c) => slotColumn(c) !== null);
      if (hasShop && hasSlot) {
        candidates.push({ name, header: i, cols });
        break;
      }
    }
  }

  if (candidates.length === 0) return null;
  return candidates.find((c) => c.name.includes("一般充足")) ?? candidates[0];
}

export function parsePlacementSheet(fileName: string, buf: ArrayBuffer): PlacementSheet {
  // cellStyles を付けないとセルの塗り（＝まだ埋まっていない印）が読めない
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", raw: false, cellStyles: true });
  const picked = pickSheet(wb);
  if (!picked) {
    return {
      rows: [],
      sheetName: "",
      sourceFileName: fileName,
      importedAt: new Date().toISOString(),
      skipped: 0,
      markedCells: 0,
    };
  }

  const ws = wb.Sheets[picked.name];
  const grid = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "", raw: false });
  /**
   * 表のマス目の位置から、シート上のセルを引く。
   * sheet_to_json は使っている範囲の左上から並べるので、その分をずらす。
   */
  const origin = XLSX.utils.decode_range(ws["!ref"] ?? "A1").s;
  const cellAt = (gridRow: number, gridCol: number) =>
    ws[XLSX.utils.encode_cell({ r: origin.r + gridRow, c: origin.c + gridCol })] as XLSX.CellObject | undefined;
  const at = (name: string) =>
    picked.cols.findIndex((c) => c.replace(/[\s　]/g, "") === name);

  const col = {
    manager: at("管理者"),
    shopStatus: at("校舎ステータス"),
    shop: at("校舎名"),
    note: at("現状と方針"),
    slots: at("合計"),
    judgement: at("一般充足判定"),
    needHires: at("不足数"),
    stop: at("全分野募集停止"),
  };
  const slotCols = picked.cols
    .map((c, i) => ({ i, hit: slotColumn(c) }))
    .filter((x): x is { i: number; hit: { field: PlacementField; slot: Slot } } => x.hit !== null);

  const rows: PlacementRow[] = [];
  let skipped = 0;
  let markedCells = 0;

  for (let gridRow = picked.header + 1; gridRow < grid.length; gridRow++) {
    const raw = grid[gridRow] ?? [];
    const shop = text(raw[col.shop]);
    if (!shop || shop === "校舎名") {
      skipped++;
      continue;
    }

    const shortage = emptyShortage();
    const marked = emptyMarks();
    for (const { i, hit } of slotCols) {
      shortage[hit.field][hit.slot] = num(raw[i]);
      const mark = markOf(cellAt(gridRow, i));
      marked[hit.field][hit.slot] = mark;
      if (mark) markedCells++;
    }
    const sum = PLACEMENT_FIELDS.reduce(
      (a, f) => a + SLOTS.reduce((b, s) => b + shortage[f][s], 0),
      0
    );

    rows.push({
      shopShortName: shop,
      marked,
      manager: col.manager >= 0 ? text(raw[col.manager]) : "",
      shopStatus: col.shopStatus >= 0 ? text(raw[col.shopStatus]) : "",
      note: col.note >= 0 ? text(raw[col.note]) : "",
      shortage,
      // 「合計」列が空でも、コマを足せば出せる
      slots: col.slots >= 0 && text(raw[col.slots]) ? num(raw[col.slots]) : sum,
      judgement: col.judgement >= 0 ? text(raw[col.judgement]) : "",
      needHires: col.needHires >= 0 ? num(raw[col.needHires]) : 0,
      stop: col.stop >= 0 ? text(raw[col.stop]) : "",
    });
  }

  return {
    rows,
    sheetName: picked.name,
    sourceFileName: fileName,
    importedAt: new Date().toISOString(),
    skipped,
    markedCells,
  };
}

/** 全校舎の不足数（これから採る人数）の合計 */
export function totalNeedHires(sheet: PlacementSheet | null): number {
  return sheet ? sheet.rows.reduce((a, r) => a + r.needHires, 0) : 0;
}

/** 全校舎の不足コマ数の合計 */
export function totalShortageSlots(sheet: PlacementSheet | null): number {
  return sheet ? sheet.rows.reduce((a, r) => a + r.slots, 0) : 0;
}
