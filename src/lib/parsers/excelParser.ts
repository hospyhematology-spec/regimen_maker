import ExcelJS from "exceljs";
import type {
  SourceTrace,
  RichBulletItem,
  AdministrationStep,
  DrugComponent,
  DiluentItem,
  FlushItem,
} from "@/types/regimen";

export interface CellInfo {
  address: string;
  value: string;
  merged: boolean;
  mergeRange?: string;
}

export interface SheetData {
  name: string;
  cells: Map<string, CellInfo>;
  rows: CellInfo[][];
}

/**
 * Excelファイルのすべてのシートを読み込み、結合セル情報を保持した中間表現を返す
 */
export async function parseExcelFile(buffer: ArrayBuffer): Promise<SheetData[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheets: SheetData[] = [];

  workbook.eachSheet((worksheet) => {
    const cells = new Map<string, CellInfo>();
    const rows: CellInfo[][] = [];

    // 結合セル情報を収集
    const mergeMap = new Map<string, string>();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const model = (worksheet as any).model;
    if (model?.merges) {
      for (const mergeRange of model.merges as string[]) {
        // mergeRange e.g. "A1:C3"
        const [start] = mergeRange.split(":");
        // すべての結合範囲セルに開始セルアドレスをマップ
        const ref = worksheet.getCell(mergeRange.replace(":", ":"));
        void ref;
        margeRangeCells(worksheet, mergeRange, start, mergeMap);
      }
    }

    worksheet.eachRow({ includeEmpty: false }, (row, rowIndex) => {
      const rowCells: CellInfo[] = [];
      row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
        const address = cell.address;
        const rawValue = cell.type === ExcelJS.ValueType.Merge
          ? ""
          : getCellText(cell);
        const masterAddress = mergeMap.get(address);
        const mergeRange = [...mergeMap.entries()]
          .find(([, v]) => v === address)?.[0];

        const info: CellInfo = {
          address,
          value: rawValue,
          merged: cell.type === ExcelJS.ValueType.Merge,
          mergeRange: masterAddress ?? mergeRange,
        };
        cells.set(address, info);
        rowCells.push(info);
        void rowIndex;
        void colIndex;
      });
      rows.push(rowCells);
    });

    sheets.push({ name: worksheet.name, cells, rows });
  });

  return sheets;
}

function margeRangeCells(
  worksheet: ExcelJS.Worksheet,
  range: string,
  masterAddr: string,
  map: Map<string, string>
) {
  // range: "A1:C3" → decode all addresses in range
  const [startRef, endRef] = range.split(":");
  const start = worksheet.getCell(startRef);
  const end = worksheet.getCell(endRef);
  const startRow = Number(start.row);
  const startCol = Number(start.col);
  const endRow = Number(end.row);
  const endCol = Number(end.col);
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const addr = colNumToLetter(c) + r;
      map.set(addr, masterAddr);
    }
  }
}

function colNumToLetter(col: number): string {
  let letter = "";
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function getCellText(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if ("richText" in v) {
      return (v as ExcelJS.CellRichTextValue).richText
        .map((r) => r.text)
        .join("");
    }
    if ("result" in v) return String((v as ExcelJS.CellFormulaValue).result ?? "");
    if ("text" in v) return String((v as ExcelJS.CellHyperlinkValue).text);
    if ("error" in v) return "";
  }
  return String(v);
}

/**
 * セルの値を取得する（結合セルの場合は先頭セルの値を参照）
 */
export function getEffectiveValue(cells: Map<string, CellInfo>, address: string): string {
  const cell = cells.get(address);
  if (!cell) return "";
  if (cell.merged && cell.mergeRange) {
    return cells.get(cell.mergeRange)?.value ?? "";
  }
  return cell.value;
}

/**
 * シートの行を返す（先頭列の番号でステップを分割するため）
 */
export function getRowValues(rows: CellInfo[][], rowIndex: number): string[] {
  const row = rows[rowIndex];
  if (!row) return [];
  return row.map((c) => c.value);
}

/**
 * 指定したシートの特定のセルレンジからテキストを取得
 */
export function getCellValue(cells: Map<string, CellInfo>, address: string): string {
  return getEffectiveValue(cells, address);
}

/**
 * 箇条書きエリアのテキストブロックから RichBulletItem[] を生成
 */
export function extractBulletItems(
  cells: Map<string, CellInfo>,
  startRow: number,
  endRow: number,
  col: string,
  fileName: string,
  sheetName: string,
  confidence: number = 0.85
): RichBulletItem[] {
  const items: RichBulletItem[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const addr = `${col}${r}`;
    const text = getEffectiveValue(cells, addr).trim();
    if (text) {
      const trace: SourceTrace = {
        sourceType: "excel",
        fileName,
        sheetName,
        cellRange: addr,
        quotedText: text,
        aiInterpretation: text,
        confidence,
      };
      items.push({ text, sourceTrace: [trace] });
    }
  }
  return items;
}

/**
 * 投与方法シートからAdministrationStep[]を抽出する
/**
 * 投与方法シートからAdministrationStep[]を抽出する
 * A列の整数を区切り点として使用する。なければ全行を1ステップとして扱う。
 */
export function extractAdministrationSteps(
  sheetData: SheetData,
  fileName: string,
  stepColIndex: number = 0
): AdministrationStep[] {
  const steps: AdministrationStep[] = [];
  let currentStepNo: number | null = null;
  let currentRows: CellInfo[][] = [];

  // ヘッダー行検索（「薬剤」「投与」「日」が行に含まれる最初の行）
  const headerKeywords = /薬剤|投与|ルート|route|dose|drug|day|単位/i;
  let skipUntil = 0;
  for (let i = 0; i < Math.min(sheetData.rows.length, 15); i++) {
    const rowStr = sheetData.rows[i].map((c: CellInfo) => c.value?.trim() ?? "").join(" ");
    if (headerKeywords.test(rowStr)) {
      skipUntil = i + 1; // ヘッダー行を含めてスキップ
      break;
    }
  }

  // データ行を処理
  for (let ri = 0; ri < sheetData.rows.length; ri++) {
    const row = sheetData.rows[ri];
    const firstCell = row[stepColIndex]?.value?.trim() ?? "";
    const stepNoMatch = /^(\d+)$/.exec(firstCell);

    // ヘッダー行はスキップ（ただし既に A列整数ステップに入った後はスキップしない）
    if (ri < skipUntil && currentStepNo === null) continue;

    if (stepNoMatch) {
      // A列に純整数→ステップ区切り
      if (currentStepNo !== null && currentRows.length > 0) {
        const step = buildStep(currentStepNo, currentRows, fileName, sheetData.name);
        if (step.mainDrugs.length > 0 || step.diluents.length > 0) {
          steps.push(step);
        }
      }
      currentStepNo = parseInt(stepNoMatch[1], 10);
      currentRows = [row];
    } else if (currentStepNo !== null) {
      currentRows.push(row);
    }
  }

  // 最後のステップを確定
  if (currentStepNo !== null && currentRows.length > 0) {
    const step = buildStep(currentStepNo, currentRows, fileName, sheetData.name);
    if (step.mainDrugs.length > 0 || step.diluents.length > 0) {
      steps.push(step);
    }
  }

  // ステップが得られなかった場合、シート全体を1ステップとして解析
  if (steps.length === 0 && sheetData.rows.length > 0) {
    const step = buildStep(1, sheetData.rows, fileName, sheetData.name);
    if (step.mainDrugs.length > 0 || step.diluents.length > 0) {
      steps.push(step);
    }
  }

  return steps;
}

function buildStep(
  stepNo: number,
  rows: CellInfo[][],
  fileName: string,
  sheetName: string
): AdministrationStep {
  const mainDrugs: DrugComponent[] = [];
  const diluents: DiluentItem[] = [];
  const flushes: FlushItem[] = [];
  const notes: RichBulletItem[] = [];
  const administrationDays: string[] = [];
  let route: string | null = null;
  let infusionRate: string | null = null;

  // 用量・単位のパターン
  const dosePattern = /^[\d.]+$/;
  const unitPattern = /mg|g|mL|μg|mcg|IU|U|unit|万/i;
  const routePattern = /静注|点滴|経口|皮下|筋注|静脈|IV|SC|PO|IM|iv/i;
  const dayPattern = /^Day\s*\d+|^\d+日目|^day/i;
  const flushPattern = /フラッシュ|生食|生理食塩|NS\b/i;
  const diluentPattern = /溶解|希釈|補液|前投薬|前処置|制吐|生理食塩|ブドウ糖|dextrose|saline/i;
  // 薬剤名らしさの判定（カタカナ含む、英数字多め、2文字以上）
  const drugNameLike = (s: string): boolean => {
    if (!s || s.length < 2) return false;
    if (/^\d+$/.test(s)) return false; // 純数値はNG
    if (dosePattern.test(s)) return false;
    if (dayPattern.test(s)) return false;
    if (routePattern.test(s)) return false;
    if (unitPattern.test(s) && s.length < 4) return false;
    // カタカナ or 英字含む
    return /[ァ-ヶA-Za-z]/.test(s);
  };

  for (const row of rows) {
    const cells = row.map((c: CellInfo) => c.value?.trim() ?? "").filter(Boolean);
    if (cells.length === 0) continue;

    // 固定列インデックスで試みる方法 ─ category列が存在する場合
    const colB = row[1]?.value?.trim() ?? "";
    const colC = row[2]?.value?.trim() ?? "";
    const colD = row[3]?.value?.trim() ?? "";
    const colE = row[4]?.value?.trim() ?? "";

    // Day値を探す
    for (const cell of cells) {
      if (dayPattern.test(cell) && !administrationDays.includes(cell)) {
        administrationDays.push(cell);
      }
    }
    // 投与経路を探す
    for (const cell of cells) {
      if (routePattern.test(cell) && !route) route = cell;
    }

    // 列B にカテゴリ、列C に薬剤名というパターンを試みる
    if (colC && drugNameLike(colC)) {
      const dose = colD && dosePattern.test(colD) ? colD : null;
      const unit = colE && unitPattern.test(colE) ? colE : null;
      const cat = colB.toLowerCase();

      if (diluentPattern.test(colB) || diluentPattern.test(colC)) {
        diluents.push({ name: colC, volume: dose, unit });
      } else if (flushPattern.test(colC)) {
        flushes.push({ name: colC, volume: dose, unit });
      } else {
        mainDrugs.push({ drugName: colC, dose, doseUnit: unit });
      }
      void cat;
    } else {
      // 薬剤名っぽいセルを行内からすべて探す（フォールバック）
      for (let ci = 0; ci < row.length; ci++) {
        const val = row[ci]?.value?.trim() ?? "";
        if (!drugNameLike(val)) continue;
        // 隣のセルから用量と単位を試みる
        const nextVal = row[ci + 1]?.value?.trim() ?? "";
        const nextNextVal = row[ci + 2]?.value?.trim() ?? "";
        const dose = dosePattern.test(nextVal) ? nextVal : null;
        const unit = unitPattern.test(nextNextVal) ? nextNextVal : (unitPattern.test(nextVal) ? nextVal : null);

        if (diluentPattern.test(val)) {
          diluents.push({ name: val, volume: dose, unit });
        } else if (flushPattern.test(val)) {
          flushes.push({ name: val, volume: dose, unit });
        } else {
          // 重複チェック
          if (!mainDrugs.some(d => d.drugName === val)) {
            mainDrugs.push({ drugName: val, dose, doseUnit: unit });
          }
        }
        break; // 1行1薬剤として扱う
      }
    }

    // 速度の検出
    const rateMatch = cells.find(c => /\d+\s*(mL\/h|ml\/h|mg\/h)/.test(c));
    if (rateMatch && !infusionRate) infusionRate = rateMatch;

    // ノート（数字・薬剤名・単位以外の長めのテキスト）
    const noteCell = cells.find(c => c.length > 10 && !drugNameLike(c) && !dosePattern.test(c) && !routePattern.test(c));
    if (noteCell) notes.push({ text: noteCell });
  }

  const firstRow = rows[0];
  const trace: SourceTrace = {
    sourceType: "excel",
    fileName,
    sheetName,
    cellRange: firstRow?.[0]?.address,
    quotedText: rows.slice(0, 3).map((r) => r.map((c) => c.value).join("\t")).join("\n"),
    aiInterpretation: `ステップ${stepNo}の投与情報`,
    confidence: 0.8,
  };

  return {
    stepNo,
    mainDrugs,
    diluents,
    flushes,
    route: route || null,
    infusionRate: infusionRate || null,
    rateUnit: null,
    administrationDays,
    notes,
    sourceTrace: [trace],
  };
}

