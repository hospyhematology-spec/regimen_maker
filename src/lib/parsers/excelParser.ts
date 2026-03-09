import ExcelJS from "exceljs";
import path from "path";
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
export async function parseExcelFile(filePath: string): Promise<SheetData[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

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
 */
export function extractAdministrationSteps(
  sheetData: SheetData,
  fileName: string,
  stepColIndex: number = 0
): AdministrationStep[] {
  const steps: AdministrationStep[] = [];
  let currentStepNo: number | null = null;
  let currentRows: CellInfo[][] = [];

  for (const row of sheetData.rows) {
    const firstCell = row[stepColIndex]?.value?.trim() ?? "";
    const stepNo = parseInt(firstCell, 10);

    if (!isNaN(stepNo) && firstCell === String(stepNo)) {
      // 前のステップを確定
      if (currentStepNo !== null && currentRows.length > 0) {
        steps.push(buildStep(currentStepNo, currentRows, fileName, sheetData.name));
      }
      currentStepNo = stepNo;
      currentRows = [row];
    } else if (currentStepNo !== null) {
      currentRows.push(row);
    }
  }

  // 最後のステップを確定
  if (currentStepNo !== null && currentRows.length > 0) {
    steps.push(buildStep(currentStepNo, currentRows, fileName, sheetData.name));
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

  for (const row of rows) {
    const category = row[1]?.value?.trim() ?? "";
    const name = row[2]?.value?.trim() ?? "";
    const dose = row[3]?.value?.trim() ?? "";
    const unit = row[4]?.value?.trim() ?? "";
    const noteText = row[5]?.value?.trim() ?? "";
    const routeVal = row[6]?.value?.trim() ?? "";
    const rateVal = row[7]?.value?.trim() ?? "";
    const dayVal = row[8]?.value?.trim() ?? "";

    if (name) {
      const cat = category.toLowerCase();
      if (cat.includes("主") || cat.includes("抗")) {
        mainDrugs.push({ drugName: name, dose: dose || null, doseUnit: unit || null });
      } else if (cat.includes("溶") || cat.includes("補液") || cat.includes("前")) {
        diluents.push({ name, volume: dose || null, unit: unit || null });
      } else if (cat.includes("フラッシュ")) {
        flushes.push({ name, volume: dose || null, unit: unit || null });
      }
    }
    if (routeVal && !route) route = routeVal;
    if (rateVal && !infusionRate) infusionRate = rateVal;
    if (dayVal && !administrationDays.includes(dayVal)) administrationDays.push(dayVal);
    if (noteText) notes.push({ text: noteText });
  }

  const firstRow = rows[0];
  const trace: SourceTrace = {
    sourceType: "excel",
    fileName,
    sheetName,
    cellRange: firstRow?.[0]?.address,
    quotedText: rows.map((r) => r.map((c) => c.value).join("\t")).join("\n"),
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
