import { v4 as uuidv4 } from "uuid";
import type { RegimenBlock } from "@/types/regimen";
import type { SheetData } from "@/lib/parsers/excelParser";
import { extractAdministrationSteps } from "@/lib/parsers/excelParser";

/**
 * 投与方法シートからレジメンブロックを抽出する
 * シート名のパターンから複数ブロックを作成する
 */
export function extractRegimenBlocks(
  sheets: SheetData[],
  fileName: string
): RegimenBlock[] {
  const blocks: RegimenBlock[] = [];

  // 投与方法シートを特定（シート名に「投与」「レジメン」「Cycle」等を含むもの）
  const regimenSheets = sheets.filter((s) =>
    /投与|レジメン|Cycle|cycle|投薬|方法/i.test(s.name)
  );

  if (regimenSheets.length === 0) {
    // 基本情報以外のシートをすべて投与シートとして扱う
    const nonBasicSheets = sheets.filter(
      (s) => !/基本|入力事項|表紙|概要/i.test(s.name)
    );
    regimenSheets.push(...nonBasicSheets);
  }

  for (let i = 0; i < regimenSheets.length; i++) {
    const sheet = regimenSheets[i];
    const steps = extractAdministrationSteps(sheet, fileName);

    // シート名からコース範囲を推定
    const courseRange = inferCourseRange(sheet.name, i, regimenSheets.length);

    blocks.push({
      regimenId: uuidv4(),
      regimenLabel: inferRegimenLabel(sheet.name, i),
      applicableCourseRange: courseRange,
      sourceSheetName: sheet.name,
      administrationSteps: steps,
      oralOrders: [],
    });
  }

  return blocks;
}

function inferRegimenLabel(sheetName: string, index: number): string {
  const courseMatch = sheetName.match(/(\d+)コース|Cycle\s*(\d+)/i);
  if (courseMatch) {
    const num = courseMatch[1] ?? courseMatch[2];
    return `レジメン${index + 1}（${num}コース目）`;
  }
  if (sheetName.includes("Cycle") || sheetName.includes("cycle")) {
    return `レジメン${index + 1}（${sheetName}）`;
  }
  return `レジメン${index + 1}（${sheetName}）`;
}

function inferCourseRange(sheetName: string, index: number, total: number): string | null {
  const rangeMatch = sheetName.match(/(\d+)[〜~\-−](\d+)コース/);
  if (rangeMatch) return `${rangeMatch[1]}〜${rangeMatch[2]}コース`;

  const singleMatch = sheetName.match(/(\d+)コース|Cycle\s*(\d+)/i);
  if (singleMatch) {
    const num = singleMatch[1] ?? singleMatch[2];
    return `${num}コース目`;
  }

  if (total === 1) return null;
  if (index === 0) return "1コース目";
  if (index === total - 1) return `${index + 1}コース目以降`;
  return `${index + 1}コース目`;
}
