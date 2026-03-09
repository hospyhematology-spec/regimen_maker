import type { BasicInfo, SourceTrace, RichBulletItem } from "@/types/regimen";
import type { SheetData } from "@/lib/parsers/excelParser";
import { getEffectiveValue } from "@/lib/parsers/excelParser";

interface BasicInfoCandidate {
  value: Partial<BasicInfo>;
  sourceTraces: SourceTrace[];
}

/**
 * 基本情報シートから BasicInfo を抽出する
 * テンプレートのセルレイアウトに依存するため、マッピングは設定可能にする
 */
export function extractBasicInfo(
  sheet: SheetData,
  fileName: string
): BasicInfoCandidate {
  const { cells } = sheet;
  const sheetName = sheet.name;

  const trace = (addr: string, text: string, interp: string, conf = 0.85): SourceTrace => ({
    sourceType: "excel",
    fileName,
    sheetName,
    cellRange: addr,
    quotedText: text,
    aiInterpretation: interp,
    confidence: conf,
  });

  // 典型的な基本情報シートのレイアウト例（実際のテンプレートに合わせて調整）
  // ラベルが A列, 値が B列/C列 というパターン
  let applicationDate: string | null = null;
  let applicantDoctor: string | null = null;
  let departmentChief: string | null = null;
  let cancerType: string | null = null;
  let regimenName: string | null = null;
  let courseLengthDays: number | null = null;
  let totalCourses: string | null = null;
  let treatmentPurposeRaw: string | null = null;
  let categoryRaw: string | null = null;

  const sourceTraces: SourceTrace[] = [];

  // 全行をスキャンしてラベルパターンマッチング
  for (const row of sheet.rows) {
    for (let ci = 0; ci < row.length - 1; ci++) {
      const label = row[ci]?.value?.trim() ?? "";
      const valueAddr = row[ci + 1]?.address ?? "";
      const value = row[ci + 1]?.value?.trim() ?? getEffectiveValue(cells, valueAddr).trim();

      if (!value) continue;

      // ラベルベースのパターンマッチ
      if (/申請日|申請年月日/.test(label) && !applicationDate) {
        applicationDate = value;
        sourceTraces.push(trace(valueAddr, value, "申請日", 0.9));
      } else if (/申請医師|申請者/.test(label) && !applicantDoctor) {
        applicantDoctor = value;
        sourceTraces.push(trace(valueAddr, value, "申請医師名", 0.9));
      } else if (/科長|診療科長|部長/.test(label) && !departmentChief) {
        departmentChief = value;
        sourceTraces.push(trace(valueAddr, value, "診療科長名", 0.9));
      } else if (/がん種|癌種|腫瘍種|病名/.test(label) && !cancerType) {
        cancerType = value;
        sourceTraces.push(trace(valueAddr, value, "がん種", 0.9));
      } else if (/レジメン名|レジメン/.test(label) && !regimenName) {
        regimenName = value;
        sourceTraces.push(trace(valueAddr, value, "レジメン名", 0.95));
      } else if (/コース.*(日|数)|投与期間/.test(label) && courseLengthDays === null) {
        const n = parseInt(value.replace(/[^\d]/g, ""), 10);
        if (!isNaN(n)) {
          courseLengthDays = n;
          sourceTraces.push(trace(valueAddr, value, "コース日数", 0.85));
        }
      } else if (/総コース数|最大コース/.test(label) && !totalCourses) {
        totalCourses = value;
        sourceTraces.push(trace(valueAddr, value, "総コース数", 0.85));
      } else if (/治療目的|適応/.test(label) && !treatmentPurposeRaw) {
        treatmentPurposeRaw = value;
        sourceTraces.push(trace(valueAddr, value, "治療目的", 0.8));
      } else if (/区分|外来|入院/.test(label) && !categoryRaw) {
        categoryRaw = value;
        sourceTraces.push(trace(valueAddr, value, "入院・外来区分", 0.8));
      }
    }
  }

  // 治療目的を正規化
  let treatmentPurpose: BasicInfo["treatmentPurpose"] = { selected: null };
  if (treatmentPurposeRaw) {
    if (treatmentPurposeRaw.includes("進行再発")) {
      treatmentPurpose = { selected: "進行再発" };
    } else if (treatmentPurposeRaw.includes("術前")) {
      treatmentPurpose = { selected: "術前補助化学療法" };
    } else if (treatmentPurposeRaw.includes("術後")) {
      treatmentPurpose = { selected: "術後補助化学療法" };
    } else {
      treatmentPurpose = { selected: "その他", otherText: treatmentPurposeRaw };
    }
  }

  // 区分を正規化
  let category: BasicInfo["category"] = null;
  if (categoryRaw) {
    if (categoryRaw.includes("入院") && categoryRaw.includes("外来")) {
      category = "入院・外来";
    } else if (categoryRaw.includes("外来")) {
      category = "外来";
    } else if (categoryRaw.includes("入院")) {
      category = "入院";
    }
  }

  // 箇条書き項目はシートの特定エリアから抽出（行範囲は実際のテンプレートに合わせる）
  // ここでは仮の行範囲を使用
  const contraindications = extractBulletItemsFromRows(sheet, fileName, "禁忌");
  const eligibilityCriteria = extractBulletItemsFromRows(sheet, fileName, "投与基準");
  const stopCriteria = extractBulletItemsFromRows(sheet, fileName, "中止基準");
  const doseReductionCriteria = extractBulletItemsFromRows(sheet, fileName, "減量基準");
  const precautions = extractBulletItemsFromRows(sheet, fileName, "注意事項");
  const popupNotes = extractBulletItemsFromRows(sheet, fileName, "ポップアップ");

  const basicInfo: Partial<BasicInfo> = {
    applicationDate,
    applicantDoctor,
    departmentChief,
    cancerType,
    regimenName,
    courseLengthDays,
    totalCourses,
    treatmentPurpose,
    category,
    contraindications,
    eligibilityCriteria,
    stopCriteria,
    doseReductionCriteria,
    precautions,
    popupNotes,
  };

  return { value: basicInfo, sourceTraces };
}

/**
 * シートの行から特定ラベルに続く箇条書き領域を抽出する
 */
function extractBulletItemsFromRows(
  sheet: SheetData,
  fileName: string,
  labelKeyword: string
): RichBulletItem[] {
  const items: RichBulletItem[] = [];
  let capturing = false;

  for (const row of sheet.rows) {
    const firstCell = row[0]?.value?.trim() ?? "";
    if (firstCell.includes(labelKeyword)) {
      capturing = true;
      continue;
    }
    // 次のラベルに到達したら停止
    if (capturing && firstCell && /^[^\s]/.test(firstCell) && !firstCell.startsWith("・") && !firstCell.startsWith("　")) {
      capturing = false;
    }
    if (capturing) {
      const text = row.map((c: { value: string }) => c.value).join("").trim();
      if (text) {
        items.push({
          text,
          sourceTrace: [{
            sourceType: "excel",
            fileName,
            sheetName: sheet.name,
            cellRange: row[0]?.address,
            quotedText: text,
            aiInterpretation: `${labelKeyword}の項目`,
            confidence: 0.75,
          }],
        });
      }
    }
  }
  return items;
}
