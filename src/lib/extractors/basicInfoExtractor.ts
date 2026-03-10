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

  // 2パス方式で全セクションを一度に抽出（振り分けの精度が高い）
  const bulletSections = extractAllBulletSections(sheet, fileName);

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
    contraindications: bulletSections.contraindications ?? [],
    eligibilityCriteria: bulletSections.eligibilityCriteria ?? [],
    stopCriteria: bulletSections.stopCriteria ?? [],
    doseReductionCriteria: bulletSections.doseReductionCriteria ?? [],
    precautions: bulletSections.precautions ?? [],
    popupNotes: bulletSections.popupNotes ?? [],
  };

  return { value: basicInfo, sourceTraces };
}

/**
 * シートの全行を走査してセクションマップを構築し、
 * 指定ラベルに対応する行範囲だけを返すことで、
 * 振り分けの精度を大幅に向上させる（2パス方式）
 */

interface SectionEntry {
  keyword: string;
  rowIndex: number;
}

// 認識対象の全セクションラベルと対応キー
const SECTION_DEFS: { key: string; patterns: RegExp; keywords: string[] }[] = [
  { key: "contraindications",     patterns: /禁忌/,                         keywords: ["禁忌"] },
  { key: "eligibilityCriteria",   patterns: /投与基準|適応基準|対象患者|適応症例/, keywords: ["投与基準", "適応基準", "対象患者", "適応症例"] },
  { key: "stopCriteria",          patterns: /中止基準|投与中止|休薬基準/,     keywords: ["中止基準", "投与中止", "休薬基準"] },
  { key: "doseReductionCriteria", patterns: /減量基準|用量調整|用量変更/,     keywords: ["減量基準", "用量調整", "用量変更"] },
  { key: "precautions",           patterns: /注意事項|注意点|副作用対策|管理方法/, keywords: ["注意事項", "注意点", "副作用対策", "管理方法"] },
  { key: "popupNotes",            patterns: /ポップアップ|コメント|注記|備考/, keywords: ["ポップアップ", "コメント", "注記", "備考"] },
];

function extractAllBulletSections(
  sheet: SheetData,
  fileName: string
): Record<string, RichBulletItem[]> {
  const result: Record<string, RichBulletItem[]> = {};
  SECTION_DEFS.forEach(d => { result[d.key] = []; });

  // パス1: 各行を走査して各セクションの開始行インデックスを記録
  const sectionBoundaries: SectionEntry[] = [];
  for (let ri = 0; ri < sheet.rows.length; ri++) {
    const row = sheet.rows[ri];
    const rowText = row.map((c: { value: string }) => c.value?.trim() ?? "").join(" ");
    for (const def of SECTION_DEFS) {
      if (def.patterns.test(rowText)) {
        sectionBoundaries.push({ keyword: def.key, rowIndex: ri });
        break; // 1行に複数ラベルがある場合は最初のみ
      }
    }
  }

  if (sectionBoundaries.length === 0) return result;

  // パス2: 各セクションの行範囲からデータを抽出
  for (let si = 0; si < sectionBoundaries.length; si++) {
    const { keyword, rowIndex } = sectionBoundaries[si];
    const nextRowIndex = sectionBoundaries[si + 1]?.rowIndex ?? sheet.rows.length;

    const items: RichBulletItem[] = [];
    // ラベル行自体（rowIndex）は飛ばして次の行から取り込む
    // ただしラベル行の右側にデータがある場合は取り込む
    const labelRow = sheet.rows[rowIndex];
    const labelRowData = labelRow?.slice(1)
      .map((c: { value: string }) => c.value?.trim() ?? "")
      .filter(Boolean).join(" ") ?? "";
    if (labelRowData) {
      items.push({
        text: labelRowData,
        sourceTrace: [{ sourceType: "excel", fileName, sheetName: sheet.name, cellRange: labelRow[0]?.address, quotedText: labelRowData, aiInterpretation: `${keyword}の項目`, confidence: 0.85 }],
      });
    }

    // ラベル行の次の行から次セクション開始行の直前まで
    for (let ri = rowIndex + 1; ri < nextRowIndex; ri++) {
      const row = sheet.rows[ri];
      const text = row
        .map((c: { value: string }) => c.value?.trim() ?? "")
        .filter(Boolean).join(" ").trim();
      if (text) {
        items.push({
          text,
          sourceTrace: [{ sourceType: "excel", fileName, sheetName: sheet.name, cellRange: row[0]?.address, quotedText: text, aiInterpretation: `${keyword}の項目`, confidence: 0.75 }],
        });
      }
    }
    result[keyword] = items;
  }

  return result;
}

// 後方互換のためのラッパー（外部から呼び出し可能）
export function extractBulletItemsByKeyword(
  sheet: SheetData,
  fileName: string,
  sectionKey: string
): RichBulletItem[] {
  const res = extractAllBulletSections(sheet, fileName);
  return res[sectionKey] ?? [];
}
