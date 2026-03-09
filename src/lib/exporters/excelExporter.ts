import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import type { RegimenMaster } from "@/types/regimen";

/**
 * テンプレートExcelにRegimenMasterのデータを書き戻す
 */
export async function exportToExcel(
  templatePath: string,
  data: RegimenMaster,
  outputPath: string
): Promise<string> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  // 基本情報シートへの書き戻し
  const basicSheet = findSheet(workbook, /基本|入力事項|概要/i);
  if (basicSheet) {
    writeBasicInfo(basicSheet, data);
  }

  // 投与方法シートへの書き戻し（ブロック毎）
  for (let i = 0; i < data.regimenBlocks.length; i++) {
    const block = data.regimenBlocks[i];
    const sheetName = block.sourceSheetName ?? "";
    const regimenSheet = sheetName
      ? workbook.getWorksheet(sheetName)
      : workbook.worksheets[i + 1];
    if (regimenSheet) {
      writeRegimenBlock(regimenSheet, block);
    }
  }

  // 出力先ディレクトリを確保
  const outDir = path.dirname(outputPath);
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

function findSheet(workbook: ExcelJS.Workbook, pattern: RegExp): ExcelJS.Worksheet | undefined {
  return workbook.worksheets.find((ws) => pattern.test(ws.name));
}

/**
 * BasicInfoをシートに書き込む（セルマッピングは実テンプレートに合わせて要調整）
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function writeBasicInfo(sheet: ExcelJS.Worksheet, data: RegimenMaster) {
  const { basicInfo } = data;

  // ラベル → セルアドレスのマッピング (テンプレートに依存)
  const fieldMap: Record<string, string | null> = {
    applicationDate: null,
    applicantDoctor: null,
    departmentChief: null,
    cancerType: null,
    regimenName: null,
    courseLengthDays: null,
    totalCourses: null,
  };

  // シートを全スキャンしてラベルとなるセルの隣に値を書く
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
      const text = String(cell.value ?? "").trim();
      const nextCell = row.getCell(colIndex + 1);

      if (/申請日|申請年月日/.test(text) && basicInfo.applicationDate) {
        nextCell.value = basicInfo.applicationDate;
      } else if (/申請医師|申請者/.test(text) && basicInfo.applicantDoctor) {
        nextCell.value = basicInfo.applicantDoctor;
      } else if (/科長|診療科長|部長/.test(text) && basicInfo.departmentChief) {
        nextCell.value = basicInfo.departmentChief;
      } else if (/がん種|癌種|病名/.test(text) && basicInfo.cancerType) {
        nextCell.value = basicInfo.cancerType;
      } else if (/レジメン名/.test(text) && basicInfo.regimenName) {
        nextCell.value = basicInfo.regimenName;
      } else if (/コース.*日|投与期間/.test(text) && basicInfo.courseLengthDays) {
        nextCell.value = basicInfo.courseLengthDays;
      } else if (/総コース数/.test(text) && basicInfo.totalCourses) {
        nextCell.value = basicInfo.totalCourses;
      } else if (/区分/.test(text) && basicInfo.category) {
        nextCell.value = basicInfo.category;
      }
    });
  });

  void fieldMap;
}

/**
 * RegimenBlockを投与方法シートに書き込む
 */
function writeRegimenBlock(sheet: ExcelJS.Worksheet, block: import("@/types/regimen").RegimenBlock) {
  // ステップ番号を元に行を特定して書き込む
  for (const step of block.administrationSteps) {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      const firstCell = String(row.getCell(1).value ?? "").trim();
      if (firstCell === String(step.stepNo)) {
        // 主薬の薬品名・用量を書き込む
        if (step.mainDrugs[0]) {
          const drug = step.mainDrugs[0];
          // カラムインデックスはテンプレートに依存
          const nameCell = row.getCell(3);
          const doseCell = row.getCell(4);
          const unitCell = row.getCell(5);
          if (!nameCell.value) nameCell.value = drug.drugName;
          if (!doseCell.value && drug.dose) doseCell.value = drug.dose;
          if (!unitCell.value && drug.doseUnit) unitCell.value = drug.doseUnit;
        }
      }
    });
  }
}
