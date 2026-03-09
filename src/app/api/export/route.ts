import { NextRequest, NextResponse } from "next/server";
import { writeFile, mkdir } from "fs/promises";
import path from "path";
import { exportToExcel } from "@/lib/exporters/excelExporter";
import type { RegimenMaster } from "@/types/regimen";

const OUTPUT_DIR = path.join(process.cwd(), "data", "outputs");

export async function POST(req: NextRequest) {
  try {
    await mkdir(OUTPUT_DIR, { recursive: true });

    const formData = await req.formData();
    const templateFile = formData.get("template") as File | null;
    const dataJson = formData.get("data") as string | null;

    if (!templateFile || !dataJson) {
      return NextResponse.json(
        { error: "テンプレートファイルとデータが必要です" },
        { status: 400 }
      );
    }

    const data: RegimenMaster = JSON.parse(dataJson);

    // テンプレートを一時保存
    const templateBytes = await templateFile.arrayBuffer();
    const templatePath = path.join(OUTPUT_DIR, `template_${Date.now()}.xlsx`);
    await writeFile(templatePath, Buffer.from(templateBytes));

    // 出力ファイルパス
    const outputPath = path.join(OUTPUT_DIR, `${data.id}_output.xlsx`);
    const auditLogPath = path.join(OUTPUT_DIR, `${data.id}_audit-log.json`);
    const normalizedPath = path.join(OUTPUT_DIR, `${data.id}_normalized.json`);

    // Excel書き戻し
    await exportToExcel(templatePath, data, outputPath);

    // JSON・監査ログ保存
    const now = new Date().toISOString();
    const auditLog = {
      regimenId: data.id,
      exportedAt: now,
      operator: data.audit?.operator ?? null,
      changes: data.audit?.changes ?? [],
    };
    await writeFile(auditLogPath, JSON.stringify(auditLog, null, 2), "utf-8");
    await writeFile(normalizedPath, JSON.stringify(data, null, 2), "utf-8");

    return NextResponse.json({
      success: true,
      files: {
        excel: `/api/download?file=${encodeURIComponent(path.basename(outputPath))}`,
        json: `/api/download?file=${encodeURIComponent(path.basename(normalizedPath))}`,
        auditLog: `/api/download?file=${encodeURIComponent(path.basename(auditLogPath))}`,
      },
    });
  } catch (err) {
    console.error("Export error:", err);
    return NextResponse.json(
      { error: "エクスポート中にエラーが発生しました", detail: String(err) },
      { status: 500 }
    );
  }
}
