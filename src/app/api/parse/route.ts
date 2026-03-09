import { NextRequest, NextResponse } from "next/server";
import { writeFile, mkdir } from "fs/promises";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import * as cheerio from "cheerio";
import { parseExcelFile } from "@/lib/parsers/excelParser";
import { extractBasicInfo } from "@/lib/extractors/basicInfoExtractor";
import { extractRegimenBlocks } from "@/lib/extractors/regimenExtractor";
import { extractReferencesFromText } from "@/lib/extractors/referenceExtractor";
import { parsePdfFile } from "@/lib/parsers/pdfParser";
import { parseWordFile } from "@/lib/parsers/wordParser";
import type { RegimenMaster, SourceFile, ReferenceItem, RichBulletItem } from "@/types/regimen";

const UPLOAD_DIR = path.join(process.cwd(), "data", "uploads");

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    await mkdir(UPLOAD_DIR, { recursive: true });

    const sourceFiles: SourceFile[] = [];
    const savedPaths: string[] = [];
    let urlInputData: string | null = null;

    // アップロードされたファイルを保存
    for (const [key, val] of formData.entries()) {
      if (val instanceof File) {
        const bytes = await val.arrayBuffer();
        const buffer = Buffer.from(bytes);
        const savedPath = path.join(UPLOAD_DIR, `${Date.now()}_${val.name}`);
        await writeFile(savedPath, buffer);
        savedPaths.push(savedPath);
        sourceFiles.push({
          name: val.name,
          type: detectFileType(val.name),
          url: null,
        });
      } else if (key === "url" && typeof val === "string") {
        urlInputData = val;
        sourceFiles.push({
          name: new URL(val).hostname,
          type: "web",
          url: val,
        });
      }
    }

    if (savedPaths.length === 0 && !urlInputData) {
      return NextResponse.json({ error: "ファイルまたはURLが見つかりません" }, { status: 400 });
    }

    // Excelファイルを解析
    const excelPaths = savedPaths.filter((p) => /\.(xlsx|xls)$/i.test(p));
    let basicInfoCandidate = null;
    let regimenBlocks: RegimenMaster["regimenBlocks"] = [];

    for (const excelPath of excelPaths) {
      const sheets = await parseExcelFile(excelPath);
      const fileName = path.basename(excelPath);

      // 基本情報シートを探す
      const basicSheet = sheets.find((s) => /基本|入力事項|概要/i.test(s.name)) ?? sheets[0];
      if (basicSheet) {
        basicInfoCandidate = extractBasicInfo(basicSheet, fileName);
      }

      // 投与方法シートからレジメンブロックを抽出
      const blocks = extractRegimenBlocks(sheets, fileName);
      regimenBlocks = [...regimenBlocks, ...blocks];
    }

    // PDF / Word ファイルを解析してテキストコンテキストを取得
    let extractedReferences: ReferenceItem[] = [];
    const textFiles = savedPaths.filter((p) => /\.(pdf|docx?)$/i.test(p));
    let combinedTextBlocks: RichBulletItem[] = [];

    for (const textPath of textFiles) {
      const fileName = path.basename(textPath);
      let paragraphs: RichBulletItem[] = [];
      if (/\.pdf$/i.test(textPath)) {
        const result = await parsePdfFile(textPath);
        paragraphs = result.pages.map((p) => ({ text: p }));
      } else if (/\.docx?$/i.test(textPath)) {
        const result = await parseWordFile(textPath);
        paragraphs = result.paragraphs.map((p) => ({ text: p }));
      }
      combinedTextBlocks = [...combinedTextBlocks, ...paragraphs];
    }

    // Web URLが指定された場合、Cheerioでテキストを抽出
    if (urlInputData) {
      try {
        const res = await fetch(urlInputData);
        if (res.ok) {
          const html = await res.text();
          const $ = cheerio.load(html);
          $("script, style, nav, footer, header").remove();
          const webText = $("body").text();
          const pTexts = webText.split(/\n+/).map((t) => t.trim()).filter((t) => t.length > 5);
          const webParagraphs = pTexts.map((p) => ({
            text: p,
            sourceTrace: [{
              sourceType: "web" as const,
              fileName: urlInputData,
              url: urlInputData,
              quotedText: p.slice(0, 200),
              aiInterpretation: "Webからの抽出テキスト",
              confidence: 0.6,
            }],
          }));
          combinedTextBlocks = [...combinedTextBlocks, ...webParagraphs];
        }
      } catch (e) {
        console.warn("Failed to fetch URL:", e);
      }
    }

    if (combinedTextBlocks.length > 0) {
      extractedReferences = extractReferencesFromText(
        combinedTextBlocks,
        "アップロードテキスト群",
        "text"
      );
    }

    // デフォルトのBasicInfo
    const defaultBasicInfo: RegimenMaster["basicInfo"] = {
      treatmentPurpose: { selected: null },
      category: null,
      contraindications: [],
      eligibilityCriteria: [],
      stopCriteria: [],
      doseReductionCriteria: [],
      precautions: [],
      popupNotes: [],
      ...basicInfoCandidate?.value,
    };

    const now = new Date().toISOString();
    const result: RegimenMaster = {
      id: uuidv4(),
      templateVersion: "1.0",
      sourceFiles,
      basicInfo: defaultBasicInfo,
      regimenBlocks,
      supplementaryOrders: [],
      references: extractedReferences,
      audit: {
        createdAt: now,
        updatedAt: now,
        changes: ["初回抽出"],
      },
    };

    // 結果をJSONファイルに保存
    const jsonPath = path.join(process.cwd(), "data", `${result.id}_normalized.json`);
    await writeFile(jsonPath, JSON.stringify(result, null, 2), "utf-8");

    return NextResponse.json({ success: true, data: result });
  } catch (err) {
    console.error("Parse error:", err);
    return NextResponse.json(
      { error: "解析中にエラーが発生しました", detail: String(err) },
      { status: 500 }
    );
  }
}

function detectFileType(filename: string): SourceFile["type"] {
  if (/\.(xlsx|xls)$/i.test(filename)) return "excel";
  if (/\.pdf$/i.test(filename)) return "pdf";
  if (/\.(docx|doc)$/i.test(filename)) return "word";
  return "text";
}
