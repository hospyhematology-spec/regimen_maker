import fs from "fs";
import type { SourceTrace, RichBulletItem } from "@/types/regimen";

interface PdfExtractResult {
  text: string;
  pages: string[];
  sourceTraces: SourceTrace[];
}

/**
 * PDFファイルからテキストを抽出する
 */
export async function parsePdfFile(filePath: string): Promise<PdfExtractResult> {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const pdfParse = require("pdf-parse");
  const dataBuffer = fs.readFileSync(filePath);
  const fileName = filePath.split(/[\\/]/).pop() ?? filePath;

  const data = await pdfParse(dataBuffer);
  const fullText: string = data.text;
  const numPages: number = data.numpages;

  // ページ単位でテキストを分割（近似）
  const lines = fullText.split("\n");
  const linesPerPage = Math.ceil(lines.length / numPages);
  const pages: string[] = [];
  for (let i = 0; i < numPages; i++) {
    pages.push(lines.slice(i * linesPerPage, (i + 1) * linesPerPage).join("\n"));
  }

  const sourceTraces: SourceTrace[] = pages.map((pageText, i) => ({
    sourceType: "pdf" as const,
    fileName,
    page: i + 1,
    quotedText: pageText.slice(0, 200),
    aiInterpretation: `PDF ${i + 1}ページ目のテキスト`,
    confidence: 0.7,
  }));

  return { text: fullText, pages, sourceTraces };
}

/**
 * PDFテキストから箇条書きを抽出する（行頭パターンベース）
 */
export function extractBulletItemsFromText(
  text: string,
  fileName: string,
  page: number,
  confidence = 0.65
): RichBulletItem[] {
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
  const bullets: RichBulletItem[] = [];

  for (const line of lines) {
    if (/^[・•●◆◇▶\-\d+\.\)]\s*/.test(line) && line.length > 2) {
      const cleanText = line.replace(/^[・•●◆◇▶\-\d+\.\)]\s*/, "").trim();
      if (cleanText) {
        bullets.push({
          text: cleanText,
          sourceTrace: [{
            sourceType: "pdf",
            fileName,
            page,
            quotedText: line,
            aiInterpretation: cleanText,
            confidence,
          }],
        });
      }
    }
  }
  return bullets;
}
