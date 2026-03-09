import * as pdfjsLib from "pdfjs-dist";
// 必要なワーカーの設定（ブラウザ環境必須）
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

import type { SourceTrace, RichBulletItem } from "@/types/regimen";

interface PdfExtractResult {
  text: string;
  pages: string[];
  sourceTraces: SourceTrace[];
}

/**
 * PDFファイルからテキストを抽出する（ブラウザ互換）
 */
export async function parsePdfFile(buffer: ArrayBuffer, fileName: string): Promise<PdfExtractResult> {
  const loadingTask = pdfjsLib.getDocument({ data: buffer });
  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;

  const pages: string[] = [];
  let fullText = "";

  for (let i = 1; i <= numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const strings = content.items.map((item: any) => item.str);
    const pageText = strings.join(" ");
    pages.push(pageText);
    fullText += pageText + "\n";
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
