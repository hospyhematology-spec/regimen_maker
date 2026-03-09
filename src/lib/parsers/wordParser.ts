import fs from "fs";
import mammoth from "mammoth";
import type { SourceTrace, RichBulletItem } from "@/types/regimen";

interface WordExtractResult {
  text: string;
  paragraphs: string[];
  sourceTraces: SourceTrace[];
}

/**
 * Word (.docx) ファイルからテキストを抽出する
 */
export async function parseWordFile(filePath: string): Promise<WordExtractResult> {
  const buffer = fs.readFileSync(filePath);
  const fileName = filePath.split(/[\\/]/).pop() ?? filePath;

  const result = await mammoth.extractRawText({ buffer });
  const text = result.value;
  const paragraphs = text.split("\n").map((p) => p.trim()).filter(Boolean);

  const sourceTraces: SourceTrace[] = paragraphs.map((p, i) => ({
    sourceType: "word" as const,
    fileName,
    paragraphIndex: i,
    quotedText: p.slice(0, 200),
    aiInterpretation: `Wordの段落 ${i + 1}`,
    confidence: 0.7,
  }));

  return { text, paragraphs, sourceTraces };
}

/**
 * Wordの段落リストから箇条書き項目を抽出する
 */
export function extractBulletItemsFromParagraphs(
  paragraphs: string[],
  fileName: string,
  confidence = 0.65
): RichBulletItem[] {
  return paragraphs
    .filter((p) => /^[・•●◆◇▶\-\d+\.\)]\s*/.test(p) && p.length > 2)
    .map((p, i) => {
      const cleanText = p.replace(/^[・•●◆◇▶\-\d+\.\)]\s*/, "").trim();
      return {
        text: cleanText,
        sourceTrace: [{
          sourceType: "word" as const,
          fileName,
          paragraphIndex: i,
          quotedText: p,
          aiInterpretation: cleanText,
          confidence,
        }],
      };
    });
}
