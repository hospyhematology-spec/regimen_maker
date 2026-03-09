import type { ReferenceItem, RichBulletItem, SourceTrace } from "@/types/regimen";

export function extractReferencesFromText(
  paragraphs: RichBulletItem[],
  fileName: string,
  sourceType: SourceTrace["sourceType"]
): ReferenceItem[] {
  const references: ReferenceItem[] = [];

  for (const p of paragraphs) {
    const text = p.text;
    if (/(PMID|PubMed|DOI|N Engl J Med|Lancet|J Clin Oncol|Ann Oncol|文献)/i.test(text)) {
      const pmidMatch = text.match(/PMID[:\s]*(\d+)/i);
      references.push({
        title: text.substring(0, 100) + (text.length > 100 ? "..." : ""),
        pmid: pmidMatch ? pmidMatch[1] : null,
        representative: false,
        sourceTrace: [
          {
            sourceType,
            fileName,
            quotedText: text,
            aiInterpretation: "参考文献の候補",
            confidence: 0.8,
          },
        ],
      } as ReferenceItem);
    }
  }

  return references;
}
