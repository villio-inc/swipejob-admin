import "server-only";
import { extractText, getDocumentProxy } from "unpdf";

export type PdfJobInfo = { jobId: string | null; company: string | null; title: string | null };

/**
 * 求人票 PDF の 1〜2 ページ目の文字から、求人 ID・会社名・求人タイトルを推測する。
 * 取れなかった項目は null（呼び出し側でファイル名などを使う）。
 */
export async function readPdfJobInfo(data: ArrayBuffer): Promise<PdfJobInfo> {
  const pdf = await getDocumentProxy(new Uint8Array(data));
  const { text } = await extractText(pdf, { mergePages: false });
  const head = (text as string[]).slice(0, 2).join("\n");
  const lines = head
    .split(/\n+/)
    .map((l) => l.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  const jobId = head.match(/hl-ap-\d+/)?.[0] ?? null;
  const companyRe = /((?:株式会社|有限会社|合同会社)[^\s　／/（(]{1,30}|[^\s　／/（(:：]{1,30}(?:株式会社|有限会社|合同会社))/;
  const company = lines.map((l) => l.match(companyRe)?.[1]).find(Boolean) ?? null;
  // タイトルは【】を含む、それなりに長い最初の行
  const title = lines.find((l) => /【.+】/.test(l) && l.length >= 10 && l.length <= 150) ?? null;
  return { jobId, company, title };
}
