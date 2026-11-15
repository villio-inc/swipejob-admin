import { NextResponse } from "next/server";
import { findSharedJob, PDF_BUCKET } from "@/lib/candidateJobs";
import { runWithProject, storageSignDownload } from "@/lib/swipeAdmin";

export const dynamic = "force-dynamic";

/** 共有リンクの PDF。毎回 1 時間だけ有効な URL を作ってそこへ転送する */
export async function GET(_req: Request, { params }: { params: Promise<{ token: string }> }) {
  const { token } = await params;
  const shared = await findSharedJob(token).catch(() => null);
  const pdfPath = shared?.job.pdf_path;
  if (!shared || !pdfPath) return new NextResponse("見つかりません", { status: 404 });
  const url = await runWithProject(shared.project, () => storageSignDownload(PDF_BUCKET, pdfPath, 3600));
  return NextResponse.redirect(url, { headers: { "Cache-Control": "no-store" } });
}
