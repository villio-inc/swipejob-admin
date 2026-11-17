"use client";

import { useRef, useState } from "react";
import { useRouter } from "next/navigation";
import { addJobsFromPdfs, getPdfUploadUrl } from "@/app/admin/actions";

type Item = { name: string; state: "待機中" | "アップロード中" | "登録中" | "完了" | "失敗"; note?: string };

/** 求人票 PDF をまとめてドラッグ＆ドロップすると、1 ファイル 1 件の提案求人として登録する */
export default function PdfDropZone({ userId }: { userId: string }) {
  const [over, setOver] = useState(false);
  const [items, setItems] = useState<Item[]>([]);
  const [busy, setBusy] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const router = useRouter();

  const set = (i: number, patch: Partial<Item>) =>
    setItems((prev) => prev.map((it, k) => (k === i ? { ...it, ...patch } : it)));

  async function handle(fileList: FileList | File[]) {
    const files = Array.from(fileList).filter((f) => /\.pdf$/i.test(f.name) || f.type === "application/pdf");
    if (!files.length || busy) return;
    setBusy(true);
    setItems(files.map((f) => ({ name: f.name, state: "待機中" })));

    // 1) Storage に直接アップロード（3 件ずつ並行）
    const uploaded: { i: number; path: string; fileName: string }[] = [];
    let next = 0;
    const worker = async () => {
      while (next < files.length) {
        const i = next++;
        const f = files[i];
        if (f.size > 20 * 1024 * 1024) {
          set(i, { state: "失敗", note: "20MB までです" });
          continue;
        }
        set(i, { state: "アップロード中" });
        const up = await getPdfUploadUrl(userId, f.name);
        if (!up.ok) {
          set(i, { state: "失敗", note: up.error });
          continue;
        }
        const res = await fetch(up.uploadUrl, { method: "PUT", headers: { "Content-Type": "application/pdf", "x-upsert": "false" }, body: f }).catch(() => null);
        if (!res?.ok) {
          set(i, { state: "失敗", note: "アップロードに失敗しました" });
          continue;
        }
        set(i, { state: "登録中" });
        uploaded.push({ i, path: up.path, fileName: f.name });
      }
    };
    await Promise.all([worker(), worker(), worker()]);

    // 2) PDF を読んで提案求人として登録
    if (uploaded.length) {
      const results = await addJobsFromPdfs(userId, uploaded.map(({ path, fileName }) => ({ path, fileName })));
      results.forEach((r, k) => set(uploaded[k].i, r.ok ? { state: "完了", note: r.label } : { state: "失敗", note: r.error }));
      router.refresh();
    }
    setBusy(false);
  }

  return (
    <div>
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setOver(false);
          handle(e.dataTransfer.files);
        }}
        onClick={() => !busy && inputRef.current?.click()}
        className={`cursor-pointer rounded-lg border-2 border-dashed px-4 py-5 text-center text-sm transition ${
          over ? "border-indigo-500 bg-indigo-50 text-indigo-700" : "border-slate-300 text-slate-500 hover:border-slate-400 hover:bg-slate-50"
        }`}
      >
        <div className="font-semibold">📄 求人票 PDF をまとめてここにドラッグ＆ドロップ</div>
        <div className="mt-0.5 text-xs">（クリックして選ぶこともできます。1 ファイル＝1 求人として登録され、共有リンクで候補者が見られます）</div>
        <input
          ref={inputRef}
          type="file"
          accept="application/pdf,.pdf"
          multiple
          className="hidden"
          onChange={(e) => {
            if (e.target.files) handle(e.target.files);
            e.target.value = "";
          }}
        />
      </div>
      {items.length > 0 && (
        <ul className="mt-2 space-y-1 text-xs">
          {items.map((it, i) => (
            <li key={i} className="flex flex-wrap gap-x-2">
              <span
                className={
                  it.state === "完了" ? "text-emerald-700" : it.state === "失敗" ? "text-rose-700" : "text-slate-500"
                }
              >
                {it.state === "完了" ? "✓" : it.state === "失敗" ? "✕" : "…"} {it.state}
              </span>
              <span className="text-slate-700">{it.name}</span>
              {it.note && <span className="text-slate-400">→ {it.note}</span>}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
