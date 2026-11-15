"use client";

import { useEffect, useTransition } from "react";
import { usePathname, useSearchParams } from "next/navigation";
import { switchProject } from "@/app/admin/actions";
import { PROJECT_PARAM } from "@/lib/projects";

type Props = { current: string; projects: { key: string; label: string }[] };

/** ヘッダーのプロジェクト切り替え。選ぶと、そのプロジェクトの候補者一覧へ移る */
export default function ProjectSwitcher({ current, projects }: Props) {
  const [pending, start] = useTransition();
  const pathname = usePathname();
  const search = useSearchParams();

  // アドレスバーの URL に ?p=<プロジェクト> を付けておく。URL をコピーして他のメンバーに送っても、
  // 相手側で同じプロジェクトの画面が開く
  useEffect(() => {
    if (projects.length < 2 || search.get(PROJECT_PARAM) === current) return;
    const params = new URLSearchParams(search.toString());
    params.set(PROJECT_PARAM, current);
    window.history.replaceState(null, "", `${pathname}?${params}`);
  }, [pathname, search, current, projects.length]);

  if (projects.length < 2) return null;
  return (
    <select
      aria-label="プロジェクト"
      value={current}
      disabled={pending}
      onChange={(e) => {
        const next = e.target.value;
        start(() => switchProject(next));
      }}
      className="max-w-[9.5rem] rounded-lg border border-slate-300 bg-white py-1 pl-2 pr-7 text-sm font-semibold text-slate-800 disabled:opacity-60"
    >
      {projects.map((p) => (
        <option key={p.key} value={p.key}>
          {p.label}
        </option>
      ))}
    </select>
  );
}
