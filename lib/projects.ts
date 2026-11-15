// 管理画面で切り替えられるプロジェクト（事業）の一覧。
// 候補者・求人・架電記録などはプロジェクトごとに別の Supabase に入っている。
// メンバー（crm_staff）と招待（crm_invites）だけは全プロジェクト共通で、先頭のプロジェクトの DB を正とする。

export type ProjectKey = "denken" | "driver";

export const PROJECTS: { key: ProjectKey; label: string; urlEnv: string; keyEnv: string }[] = [
  { key: "denken", label: "電験・設備管理", urlEnv: "SUPABASE_URL", keyEnv: "SUPABASE_SERVICE_ROLE_KEY" },
  { key: "driver", label: "ドライバー", urlEnv: "DRIVER_SUPABASE_URL", keyEnv: "DRIVER_SUPABASE_SERVICE_ROLE_KEY" },
];

export const PRIMARY_PROJECT: ProjectKey = PROJECTS[0].key;

/** 選択中のプロジェクトを覚えておく Cookie */
export const PROJECT_COOKIE = "crm_project";
/** URL に付けるプロジェクト指定（?p=driver）。URL を共有した相手にも同じプロジェクトで開かせるため */
export const PROJECT_PARAM = "p";
/** proxy.ts が URL の指定をサーバー側へ渡すためのリクエストヘッダー */
export const PROJECT_HEADER = "x-crm-project";

export function isProjectKey(v: unknown): v is ProjectKey {
  return typeof v === "string" && PROJECTS.some((p) => p.key === v);
}
