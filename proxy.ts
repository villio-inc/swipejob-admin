import { NextResponse, type NextRequest } from "next/server";
import { isProjectKey, PROJECT_COOKIE, PROJECT_HEADER, PROJECT_PARAM } from "@/lib/projects";
import { SESSION_COOKIE, verifySession } from "@/lib/session";

// /admin 以下（SwipeJob 顧客管理）は個人情報を含むため、ログインした担当者だけが見られる。
const PUBLIC = [/^\/admin\/login\/?$/, /^\/admin\/setup\/?$/, /^\/admin\/invite\/[^/]+\/?$/];

export async function proxy(req: NextRequest) {
  const { pathname, search } = req.nextUrl;

  // URL に ?p=driver のようなプロジェクト指定があれば、その表示に使い、Cookie にも覚えておく。
  // （URL を受け取った人が別のプロジェクトを選んでいても、送った人と同じ画面が開くようにする）
  const project = req.nextUrl.searchParams.get(PROJECT_PARAM);
  const headers = new Headers(req.headers);
  headers.delete(PROJECT_HEADER);
  if (isProjectKey(project)) headers.set(PROJECT_HEADER, project);
  const pass = () => {
    const res = NextResponse.next({ request: { headers } });
    if (isProjectKey(project) && req.cookies.get(PROJECT_COOKIE)?.value !== project) {
      res.cookies.set(PROJECT_COOKIE, project, { path: "/", sameSite: "lax", maxAge: 60 * 60 * 24 * 365 });
    }
    return res;
  };

  if (PUBLIC.some((re) => re.test(pathname))) return pass();

  let ok = false;
  try {
    ok = !!(await verifySession(req.cookies.get(SESSION_COOKIE)?.value));
  } catch {
    ok = false;
  }
  if (ok) return pass();

  const url = req.nextUrl.clone();
  url.pathname = "/admin/login";
  url.search = `?next=${encodeURIComponent(pathname + search)}`;
  return NextResponse.redirect(url);
}

export const config = { matcher: ["/admin", "/admin/:path*"] };
