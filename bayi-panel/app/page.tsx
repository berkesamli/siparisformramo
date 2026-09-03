import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";

export const dynamic = "force-dynamic";

export default async function Home() {
  const user = await getSessionUser();
  if (!user) redirect("/giris");
  redirect(user.kind === "admin" ? "/yonetim" : "/panel");
}
