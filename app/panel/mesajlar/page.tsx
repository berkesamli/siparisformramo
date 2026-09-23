import { redirect } from "next/navigation";
import { getSessionUser } from "@/lib/auth";
import { isMesajci, isOwner } from "@/data/users";
import { dbConfigured } from "@/lib/mesaj/db";
import { kanalDurumu } from "@/lib/mesaj/gonder";
import { taslakHazir } from "@/lib/mesaj/taslak";
import Inbox from "@/components/mesaj/Inbox";

export const dynamic = "force-dynamic";

// Gelen kutusu: WhatsApp (Cloud API numarası), Instagram (olga.cerceve) ve
// Gmail hesapları tek listede. Yalnızca MESAJ_USERNAMES'teki çalışanlar (ve sahipler).
export default async function MesajlarPage({ searchParams }: { searchParams?: { k?: string } }) {
  const user = await getSessionUser();
  if (!user) redirect("/giris?next=/panel/mesajlar");
  if (user.role !== "staff") redirect("/portal");
  if (!isMesajci(user.username)) redirect("/");

  return (
    <main className="container ib-container">
      <Inbox
        me={{ username: user.username, name: user.name, owner: isOwner(user.username) }}
        dbHazir={dbConfigured()}
        kanallar={await kanalDurumu()}
        taslak={taslakHazir()}
        ilkKonusma={searchParams?.k || ""}
      />
    </main>
  );
}
