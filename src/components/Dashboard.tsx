import { useState, useEffect, useCallback } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { loginRequest, groups } from "@/lib/msalConfig";
import { fetchGroupMembers, fetchMe, getTeamsCallLink, type GroupMember } from "@/lib/graphApi";
import { Phone, LogOut, Users, Loader2, ShieldCheck, UserRound } from "lucide-react";

const STATIC_CONTACTS = [
  { name: "Agenti/Corrieri", email: "monica.pillon@astidental.com" },
  { name: "Acquisti", email: "luca.nicodemo@astidental.com" },
  { name: "Customer Care", email: "althea.borio@astidental.com" },
  { name: "CED", email: "florentin.rachita@astidental.com" },
  { name: "Assistenza Tecnica", email: "bianca.lucaci@astidental.com" },
  { name: "Commerciale e Marketing", email: "sergio.campini@astidental.com" },
];

interface GroupData {
  name: string;
  members: GroupMember[];
}

export default function Dashboard() {
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [groupsData, setGroupsData] = useState<GroupData[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [userName, setUserName] = useState("");

  const getAccessToken = useCallback(async () => {
    const account = accounts[0];
    if (!account) throw new Error("No account");
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return response.accessToken;
  }, [instance, accounts]);

  useEffect(() => {
    if (!isAuthenticated || accounts.length === 0) return;

    const loadData = async () => {
      setLoading(true);
      setError(null);
      try {
        const token = await getAccessToken();
        const [me, ...memberResults] = await Promise.all([
          fetchMe(token),
          ...groups.map((g) => fetchGroupMembers(token, g.endpoint)),
        ]);
        setUserName(me.displayName);
        setGroupsData(
          groups.map((g, i) => ({ name: g.name, members: memberResults[i] }))
        );
      } catch (err: any) {
        setError(err.message || "Errore nel caricamento dei dati");
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, [isAuthenticated, accounts, getAccessToken]);

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch(console.error);
  };

  const handleLogout = () => {
    instance.logoutRedirect().catch(console.error);
  };

  if (inProgress === InteractionStatus.Startup || inProgress === InteractionStatus.HandleRedirect) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-background">
        <Loader2 className="h-8 w-8 animate-spin text-primary" />
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-background">
        <div className="w-full max-w-sm rounded-lg border border-border bg-card p-8 text-center shadow-sm">
          <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-full bg-primary/10">
            <ShieldCheck className="h-7 w-7 text-primary" />
          </div>
          <h1 className="mb-1 text-xl font-semibold text-card-foreground">Benvenuto in Astidental</h1>
          <p className="mb-6 text-sm text-muted-foreground">Accedi con il tuo account aziendale Microsoft</p>
          <button
            onClick={handleLogin}
            className="w-full rounded-md bg-primary px-4 py-2.5 text-sm font-semibold text-primary-foreground transition-colors hover:bg-primary/90"
          >
            Accedi con Microsoft
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background">
      <header className="sticky top-0 z-10 border-b border-border bg-card shadow-sm">
        <div className="mx-auto flex max-w-4xl items-center justify-between px-4 py-3">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-md bg-primary">
              <Users className="h-5 w-5 text-primary-foreground" />
            </div>
            <div>
              <h1 className="text-base font-semibold text-card-foreground leading-tight">Call Dashboard</h1>
              <p className="text-xs text-muted-foreground">{userName}</p>
            </div>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-4xl px-4 py-6">
        <div className="mb-8 text-center">
          <h2 className="text-xl font-semibold text-card-foreground">Benvenuti in Astidental</h2>
          <p className="mt-1 text-sm text-muted-foreground">Selezionate il reparto desiderato per parlare con un operatore.</p>
        </div>

        {loading && (
          <div className="flex items-center justify-center py-20">
            <Loader2 className="h-6 w-6 animate-spin text-primary" />
            <span className="ml-2 text-sm text-muted-foreground">Caricamento gruppi…</span>
          </div>
        )}

        {error && (
          <div className="rounded-md bg-destructive/10 p-4 text-sm text-destructive">
            {error}
          </div>
        )}

        <div className="grid gap-6 sm:grid-cols-2">
          {/* Primo: Agenti/Corrieri */}
          {(() => {
            const contact = STATIC_CONTACTS.find((c) => c.name === "Agenti/Corrieri");
            if (!contact) return null;
            return (
              <div
                key={contact.name}
                className="rounded-lg border border-border bg-card p-8 text-center shadow-sm"
              >
                <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-full bg-primary/10">
                  <UserRound className="h-8 w-8 text-primary" />
                </div>
                <h2 className="mb-6 text-lg font-semibold text-card-foreground">{contact.name}</h2>
                <button
                  onClick={() => { window.open(getTeamsCallLink(contact.email), "_blank"); }}
                  className="inline-flex w-full items-center justify-center gap-2 rounded-md bg-primary px-4 py-3 text-sm font-semibold text-primary-foreground transition-colors hover:bg-primary/90"
                >
                  <Phone className="h-4 w-4" />
                  Chiama {contact.name}
                </button>
              </div>
            );
          })()}

          {/* Secondo: gruppi dinamici (Amministrazione primo) */}
          {!loading && !error && groupsData.map((group) => (
            <div
              key={group.name}
              className="rounded-lg border border-border bg-card p-8 text-center shadow-sm"
            >
              <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-full bg-primary/10">
                <Users className="h-8 w-8 text-primary" />
              </div>
              <h2 className="mb-1 text-lg font-semibold text-card-foreground">{group.name}</h2>
              <p className="mb-6 text-sm text-muted-foreground">{group.members.length} membri</p>
              <a
                href={getTeamsCallLink(group.members.map((m) => m.mail || m.userPrincipalName).join(","))}
                target="_blank"
                className="inline-flex w-full items-center justify-center gap-2 rounded-md bg-primary px-4 py-3 text-sm font-semibold text-primary-foreground transition-colors hover:bg-primary/90"
              >
                <Phone className="h-4 w-4" />
                Chiama {group.name}
              </a>
            </div>
          ))}

          {/* Poi gli altri contatti statici */}
          {STATIC_CONTACTS.filter((c) => c.name !== "Agenti/Corrieri").map((contact) => (
            <div
              key={contact.name}
              className="rounded-lg border border-border bg-card p-8 text-center shadow-sm"
            >
              <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-full bg-primary/10">
                <UserRound className="h-8 w-8 text-primary" />
              </div>
              <h2 className="mb-6 text-lg font-semibold text-card-foreground">{contact.name}</h2>
              <button
                onClick={() => { window.open(getTeamsCallLink(contact.email), "_blank"); }}
                className="inline-flex w-full items-center justify-center gap-2 rounded-md bg-primary px-4 py-3 text-sm font-semibold text-primary-foreground transition-colors hover:bg-primary/90"
              >
                <Phone className="h-4 w-4" />
                Chiama {contact.name}
              </button>
            </div>
          ))}
        </div>
      </main>
    </div>
  );
}
