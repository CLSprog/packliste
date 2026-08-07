import { useEffect, useMemo, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import seedData from "./data/schema-data.json";
import { logout } from "./auth";
import { loadState, saveState } from "./onedrive";
import type { SchemaData, Tk04GegenstandPerson } from "./data/schema";
import {
  gegenstaendeFuerReise,
  tk03FuerGegenstand,
  personenFuerTk03,
  kathegorieName,
  personName,
} from "./data/schema";

const SCHEMA_FILE = "P03_Packliste_Schema_AI.json";

export default function SchemaApp({ account }: { account: AccountInfo }) {
  const [data, setData] = useState<SchemaData | null>(null);
  const [loadStatus, setLoadStatus] = useState<"loading" | "ready" | "error">("loading");
  const [saveStatus, setSaveStatus] = useState<string>("");
  const [selectedReiseId, setSelectedReiseId] = useState<string | null>(null);
  const [personFilter, setPersonFilter] = useState<string>("alle");

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const remote = await loadState(SCHEMA_FILE);
        if (cancelled) return;
        const initial = (remote as SchemaData) ?? (seedData as unknown as SchemaData);
        setData(initial);
        setSelectedReiseId(initial.t01_reise[0]?.id ?? null);
        setLoadStatus("ready");
      } catch (error) {
        if (cancelled) return;
        console.error(error);
        setLoadStatus("error");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!data) return;
    setSaveStatus("Änderungen werden gespeichert …");
    const t = setTimeout(async () => {
      try {
        await saveState(data, SCHEMA_FILE);
        const now = new Date().toLocaleTimeString("de-AT", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
        setSaveStatus(`Gespeichert um ${now}`);
      } catch (error) {
        console.error(error);
        setSaveStatus("Speichern fehlgeschlagen – bitte Verbindung prüfen.");
      }
    }, 600);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data]);

  const reise = useMemo(
    () => data?.t01_reise.find((r) => r.id === selectedReiseId) ?? null,
    [data, selectedReiseId]
  );

  const gegenstaende = useMemo(() => {
    if (!data || !reise) return [];
    return gegenstaendeFuerReise(data, reise.id);
  }, [data, reise]);

  const beteiligtePersonen = useMemo(() => {
    if (!data || !reise) return [];
    const ids = new Set<string>();
    for (const g of gegenstaende) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      for (const p of personenFuerTk03(data, tk03.id)) ids.add(p.id_t05);
    }
    return data.t05_namen.filter((n) => ids.has(n.id));
  }, [data, reise, gegenstaende]);

  const gruppiert = useMemo(() => {
    if (!data) return [];
    const byKat = new Map<string, typeof gegenstaende>();
    for (const g of gegenstaende) {
      if (personFilter !== "alle") {
        const tk03 = tk03FuerGegenstand(data, reise!.id, g.id);
        const personen = tk03 ? personenFuerTk03(data, tk03.id) : [];
        if (!personen.some((p) => p.id_t05 === personFilter)) continue;
      }
      const kat = kathegorieName(data, g.id_kathegorie);
      if (!byKat.has(kat)) byKat.set(kat, []);
      byKat.get(kat)!.push(g);
    }
    return Array.from(byKat.entries()).sort((a, b) => a[0].localeCompare(b[0], "de"));
  }, [data, gegenstaende, personFilter, reise]);

  const fortschritt = useMemo(() => {
    if (!data || !reise) return { done: 0, total: 0 };
    let done = 0;
    let total = 0;
    for (const g of gegenstaende) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      for (const p of personenFuerTk03(data, tk03.id)) {
        total++;
        if (p.eingepackt !== null && p.eingepackt !== undefined && p.eingepackt > 0) done++;
      }
    }
    return { done, total };
  }, [data, reise, gegenstaende]);

  function advanceStage(row: Tk04GegenstandPerson) {
    // Ein Tippen wandert durch: nichts -> hergerichtet -> eingepackt -> verwendet -> nichts
    setData((prev) => {
      if (!prev) return prev;
      const updated = prev.tk04_tk03_t05.map((r) => {
        if (r.id !== row.id) return r;
        const menge = r.ausgewaehlt ?? 1;
        const stageOf = (v: number | null | undefined) => (v !== null && v !== undefined && v > 0);
        if (!stageOf(r.hergerichtet)) {
          return { ...r, hergerichtet: menge };
        }
        if (!stageOf(r.eingepackt)) {
          return { ...r, eingepackt: menge };
        }
        if (!stageOf(r.verwendet)) {
          return { ...r, verwendet: menge };
        }
        // Kreislauf: von vorne, alles zurücksetzen
        return { ...r, hergerichtet: null, eingepackt: null, verwendet: null };
      });
      return { ...prev, tk04_tk03_t05: updated };
    });
  }

  function stageLabelFor(row: Tk04GegenstandPerson): { cls: string; text: string } {
    const has = (v: number | null | undefined) => v !== null && v !== undefined && v > 0;
    if (has(row.verwendet)) return { cls: "stage-verwendet", text: "V" };
    if (has(row.eingepackt)) return { cls: "stage-eingepackt", text: "E" };
    if (has(row.hergerichtet)) return { cls: "stage-hergerichtet", text: "H" };
    return { cls: "", text: "–" };
  }

  if (loadStatus === "loading") {
    return <div className="pl-loading">Packliste wird geladen …</div>;
  }
  if (loadStatus === "error" || !data) {
    return <div className="pl-loading">Fehler beim Laden der Packliste. Bitte Seite neu laden.</div>;
  }

  const pct = fortschritt.total > 0 ? Math.round((fortschritt.done / fortschritt.total) * 100) : 0;

  return (
    <div className="pl-shell">
      <header className="pl-header">
        <div className="pl-header-top">
          <h1>{reise?.reise ?? "P03 Packliste"}</h1>
          <button className="pl-logout" onClick={() => logout()}>Abmelden</button>
        </div>
        <p className="pl-header-sub">
          {reise?.von && reise?.bis
            ? `${new Date(reise.von).toLocaleDateString("de-AT")} – ${new Date(reise.bis).toLocaleDateString("de-AT")}`
            : "Zeitraum noch offen"}
          {" · "}
          {account.name ?? account.username}
        </p>
        <div className="pl-progress-wrap">
          <div className="pl-progress-labels">
            <span>Eingepackt</span>
            <span>{fortschritt.done} / {fortschritt.total}</span>
          </div>
          <div className="pl-progress-track">
            <div className="pl-progress-fill" style={{ width: `${pct}%` }} />
          </div>
        </div>
        <p className="pl-save">{saveStatus}</p>
      </header>

      {data.t01_reise.length > 1 && (
        <div className="pl-select-row">
          <select value={selectedReiseId ?? ""} onChange={(e) => setSelectedReiseId(e.target.value)}>
            {data.t01_reise.map((r) => (
              <option key={r.id} value={r.id}>{r.reise}</option>
            ))}
          </select>
        </div>
      )}

      <div className="pl-filterbar">
        <button
          className={"pl-filter-chip" + (personFilter === "alle" ? " active" : "")}
          onClick={() => setPersonFilter("alle")}
        >
          Alle
        </button>
        {beteiligtePersonen.map((p) => (
          <button
            key={p.id}
            className={"pl-filter-chip" + (personFilter === p.id ? " active" : "")}
            onClick={() => setPersonFilter(p.id)}
          >
            {p.namen}
          </button>
        ))}
      </div>

      <div className="pl-body">
        {gruppiert.map(([katName, items]) => (
          <div className="pl-category" key={katName}>
            <div className="pl-category-head">
              <span className="pl-category-tag">{katName}</span>
              <span className="pl-category-count">{items.length}</span>
            </div>
            {items.map((g) => {
              const tk03 = tk03FuerGegenstand(data, reise!.id, g.id);
              const personen = tk03 ? personenFuerTk03(data, tk03.id) : [];
              const gefiltertePersonen =
                personFilter === "alle" ? personen : personen.filter((p) => p.id_t05 === personFilter);
              return (
                <div className="pl-item" key={g.id}>
                  <span className="pl-item-name">{g.gegenstand}</span>
                  {gefiltertePersonen.length === 0 ? (
                    <span className="pl-item-empty">–</span>
                  ) : (
                    <div className="pl-badges">
                      {gefiltertePersonen.map((p) => {
                        const { cls, text } = stageLabelFor(p);
                        const initiale = personName(data, p.id_t05).slice(0, 1).toUpperCase();
                        return (
                          <button
                            key={p.id}
                            className={"pl-badge" + (cls ? " " + cls : "")}
                            title={`${personName(data, p.id_t05)} · ${p.ausgewaehlt ?? 0}×`}
                            onClick={() => advanceStage(p)}
                          >
                            <span>{initiale}{text !== "–" ? " " + text : ""}</span>
                            <span className="qty">{p.ausgewaehlt ?? 0}×</span>
                          </button>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </div>
  );
}
