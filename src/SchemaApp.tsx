import { useEffect, useMemo, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import seedData from "./data/schema-data.json";
import { logout } from "./auth";
import { loadState, saveState } from "./onedrive";
import type { SchemaData, Tk04GegenstandPerson, T01Reise, T04Gegenstand } from "./data/schema";
import {
  gegenstaendeFuerReise,
  tk03FuerGegenstand,
  personenFuerTk03,
  kathegorieName,
  newId,
} from "./data/schema";

const SCHEMA_FILE = "P03_Packliste_Schema_AI.json";

export default function SchemaApp({ account }: { account: AccountInfo }) {
  const [data, setData] = useState<SchemaData | null>(null);
  const [loadStatus, setLoadStatus] = useState<"loading" | "ready" | "error">("loading");
  const [saveStatus, setSaveStatus] = useState<string>("");
  const [selectedReiseId, setSelectedReiseId] = useState<string | null>(null);
  const [personFilter, setPersonFilter] = useState<string | null>(null);
  const [mode, setMode] = useState<"liste" | "bearbeiten" | "neueReise">("liste");
  const [editSearch, setEditSearch] = useState("");
  const [neuReiseName, setNeuReiseName] = useState("");
  const [neuReiseVon, setNeuReiseVon] = useState("");
  const [neuReiseBis, setNeuReiseBis] = useState("");

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

  useEffect(() => {
    if (personFilter === null && beteiligtePersonen.length > 0) {
      setPersonFilter(beteiligtePersonen[0].id);
    }
  }, [personFilter, beteiligtePersonen]);

  const gruppiert = useMemo(() => {
    if (!data || !reise || !personFilter) return [];
    const byKat = new Map<string, { g: (typeof gegenstaende)[number]; row: Tk04GegenstandPerson }[]>();
    for (const g of gegenstaende) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      const personen = personenFuerTk03(data, tk03.id);
      const row = personen.find((p) => p.id_t05 === personFilter);
      if (!row) continue; // diese Person braucht den Gegenstand nicht
      const kat = kathegorieName(data, g.id_kathegorie);
      if (!byKat.has(kat)) byKat.set(kat, []);
      byKat.get(kat)!.push({ g, row });
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

  function bumpField(
    row: Tk04GegenstandPerson,
    field: "ausgewaehlt" | "hergerichtet" | "eingepackt" | "verwendet",
    direction: 1 | -1
  ) {
    setData((prev) => {
      if (!prev) return prev;
      // Vorläufig: bei "ausgewählt" auf 0 NICHT automatisch aus der Liste entfernen
      // (Unfallgefahr bei versehentlichem Wischen) - Sicherheitsnetz/Rückgängig folgt noch.
      const updated = prev.tk04_tk03_t05.map((r) => {
        if (r.id !== row.id) return r;
        if (field === "ausgewaehlt") {
          const next = Math.max(0, (r.ausgewaehlt ?? 0) + direction);
          return { ...r, ausgewaehlt: next };
        }
        const current = r[field];
        const has = current !== null && current !== undefined;
        if (direction === 1) {
          if (!has) return { ...r, [field]: r.ausgewaehlt ?? 1 };
          return { ...r, [field]: current + 1 };
        }
        // runterzählen
        if (!has) return r; // schon leer, nichts zu tun
        if (current <= 1) return { ...r, [field]: null }; // zurück auf "nicht angetippt"
        return { ...r, [field]: current - 1 };
      });
      return { ...prev, tk04_tk03_t05: updated };
    });
  }

  // Wisch-Erkennung: nach oben = hochzählen, nach unten = runterzählen.
  // Reine Tippen (ohne nennenswerte Bewegung) zählt ebenfalls als "hoch" (schnellster Standardfall).
  const swipeStart = { current: null as { y: number } | null };
  function handlePointerDown(e: React.PointerEvent<HTMLButtonElement>) {
    swipeStart.current = { y: e.clientY };
  }
  function handlePointerUp(
    e: React.PointerEvent<HTMLButtonElement>,
    row: Tk04GegenstandPerson,
    field: "ausgewaehlt" | "hergerichtet" | "eingepackt" | "verwendet"
  ) {
    const start = swipeStart.current;
    swipeStart.current = null;
    const deltaY = start ? e.clientY - start.y : 0;
    const THRESHOLD = 12; // Pixel, ab wann es als Wisch gilt
    if (deltaY < -THRESHOLD) {
      bumpField(row, field, 1); // nach oben gewischt
    } else if (deltaY > THRESHOLD) {
      bumpField(row, field, -1); // nach unten gewischt
    } else {
      bumpField(row, field, 1); // einfaches Antippen = hochzählen
    }
  }

  function collectAllIds(d: SchemaData): Set<string> {
    const ids = new Set<string>();
    const add = (arr: { id: string }[]) => arr.forEach((r) => ids.add(r.id));
    add(d.t01_reise);
    add(d.t02_aktivitaet);
    add(d.t03_kathegorie);
    add(d.t04_gegenstand);
    add(d.t05_namen);
    add(d.tk01_t01_t02);
    add(d.tk02_t02_t04);
    add(d.tk03_tk01_t04);
    add(d.tk04_tk03_t05);
    return ids;
  }

  function erstelleNeueReise() {
    if (!data || !neuReiseName.trim()) return;
    setData((prev) => {
      if (!prev) return prev;
      const ids = collectAllIds(prev);
      const reiseId = newId("t01", ids);
      ids.add(reiseId);
      const neu: T01Reise = {
        id: reiseId,
        reise: neuReiseName.trim(),
        von: neuReiseVon || null,
        bis: neuReiseBis || null,
        notiz: "",
      };
      const basics = prev.t02_aktivitaet.find((a) => a.aktivitaet === "Basics");
      let tk01Neu = prev.tk01_t01_t02;
      if (basics) {
        const tk01Id = newId("tk01", ids);
        ids.add(tk01Id);
        tk01Neu = [...prev.tk01_t01_t02, { id: tk01Id, id_t01: reiseId, id_t02: basics.id }];
      }
      return { ...prev, t01_reise: [...prev.t01_reise, neu], tk01_t01_t02: tk01Neu };
    });
    setNeuReiseName("");
    setNeuReiseVon("");
    setNeuReiseBis("");
    setMode("bearbeiten");
  }

  function ankerTk01(d: SchemaData, reiseId: string): string | null {
    const basics = d.t02_aktivitaet.find((a) => a.aktivitaet === "Basics");
    const tk01Rows = d.tk01_t01_t02.filter((r) => r.id_t01 === reiseId);
    if (basics) {
      const found = tk01Rows.find((r) => r.id_t02 === basics.id);
      if (found) return found.id;
    }
    return tk01Rows[0]?.id ?? null;
  }

  function toggleGegenstandInReise(gegenstandId: string) {
    if (!data || !reise) return;
    setData((prev) => {
      if (!prev) return prev;
      const bestehende = tk03FuerGegenstand(prev, reise.id, gegenstandId);
      if (bestehende) {
        return {
          ...prev,
          tk03_tk01_t04: prev.tk03_tk01_t04.filter((r) => r.id !== bestehende.id),
          tk04_tk03_t05: prev.tk04_tk03_t05.filter((r) => r.id_tk03 !== bestehende.id),
        };
      }
      const anker = ankerTk01(prev, reise.id);
      if (!anker) return prev;
      const ids = collectAllIds(prev);
      const neuId = newId("tk03", ids);
      return {
        ...prev,
        tk03_tk01_t04: [
          ...prev.tk03_tk01_t04,
          { id: neuId, id_tk01: anker, id_t04: gegenstandId, notiz: "manuell ausgewählt" },
        ],
      };
    });
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

      <div className="pl-editbar">
        <button
          className={"pl-edit-toggle" + (mode === "neueReise" ? " active" : "")}
          onClick={() => setMode("neueReise")}
        >
          + Neue Reise
        </button>{" "}
        <button
          className={"pl-edit-toggle" + (mode === "bearbeiten" ? " active" : "")}
          onClick={() => setMode(mode === "bearbeiten" ? "liste" : "bearbeiten")}
        >
          {mode === "bearbeiten" ? "Fertig" : "Liste bearbeiten"}
        </button>
      </div>

      {mode === "neueReise" && (
        <div className="pl-newreise">
          <label>Name der Reise</label>
          <input value={neuReiseName} onChange={(e) => setNeuReiseName(e.target.value)} placeholder="z.B. Grimming 2026" />
          <label>Von</label>
          <input type="date" value={neuReiseVon} onChange={(e) => setNeuReiseVon(e.target.value)} />
          <label>Bis</label>
          <input type="date" value={neuReiseBis} onChange={(e) => setNeuReiseBis(e.target.value)} />
          <button onClick={erstelleNeueReise} disabled={!neuReiseName.trim()}>
            Reise anlegen &amp; Gegenstände auswählen
          </button>
        </div>
      )}

      {mode === "bearbeiten" && reise && (
        <>
          <div className="pl-select-row">
            <input
              value={editSearch}
              onChange={(e) => setEditSearch(e.target.value)}
              placeholder="Gegenstand suchen…"
              style={{ width: "100%", padding: 9, borderRadius: 9, border: "1px solid var(--line)", fontSize: 14 }}
            />
          </div>
          <div className="pl-body">
            {(() => {
              const byKat = new Map<string, T04Gegenstand[]>();
              for (const g of data.t04_gegenstand) {
                if (editSearch && !g.gegenstand.toLowerCase().includes(editSearch.toLowerCase())) continue;
                const kat = kathegorieName(data, g.id_kathegorie);
                if (!byKat.has(kat)) byKat.set(kat, []);
                byKat.get(kat)!.push(g);
              }
              const sorted = Array.from(byKat.entries()).sort((a, b) => a[0].localeCompare(b[0], "de"));
              return sorted.map(([katName, items]) => (
                <div className="pl-category" key={katName}>
                  <div className="pl-category-head">
                    <span className="pl-category-tag">{katName}</span>
                    <span className="pl-category-count">{items.length}</span>
                  </div>
                  {items.map((g) => {
                    const drin = !!tk03FuerGegenstand(data, reise.id, g.id);
                    return (
                      <div className="pl-edit-item" key={g.id}>
                        <span className="pl-edit-item-name">{g.gegenstand}</span>
                        <button
                          className={"pl-edit-check" + (drin ? " on" : "")}
                          onClick={() => toggleGegenstandInReise(g.id)}
                        >
                          {drin ? "1" : "0"}
                        </button>
                      </div>
                    );
                  })}
                </div>
              ));
            })()}
          </div>
        </>
      )}

      {mode === "liste" && (
        <>
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

      <div className="pl-legend">
        <span style={{ width: 16 }}></span>
        <span title="Ausgewählt (geplant)">A</span>
        <span title="Hergerichtet">H</span>
        <span title="Eingepackt">E</span>
        <span title="Verwendet">V</span>
      </div>

      <div className="pl-body">
        {gruppiert.map(([katName, entries]) => (
          <div className="pl-category" key={katName}>
            <div className="pl-category-head">
              <span className="pl-category-tag">{katName}</span>
              <span className="pl-category-count">{entries.length}</span>
            </div>
            {entries.map(({ g, row }) => (
              <div className="pl-item pl-item-single" key={g.id}>
                <span className="pl-item-name">{g.gegenstand}</span>
                <div className="pl-qrow">
                  <button
                    className="pl-qbox q-ausgewaehlt"
                    onPointerDown={handlePointerDown}
                    onPointerUp={(e) => handlePointerUp(e, row, "ausgewaehlt")}
                  >
                    {row.ausgewaehlt ?? 0}
                  </button>
                  {(["hergerichtet", "eingepackt", "verwendet"] as const).map((field) => {
                    const val = row[field];
                    const has = val !== null && val !== undefined;
                    return (
                      <button
                        key={field}
                        className={"pl-qbox q-" + field + (has ? " q-set" : "")}
                        onPointerDown={handlePointerDown}
                        onPointerUp={(e) => handlePointerUp(e, row, field)}
                      >
                        {has ? val : "–"}
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        ))}
      </div>
        </>
      )}
    </div>
  );
}
