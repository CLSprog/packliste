import { useEffect, useMemo, useRef, useState } from "react";
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
import { diffAndMerge, applyConflictResolutions, hashData, schemaEqual, type RowConflict, type AutoMergedChange, type SyncResult } from "./sync";
import { readBaseline, writeBaseline, readPending, writePending, clearPending, istVerbindungsfehler } from "./syncStore";
import ConflictModal from "./ConflictModal";

// Von Clemens gewünscht (2026-08-27): sichtbare Versionsnummer im Kopfbereich,
// damit jederzeit erkennbar ist, ob GitHub Pages wirklich den aktuellsten Stand
// ausliefert. Bei jeder Auslieferung hier mitziehen.
const APP_VERSION = "V02-02";

// Einziger Speicherort/Dateiname (von Clemens am 2026-08-28 bestätigt: kein
// Fallback mehr auf alte Ordner/Dateinamen, die werden von ihm manuell aus
// OneDrive gelöscht). Siehe onedrive.ts für den Ordnerpfad.
const SCHEMA_FILE = "P03_Packliste_AI.json";

// Ergänzt in "remote" (den echten, live gespeicherten Daten eines Nutzers)
// alle Zeilen aus "seed" (der im ZIP mitgelieferten, ggf. neuer importierten
// Katalog-/Reise-Daten), die anhand ihrer ID dort noch fehlen. Bereits
// vorhandene Zeilen in remote werden nie verändert oder überschrieben –
// so gehen keine bereits eingetragenen Packzustände verloren, aber neu
// importierte Reisen/Gegenstände (die sonst nur in der ZIP-Seed-Datei
// stecken würden, weil OneDrive schon vorher Daten hatte) kommen trotzdem an.
function mergeSeedInto(remote: SchemaData, seed: SchemaData): SchemaData {
  function ergaenzt<T extends { id: string }>(remoteZeilen: T[], seedZeilen: T[]): T[] {
    const vorhandeneIds = new Set(remoteZeilen.map((zeile) => zeile.id));
    const fehlende = seedZeilen.filter((zeile) => !vorhandeneIds.has(zeile.id));
    return fehlende.length > 0 ? [...remoteZeilen, ...fehlende] : remoteZeilen;
  }
  return {
    t01_reise: ergaenzt(remote.t01_reise, seed.t01_reise),
    t02_aktivitaet: ergaenzt(remote.t02_aktivitaet, seed.t02_aktivitaet),
    t03_kathegorie: ergaenzt(remote.t03_kathegorie, seed.t03_kathegorie),
    t04_gegenstand: ergaenzt(remote.t04_gegenstand, seed.t04_gegenstand),
    t05_namen: ergaenzt(remote.t05_namen, seed.t05_namen),
    tk01_t01_t02: ergaenzt(remote.tk01_t01_t02, seed.tk01_t01_t02),
    tk02_t02_t04: ergaenzt(remote.tk02_t02_t04, seed.tk02_t02_t04),
    tk03_tk01_t04: ergaenzt(remote.tk03_tk01_t04, seed.tk03_tk01_t04),
    tk04_tk03_t05: ergaenzt(remote.tk04_tk03_t05, seed.tk04_tk03_t05),
    // "Neu hinzugefügt"-Markierungen kommen nur aus den echten Live-Daten, nie aus der
    // Seed-Datei (die kennt dieses Feld nicht) - sonst blieben alte ZIP-Importe für immer
    // fälschlich als "neu" markiert.
    neu_hinzugefuegt: remote.neu_hinzugefuegt ?? [],
  };
}

// Einmaliger, gezielter Nachtrag für Grimming 2026 (2026-08-28): Die 114 Gegenstände
// waren korrekt mit der Reise verknüpft (tk03), aber für niemanden gab es je eine
// persönliche Mengen-Zeile (tk04) - anders als bei China 2024 (siehe V01-17). Dadurch
// gab es keinen anklickbaren Personen-Tab und die Reise wirkte fälschlich leer (Bug,
// siehe Fix bei der Personen-Tab-Leiste weiter unten). Legt für Clemens und Sonja
// (von Clemens am 2026-08-28 bestätigt) je eine tk04-Zeile mit demselben Standardwert
// an, den auch der "+"-Knopf in "Liste bearbeiten" verwendet ("0 – unsicher"). Läuft
// bei jedem Laden, legt aber wegen der Vorhanden-Prüfung nie doppelte Zeilen an -
// sobald echte Mengen eingetragen sind, passiert hier nichts mehr.
function backfillPersonenOhneMenge(data: SchemaData, reiseName: string, personenNamen: string[]): SchemaData {
  const reise = data.t01_reise.find((r) => r.reise === reiseName);
  if (!reise) return data;
  const personen = data.t05_namen.filter((n) => personenNamen.includes(n.namen));
  if (personen.length === 0) return data;
  const tk01Ids = new Set(data.tk01_t01_t02.filter((r) => r.id_t01 === reise.id).map((r) => r.id));
  const tk03Rows = data.tk03_tk01_t04.filter((r) => tk01Ids.has(r.id_tk01));
  if (tk03Rows.length === 0) return data;

  const existingIds = new Set<string>();
  const addIds = (arr: { id: string }[]) => arr.forEach((r) => existingIds.add(r.id));
  addIds(data.t01_reise);
  addIds(data.t02_aktivitaet);
  addIds(data.t03_kathegorie);
  addIds(data.t04_gegenstand);
  addIds(data.t05_namen);
  addIds(data.tk01_t01_t02);
  addIds(data.tk02_t02_t04);
  addIds(data.tk03_tk01_t04);
  addIds(data.tk04_tk03_t05);

  const neueZeilen: Tk04GegenstandPerson[] = [];
  for (const tk03 of tk03Rows) {
    for (const person of personen) {
      const vorhanden = data.tk04_tk03_t05.some((r) => r.id_tk03 === tk03.id && r.id_t05 === person.id);
      if (vorhanden) continue;
      const neuId = newId("tk04", existingIds);
      existingIds.add(neuId);
      neueZeilen.push({
        id: neuId,
        id_tk03: tk03.id,
        id_t05: person.id,
        ausgewaehlt: 0,
        hergerichtet: null,
        eingepackt: null,
        verwendet: null,
      });
    }
  }
  if (neueZeilen.length === 0) return data;
  return { ...data, tk04_tk03_t05: [...data.tk04_tk03_t05, ...neueZeilen] };
}

function safeFilename(value: string) {
  return value.trim().replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-").replace(/\s+/g, "_").replace(/_+/g, "_") || "Packliste";
}

function dateStamp() {
  const now = new Date();
  const part = (value: number) => String(value).padStart(2, "0");
  return `${now.getFullYear()}${part(now.getMonth() + 1)}${part(now.getDate())}-${part(now.getHours())}${part(now.getMinutes())}`;
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

type ExportZeile = {
  kategorie: string;
  gegenstand: string;
  person: string;
  ausgewaehlt: number;
  hergerichtet: number | null;
  eingepackt: number | null;
  verwendet: number | null;
};

const HISTORY_FILE = "P03_Packliste_Verlauf_AI.json";
const MAX_HISTORY = 20;

// Von Clemens gemeldet (2026-08-28): nach jedem Aktualisieren/Neuladen sprang die App
// immer auf die erste Reise (Schottland) zurück, statt bei der zuletzt angesehenen zu
// bleiben (z.B. Grimming). Rein pro Gerät im Browser gemerkt, nicht Teil der in OneDrive
// gespeicherten Reise-Daten - jedes Gerät merkt sich seine eigene zuletzt offene Reise.
const LETZTE_REISE_KEY = "p03_letzte_reise_id";

function ermittleStartReise(data: SchemaData): string | null {
  try {
    const gespeichert = localStorage.getItem(LETZTE_REISE_KEY);
    if (gespeichert && data.t01_reise.some((r) => r.id === gespeichert)) return gespeichert;
  } catch {
    // z.B. privates Fenster ohne Speicherzugriff - dann einfach die erste Reise
  }
  return data.t01_reise[0]?.id ?? null;
}

export default function SchemaApp({ account }: { account: AccountInfo }) {
  const [data, setData] = useState<SchemaData | null>(null);
  // Rückgängig-Verlauf: die letzten Datenstände VOR jeder Änderung, älteste zuerst.
  // Wird zusätzlich in OneDrive gesichert, damit "Rückgängig" auch nach einem
  // Neuladen der Seite noch funktioniert (z.B. wenn man die App zwischendurch schließt).
  const [history, setHistory] = useState<SchemaData[]>([]);
  const [loadStatus, setLoadStatus] = useState<"loading" | "ready" | "error">("loading");
  const [saveStatus, setSaveStatus] = useState<string>("");
  const [exportStatus, setExportStatus] = useState<string>("");
  const [selectedReiseId, setSelectedReiseId] = useState<string | null>(null);
  const [personFilter, setPersonFilter] = useState<string | null>(null);
  const [offenFilter, setOffenFilter] = useState<"hergerichtet" | "eingepackt" | "neu" | null>(null);
  // Welche tk04-Zeilen über "Liste bearbeiten" hinzugefügt/reaktiviert wurden und noch
  // nicht angetippt sind, für den Filter "Neu hinzugefügt". Seit V01-29 Teil von `data`
  // (Feld `neu_hinzugefuegt`) und damit in OneDrive gesichert - bleibt also über ein
  // Neuladen und über Geräte hinweg erhalten, bis der jeweilige Eintrag angetippt wird
  // (von Clemens am 2026-08-28 so gewünscht: "nicht nur lokal gespeichert").
  const neuHinzugefuegt = useMemo(
    () => new Set(data?.neu_hinzugefuegt ?? []),
    [data?.neu_hinzugefuegt]
  );
  const [mode, setMode] = useState<"liste" | "bearbeiten" | "neueReise">("liste");
  const [editSearch, setEditSearch] = useState("");
  const [neuReiseName, setNeuReiseName] = useState("");
  const [neuReiseVon, setNeuReiseVon] = useState("");
  const [neuReiseBis, setNeuReiseBis] = useState("");
  const [moveFor, setMoveFor] = useState<{ tk03Id: string; tk04Id: string; vonPerson: string } | null>(null);
  const [showNeuGegenstand, setShowNeuGegenstand] = useState(false);
  const [neuGegenstandName, setNeuGegenstandName] = useState("");
  const [neuGegenstandKat, setNeuGegenstandKat] = useState("");
  const [neuGegenstandKatNeu, setNeuGegenstandKatNeu] = useState("");
  const [editingQty, setEditingQty] = useState<{
    rowId: string;
    field: "ausgewaehlt" | "hergerichtet" | "eingepackt" | "verwendet";
  } | null>(null);
  const [editingQtyValue, setEditingQtyValue] = useState("");
  const longPressTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const longPressFired = useRef(false);

  // ---- Sync/Offline (Multiuser-Synchronisation, Phase 1+2) ----
  // "baseline" = der letzte Stand, der nachweislich mit dem Server abgeglichen wurde -
  // Grundlage für den Zeilen-für-Zeile-Vergleich in sync.ts. In einem Ref, weil er sich
  // unabhängig von React-Re-Renders sofort nach jedem erfolgreichen Sync ändern muss.
  const baselineRef = useRef<SchemaData | null>(null);
  // Der volle Vergleichs-Vorschlag (inkl. der noch offenen Konflikt-Zeilen), solange der
  // Nutzer über das Konflikt-Fenster noch nicht entschieden hat.
  const pendingSyncResultRef = useRef<SyncResult | null>(null);
  const [conflicts, setConflicts] = useState<RowConflict[]>([]);
  const [autoMerged, setAutoMerged] = useState<AutoMergedChange[]>([]);
  const [showAutoMerged, setShowAutoMerged] = useState(false);
  const [offline, setOffline] = useState(false);
  const [retryTick, setRetryTick] = useState(0);

  function alsServerstand(remote: unknown): SchemaData {
    const merged = remote
      ? mergeSeedInto(remote as SchemaData, seedData as unknown as SchemaData)
      : (seedData as unknown as SchemaData);
    return backfillPersonenOhneMenge(merged, "Grimming 2026", ["Clemens", "Sonja"]);
  }

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        // Einziger Speicherort, kein Fallback mehr auf alte Ordner/Dateinamen
        // (siehe onedrive.ts / Projektstand, ab V01-27 auf Clemens' Wunsch entfernt).
        const remote = await loadState(SCHEMA_FILE);
        const remoteHistory = await loadState(HISTORY_FILE);
        if (cancelled) return;
        const server = alsServerstand(remote);

        // Prüfen, ob von einer früheren Offline-Sitzung noch ein nicht synchronisierter
        // Stand im Browser liegt (z.B. App wurde ohne Verbindung geschlossen, bevor die
        // Änderung nach OneDrive geschrieben werden konnte).
        const pending = readPending(SCHEMA_FILE);
        const storedBaseline = readBaseline(SCHEMA_FILE);
        const pendingIstNeu = pending && (!storedBaseline || JSON.stringify(pending) !== JSON.stringify(storedBaseline));

        if (pendingIstNeu && pending) {
          const result = diffAndMerge(storedBaseline ?? server, pending, server);
          if (result.conflicts.length > 0) {
            pendingSyncResultRef.current = result;
            setConflicts(result.conflicts);
            setAutoMerged(result.autoMerged);
            setData(result.merged);
            setSelectedReiseId(ermittleStartReise(result.merged));
          } else {
            await saveState(result.merged, SCHEMA_FILE);
            baselineRef.current = result.merged;
            writeBaseline(SCHEMA_FILE, result.merged);
            clearPending(SCHEMA_FILE);
            setData(result.merged);
            setSelectedReiseId(ermittleStartReise(result.merged));
            if (result.autoMerged.length > 0) {
              setAutoMerged(result.autoMerged);
              setShowAutoMerged(true);
            }
          }
        } else {
          baselineRef.current = server;
          writeBaseline(SCHEMA_FILE, server);
          clearPending(SCHEMA_FILE);
          setData(server);
          setSelectedReiseId(ermittleStartReise(server));
        }
        setHistory((remoteHistory as SchemaData[]) ?? []);
        setOffline(false);
        setLoadStatus("ready");
      } catch (error) {
        if (cancelled) return;
        console.error(error);
        if (istVerbindungsfehler(error)) {
          // Ohne Verbindung: mit dem letzten lokal gemerkten Stand weiterarbeiten,
          // falls vorhanden (offline gemachte Änderungen zuerst, sonst der letzte
          // erfolgreich synchronisierte Stand).
          const lokal = readPending(SCHEMA_FILE) ?? readBaseline(SCHEMA_FILE);
          if (lokal) {
            baselineRef.current = readBaseline(SCHEMA_FILE) ?? lokal;
            setData(lokal);
            setSelectedReiseId(ermittleStartReise(lokal));
            setHistory([]);
            setOffline(true);
            setLoadStatus("ready");
            return;
          }
        }
        setLoadStatus("error");
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Verbindung kommt zurück (Browser-Ereignis) bzw. fällt aus - Sync-Versuch anstoßen bzw.
  // Status setzen. Zusätzlich alle 30s ein erneuter Versuch, solange offline, für den Fall,
  // dass das Browser-Ereignis nicht zuverlässig feuert (z.B. WLAN ohne echtes Internet).
  useEffect(() => {
    function handleOnline() {
      setOffline(false);
      setRetryTick((t) => t + 1);
    }
    function handleOffline() {
      setOffline(true);
    }
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);

  // Läuft immer (nicht nur offline): stößt regelmäßig einen Sync-Check an, damit ein
  // Gerät, an dem gerade NICHT selbst getippt wird, fremde Änderungen auch ohne eigenes
  // Antippen oder manuelles Neuladen mitbekommt (von Clemens gemeldet: Änderung am Laptop
  // erschien am Handy erst nach manuellem Aktualisieren). Offline richtet das nichts an -
  // der Versuch schlägt einfach mit einem Verbindungsfehler fehl (siehe unten).
  useEffect(() => {
    const interval = setInterval(() => setRetryTick((t) => t + 1), 20000);
    return () => clearInterval(interval);
  }, []);

  // Speichert die aktuelle Änderung nach OneDrive - aber erst, nachdem geprüft wurde, ob
  // sich der Server-Stand seit dem letzten eigenen Sync verändert hat (anderes Gerät hat
  // gespeichert). Ist das der Fall, wird Zeile für Zeile abgeglichen (sync.ts): eindeutige
  // fremde Änderungen werden automatisch übernommen und angezeigt, echte Konflikte (gleiche
  // Zeile beidseitig unterschiedlich geändert) werden dem Nutzer einzeln vorgelegt (Konflikt-
  // Fenster) - erst danach wird tatsächlich gespeichert. Dieselbe Logik greift unverändert,
  // wenn die Änderung offline gemacht wurde und erst jetzt synchronisiert werden kann.
  useEffect(() => {
    if (!data) return;
    if (conflicts.length > 0) return; // erst entscheiden lassen, bevor weitergespeichert wird
    // Gibt es überhaupt eine eigene, noch nicht gesicherte Änderung? Bei einem rein
    // periodischen Check (retryTick, alle 20s bzw. beim Wiederverbinden, siehe oben) ist
    // das meist nicht der Fall - dann wird nur "gepullt" (fremde Änderungen übernommen),
    // ohne unnötig nach OneDrive zu schreiben.
    const habenWirWasZuSpeichern = !baselineRef.current || !schemaEqual(data, baselineRef.current);
    if (habenWirWasZuSpeichern) {
      setSaveStatus(offline ? "Offline – Änderung wird lokal gemerkt …" : "Änderungen werden gespeichert …");
    }
    const t = setTimeout(async () => {
      // Immer zuerst lokal merken, damit bei einem Absturz/Schließen mitten im Offline-
      // Betrieb nichts verloren geht (siehe syncStore.ts) - unabhängig davon, ob der
      // anschließende Speicherversuch klappt.
      if (habenWirWasZuSpeichern) writePending(SCHEMA_FILE, data);
      try {
        const remote = await loadState(SCHEMA_FILE);
        const remoteData = remote ? alsServerstand(remote) : data;
        const baseline = baselineRef.current ?? remoteData;
        const remoteHash = await hashData(remoteData);
        const baselineHash = await hashData(baseline);
        const jetzt = () =>
          new Date().toLocaleTimeString("de-AT", { hour: "2-digit", minute: "2-digit", second: "2-digit" });

        if (remoteHash === baselineHash) {
          // Niemand sonst hat seit unserem letzten Sync gespeichert.
          setOffline(false);
          if (!habenWirWasZuSpeichern) return; // reiner Pull-Check, nichts zu tun
          await saveState(data, SCHEMA_FILE);
          baselineRef.current = data;
          writeBaseline(SCHEMA_FILE, data);
          clearPending(SCHEMA_FILE);
          setSaveStatus(`Gespeichert um ${jetzt()}`);
          return;
        }

        // Es gibt fremde Änderungen seit unserem letzten Sync - Zeilen-für-Zeile abgleichen
        // (auch relevant, wenn wir selbst gerade nichts zu speichern haben - dann sind es
        // reine "Pull"-Änderungen eines anderen Geräts, die wir trotzdem übernehmen wollen,
        // ohne selbst etwas zurückzuschreiben).
        const result = diffAndMerge(baseline, data, remoteData);
        if (result.conflicts.length > 0) {
          pendingSyncResultRef.current = result;
          setConflicts(result.conflicts);
          setAutoMerged(result.autoMerged);
          setData(result.merged);
          setSaveStatus("Es gibt widersprüchliche Änderungen von einem anderen Gerät – bitte entscheiden.");
          return;
        }
        if (habenWirWasZuSpeichern) await saveState(result.merged, SCHEMA_FILE);
        baselineRef.current = result.merged;
        writeBaseline(SCHEMA_FILE, result.merged);
        if (habenWirWasZuSpeichern) clearPending(SCHEMA_FILE);
        setOffline(false);
        setData(result.merged);
        if (result.autoMerged.length > 0) {
          setAutoMerged(result.autoMerged);
          setShowAutoMerged(true);
        }
        if (habenWirWasZuSpeichern) setSaveStatus(`Gespeichert um ${jetzt()}`);
      } catch (error) {
        console.error(error);
        if (istVerbindungsfehler(error)) {
          setOffline(true);
          if (habenWirWasZuSpeichern) {
            setSaveStatus("Offline – Änderungen werden gespeichert, sobald wieder online.");
          }
        } else if (habenWirWasZuSpeichern) {
          setSaveStatus("Speichern fehlgeschlagen – bitte Verbindung prüfen.");
        }
      }
    }, 600);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, conflicts.length, offline, retryTick]);

  // Rückgängig-Verlauf separat in OneDrive sichern, damit er auch nach einem
  // Neuladen der Seite (z.B. App am Handy zwischendurch geschlossen) erhalten bleibt.
  // Unkritisch genug (reine Komfortfunktion), um offline einfach ausgesetzt zu werden,
  // statt dieselbe Konflikt-Logik wie beim eigentlichen Datenstand zu durchlaufen - wird
  // beim nächsten Online-Speichern automatisch nachgezogen.
  useEffect(() => {
    if (loadStatus !== "ready" || offline) return;
    const t = setTimeout(() => {
      saveState(history, HISTORY_FILE).catch((error) => {
        console.error(error);
        if (istVerbindungsfehler(error)) setOffline(true);
      });
    }, 600);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [history, loadStatus, offline, retryTick]);

  // Nutzer hat für jede Konflikt-Zeile einzeln entschieden (Konflikt-Fenster) - Ergebnis
  // einarbeiten und speichern.
  async function konflikteEntschieden(resolutions: Map<string, "local" | "remote">) {
    const result = pendingSyncResultRef.current;
    if (!result) return;
    const finalData = applyConflictResolutions(result, resolutions);
    try {
      await saveState(finalData, SCHEMA_FILE);
      pendingSyncResultRef.current = null;
      baselineRef.current = finalData;
      writeBaseline(SCHEMA_FILE, finalData);
      clearPending(SCHEMA_FILE);
      setConflicts([]);
      setOffline(false);
      setData(finalData);
      const now = new Date().toLocaleTimeString("de-AT", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
      setSaveStatus(`Gespeichert um ${now}`);
      if (result.autoMerged.length > 0) setShowAutoMerged(true);
    } catch (error) {
      console.error(error);
      if (istVerbindungsfehler(error)) {
        // Entscheidung lokal übernehmen und merken, wird gespeichert sobald wieder online.
        pendingSyncResultRef.current = null;
        setConflicts([]);
        setOffline(true);
        setData(finalData);
        writePending(SCHEMA_FILE, finalData);
        setSaveStatus("Offline – Entscheidung gemerkt, wird gespeichert sobald wieder online.");
      } else {
        setSaveStatus("Speichern fehlgeschlagen – bitte Verbindung prüfen, dann erneut versuchen.");
      }
    }
  }

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

  // Personen-Tabs: normalerweise nur die, die für DIESE Reise tatsächlich schon
  // Gegenstände/Mengen haben (beteiligtePersonen) - von Clemens ausdrücklich so
  // gewünscht (2026-08-28), z.B. bei Schottland nur Clemens/Florian/Carina/Allgemein,
  // bei Grimming nur Clemens/Sonja, nicht alle bekannten Personen. Fällt NUR dann auf
  // alle bekannten Personen zurück, wenn eine Reise wirklich noch niemanden hat (frisch
  // importiert, noch nie bearbeitet) - sonst wäre man wieder in der Sackgasse von vorher
  // gefangen (siehe Grimming-Fix V01-27).
  const sichtbarePersonen = useMemo(() => {
    if (!data) return [];
    return beteiligtePersonen.length > 0 ? beteiligtePersonen : data.t05_namen;
  }, [data, beteiligtePersonen]);

  // Vorbelegung/Umschalten der ausgewählten Person: beim Laden UND beim Wechsel der
  // Reise wird geprüft, ob die aktuell gewählte Person für die neue Reise überhaupt
  // sichtbar ist - wenn nicht (z.B. Sonja war bei Grimming gewählt, jetzt wechselt man
  // zu China 2024, wo nur Clemens sichtbar ist), wird automatisch auf die erste
  // sichtbare Person umgeschaltet, statt eine unsichtbare Person aktiv zu lassen.
  useEffect(() => {
    if (sichtbarePersonen.length === 0) return;
    if (personFilter === null || !sichtbarePersonen.some((p) => p.id === personFilter)) {
      setPersonFilter(sichtbarePersonen[0].id);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reise?.id, sichtbarePersonen]);

  // Zuletzt angesehene Reise pro Gerät merken (siehe ermittleStartReise oben), damit ein
  // Neuladen/Aktualisieren nicht mehr auf die erste Reise zurückspringt.
  useEffect(() => {
    if (!selectedReiseId) return;
    try {
      localStorage.setItem(LETZTE_REISE_KEY, selectedReiseId);
    } catch {
      // ignorieren - reiner Komfort, kein Muss
    }
  }, [selectedReiseId]);

  const gruppiert = useMemo(() => {
    if (!data || !reise || !personFilter) return [];
    const byKat = new Map<string, { g: (typeof gegenstaende)[number]; row: Tk04GegenstandPerson }[]>();
    for (const g of gegenstaende) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      const personen = personenFuerTk03(data, tk03.id);
      const row = personen.find((p) => p.id_t05 === personFilter);
      if (!row) continue; // diese Person braucht den Gegenstand nicht
      if (row.ausgewaehlt === -1) continue; // Strich = sicher nicht mitgenommen, aus der Liste raus
      if (offenFilter === "neu") {
        if (!neuHinzugefuegt.has(row.id)) continue;
      } else if (offenFilter) {
        const feldWert = row[offenFilter];
        const istOffen = feldWert === null || feldWert === undefined;
        if (!istOffen) continue;
      }
      const kat = kathegorieName(data, g.id_kathegorie);
      if (!byKat.has(kat)) byKat.set(kat, []);
      byKat.get(kat)!.push({ g, row });
    }
    return Array.from(byKat.entries()).sort((a, b) => a[0].localeCompare(b[0], "de"));
  }, [data, gegenstaende, personFilter, offenFilter, reise, neuHinzugefuegt]);

  // Zaehlt, wie viele "neu hinzugefügt"-Zeilen es fuer die aktuelle Person/Reise gibt -
  // fuer die Anzeige "(n)" am Filter-Chip. Nur sichtbar, wenn > 0.
  const neuCount = useMemo(() => {
    if (!data || !reise || !personFilter || neuHinzugefuegt.size === 0) return 0;
    let n = 0;
    for (const g of gegenstaende) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      const row = data.tk04_tk03_t05.find((r) => r.id_tk03 === tk03.id && r.id_t05 === personFilter);
      if (row && neuHinzugefuegt.has(row.id) && row.ausgewaehlt !== -1) n++;
    }
    return n;
  }, [data, reise, gegenstaende, personFilter, neuHinzugefuegt]);

  const fortschritt = useMemo(() => {
    if (!data || !reise || !personFilter) return { done: 0, total: 0 };
    let done = 0;
    let total = 0;
    for (const g of gegenstaende) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      const row = data.tk04_tk03_t05.find((r) => r.id_tk03 === tk03.id && r.id_t05 === personFilter);
      if (!row) continue; // Gegenstand betrifft diese Person nicht
      if (row.ausgewaehlt === -1) continue; // Strich zählt nicht mit
      total++;
      if (row.eingepackt !== null && row.eingepackt !== undefined && row.eingepackt > 0) done++;
    }
    return { done, total };
  }, [data, reise, gegenstaende, personFilter]);

  // Alle inhaltlichen Änderungen laufen über diese Funktion statt über setData direkt,
  // damit vor jeder Änderung der bisherige Stand im Rückgängig-Verlauf gesichert wird.
  // Der updater bekommt garantiert einen nicht-null Stand und gibt entweder eine
  // geänderte Kopie zurück, oder exakt dasselbe Objekt (===), wenn nichts zu tun war -
  // in dem Fall wird auch nichts im Verlauf vermerkt.
  function updateData(updater: (prev: SchemaData) => SchemaData) {
    setData((prev) => {
      if (!prev) return prev;
      const next = updater(prev);
      if (next === prev) return prev;
      setHistory((h) => {
        const erweitert = [...h, prev];
        return erweitert.length > MAX_HISTORY ? erweitert.slice(erweitert.length - MAX_HISTORY) : erweitert;
      });
      return next;
    });
  }

  // Letzten Schritt rückgängig machen: den zuletzt gesicherten Stand vor der Änderung
  // wiederherstellen. Mehrfach hintereinander drückbar, bis der Verlauf leer ist.
  function undo() {
    setHistory((h) => {
      if (h.length === 0) return h;
      const letzter = h[h.length - 1];
      setData(letzter);
      return h.slice(0, h.length - 1);
    });
  }

  function bumpField(
    row: Tk04GegenstandPerson,
    field: "ausgewaehlt" | "hergerichtet" | "eingepackt" | "verwendet",
    direction: 1 | -1
  ) {
    updateData((prev) => {
      const updated = prev.tk04_tk03_t05.map((r) => {
        if (r.id !== row.id) return r;

        if (field === "eingepackt") {
          // Eingepackt bleibt wie bisher: nur "leer" (–) oder Zahl, kein Strich-Zustand.
          const current = r.eingepackt;
          const has = current !== null && current !== undefined;
          if (direction === 1) {
            if (!has) return { ...r, eingepackt: r.ausgewaehlt ?? 1 };
            return { ...r, eingepackt: current + 1 };
          }
          if (!has) return r;
          if (current <= 1) return { ...r, eingepackt: null };
          return { ...r, eingepackt: current - 1 };
        }

        // Ausgewählt/Hergerichtet/Verwendet: -1 = Strich (bewusst ausgeschlossen),
        // 0 = unsicher, Zahl > 0 = Menge. In 1er-Schritten, Untergrenze -1.
        if (field === "ausgewaehlt") {
          const next = Math.max(-1, (r.ausgewaehlt ?? 0) + direction);
          if (next === -1 && (r.ausgewaehlt ?? 0) !== -1) {
            // Beim Wechsel auf Strich: Hergerichtet/Eingepackt/Verwendet zurücksetzen,
            // die gehören nicht mehr dazu, wenn der Gegenstand sicher nicht mitkommt.
            return { ...r, ausgewaehlt: next, hergerichtet: null, eingepackt: null, verwendet: null };
          }
          return { ...r, ausgewaehlt: next };
        }

        // hergerichtet oder verwendet
        const current = r[field] ?? 0;
        const next = Math.max(-1, current + direction);
        return { ...r, [field]: next };
      });
      // Sobald an einer frisch hinzugefügten Zeile etwas eingetragen wird, gilt sie
      // nicht mehr als "neu" - verschwindet also aus dem "Neu hinzugefügt"-Filter.
      // Im selben updateData-Aufruf wie die eigentliche Änderung, damit es nur einen
      // Rückgängig-Schritt gibt und beides zusammen gespeichert wird.
      const neuNeuHinzugefuegt = (prev.neu_hinzugefuegt ?? []).filter((id) => id !== row.id);
      return { ...prev, tk04_tk03_t05: updated, neu_hinzugefuegt: neuNeuHinzugefuegt };
    });
  }

  // Antippen der linken Hälfte eines Feldes = runterzählen, rechte Hälfte = hochzählen.
  // (Wischen war auf dem Handy unzuverlässig, weil die Seite dabei mitscrollt.)
  function handleFieldTap(
    e: React.MouseEvent<HTMLButtonElement>,
    row: Tk04GegenstandPerson,
    field: "ausgewaehlt" | "hergerichtet" | "eingepackt" | "verwendet"
  ) {
    if (longPressFired.current) {
      // Der Tap kam direkt nach einem langen Druck (öffnet Direkteingabe) - nicht mehr zählen.
      longPressFired.current = false;
      return;
    }
    const rect = e.currentTarget.getBoundingClientRect();
    const isRightHalf = e.clientX - rect.left > rect.width / 2;
    bumpField(row, field, isRightHalf ? 1 : -1);
  }

  // Langer Druck (ca. 450ms halten) auf ein Mengenfeld öffnet die Direkteingabe der Zahl.
  function startLongPress(
    row: Tk04GegenstandPerson,
    field: "ausgewaehlt" | "hergerichtet" | "eingepackt" | "verwendet"
  ) {
    longPressFired.current = false;
    if (longPressTimer.current) clearTimeout(longPressTimer.current);
    longPressTimer.current = setTimeout(() => {
      longPressFired.current = true;
      const current = row[field];
      setEditingQty({ rowId: row.id, field });
      setEditingQtyValue(current !== null && current !== undefined ? String(current) : "");
    }, 450);
  }

  function cancelLongPress() {
    if (longPressTimer.current) {
      clearTimeout(longPressTimer.current);
      longPressTimer.current = null;
    }
  }

  // Direkt eingegebene Zahl speichern. Leeres Feld = "ausgewählt" zurück auf 0,
  // bei den anderen drei Feldern zurück auf "nicht angetippt" (null), wie beim Antippen.
  function speichereQtyEingabe() {
    if (!editingQty) return;
    const { rowId, field } = editingQty;
    const trimmed = editingQtyValue.trim();
    updateData((prev) => {
      return {
        ...prev,
        tk04_tk03_t05: prev.tk04_tk03_t05.map((r) => {
          if (r.id !== rowId) return r;
          if (field === "eingepackt") {
            if (trimmed === "") return { ...r, eingepackt: null };
            const n = Math.max(0, Math.floor(Number(trimmed)));
            if (Number.isNaN(n)) return r;
            return { ...r, eingepackt: n };
          }
          // ausgewählt/hergerichtet/verwendet: -1 (Strich) bis beliebig hoch
          if (trimmed === "" || trimmed === "-") {
            const reset = { ...r, [field]: 0 };
            return reset;
          }
          const n = Math.max(-1, Math.floor(Number(trimmed)));
          if (Number.isNaN(n)) return r;
          if (field === "ausgewaehlt" && n === -1 && (r.ausgewaehlt ?? 0) !== -1) {
            return { ...r, ausgewaehlt: -1, hergerichtet: null, eingepackt: null, verwendet: null };
          }
          return { ...r, [field]: n };
        }),
        // Wie bei bumpField: Zeile verliert die "neu"-Markierung, sobald sie bearbeitet wird.
        neu_hinzugefuegt: (prev.neu_hinzugefuegt ?? []).filter((id) => id !== rowId),
      };
    });
    setEditingQty(null);
    setEditingQtyValue("");
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
    updateData((prev) => {
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

  // V01-20: Gegenstand für die aktuell ausgewählte Person zur Reise hinzufügen bzw.
  // reaktivieren - der einzige Weg aus der kompakten Auswahl-Liste in "Liste
  // bearbeiten". Löscht nie etwas, nur hinzufügen/reaktivieren:
  // - Gegenstand noch nicht auf der Reise: tk03 (Reise-Zuordnung) + tk04 (für diese
  //   Person, Menge "0 - unsicher") werden neu angelegt.
  // - Gegenstand ist schon auf der Reise (z.B. bei einer anderen Person), aber diese
  //   Person hat noch keine Zeile: nur eine neue tk04-Zeile für sie.
  // - Diese Person hatte den Gegenstand schon mal (Strich, ausgewählt = -1): die
  //   bestehende Zeile wird auf "0 - unsicher" zurückgesetzt statt eine zweite anzulegen.
  function fuegeGegenstandHinzu(gegenstandId: string) {
    if (!data || !reise || !personFilter) return;
    updateData((prev) => {
      const ids = collectAllIds(prev);
      let tk03 = tk03FuerGegenstand(prev, reise.id, gegenstandId);
      let tk03Rows = prev.tk03_tk01_t04;
      if (!tk03) {
        const anker = ankerTk01(prev, reise.id);
        if (!anker) return prev;
        const neuTk03Id = newId("tk03", ids);
        ids.add(neuTk03Id);
        tk03 = { id: neuTk03Id, id_tk01: anker, id_t04: gegenstandId, notiz: "manuell ausgewählt" };
        tk03Rows = [...prev.tk03_tk01_t04, tk03];
      }
      const bestehendeZeile = prev.tk04_tk03_t05.find(
        (r) => r.id_tk03 === tk03!.id && r.id_t05 === personFilter
      );
      let tk04Rows = prev.tk04_tk03_t05;
      let betroffeneZeileId: string;
      if (bestehendeZeile) {
        betroffeneZeileId = bestehendeZeile.id;
        tk04Rows = prev.tk04_tk03_t05.map((r) =>
          r.id === bestehendeZeile.id ? { ...r, ausgewaehlt: 0 } : r
        );
      } else {
        const neuId = newId("tk04", ids);
        betroffeneZeileId = neuId;
        const neu: Tk04GegenstandPerson = {
          id: neuId,
          id_tk03: tk03.id,
          id_t05: personFilter,
          ausgewaehlt: 0,
          hergerichtet: null,
          eingepackt: null,
          verwendet: null,
        };
        tk04Rows = [...prev.tk04_tk03_t05, neu];
      }
      // Als "neu" markieren, im selben updateData-Aufruf wie das eigentliche Anlegen/
      // Reaktivieren, damit beides zusammen (ein Rückgängig-Schritt) gespeichert wird.
      // Seit V01-29 Teil von `data` -> übersteht Neuladen und Geräte-Wechsel.
      const bisherige = prev.neu_hinzugefuegt ?? [];
      const neuHinzugefuegtNext = bisherige.includes(betroffeneZeileId)
        ? bisherige
        : [...bisherige, betroffeneZeileId];
      return {
        ...prev,
        tk03_tk01_t04: tk03Rows,
        tk04_tk03_t05: tk04Rows,
        neu_hinzugefuegt: neuHinzugefuegtNext,
      };
    });
  }

  function toggleGegenstandInReise(gegenstandId: string) {
    if (!data || !reise) return;
    updateData((prev) => {
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

  // Person einem Gegenstand (tk03-Zeile) zuweisen (an) oder wieder entfernen (aus).
  // Entfernen löscht die tk04-Zeile komplett inkl. aller vier Mengenwerte.
  function togglePersonZuweisung(tk03Id: string, personId: string) {
    updateData((prev) => {
      const bestehende = prev.tk04_tk03_t05.find(
        (r) => r.id_tk03 === tk03Id && r.id_t05 === personId
      );
      if (bestehende) {
        return {
          ...prev,
          tk04_tk03_t05: prev.tk04_tk03_t05.filter((r) => r.id !== bestehende.id),
        };
      }
      const ids = collectAllIds(prev);
      const neuId = newId("tk04", ids);
      const neu: Tk04GegenstandPerson = {
        id: neuId,
        id_tk03: tk03Id,
        id_t05: personId,
        ausgewaehlt: null,
        hergerichtet: null,
        eingepackt: null,
        verwendet: null,
      };
      return { ...prev, tk04_tk03_t05: [...prev.tk04_tk03_t05, neu] };
    });
  }

  // Eine bestehende Personen-Zuordnung auf eine andere Person umhängen.
  // Alle vier Mengenwerte (A/H/E/V) bleiben dabei unverändert erhalten.
  function verschiebePersonZuweisung(tk04Id: string, neuePersonId: string) {
    updateData((prev) => {
      return {
        ...prev,
        tk04_tk03_t05: prev.tk04_tk03_t05.map((r) =>
          r.id === tk04Id ? { ...r, id_t05: neuePersonId } : r
        ),
      };
    });
    setMoveFor(null);
  }

  // Komplett neuen Gegenstand im Katalog (t04) anlegen, ggf. mit neuer Kategorie (t03).
  function erstelleNeuenGegenstand() {
    if (!data || !neuGegenstandName.trim()) return;
    const neueKatName = neuGegenstandKatNeu.trim();
    if (!neuGegenstandKat && !neueKatName) return; // Kategorie fehlt
    updateData((prev) => {
      const ids = collectAllIds(prev);
      let katId = neuGegenstandKat;
      let t03Neu = prev.t03_kathegorie;
      if (!katId && neueKatName) {
        katId = newId("t03", ids);
        ids.add(katId);
        t03Neu = [...prev.t03_kathegorie, { id: katId, kathegorie: neueKatName, notiz: "" }];
      }
      const gegenstandId = newId("t04", ids);
      const neuesGegenstand: T04Gegenstand = {
        id: gegenstandId,
        gegenstand: neuGegenstandName.trim(),
        id_kathegorie: katId,
        notiz: "",
      };
      return {
        ...prev,
        t03_kathegorie: t03Neu,
        t04_gegenstand: [...prev.t04_gegenstand, neuesGegenstand],
      };
    });
    setNeuGegenstandName("");
    setNeuGegenstandKat("");
    setNeuGegenstandKatNeu("");
    setShowNeuGegenstand(false);
  }

  // Alle Zeilen (alle Personen, ganze Reise) für Export/Druck aufbereiten - sortiert nach Kategorie, Gegenstand, Person.
  const exportZeilen = useMemo<ExportZeile[]>(() => {
    if (!data || !reise) return [];
    const rows: ExportZeile[] = [];
    for (const g of gegenstaende) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      const kat = kathegorieName(data, g.id_kathegorie);
      for (const row of data.tk04_tk03_t05.filter((r) => r.id_tk03 === tk03.id)) {
        const person = data.t05_namen.find((p) => p.id === row.id_t05);
        rows.push({
          kategorie: kat,
          gegenstand: g.gegenstand,
          person: person?.namen ?? "?",
          ausgewaehlt: row.ausgewaehlt ?? 0,
          hergerichtet: row.hergerichtet ?? null,
          eingepackt: row.eingepackt ?? null,
          verwendet: row.verwendet ?? null,
        });
      }
    }
    rows.sort(
      (a, b) =>
        a.kategorie.localeCompare(b.kategorie, "de") ||
        a.gegenstand.localeCompare(b.gegenstand, "de") ||
        a.person.localeCompare(b.person, "de")
    );
    return rows;
  }, [data, reise, gegenstaende]);

  async function createPdf() {
    if (!reise) return;
    setExportStatus("PDF wird erstellt …");
    try {
      const [{ jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
      const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
      const pageWidth = doc.internal.pageSize.getWidth();
      doc.setTextColor(8, 45, 73);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(18);
      doc.text(reise.reise || "Packliste", 12, 14);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(80, 92, 98);
      const zeitraum =
        reise.von && reise.bis
          ? `${new Date(reise.von).toLocaleDateString("de-AT")} – ${new Date(reise.bis).toLocaleDateString("de-AT")}`
          : "Zeitraum noch offen";
      doc.text(`Zeitraum: ${zeitraum}   |   Personen: ${beteiligtePersonen.map((p) => p.namen).join(", ")}`, 12, 20);
      doc.text("A = ausgewählt (geplant)   H = hergerichtet   E = eingepackt   V = verwendet", 12, 25);

      autoTable(doc, {
        startY: 29,
        margin: { top: 15, right: 12, bottom: 13, left: 12 },
        theme: "grid",
        head: [["Kategorie", "Gegenstand", "Person", "A", "H", "E", "V"]],
        body: exportZeilen.map((r) => [
          r.kategorie,
          r.gegenstand,
          r.person,
          String(r.ausgewaehlt),
          r.hergerichtet ?? "",
          r.eingepackt ?? "",
          r.verwendet ?? "",
        ]),
        styles: { font: "helvetica", fontSize: 8, cellPadding: 1.8, lineColor: [210, 214, 207], lineWidth: 0.2, textColor: [24, 38, 45], overflow: "linebreak", valign: "middle" },
        headStyles: { fillColor: [8, 45, 73], textColor: [255, 255, 255], fontStyle: "bold", halign: "left" },
        alternateRowStyles: { fillColor: [247, 249, 244] },
        columnStyles: {
          0: { cellWidth: 55 }, 1: { cellWidth: 90 }, 2: { cellWidth: 45 },
          3: { cellWidth: 20, halign: "center" }, 4: { cellWidth: 20, halign: "center" },
          5: { cellWidth: 20, halign: "center" }, 6: { cellWidth: 20, halign: "center" },
        },
        didDrawPage: () => {
          const pageNumber = doc.getNumberOfPages();
          doc.setFontSize(7.5);
          doc.setTextColor(105, 110, 112);
          doc.text(`P03 Packliste · Seite ${pageNumber}`, pageWidth - 12, doc.internal.pageSize.getHeight() - 6, { align: "right" });
        },
      });

      doc.save(`P03_${safeFilename(reise.reise)}_Packliste_${dateStamp()}.pdf`);
      setExportStatus(`PDF für „${reise.reise}“ wurde heruntergeladen.`);
    } catch {
      setExportStatus("Das PDF konnte nicht erstellt werden.");
    }
  }

  async function createExcel() {
    if (!reise || !data) return;
    setExportStatus("Excel-Datei wird erstellt …");
    try {
      const ExcelJS = (await import("exceljs")).default;
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "P03 Packliste";
      workbook.created = new Date();
      const navy = "082D49";
      const green = "39734B";
      const borderColor = "DDD8CF";

      const overview = workbook.addWorksheet("Reise", { views: [{ showGridLines: false }] });
      overview.columns = [{ width: 24 }, { width: 72 }];
      overview.mergeCells("A1:B1");
      overview.getCell("A1").value = reise.reise || "Packliste";
      overview.getCell("A1").font = { bold: true, size: 18, color: { argb: "FFFFFFFF" } };
      overview.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${navy}` } };
      overview.getCell("A1").alignment = { vertical: "middle" };
      overview.getRow(1).height = 30;
      overview.addRows([
        ["Von", reise.von ? new Date(reise.von).toLocaleDateString("de-AT") : "Noch offen"],
        ["Bis", reise.bis ? new Date(reise.bis).toLocaleDateString("de-AT") : "Noch offen"],
        ["Personen", beteiligtePersonen.map((p) => p.namen).join(", ") || "Keine"],
        ["Gegenstände", gegenstaende.length],
        ["Exportiert", new Date()],
      ]);
      overview.getColumn(1).font = { bold: true, color: { argb: `FF${green}` } };
      overview.getCell("B5").numFmt = "yyyy-mm-dd hh:mm";
      overview.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          row.alignment = { vertical: "top", wrapText: true };
          row.eachCell((cell) => { cell.border = { bottom: { style: "thin", color: { argb: `FF${borderColor}` } } }; });
        }
      });

      const list = workbook.addWorksheet("Packliste", { views: [{ state: "frozen", ySplit: 1, showGridLines: false }] });
      list.autoFilter = "A1:G1";
      list.columns = [
        { header: "Kategorie", key: "kategorie", width: 24 },
        { header: "Gegenstand", key: "gegenstand", width: 34 },
        { header: "Person", key: "person", width: 20 },
        { header: "Ausgewählt", key: "ausgewaehlt", width: 13 },
        { header: "Hergerichtet", key: "hergerichtet", width: 14 },
        { header: "Eingepackt", key: "eingepackt", width: 13 },
        { header: "Verwendet", key: "verwendet", width: 13 },
      ];
      list.addRows(exportZeilen);
      list.getRow(1).height = 28;
      list.getRow(1).eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${navy}` } };
        cell.alignment = { vertical: "middle", wrapText: true };
      });
      list.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        row.eachCell((cell) => { cell.border = { bottom: { style: "thin", color: { argb: `FF${borderColor}` } } }; });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      downloadBlob(
        new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
        `P03_${safeFilename(reise.reise)}_Packliste_${dateStamp()}.xlsx`
      );
      setExportStatus(`Excel-Datei für „${reise.reise}“ wurde heruntergeladen.`);
    } catch {
      setExportStatus("Die Excel-Datei konnte nicht erstellt werden.");
    }
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
      {/* Fixierter Kopfbereich: bleibt beim Scrollen der Gegenstands-Liste
          oben stehen (Excel-artige "eingefrorene Zeilen"), siehe .pl-sticky-top. */}
      <div className="pl-sticky-top">
        <header className="pl-header">
          <div className="pl-header-top">
            <h1>{reise?.reise ?? "P03 Packliste"}</h1>
            <div className="pl-header-actions">
              <span className="pl-version">{APP_VERSION}</span>
              <button className="pl-logout" onClick={() => logout()}>Abmelden</button>
            </div>
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
          <p className={"pl-save" + (offline ? " offline" : "")}>{offline ? "🔌 " : ""}{saveStatus}</p>
        </header>

        {showAutoMerged && autoMerged.length > 0 && (
          <div className="pl-automerge-banner">
            Ein anderes Gerät hat inzwischen ebenfalls gespeichert. Automatisch übernommen:
            <ul>
              {autoMerged.map((c, i) => (
                <li key={`${c.table}-${c.id}-${i}`}>{c.label} ({c.art})</li>
              ))}
            </ul>
            <button type="button" onClick={() => setShowAutoMerged(false)}>Verstanden</button>
          </div>
        )}

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
          </button>{" "}
          {reise && (
            <>
              <button className="pl-edit-toggle" onClick={createPdf}>Als PDF drucken</button>{" "}
              <button className="pl-edit-toggle" onClick={createExcel}>Als Excel sichern</button>{" "}
            </>
          )}
          <button
            className="pl-edit-toggle"
            onClick={undo}
            disabled={history.length === 0}
            title={history.length > 0 ? `${history.length} Schritt(e) verfügbar` : "Kein Verlauf vorhanden"}
          >
            ↩ Rückgängig{history.length > 0 ? ` (${history.length})` : ""}
          </button>
        </div>
        {exportStatus && <p className="pl-save">{exportStatus}</p>}

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
            <div className="pl-select-row">
              <button
                className={"pl-edit-toggle" + (showNeuGegenstand ? " active" : "")}
                onClick={() => setShowNeuGegenstand((v) => !v)}
              >
                {showNeuGegenstand ? "Abbrechen" : "+ Neuer Gegenstand"}
              </button>
            </div>
            {showNeuGegenstand && (
              <div className="pl-newreise">
                <label>Name des Gegenstands</label>
                <input
                  value={neuGegenstandName}
                  onChange={(e) => setNeuGegenstandName(e.target.value)}
                  placeholder="z.B. Regenjacke"
                />
                <label>Kategorie (bestehende wählen)</label>
                <select
                  value={neuGegenstandKat}
                  onChange={(e) => {
                    setNeuGegenstandKat(e.target.value);
                    if (e.target.value) setNeuGegenstandKatNeu("");
                  }}
                >
                  <option value="">– bitte wählen –</option>
                  {data.t03_kathegorie
                    .slice()
                    .sort((a, b) => a.kathegorie.localeCompare(b.kathegorie, "de"))
                    .map((k) => (
                      <option key={k.id} value={k.id}>{k.kathegorie}</option>
                    ))}
                </select>
                <label>… oder neue Kategorie eintippen</label>
                <input
                  value={neuGegenstandKatNeu}
                  onChange={(e) => {
                    setNeuGegenstandKatNeu(e.target.value);
                    if (e.target.value) setNeuGegenstandKat("");
                  }}
                  placeholder="z.B. Regenschutz"
                />
                {(() => {
                  const name = neuGegenstandName.trim().toLowerCase();
                  if (!name) return null;
                  const treffer = data.t04_gegenstand.find((g) => g.gegenstand.trim().toLowerCase() === name);
                  if (!treffer) return null;
                  return (
                    <p className="pl-save" style={{ color: "#7a5c21" }}>
                      Gibt es schon in „{kathegorieName(data, treffer.id_kathegorie)}“ – lieber über die Auswahl-Liste
                      unten hinzufügen statt doppelt anzulegen?
                    </p>
                  );
                })()}
                <button
                  onClick={erstelleNeuenGegenstand}
                  disabled={!neuGegenstandName.trim() || (!neuGegenstandKat && !neuGegenstandKatNeu.trim())}
                >
                  Gegenstand anlegen
                </button>
              </div>
            )}
            {!personFilter && <p className="pl-save">Bitte zuerst oben eine Person auswählen.</p>}
            {personFilter && (
              <div className="pl-add-legend">
                <span><i className="rot"></i>wieder aufnehmen</span>
                <span><i className="blau"></i>andere Person</span>
                <span><i className="gruen"></i>neu hinzufügen</span>
              </div>
            )}
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
              {/* Nur die für DIESE Reise tatsächlich beteiligten Personen (siehe
                  sichtbarePersonen oben) - nicht mehr immer alle, das war zu viel. */}
              {sichtbarePersonen.map((p) => (
                <button
                  key={p.id}
                  className={"pl-filter-chip" + (personFilter === p.id ? " active" : "")}
                  onClick={() => setPersonFilter(p.id)}
                >
                  {p.namen}
                </button>
              ))}
            </div>

            <div className="pl-filterbar pl-offenbar">
              <button
                className={"pl-filter-chip pl-offen-chip" + (offenFilter === "hergerichtet" ? " active" : "")}
                onClick={() => setOffenFilter(offenFilter === "hergerichtet" ? null : "hergerichtet")}
              >
                Herrichten offen
              </button>
              <button
                className={"pl-filter-chip pl-offen-chip" + (offenFilter === "eingepackt" ? " active" : "")}
                onClick={() => setOffenFilter(offenFilter === "eingepackt" ? null : "eingepackt")}
              >
                Einpacken offen
              </button>
              {neuCount > 0 && (
                <button
                  className={"pl-filter-chip pl-offen-chip" + (offenFilter === "neu" ? " active" : "")}
                  onClick={() => setOffenFilter(offenFilter === "neu" ? null : "neu")}
                >
                  Neu hinzugefügt ({neuCount})
                </button>
              )}
            </div>

            <div className="pl-legend">
              <span title="Ausgewählt (geplant)">A</span>
              <span title="Hergerichtet">H</span>
              <span title="Eingepackt">E</span>
              <span title="Verwendet">V</span>
            </div>
          </>
        )}
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

      {mode === "bearbeiten" && reise && personFilter && (
        <div className="pl-body">
          {(() => {
            type Prio = "rot" | "blau" | "gruen";
            type Eintrag = { g: T04Gegenstand; prio: Prio; hint: string | null };
            const byKat = new Map<string, Eintrag[]>();
            for (const g of data.t04_gegenstand) {
              if (editSearch && !g.gegenstand.toLowerCase().includes(editSearch.toLowerCase())) continue;
              const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
              const meineZeile = tk03
                ? data.tk04_tk03_t05.find((r) => r.id_tk03 === tk03.id && r.id_t05 === personFilter)
                : undefined;
              if (meineZeile && meineZeile.ausgewaehlt !== -1) continue; // hab ich schon aktiv - nicht nochmal anzeigen

              let prio: Prio;
              let hint: string | null = null;
              if (meineZeile && meineZeile.ausgewaehlt === -1) {
                prio = "rot"; // war bei mir, ist rausgefallen
              } else {
                const andere = tk03
                  ? data.tk04_tk03_t05.filter(
                      (r) => r.id_tk03 === tk03!.id && r.id_t05 !== personFilter && r.ausgewaehlt !== -1
                    )
                  : [];
                if (andere.length > 0) {
                  prio = "blau";
                  hint = andere
                    .map((r) => data.t05_namen.find((p) => p.id === r.id_t05)?.namen ?? "?")
                    .join(", ");
                } else {
                  prio = "gruen";
                }
              }
              const kat = kathegorieName(data, g.id_kathegorie);
              if (!byKat.has(kat)) byKat.set(kat, []);
              byKat.get(kat)!.push({ g, prio, hint });
            }
            const prioRang: Record<Prio, number> = { rot: 0, blau: 1, gruen: 2 };
            const sorted = Array.from(byKat.entries()).sort((a, b) => a[0].localeCompare(b[0], "de"));
            return sorted.map(([katName, items]) => {
              items.sort(
                (a, b) => prioRang[a.prio] - prioRang[b.prio] || a.g.gegenstand.localeCompare(b.g.gegenstand, "de")
              );
              return (
                <div className="pl-category" key={katName}>
                  <div className="pl-category-head">
                    <span className="pl-category-tag">{katName}</span>
                    <span className="pl-category-count">{items.length}</span>
                  </div>
                  {items.map(({ g, prio, hint }) => (
                    <div className={"pl-add-item pri-" + prio} key={g.id}>
                      <div className="pl-add-item-row">
                        <div className="pl-add-item-text">
                          <span className="pl-add-item-name">{g.gegenstand}</span>
                          {hint && <span className="pl-add-item-hint">{hint}</span>}
                        </div>
                        <button className="pl-add-btn" onClick={() => fuegeGegenstandHinzu(g.id)}>
                          +
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              );
            });
          })()}
        </div>
      )}

      {mode === "liste" && (
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
                  {editingQty && editingQty.rowId === row.id && editingQty.field === "ausgewaehlt" ? (
                    <input
                      className="pl-qbox pl-qbox-input q-ausgewaehlt"
                      type="number"
                      inputMode="numeric"
                      autoFocus
                      value={editingQtyValue}
                      onChange={(e) => setEditingQtyValue(e.target.value)}
                      onBlur={speichereQtyEingabe}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") (e.target as HTMLInputElement).blur();
                        if (e.key === "Escape") setEditingQty(null);
                      }}
                    />
                  ) : (
                    <button
                      className="pl-qbox q-ausgewaehlt"
                      onClick={(e) => handleFieldTap(e, row, "ausgewaehlt")}
                      onMouseDown={() => startLongPress(row, "ausgewaehlt")}
                      onMouseUp={cancelLongPress}
                      onMouseLeave={cancelLongPress}
                      onTouchStart={() => startLongPress(row, "ausgewaehlt")}
                      onTouchEnd={cancelLongPress}
                    >
                      {row.ausgewaehlt === -1 ? "–" : row.ausgewaehlt ?? 0}
                    </button>
                  )}
                  {(["hergerichtet", "eingepackt", "verwendet"] as const).map((field) => {
                    const val = row[field];
                    const hatStrich = field !== "eingepackt" && val === -1;
                    const has = val !== null && val !== undefined && !hatStrich;
                    const istGesetzt = has && (val as number) > 0;
                    if (editingQty && editingQty.rowId === row.id && editingQty.field === field) {
                      return (
                        <input
                          key={field}
                          className={"pl-qbox pl-qbox-input q-" + field}
                          type="number"
                          inputMode="numeric"
                          autoFocus
                          value={editingQtyValue}
                          onChange={(e) => setEditingQtyValue(e.target.value)}
                          onBlur={speichereQtyEingabe}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") (e.target as HTMLInputElement).blur();
                            if (e.key === "Escape") setEditingQty(null);
                          }}
                        />
                      );
                    }
                    return (
                      <button
                        key={field}
                        className={"pl-qbox q-" + field + (istGesetzt ? " q-set" : "")}
                        onClick={(e) => handleFieldTap(e, row, field)}
                        onMouseDown={() => startLongPress(row, field)}
                        onMouseUp={cancelLongPress}
                        onMouseLeave={cancelLongPress}
                        onTouchStart={() => startLongPress(row, field)}
                        onTouchEnd={cancelLongPress}
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
      )}
      {conflicts.length > 0 && <ConflictModal conflicts={conflicts} onResolve={konflikteEntschieden} />}
    </div>
  );
}
