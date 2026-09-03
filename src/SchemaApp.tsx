import { useEffect, useMemo, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import seedData from "./data/schema-data.json";
import { logout } from "./auth";
import { loadState, saveState, listFiles, deleteState } from "./onedrive";
import type {
  SchemaData,
  Tk04GegenstandPerson,
  Tk01ReiseAktivitaet,
  Tk03ReiseaktivitaetGegenstand,
  T01Reise,
  T02Aktivitaet,
  T04Gegenstand,
  ID,
} from "./data/schema";
import {
  gegenstaendeFuerReise,
  tk03FuerGegenstand,
  personenFuerTk03,
  kathegorieName,
  newId,
} from "./data/schema";
import {
  STAMMDATEN_DATEI,
  reiseDateiname,
  reiseIdAusDateiname,
  splitSchemaData,
  mergeSplitData,
  stammdatenSnapshotFuerReise,
  stammdatenAusSnapshotsRekonstruieren,
  stammdatenAlsSchemaData,
  schemaDataAlsStammdatenKern,
  reiseAlsSchemaData,
  schemaDataAlsReiseKern,
  verlaufAufteilen,
  historyRekonstruieren,
  type StammdatenKern,
  type ReiseKern,
  type StammdatenDatei,
  type ReiseDatei,
} from "./data/splitSchema";
import { diffAndMerge, applyConflictResolutions, hashData, schemaEqual, type RowConflict, type AutoMergedChange, type SyncResult } from "./sync";
import { readBaseline, writeBaseline, readPending, writePending, clearPending, istVerbindungsfehler } from "./syncStore";
import ConflictModal from "./ConflictModal";

// Von Clemens gewünscht (2026-08-27): sichtbare Versionsnummer im Kopfbereich,
// damit jederzeit erkennbar ist, ob GitHub Pages wirklich den aktuellsten Stand
// ausliefert. Bei jeder Auslieferung hier mitziehen.
const APP_VERSION = "V03-05";

// Bis V02-02 einziger Speicherort/Dateiname. Seit Paket A (Aufteilung in Stammdaten- +
// Reise-Einzeldateien, siehe splitSchema.ts) nur noch als LESE-Quelle für die einmalige
// Migration beim ersten Start mit dem neuen Code relevant - wird danach nicht mehr
// beschrieben. Bleibt unangetastet in OneDrive liegen, Clemens löscht sie bei Gelegenheit
// selbst (so mit ihm abgestimmt, 2026-08-29), der Code greift nicht mehr automatisch
// darauf zu.
const LEGACY_SCHEMA_FILE = "P03_Packliste_AI.json";
// Lokaler Cache-Schlüssel (localStorage, siehe syncStore.ts) für den zusammengeführten
// Gesamtstand - bewusst derselbe String wie früher der Dateiname, damit ein evtl. noch
// nicht synchronisierter Offline-Stand aus einer Sitzung vor Paket A nicht verloren geht.
const LOKALER_CACHE_SCHLUESSEL = "P03_Packliste_AI.json";

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

// Allgemeiner Nachtrag (gefunden 2026-08-30 beim Testen von "China 2025", siehe
// Projektstand): Jede Reise braucht mindestens eine tk01-Zeile (Verknüpfung zu einer
// Aktivität, meist "Basics"), sonst hat "Liste bearbeiten"/"Gegenstand hinzufügen"
// keinen Anker-Punkt zum Anlegen neuer Zuordnungen und tut buchstäblich gar nichts
// (siehe ankerTk01 weiter unten - lieferte für so eine Reise `null`, "+"-Knopf blieb
// wirkungslos). "China 2025" wurde offenbar angelegt, ohne dass ihr je die "Basics"-
// Aktivität zugeordnet wurde (0 tk01-Zeilen, anders als bei den anderen drei Reisen -
// vermutlich noch nie richtig benutzt). Legt wie bei backfillPersonenOhneMenge nie
// doppelte Zeilen an, sobald einmal eine tk01-Zeile existiert passiert hier nichts mehr.
export function backfillReisenOhneAnker(data: SchemaData): SchemaData {
  const basics = data.t02_aktivitaet.find((a) => a.aktivitaet === "Basics");
  if (!basics) return data;
  const reisenOhneAnker = data.t01_reise.filter(
    (r) => !data.tk01_t01_t02.some((tk01) => tk01.id_t01 === r.id)
  );
  if (reisenOhneAnker.length === 0) return data;

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

  const neueZeilen = reisenOhneAnker.map((reise) => {
    const neuId = newId("tk01", existingIds);
    existingIds.add(neuId);
    return { id: neuId, id_t01: reise.id, id_t02: basics.id };
  });
  return { ...data, tk01_t01_t02: [...data.tk01_t01_t02, ...neueZeilen] };
}

// Nachtrag für bestehende Reisen ohne explizite Teilnehmerliste (V03-02, siehe
// T01Reise.teilnehmer in schema.ts): leitet sie einmalig aus den tatsächlich
// vorhandenen Personen-Zuordnungen ab, damit Schottland/China 2024/Grimming & Co. nach
// dem Update genau dieselben Personen-Reiter zeigen wie bisher - keine sichtbare
// Änderung für bestehende Reisen. Neu angelegte Reisen bekommen ihre Teilnehmerliste ab
// V03-02 immer schon beim Anlegen explizit mit (siehe erstelleNeueReise), brauchen
// diesen Nachtrag also nicht. Läuft nie doppelt (nur Reisen ohne das Feld betroffen).
export function backfillFehlendeTeilnehmer(data: SchemaData): SchemaData {
  let geaendert = false;
  const neueReisen = data.t01_reise.map((reise) => {
    if (reise.teilnehmer !== undefined) return reise;
    const ids = new Set<ID>();
    for (const g of gegenstaendeFuerReise(data, reise.id)) {
      const tk03 = tk03FuerGegenstand(data, reise.id, g.id);
      if (!tk03) continue;
      for (const p of personenFuerTk03(data, tk03.id)) ids.add(p.id_t05);
    }
    geaendert = true;
    return { ...reise, teilnehmer: Array.from(ids) };
  });
  if (!geaendert) return data;
  return { ...data, t01_reise: neueReisen };
}

// Bündelt alle "Nachträge" (Selbstheilung von historisch unvollständigen Daten) an
// einer Stelle. Bis V03-00 liefen diese nur einmalig während der Migration von der
// alten Kombi-Datei (über alsServerstand) - seit Paket A wird diese Funktion beim
// normalen Laden aus den aufgeteilten Dateien aber nicht mehr automatisch durchlaufen,
// weil es dafür keine alte Kombi-Datei mehr zu lesen gibt. Deshalb wird
// wendeNachtraegeAn jetzt zusätzlich direkt im Ladeeffekt aufgerufen (siehe unten),
// damit die Selbstheilung wie beabsichtigt bei jedem Laden greift, nicht nur einmalig.
export function wendeNachtraegeAn(data: SchemaData): SchemaData {
  return backfillFehlendeTeilnehmer(
    backfillPersonenOhneMenge(backfillReisenOhneAnker(data), "Grimming 2026", ["Clemens", "Sonja"])
  );
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

const MAX_HISTORY = 20;

// ---- Paket A: Aufteilung in Stammdaten- + Reise-Einzeldateien ----
// Die folgenden Funktionen brauchen keinen Komponenten-Zustand und stehen deshalb auf
// Modul-Ebene. Grundidee (siehe splitSchema.ts und Konzeptdokument): Innerhalb der
// laufenden Sitzung arbeitet die Komponente weiterhin mit EINEM zusammengeführten
// `SchemaData`-Objekt wie bisher (siehe `data`/`updateData`/`undo` unten, unverändert) -
// nur beim Laden/Speichern wird das in mehrere OneDrive-Dateien auseinandergenommen bzw.
// wieder zusammengesetzt.

type OffenerFileSync =
  | { art: "stammdaten"; result: SyncResult }
  | { art: "reise"; reiseId: ID; result: SyncResult };

export function stammdatenGleich(a: StammdatenKern, b: StammdatenKern): boolean {
  return schemaEqual(stammdatenAlsSchemaData(a), stammdatenAlsSchemaData(b));
}
export function reiseGleich(a: ReiseKern, b: ReiseKern): boolean {
  return schemaEqual(reiseAlsSchemaData(a), reiseAlsSchemaData(b));
}

export async function schreibeStammdatenDatei(kern: StammdatenKern, verlauf: StammdatenKern[]): Promise<void> {
  const datei: StammdatenDatei = { ...kern, verlauf };
  await saveState(datei, STAMMDATEN_DATEI);
}

export async function schreibeReiseDatei(
  reiseId: ID,
  kern: ReiseKern,
  stammdaten: StammdatenKern,
  verlauf: ReiseKern[]
): Promise<void> {
  const datei: ReiseDatei = {
    ...kern,
    stammdaten_snapshot: stammdatenSnapshotFuerReise(stammdaten, kern),
    verlauf,
  };
  await saveState(datei, reiseDateiname(reiseId));
}

/** Schreibt einen kompletten Datenstand erstmalig/komplett neu als aufgeteilte Dateien
 *  (Migration von der alten Kombi-Datei, oder Wiederherstellung nach einem Offline-
 *  Konflikt beim allerersten Laden - siehe ladeAufgeteiltenStand/Ladeeffekt unten). */
export async function schreibeGesamtenStandNeu(
  merged: SchemaData,
  historyFuerAufteilung: SchemaData[]
): Promise<{ stammdaten: StammdatenKern; reisen: Map<ID, ReiseKern> }> {
  const { stammdaten, reisen } = splitSchemaData(merged);
  const { stammdatenVerlauf, reiseVerlaufMap } = verlaufAufteilen(historyFuerAufteilung);
  await schreibeStammdatenDatei(stammdaten, stammdatenVerlauf);
  await Promise.all(
    Array.from(reisen.entries()).map(([reiseId, kern]) =>
      schreibeReiseDatei(reiseId, kern, stammdaten, reiseVerlaufMap.get(reiseId) ?? [])
    )
  );
  return { stammdaten, reisen };
}

/** Lädt den aktuellen Stand aus den aufgeteilten Dateien. Existiert die zentrale
 *  Stammdaten-Datei noch nicht, wird einmalig aus der alten Kombi-Datei (LEGACY_SCHEMA_FILE,
 *  falls vorhanden, sonst der mitgelieferten Seed-Datei) migriert und sofort als neue
 *  Dateien geschrieben - die alte Datei bleibt dabei unangetastet liegen (Clemens löscht
 *  sie bei Gelegenheit selbst, siehe LEGACY_SCHEMA_FILE oben). */
export async function ladeAufgeteiltenStand(alsServerstand: (remote: unknown) => SchemaData): Promise<{
  stammdaten: StammdatenKern;
  reisen: Map<ID, ReiseKern>;
  stammdatenVerlauf: StammdatenKern[];
  reiseVerlaufMap: Map<ID, ReiseKern[]>;
}> {
  const dateien = await listFiles();

  if (dateien.includes(STAMMDATEN_DATEI)) {
    const stammdatenDatei = (await loadState(STAMMDATEN_DATEI)) as StammdatenDatei;
    const reiseDateinamen = dateien.filter((name) => reiseIdAusDateiname(name) !== null);
    const geladeneReisen = await Promise.all(
      reiseDateinamen.map(async (name) => ({
        reiseId: reiseIdAusDateiname(name)!,
        roh: (await loadState(name)) as ReiseDatei | null,
      }))
    );
    const reisen = new Map<ID, ReiseKern>();
    const reiseVerlaufMap = new Map<ID, ReiseKern[]>();
    const snapshotsAlsFallback: StammdatenKern[] = [];
    for (const { reiseId, roh } of geladeneReisen) {
      if (!roh) continue;
      reisen.set(reiseId, {
        reise: roh.reise,
        tk01_t01_t02: roh.tk01_t01_t02,
        tk03_tk01_t04: roh.tk03_tk01_t04,
        tk04_tk03_t05: roh.tk04_tk03_t05,
        neu_hinzugefuegt: roh.neu_hinzugefuegt ?? [],
      });
      reiseVerlaufMap.set(reiseId, roh.verlauf ?? []);
      if (roh.stammdaten_snapshot) snapshotsAlsFallback.push(roh.stammdaten_snapshot);
    }
    // Normalfall: Stammdaten-Datei ist da und wird verwendet. Nur falls sie ausnahmsweise
    // leer/kaputt wäre, aber Reise-Dateien mit Snapshots existieren, aus denen
    // notdürftig rekonstruieren (siehe splitSchema.ts) statt mit leeren Stammdaten zu starten.
    const stammdatenLeer =
      stammdatenDatei.t02_aktivitaet.length === 0 &&
      stammdatenDatei.t03_kathegorie.length === 0 &&
      stammdatenDatei.t04_gegenstand.length === 0 &&
      stammdatenDatei.t05_namen.length === 0;
    const stammdaten: StammdatenKern =
      stammdatenLeer && snapshotsAlsFallback.length > 0
        ? stammdatenAusSnapshotsRekonstruieren(snapshotsAlsFallback)
        : {
            t02_aktivitaet: stammdatenDatei.t02_aktivitaet,
            t03_kathegorie: stammdatenDatei.t03_kathegorie,
            t04_gegenstand: stammdatenDatei.t04_gegenstand,
            t05_namen: stammdatenDatei.t05_namen,
            tk02_t02_t04: stammdatenDatei.tk02_t02_t04,
          };
    return { stammdaten, reisen, stammdatenVerlauf: stammdatenDatei.verlauf ?? [], reiseVerlaufMap };
  }

  // Noch nicht aufgeteilt: alte Kombi-Datei (bzw. beim allerersten Start die Seed-Datei)
  // lesen, einmalig aufteilen und sofort als neue Dateien schreiben.
  const alt = await loadState(LEGACY_SCHEMA_FILE);
  const server = alsServerstand(alt);
  // Bisherige Array-Reihenfolge der Reisen als explizite `reihenfolge` übernehmen (siehe
  // schema.ts), damit die Reise-Reiter nach der Migration nicht plötzlich anders sortiert
  // erscheinen.
  const serverMitReihenfolge: SchemaData = {
    ...server,
    t01_reise: server.t01_reise.map((r, i) => (r.reihenfolge === undefined ? { ...r, reihenfolge: i } : r)),
  };
  // Der alte, globale Rückgängig-Verlauf wird bewusst NICHT mit übernommen (mit Clemens
  // abgestimmt, 2026-08-29: die bisherigen Schritte werden nicht mehr gebraucht) - jede
  // neue Datei startet mit einem leeren eigenen Verlauf.
  const { stammdaten, reisen } = await schreibeGesamtenStandNeu(serverMitReihenfolge, []);
  return { stammdaten, reisen, stammdatenVerlauf: [], reiseVerlaufMap: new Map() };
}

/** Vergleicht Stammdaten + alle Reisen je einzeln (baseline/lokal/remote) und liefert den
 *  Merge-Vorschlag - reine Vergleichslogik ohne Netzwerkzugriff, wiederverwendet sowohl
 *  beim allerersten Laden (Abgleich eines evtl. noch nicht synchronisierten Offline-
 *  Stands) als auch beim periodischen Sync-Tick weiter unten. */
export function vergleicheAlleDateien(
  baselineStammdaten: StammdatenKern,
  lokalStammdaten: StammdatenKern,
  remoteStammdaten: StammdatenKern,
  baselineReisen: Map<ID, ReiseKern>,
  lokalReisen: Map<ID, ReiseKern>,
  remoteReisen: Map<ID, ReiseKern>
): {
  offeneSyncs: OffenerFileSync[];
  neueStammdaten: StammdatenKern;
  neueReisen: Map<ID, ReiseKern>;
  autoMerged: AutoMergedChange[];
} {
  const offeneSyncs: OffenerFileSync[] = [];
  const autoMerged: AutoMergedChange[] = [];

  let neueStammdaten = lokalStammdaten;
  if (!stammdatenGleich(remoteStammdaten, baselineStammdaten)) {
    const result = diffAndMerge(
      stammdatenAlsSchemaData(baselineStammdaten),
      stammdatenAlsSchemaData(lokalStammdaten),
      stammdatenAlsSchemaData(remoteStammdaten)
    );
    if (result.conflicts.length > 0) {
      offeneSyncs.push({ art: "stammdaten", result });
    } else {
      neueStammdaten = schemaDataAlsStammdatenKern(result.merged);
      autoMerged.push(...result.autoMerged);
    }
  }

  const neueReisen = new Map<ID, ReiseKern>();
  const alleReiseIds = new Set<ID>([...baselineReisen.keys(), ...lokalReisen.keys(), ...remoteReisen.keys()]);
  for (const reiseId of alleReiseIds) {
    const lokal = lokalReisen.get(reiseId) ?? baselineReisen.get(reiseId) ?? remoteReisen.get(reiseId)!;
    const baseline = baselineReisen.get(reiseId) ?? lokal;
    const remote = remoteReisen.get(reiseId) ?? lokal;
    if (reiseGleich(remote, baseline)) {
      neueReisen.set(reiseId, lokal);
      continue;
    }
    const result = diffAndMerge(reiseAlsSchemaData(baseline), reiseAlsSchemaData(lokal), reiseAlsSchemaData(remote));
    if (result.conflicts.length > 0) {
      offeneSyncs.push({ art: "reise", reiseId, result });
      neueReisen.set(reiseId, lokal); // vorläufig, bis der Nutzer entschieden hat
    } else {
      neueReisen.set(reiseId, schemaDataAlsReiseKern(result.merged));
      autoMerged.push(...result.autoMerged);
    }
  }
  return { offeneSyncs, neueStammdaten, neueReisen, autoMerged };
}

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
  // Konkreter Fehlertext beim Laden - seit V03-01 sichtbar in der Fehleranzeige, damit
  // Clemens uns die genaue Meldung mitteilen kann, statt nur "Fehler beim Laden" zu sehen
  // (siehe istVerbindungsfehler in syncStore.ts: TypeErrors werden nicht mehr pauschal
  // als "offline" verschluckt, sondern hier sichtbar gemacht).
  const [loadErrorMessage, setLoadErrorMessage] = useState<string>("");
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
  const [mode, setMode] = useState<
    | "liste"
    | "bearbeiten"
    | "neueReise"
    | "neueReiseTeilnehmer"
    | "neueReiseAktivitaeten"
    | "teilnehmerBearbeiten"
    | "aktivitaetenVerwalten"
  >("liste");
  const [editSearch, setEditSearch] = useState("");
  const [neuReiseName, setNeuReiseName] = useState("");
  const [neuReiseVon, setNeuReiseVon] = useState("");
  const [neuReiseBis, setNeuReiseBis] = useState("");
  // Teilnehmer-Auswahl (V03-02, von Clemens gewünscht 2026-08-30): wird sowohl beim
  // Anlegen einer neuen Reise (Modus "neueReiseTeilnehmer") als auch beim nachträglichen
  // Ändern einer bestehenden Reise (Modus "teilnehmerBearbeiten") wiederverwendet.
  const [teilnehmerAuswahl, setTeilnehmerAuswahl] = useState<Set<ID>>(new Set());
  // Name-Eingabe für "+ neue Person" (V03-03, von Clemens gewünscht 2026-08-31) - direkt
  // in der Teilnehmer-Auswahl nutzbar, egal ob beim Anlegen oder beim Bearbeiten.
  const [neuPersonName, setNeuPersonName] = useState("");
  // Sicherheitsabfrage für "Reise löschen" (V03-02) - eigener Bestätigungsschritt statt
  // eines nativen Browser-Dialogs, damit es sich in den Rest der Oberfläche einfügt.
  const [reiseLoeschenBestaetigen, setReiseLoeschenBestaetigen] = useState(false);
  // Aktivitäten-Auswahl beim Anlegen einer neuen Reise (V03-03, dritter Schritt nach
  // Teilnehmern) - steuert, welche Standard-Gegenstände automatisch vorgeschlagen werden
  // (siehe erstelleNeueReise unten). "Basics" ist immer automatisch dabei und taucht
  // deshalb hier nicht als Auswahlmöglichkeit auf.
  const [aktivitaetenAuswahl, setAktivitaetenAuswahl] = useState<Set<ID>>(new Set());
  // Verwaltung "Standard-Gegenstände je Aktivität" (V03-03, eigener Bereich, reiseunabhängig
  // wie die übrigen Stammdaten) - welche Aktivität gerade bearbeitet wird, plus Suchfeld und
  // Eingabe für eine neu anzulegende Aktivität.
  const [verwaltungAktivitaetId, setVerwaltungAktivitaetId] = useState<ID | null>(null);
  const [verwaltungSearch, setVerwaltungSearch] = useState("");
  const [neuAktivitaetName, setNeuAktivitaetName] = useState("");
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
  // Seit Paket A zusätzlich: der letzte je Datei mit dem Server abgeglichene Stand
  // (Stammdaten-Datei bzw. je eine Reise-Datei) - Grundlage für den Zeilen-für-Zeile-
  // Vergleich pro Datei (siehe vergleicheAlleDateien oben). `baselineRef` bleibt daneben
  // bestehen und wirkt weiterhin auf den zusammengeführten Gesamtstand, damit der
  // schnelle "habe ich überhaupt etwas geändert"-Check unten unverändert bleibt.
  const stammdatenBaselineRef = useRef<StammdatenKern | null>(null);
  const reiseBaselinesRef = useRef<Map<ID, ReiseKern>>(new Map());
  // Die noch offenen Konflikt-Abgleiche (kann mehrere Dateien gleichzeitig betreffen,
  // z.B. Stammdaten UND eine Reise), solange der Nutzer über das Konflikt-Fenster noch
  // nicht entschieden hat.
  const pendingSyncResultsRef = useRef<OffenerFileSync[]>([]);
  const [conflicts, setConflicts] = useState<RowConflict[]>([]);
  const [autoMerged, setAutoMerged] = useState<AutoMergedChange[]>([]);
  const [showAutoMerged, setShowAutoMerged] = useState(false);
  const [offline, setOffline] = useState(false);
  const [retryTick, setRetryTick] = useState(0);

  function alsServerstand(remote: unknown): SchemaData {
    const merged = remote
      ? mergeSeedInto(remote as SchemaData, seedData as unknown as SchemaData)
      : (seedData as unknown as SchemaData);
    return wendeNachtraegeAn(merged);
  }

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        // Lädt bzw. migriert einmalig die Stammdaten-Datei + je eine Reise-Datei (Paket A,
        // siehe ladeAufgeteiltenStand oben). Kein Fallback mehr auf die alte Kombi-Datei
        // danach - die bleibt nur noch als einmalige Migrationsquelle relevant.
        const geladen = await ladeAufgeteiltenStand(alsServerstand);
        if (cancelled) return;
        const server = mergeSplitData(geladen.stammdaten, Array.from(geladen.reisen.values()));
        const historyGeladen = historyRekonstruieren(
          geladen.stammdatenVerlauf,
          geladen.reiseVerlaufMap,
          Array.from(geladen.reisen.values())
        );

        // Prüfen, ob von einer früheren Offline-Sitzung noch ein nicht synchronisierter
        // Stand im Browser liegt (z.B. App wurde ohne Verbindung geschlossen, bevor die
        // Änderung nach OneDrive geschrieben werden konnte). Arbeitet weiterhin auf dem
        // zusammengeführten Gesamtstand (unverändert, wie vor Paket A).
        const pending = readPending(LOKALER_CACHE_SCHLUESSEL);
        const storedBaseline = readBaseline(LOKALER_CACHE_SCHLUESSEL);
        const pendingIstNeu = pending && (!storedBaseline || JSON.stringify(pending) !== JSON.stringify(storedBaseline));

        if (pendingIstNeu && pending) {
          const { stammdaten: pendingStammdaten, reisen: pendingReisen } = splitSchemaData(pending);
          const baselineVoll = storedBaseline ?? server;
          const { stammdaten: baselineStammdaten, reisen: baselineReisen } = splitSchemaData(baselineVoll);
          const { offeneSyncs, neueStammdaten, neueReisen, autoMerged } = vergleicheAlleDateien(
            baselineStammdaten,
            pendingStammdaten,
            geladen.stammdaten,
            baselineReisen,
            pendingReisen,
            geladen.reisen
          );
          if (offeneSyncs.length > 0) {
            pendingSyncResultsRef.current = offeneSyncs;
            setConflicts(offeneSyncs.flatMap((s) => s.result.conflicts));
            setAutoMerged(offeneSyncs.flatMap((s) => s.result.autoMerged));
            const vorlaeufig = mergeSplitData(neueStammdaten, Array.from(neueReisen.values()));
            setData(vorlaeufig);
            setSelectedReiseId(ermittleStartReise(vorlaeufig));
          } else {
            // Kein Konflikt - Offline-Stand ist jetzt maßgeblich, wird gleich als
            // aufgeteilte Dateien geschrieben (frischer eigener Verlauf je Datei, siehe
            // ladeAufgeteiltenStand oben).
            const merged = mergeSplitData(neueStammdaten, Array.from(neueReisen.values()));
            const geschrieben = await schreibeGesamtenStandNeu(merged, []);
            stammdatenBaselineRef.current = geschrieben.stammdaten;
            reiseBaselinesRef.current = geschrieben.reisen;
            baselineRef.current = merged;
            writeBaseline(LOKALER_CACHE_SCHLUESSEL, merged);
            clearPending(LOKALER_CACHE_SCHLUESSEL);
            setData(merged);
            setSelectedReiseId(ermittleStartReise(merged));
            if (autoMerged.length > 0) {
              setAutoMerged(autoMerged);
              setShowAutoMerged(true);
            }
          }
        } else {
          // Nachträge (Selbstheilung historisch unvollständiger Daten, siehe
          // wendeNachtraegeAn oben) laufen seit V03-01 bei JEDEM normalen Laden, nicht
          // nur einmalig bei der Migration - sonst würde z.B. eine Reise ohne "Basics"-
          // Anker (siehe China 2025) für immer kaputt bleiben, weil die alte Kombi-Datei
          // nach der Migration nie wieder gelesen wird.
          const serverMitNachtraegen = wendeNachtraegeAn(server);
          if (!schemaEqual(serverMitNachtraegen, server)) {
            // Ein Nachtrag hat tatsächlich etwas ergänzt - gleich mitschreiben, sonst
            // würde die Ergänzung nur lokal existieren und beim nächsten Laden auf einem
            // anderen Gerät wieder fehlen.
            const geschrieben = await schreibeGesamtenStandNeu(serverMitNachtraegen, historyGeladen);
            stammdatenBaselineRef.current = geschrieben.stammdaten;
            reiseBaselinesRef.current = geschrieben.reisen;
          } else {
            stammdatenBaselineRef.current = geladen.stammdaten;
            reiseBaselinesRef.current = geladen.reisen;
          }
          baselineRef.current = serverMitNachtraegen;
          writeBaseline(LOKALER_CACHE_SCHLUESSEL, serverMitNachtraegen);
          clearPending(LOKALER_CACHE_SCHLUESSEL);
          setData(serverMitNachtraegen);
          setSelectedReiseId(ermittleStartReise(serverMitNachtraegen));
        }
        setHistory(historyGeladen);
        setOffline(false);
        setLoadStatus("ready");
      } catch (error) {
        if (cancelled) return;
        console.error(error);
        if (istVerbindungsfehler(error)) {
          // Ohne Verbindung: mit dem letzten lokal gemerkten (zusammengeführten) Stand
          // weiterarbeiten, falls vorhanden (offline gemachte Änderungen zuerst, sonst der
          // letzte erfolgreich synchronisierte Stand). Die Datei-Baselines bleiben in
          // diesem Fall leer - der nächste erfolgreiche Sync-Tick holt sie automatisch nach.
          const lokal = readPending(LOKALER_CACHE_SCHLUESSEL) ?? readBaseline(LOKALER_CACHE_SCHLUESSEL);
          if (lokal) {
            baselineRef.current = readBaseline(LOKALER_CACHE_SCHLUESSEL) ?? lokal;
            setData(lokal);
            setSelectedReiseId(ermittleStartReise(lokal));
            setHistory([]);
            setOffline(true);
            setLoadStatus("ready");
            return;
          }
        }
        setLoadErrorMessage(error instanceof Error ? error.message : String(error));
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
    // ohne unnötig nach OneDrive zu schreiben. Bleibt unverändert auf dem zusammen-
    // geführten Gesamtstand - der eigentliche Abgleich läuft seit Paket A darunter je
    // Datei (Stammdaten + jede Reise einzeln, siehe vergleicheAlleDateien oben).
    const habenWirWasZuSpeichern = !baselineRef.current || !schemaEqual(data, baselineRef.current);
    if (habenWirWasZuSpeichern) {
      setSaveStatus(offline ? "Offline – Änderung wird lokal gemerkt …" : "Änderungen werden gespeichert …");
    }
    const t = setTimeout(async () => {
      // Immer zuerst lokal merken, damit bei einem Absturz/Schließen mitten im Offline-
      // Betrieb nichts verloren geht (siehe syncStore.ts) - unabhängig davon, ob der
      // anschließende Speicherversuch klappt.
      if (habenWirWasZuSpeichern) writePending(LOKALER_CACHE_SCHLUESSEL, data);
      try {
        const jetzt = () =>
          new Date().toLocaleTimeString("de-AT", { hour: "2-digit", minute: "2-digit", second: "2-digit" });

        const { stammdaten: lokalStammdaten, reisen: lokalReisenMap } = splitSchemaData(data);
        const { stammdatenVerlauf, reiseVerlaufMap } = verlaufAufteilen(history);

        const stammdatenRemoteRoh = (await loadState(STAMMDATEN_DATEI)) as StammdatenDatei | null;
        const remoteStammdaten: StammdatenKern = stammdatenRemoteRoh ?? lokalStammdaten;

        const alleReiseIds = new Set<ID>([...reiseBaselinesRef.current.keys(), ...lokalReisenMap.keys()]);
        const remoteReisenEintraege = await Promise.all(
          Array.from(alleReiseIds).map(async (reiseId) => {
            const roh = (await loadState(reiseDateiname(reiseId))) as ReiseDatei | null;
            return [reiseId, roh] as const;
          })
        );
        const remoteReisenMap = new Map<ID, ReiseKern>();
        for (const [reiseId, roh] of remoteReisenEintraege) if (roh) remoteReisenMap.set(reiseId, roh);

        const { offeneSyncs, neueStammdaten, neueReisen, autoMerged } = vergleicheAlleDateien(
          stammdatenBaselineRef.current ?? lokalStammdaten,
          lokalStammdaten,
          remoteStammdaten,
          reiseBaselinesRef.current,
          lokalReisenMap,
          remoteReisenMap
        );

        if (offeneSyncs.length > 0) {
          pendingSyncResultsRef.current = offeneSyncs;
          setConflicts(offeneSyncs.flatMap((s) => s.result.conflicts));
          setAutoMerged(offeneSyncs.flatMap((s) => s.result.autoMerged));
          setData(mergeSplitData(neueStammdaten, Array.from(neueReisen.values())));
          setSaveStatus("Es gibt widersprüchliche Änderungen von einem anderen Gerät – bitte entscheiden.");
          return;
        }

        // Welche Dateien tatsächlich geschrieben werden müssen: entweder weil sie remote
        // noch gar nicht existieren (z.B. gerade erst angelegte Reise) oder weil sich der
        // lokale Stand gegenüber der letzten eigenen Baseline geändert hat.
        const stammdatenSchreiben =
          !stammdatenRemoteRoh || !stammdatenGleich(neueStammdaten, stammdatenBaselineRef.current ?? lokalStammdaten);
        const reisenSchreiben = Array.from(alleReiseIds).filter((reiseId) => {
          const neuerKern = neueReisen.get(reiseId);
          if (!neuerKern) return false;
          const baselineKern = reiseBaselinesRef.current.get(reiseId);
          return !remoteReisenMap.has(reiseId) || !baselineKern || !reiseGleich(neuerKern, baselineKern);
        });

        if (stammdatenSchreiben) await schreibeStammdatenDatei(neueStammdaten, stammdatenVerlauf);
        await Promise.all(
          reisenSchreiben.map((reiseId) =>
            schreibeReiseDatei(reiseId, neueReisen.get(reiseId)!, neueStammdaten, reiseVerlaufMap.get(reiseId) ?? [])
          )
        );
        const irgendwasGeschrieben = stammdatenSchreiben || reisenSchreiben.length > 0;

        stammdatenBaselineRef.current = neueStammdaten;
        reiseBaselinesRef.current = neueReisen;
        setOffline(false);

        if (!irgendwasGeschrieben && autoMerged.length === 0) return; // reiner Leerlauf-Tick

        if (habenWirWasZuSpeichern) clearPending(LOKALER_CACHE_SCHLUESSEL);

        if (autoMerged.length > 0) {
          // Es ist etwas von einem anderen Gerät dazugekommen - `data` muss das jetzt
          // widerspiegeln.
          const neuerGesamtstand = mergeSplitData(neueStammdaten, Array.from(neueReisen.values()));
          baselineRef.current = neuerGesamtstand;
          writeBaseline(LOKALER_CACHE_SCHLUESSEL, neuerGesamtstand);
          setData(neuerGesamtstand);
          setAutoMerged(autoMerged);
          setShowAutoMerged(true);
        } else {
          // Nur die eigene(n) Änderung(en) geschrieben, kein fremder Inhalt dazugekommen -
          // `data` entspricht bereits dem geschriebenen Stand, kein setData nötig (würde
          // sonst unnötig einen Re-Render/erneuten Effekt-Durchlauf auslösen).
          baselineRef.current = data;
          writeBaseline(LOKALER_CACHE_SCHLUESSEL, data);
        }
        if (irgendwasGeschrieben) setSaveStatus(`Gespeichert um ${jetzt()}`);
      } catch (error) {
        console.error(error);
        if (istVerbindungsfehler(error)) {
          setOffline(true);
          if (habenWirWasZuSpeichern) {
            setSaveStatus("Offline – Änderungen werden gespeichert, sobald wieder online.");
          }
        } else if (habenWirWasZuSpeichern) {
          const meldung = error instanceof Error ? error.message : String(error);
          setSaveStatus(`Speichern fehlgeschlagen: ${meldung}`);
        }
      }
    }, 600);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, history, conflicts.length, offline, retryTick]);

  // Nutzer hat für jede Konflikt-Zeile einzeln entschieden (Konflikt-Fenster) - Ergebnis
  // einarbeiten und speichern. Kann seit Paket A mehrere Dateien gleichzeitig betreffen
  // (z.B. ein Stammdaten- UND ein Reise-Konflikt) - alle in pendingSyncResultsRef
  // gesammelten Abgleiche werden hier gemeinsam aufgelöst.
  async function konflikteEntschieden(resolutions: Map<string, "local" | "remote">) {
    const offeneSyncs = pendingSyncResultsRef.current;
    if (offeneSyncs.length === 0 || !data) return;
    const { stammdatenVerlauf, reiseVerlaufMap } = verlaufAufteilen(history);
    let neueStammdaten = stammdatenBaselineRef.current ?? schemaDataAlsStammdatenKern(data);
    const neueReisen = new Map<ID, ReiseKern>(reiseBaselinesRef.current);
    const irgendeinAutoMerge = offeneSyncs.some((s) => s.result.autoMerged.length > 0);
    try {
      for (const sync of offeneSyncs) {
        const finalData = applyConflictResolutions(sync.result, resolutions);
        if (sync.art === "stammdaten") {
          neueStammdaten = schemaDataAlsStammdatenKern(finalData);
        } else {
          neueReisen.set(sync.reiseId, schemaDataAlsReiseKern(finalData));
        }
      }
      await schreibeStammdatenDatei(neueStammdaten, stammdatenVerlauf);
      await Promise.all(
        offeneSyncs
          .filter((s): s is { art: "reise"; reiseId: ID; result: SyncResult } => s.art === "reise")
          .map((s) =>
            schreibeReiseDatei(s.reiseId, neueReisen.get(s.reiseId)!, neueStammdaten, reiseVerlaufMap.get(s.reiseId) ?? [])
          )
      );

      pendingSyncResultsRef.current = [];
      stammdatenBaselineRef.current = neueStammdaten;
      reiseBaselinesRef.current = neueReisen;
      const finalDataGesamt = mergeSplitData(neueStammdaten, Array.from(neueReisen.values()));
      baselineRef.current = finalDataGesamt;
      writeBaseline(LOKALER_CACHE_SCHLUESSEL, finalDataGesamt);
      clearPending(LOKALER_CACHE_SCHLUESSEL);
      setConflicts([]);
      setOffline(false);
      setData(finalDataGesamt);
      const now = new Date().toLocaleTimeString("de-AT", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
      setSaveStatus(`Gespeichert um ${now}`);
      if (irgendeinAutoMerge) setShowAutoMerged(true);
    } catch (error) {
      console.error(error);
      if (istVerbindungsfehler(error)) {
        // Entscheidung lokal übernehmen und merken, wird gespeichert sobald wieder online.
        pendingSyncResultsRef.current = [];
        setConflicts([]);
        setOffline(true);
        const finalDataGesamt = mergeSplitData(neueStammdaten, Array.from(neueReisen.values()));
        setData(finalDataGesamt);
        writePending(LOKALER_CACHE_SCHLUESSEL, finalDataGesamt);
        setSaveStatus("Offline – Entscheidung gemerkt, wird gespeichert sobald wieder online.");
      } else {
        const meldung = error instanceof Error ? error.message : String(error);
        setSaveStatus(`Speichern fehlgeschlagen: ${meldung} – bitte erneut versuchen.`);
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

  // Personen-Tabs: seit V03-02 primär aus reise.teilnehmer (explizite Auswahl beim
  // Anlegen der Reise bzw. über "Teilnehmer" bearbeitet) - das ist jetzt die
  // maßgebliche Liste, unabhängig davon, ob schon Mengen eingetragen sind. Nur wenn
  // eine Reise (noch) kein teilnehmer-Feld hat oder es leer ist (alte Reise vor
  // V03-02, noch nicht durch backfillFehlendeTeilnehmer abgedeckt, oder eine ganz neue
  // Reise ohne jede Zuordnung), wird auf die alte Logik zurückgefallen: erst
  // beteiligtePersonen (wer tatsächlich schon Gegenstände/Mengen hat), sonst alle
  // bekannten Personen (siehe Grimming-Fix V01-27).
  const sichtbarePersonen = useMemo(() => {
    if (!data || !reise) return [];
    if (reise.teilnehmer && reise.teilnehmer.length > 0) {
      const ids = new Set(reise.teilnehmer);
      return data.t05_namen.filter((n) => ids.has(n.id));
    }
    return beteiligtePersonen.length > 0 ? beteiligtePersonen : data.t05_namen;
  }, [data, reise, beteiligtePersonen]);

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

  // Sicherheitsabfrage für "Reise löschen" beim Wechsel der Reise wieder zurücksetzen -
  // sonst könnte man sich versehentlich auf einer anderen Reise befinden, während die
  // Bestätigung noch "scharf" ist (V03-02).
  useEffect(() => {
    setReiseLoeschenBestaetigen(false);
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
      // V03-05 (Clemens, 2026-09-03): Beim Ändern auf "–" (Strich) in der "Neu
      // hinzugefügt"-Ansicht verschwand der Gegenstand trotz V03-04 sofort - die
      // allgemeine Strich-Ausblendung griff hier noch VOR der "neu"-Prüfung. Solange ein
      // noch nicht per "Fertig" bestätigter Gegenstand in dieser Ansicht offen ist, zählt
      // die Strich-Ausblendung deshalb nicht - man soll auch einen versehentlichen Strich
      // noch sehen und korrigieren können. Außerhalb dieser Ansicht bleibt "Strich = raus
      // aus der Liste" unverändert.
      const nochOffenAlsNeu = offenFilter === "neu" && neuHinzugefuegt.has(row.id);
      if (row.ausgewaehlt === -1 && !nochOffenAlsNeu) continue; // Strich = sicher nicht mitgenommen, aus der Liste raus
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
      // V03-05: nicht mehr row.ausgewaehlt !== -1 ausschließen - ein auf "Strich" gesetzter,
      // aber noch nicht per "Fertig" bestätigter Gegenstand zählt jetzt mit (siehe gruppiert
      // oben), damit die Zahl am Chip zur tatsächlich angezeigten Liste passt.
      if (row && neuHinzugefuegt.has(row.id)) n++;
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
      // Bis V03-03: Sobald an einer frisch hinzugefügten Zeile etwas eingetragen wurde,
      // verschwand sie sofort aus dem "Neu hinzugefügt"-Filter. Clemens meldete (2026-09-03):
      // genau beim Anpassen der Menge in dieser Ansicht ist das verwirrend - man sieht nicht
      // mehr, ob der Klick richtig war, und kann nicht mehr mehrere Positionen hintereinander
      // in Ruhe korrigieren. Seit V03-04 bleibt die "neu"-Markierung deshalb beim Antippen
      // bestehen und fällt erst gesammelt über den "Fertig"-Knopf weg (siehe
      // bestaetigeNeuHinzugefuegt()).
      return { ...prev, tk04_tk03_t05: updated };
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
        // Seit V03-04 wie bei bumpField: Die "neu"-Markierung bleibt beim Bearbeiten
        // bestehen, fällt erst gesammelt über den "Fertig"-Knopf weg.
      };
    });
    setEditingQty(null);
    setEditingQtyValue("");
  }

  // Schließt die "Neu hinzugefügt"-Ansicht der aktuellen Person bewusst ab: alle dort
  // gerade sichtbaren (also noch als "neu" markierten) Zeilen verlieren gemeinsam die
  // Markierung, die Ansicht springt danach zurück zur normalen Liste. Ersetzt seit V03-04
  // das automatische, sofortige Verschwinden einzelner Zeilen beim bloßen Antippen einer
  // Menge (siehe bumpField/speichereQtyEingabe) - auf ausdrücklichen Wunsch von Clemens
  // (2026-09-03): "nicht gleich aus der Liste nehmen, sondern einen Fertig-Knopf dazugeben,
  // so kann man die Liste auch noch korrigieren, bis alles fertig ist".
  function bestaetigeNeuHinzugefuegt() {
    if (!data) return;
    const idsInAnsicht = new Set<string>();
    for (const [, entries] of gruppiert) {
      for (const { row } of entries) idsInAnsicht.add(row.id);
    }
    if (idsInAnsicht.size === 0) return;
    updateData((prev) => ({
      ...prev,
      neu_hinzugefuegt: (prev.neu_hinzugefuegt ?? []).filter((id) => !idsInAnsicht.has(id)),
    }));
    setOffenFilter(null);
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

  // Zu welcher Reise eine tk03-Zeile gehört (Umkehrung von gegenstaendeFuerReise) - seit
  // V03-03 gebraucht, um beim Hinzufügen einer neuen Teilnehmerin nur Clemens' Zeilen
  // DIESER Reise zu kopieren, nicht die aus anderen Reisen.
  function reiseIdFuerTk03(d: SchemaData, tk03Id: ID): ID | null {
    const tk03 = d.tk03_tk01_t04.find((r) => r.id === tk03Id);
    if (!tk03) return null;
    const tk01 = d.tk01_t01_t02.find((r) => r.id === tk03.id_tk01);
    return tk01?.id_t01 ?? null;
  }

  // Teilnehmer und Aktivitäten werden seit V03-02/V03-03 als eigene Schritte VOR dem
  // Anlegen abgefragt (siehe Modi "neueReiseTeilnehmer"/"neueReiseAktivitaeten" unten) -
  // die ID wird deshalb schon hier (außerhalb von updateData) erzeugt, damit sie direkt
  // danach synchron für setSelectedReiseId verwendet werden kann (updateData selbst gibt
  // nichts zurück).
  function erstelleNeueReise(teilnehmerIds: ID[], aktivitaetIds: ID[]) {
    if (!data || !neuReiseName.trim()) return;
    const ids = collectAllIds(data);
    const reiseId = newId("t01", ids);
    ids.add(reiseId);
    // Reihenfolge der Reise-Reiter (seit Paket A, siehe schema.ts/splitSchema.ts) -
    // neue Reise kommt hinten an.
    const naechsteReihenfolge =
      data.t01_reise.reduce((max, r) => Math.max(max, r.reihenfolge ?? 0), 0) + 1;
    const basics = data.t02_aktivitaet.find((a) => a.aktivitaet === "Basics");
    const neu: T01Reise = {
      id: reiseId,
      reise: neuReiseName.trim(),
      von: neuReiseVon || null,
      bis: neuReiseBis || null,
      notiz: "",
      reihenfolge: naechsteReihenfolge,
      // Explizit gewählte Teilnehmer (V03-02) - von Clemens ausdrücklich als erster
      // Schritt gewünscht (2026-08-30), damit nicht mehr automatisch alle bekannten
      // Personen angeboten werden.
      teilnehmer: teilnehmerIds,
    };

    // Aktivitäten (V03-03): "Basics" ist wie bisher immer automatisch dabei, dazu alle im
    // dritten Schritt ausgewählten Aktivitäten - jede bekommt eine eigene tk01-Zeile.
    const aktivitaetenFuerReise = [
      ...(basics ? [basics.id] : []),
      ...aktivitaetIds.filter((id) => id !== basics?.id),
    ];
    const tk01Neu: Tk01ReiseAktivitaet[] = aktivitaetenFuerReise.map((aktId) => {
      const tk01Id = newId("tk01", ids);
      ids.add(tk01Id);
      return { id: tk01Id, id_t01: reiseId, id_t02: aktId };
    });

    // Standard-Gegenstände der gewählten Aktivitäten automatisch zur Liste hinzufügen und
    // den anfänglichen Teilnehmern zuweisen (Status "0 - unsicher", genau wie beim
    // manuellen "+" in "Liste bearbeiten") - von Clemens ausdrücklich gewünscht
    // (2026-08-31), siehe tk02_t02_t04 / "Standard-Gegenstände verwalten" unten. Ein
    // Gegenstand, der in mehreren gewählten Aktivitäten Standard ist, wird trotzdem nur
    // einmal auf die Liste gesetzt (bereitsTk03-Prüfung).
    const tk03Neu: Tk03ReiseaktivitaetGegenstand[] = [];
    const tk04Neu: Tk04GegenstandPerson[] = [];
    const neuHinzugefuegtNeu: ID[] = [];
    const bereitsTk03 = new Set<ID>();
    for (const tk01 of tk01Neu) {
      const standardItems = data.tk02_t02_t04
        .filter((r) => r.id_t02 === tk01.id_t02)
        .map((r) => r.id_t04);
      for (const gegenstandId of standardItems) {
        if (bereitsTk03.has(gegenstandId)) continue;
        bereitsTk03.add(gegenstandId);
        const tk03Id = newId("tk03", ids);
        ids.add(tk03Id);
        tk03Neu.push({ id: tk03Id, id_tk01: tk01.id, id_t04: gegenstandId, notiz: "Vorschlag aus Aktivität" });
        for (const personId of teilnehmerIds) {
          const tk04Id = newId("tk04", ids);
          ids.add(tk04Id);
          tk04Neu.push({
            id: tk04Id,
            id_tk03: tk03Id,
            id_t05: personId,
            ausgewaehlt: 0,
            hergerichtet: null,
            eingepackt: null,
            verwendet: null,
          });
          neuHinzugefuegtNeu.push(tk04Id);
        }
      }
    }

    updateData((prev) => ({
      ...prev,
      t01_reise: [...prev.t01_reise, neu],
      tk01_t01_t02: [...prev.tk01_t01_t02, ...tk01Neu],
      tk03_tk01_t04: [...prev.tk03_tk01_t04, ...tk03Neu],
      tk04_tk03_t05: [...prev.tk04_tk03_t05, ...tk04Neu],
      neu_hinzugefuegt: [...(prev.neu_hinzugefuegt ?? []), ...neuHinzugefuegtNeu],
    }));
    setNeuReiseName("");
    setNeuReiseVon("");
    setNeuReiseBis("");
    setTeilnehmerAuswahl(new Set());
    setAktivitaetenAuswahl(new Set());
    setSelectedReiseId(reiseId);
    setMode("bearbeiten");
  }

  function toggleTeilnehmer(personId: ID) {
    setTeilnehmerAuswahl((prev) => {
      const next = new Set(prev);
      if (next.has(personId)) next.delete(personId);
      else next.add(personId);
      return next;
    });
  }

  function toggleAktivitaet(aktivitaetId: ID) {
    setAktivitaetenAuswahl((prev) => {
      const next = new Set(prev);
      if (next.has(aktivitaetId)) next.delete(aktivitaetId);
      else next.add(aktivitaetId);
      return next;
    });
  }

  // Neue Person anlegen (V03-03, von Clemens gewünscht 2026-08-31) - direkt aus der
  // Teilnehmer-Auswahl heraus nutzbar (Anlegen einer Reise wie auch nachträgliches
  // Bearbeiten). Gibt die neue ID zurück, damit der Aufrufer sie sofort ankreuzen kann.
  function erstelleNeuePerson(name: string): ID | null {
    if (!data || !name.trim()) return null;
    const ids = collectAllIds(data);
    const neuId = newId("t05", ids);
    updateData((prev) => ({
      ...prev,
      t05_namen: [...prev.t05_namen, { id: neuId, namen: name.trim(), notiz: "" }],
    }));
    return neuId;
  }

  // Teilnehmerliste einer bestehenden Reise übernehmen (Modus "teilnehmerBearbeiten",
  // vorbelegt über den "Teilnehmer"-Button in der Werkzeugleiste unten). Entfernte
  // Personen verlieren dadurch NICHT ihre bereits eingetragenen Mengen/Zuordnungen (auf
  // Wunsch von Clemens, 2026-08-30) - sie werden nur ausgeblendet, weil sichtbarePersonen
  // oben ausschließlich reise.teilnehmer zeigt. Die Daten bleiben in tk04 erhalten und
  // erscheinen sofort wieder, sobald die Person erneut hinzugefügt wird.
  //
  // Seit V03-03 zusätzlich: Wer hier wirklich NEU zur Reise dazukommt (noch nie eine
  // eigene tk04-Zeile in dieser Reise hatte, also nicht nur wieder sichtbar gemacht wird),
  // bekommt automatisch Clemens' aktuelle Gegenstände dieser Reise einmalig zugewiesen
  // (gleiche Menge wie bei Clemens, aber ohne dessen Herrichten/Einpacken/Verwendet-
  // Fortschritt) - von Clemens ausdrücklich so gewünscht (2026-08-31), damit die neue
  // Person nicht bei null anfängt, sondern nur noch anpasst, was sie mehr oder weniger
  // braucht.
  function speichereTeilnehmer() {
    if (!data || !reise) return;
    const teilnehmerIds = Array.from(teilnehmerAuswahl);
    const reiseId = reise.id;
    const vorherigeIds = new Set(reise.teilnehmer ?? []);
    const neuHinzugekommen = teilnehmerIds.filter((id) => !vorherigeIds.has(id));

    const ids = collectAllIds(data);
    const clemens = data.t05_namen.find((p) => p.namen === "Clemens");
    const tk04Kopien: Tk04GegenstandPerson[] = [];
    const neuHinzugefuegtKopien: ID[] = [];
    if (clemens && neuHinzugekommen.length > 0) {
      const clemensZeilenDieserReise = data.tk04_tk03_t05.filter(
        (r) => r.id_t05 === clemens.id && r.ausgewaehlt !== -1 && reiseIdFuerTk03(data, r.id_tk03) === reiseId
      );
      for (const personId of neuHinzugekommen) {
        if (personId === clemens.id) continue;
        // Nur kopieren, wenn diese Person für diese Reise wirklich noch nie etwas hatte -
        // sonst würde eine wieder hinzugefügte (nur ausgeblendete) Person ihre bisherigen
        // Einträge doppelt bzw. überschrieben bekommen.
        const hatSchonEintraege = data.tk04_tk03_t05.some(
          (r) => r.id_t05 === personId && reiseIdFuerTk03(data, r.id_tk03) === reiseId
        );
        if (hatSchonEintraege) continue;
        for (const z of clemensZeilenDieserReise) {
          const tk04Id = newId("tk04", ids);
          ids.add(tk04Id);
          tk04Kopien.push({
            id: tk04Id,
            id_tk03: z.id_tk03,
            id_t05: personId,
            ausgewaehlt: z.ausgewaehlt,
            hergerichtet: null,
            eingepackt: null,
            verwendet: null,
          });
          neuHinzugefuegtKopien.push(tk04Id);
        }
      }
    }

    updateData((prev) => ({
      ...prev,
      t01_reise: prev.t01_reise.map((r) => (r.id === reiseId ? { ...r, teilnehmer: teilnehmerIds } : r)),
      tk04_tk03_t05: [...prev.tk04_tk03_t05, ...tk04Kopien],
      neu_hinzugefuegt: [...(prev.neu_hinzugefuegt ?? []), ...neuHinzugefuegtKopien],
    }));
    setMode("liste");
  }

  // Neue Aktivität anlegen (V03-03) - z.B. "SUP", wie von Clemens erwähnt. Nutzbar sowohl
  // im dritten Schritt beim Anlegen einer Reise als auch im Verwaltungsbereich "Standard-
  // Gegenstände" unten.
  function erstelleNeueAktivitaet() {
    if (!data || !neuAktivitaetName.trim()) return;
    const ids = collectAllIds(data);
    const neuId = newId("t02", ids);
    updateData((prev) => ({
      ...prev,
      t02_aktivitaet: [...prev.t02_aktivitaet, { id: neuId, aktivitaet: neuAktivitaetName.trim(), notiz: "" }],
    }));
    setNeuAktivitaetName("");
    setVerwaltungAktivitaetId(neuId);
  }

  // Standard-Gegenstand für eine Aktivität an-/abwählen (V03-03, Verwaltungsbereich
  // "Standard-Gegenstände") - steuert, was beim Anlegen einer neuen Reise mit dieser
  // Aktivität automatisch vorgeschlagen wird (siehe erstelleNeueReise oben).
  function toggleStandardGegenstand(aktivitaetId: ID, gegenstandId: ID) {
    updateData((prev) => {
      const bestehende = prev.tk02_t02_t04.find(
        (r) => r.id_t02 === aktivitaetId && r.id_t04 === gegenstandId
      );
      if (bestehende) {
        return { ...prev, tk02_t02_t04: prev.tk02_t02_t04.filter((r) => r.id !== bestehende.id) };
      }
      const ids = collectAllIds(prev);
      const neuId = newId("tk02", ids);
      return {
        ...prev,
        tk02_t02_t04: [...prev.tk02_t02_t04, { id: neuId, id_t02: aktivitaetId, id_t04: gegenstandId }],
      };
    });
  }

  // Reise unwiderruflich löschen (V03-02, von Clemens ausdrücklich als echte Löschung
  // gewünscht - bewusst eine Ausnahme vom sonstigen "nie löschen, nur ausblenden"-Prinzip
  // dieser App). Wird erst nach der Sicherheitsabfrage (reiseLoeschenBestaetigen) über den
  // Button unten aufgerufen.
  async function loescheReise() {
    if (!data || !reise) return;
    const reiseId = reise.id;
    const dateiname = reiseDateiname(reiseId);
    setReiseLoeschenBestaetigen(false);
    // Baseline sofort (synchron) entfernen, damit der Autospeicher-Effekt oben diese
    // Reise-Datei nicht versehentlich wiederherstellt, falls er zwischendurch mit noch
    // altem Stand anläuft, während das Löschen in OneDrive noch unterwegs ist.
    reiseBaselinesRef.current.delete(reiseId);
    const verbleibendeReisen = data.t01_reise.filter((r) => r.id !== reiseId);
    const naechsteReiseId = verbleibendeReisen[0]?.id ?? null;
    updateData((prev) => {
      const tk01Ids = new Set(
        prev.tk01_t01_t02.filter((r) => r.id_t01 === reiseId).map((r) => r.id)
      );
      const tk03Ids = new Set(
        prev.tk03_tk01_t04.filter((r) => tk01Ids.has(r.id_tk01)).map((r) => r.id)
      );
      return {
        ...prev,
        t01_reise: prev.t01_reise.filter((r) => r.id !== reiseId),
        tk01_t01_t02: prev.tk01_t01_t02.filter((r) => r.id_t01 !== reiseId),
        tk03_tk01_t04: prev.tk03_tk01_t04.filter((r) => !tk01Ids.has(r.id_tk01)),
        tk04_tk03_t05: prev.tk04_tk03_t05.filter((r) => !tk03Ids.has(r.id_tk03)),
      };
    });
    setSelectedReiseId(naechsteReiseId);
    setMode("liste");
    setSaveStatus("Reise wird gelöscht …");
    try {
      await deleteState(dateiname);
      setSaveStatus("Reise gelöscht.");
    } catch (error) {
      console.error(error);
      const meldung = error instanceof Error ? error.message : String(error);
      setSaveStatus(`Löschen der Reise-Datei fehlgeschlagen: ${meldung}`);
    }
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
    return (
      <div className="pl-loading">
        Fehler beim Laden der Packliste. Bitte Seite neu laden.
        {loadErrorMessage && (
          <>
            <br />
            <span style={{ fontSize: "0.85em", opacity: 0.8 }}>({loadErrorMessage})</span>
          </>
        )}
      </div>
    );
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
            onClick={() => {
              // Frisch beginnen - falls zuvor noch eine Teilnehmer-/Aktivitäten-Auswahl
              // einer anderen Reise "hängen geblieben" ist (z.B. Abbrechen bei "Teilnehmer
              // bearbeiten"), soll die nicht ungewollt in die neue Reise übernommen werden.
              // Seit V03-03 (von Clemens gewünscht, 2026-08-31): Clemens und Allgemein sind
              // von Anfang an vorausgewählt, statt dass man bei null anfängt.
              const vorbelegung = new Set<ID>();
              const clemens = data?.t05_namen.find((p) => p.namen === "Clemens");
              const allgemein = data?.t05_namen.find((p) => p.namen === "Allgemein");
              if (clemens) vorbelegung.add(clemens.id);
              if (allgemein) vorbelegung.add(allgemein.id);
              setTeilnehmerAuswahl(vorbelegung);
              setAktivitaetenAuswahl(new Set());
              setNeuPersonName("");
              setMode("neueReise");
            }}
          >
            + Neue Reise
          </button>{" "}
          <button
            className={"pl-edit-toggle" + (mode === "bearbeiten" ? " active" : "")}
            onClick={() => {
              const wechseltZurListe = mode === "bearbeiten";
              setMode(wechseltZurListe ? "liste" : "bearbeiten");
              // Von Clemens gemeldet (2026-08-30): nach dem Zuordnen fand er die frisch
              // hinzugefügten Gegenstände in der Liste nicht wieder, weil der "Neu
              // hinzugefügt"-Filter erst manuell angetippt werden musste. Beim
              // Zurückwechseln aus "Liste bearbeiten" jetzt automatisch aktivieren,
              // wenn es für die aktuelle Person welche gibt.
              if (wechseltZurListe && neuCount > 0) setOffenFilter("neu");
            }}
          >
            {mode === "bearbeiten" ? "Fertig" : "Liste bearbeiten"}
          </button>{" "}
          <button
            className={"pl-edit-toggle" + (mode === "aktivitaetenVerwalten" ? " active" : "")}
            onClick={() => {
              if (mode === "aktivitaetenVerwalten") {
                setMode("liste");
                return;
              }
              setVerwaltungSearch("");
              setMode("aktivitaetenVerwalten");
            }}
          >
            Standard-Gegenstände
          </button>{" "}
          {reise && (
            <>
              <button className="pl-edit-toggle" onClick={createPdf}>Als PDF drucken</button>{" "}
              <button className="pl-edit-toggle" onClick={createExcel}>Als Excel sichern</button>{" "}
              <button
                className={"pl-edit-toggle" + (mode === "teilnehmerBearbeiten" ? " active" : "")}
                onClick={() => {
                  if (mode === "teilnehmerBearbeiten") {
                    setMode("liste");
                    return;
                  }
                  // Vorbelegung mit den aktuell sichtbaren Personen (siehe
                  // sichtbarePersonen oben), nicht einfach leer - sonst sähe es so aus,
                  // als würden gerade alle abgewählt.
                  setTeilnehmerAuswahl(new Set(sichtbarePersonen.map((p) => p.id)));
                  setMode("teilnehmerBearbeiten");
                }}
              >
                Teilnehmer
              </button>{" "}
            </>
          )}
          <button
            className="pl-edit-toggle"
            onClick={undo}
            disabled={history.length === 0}
            title={history.length > 0 ? `${history.length} Schritt(e) verfügbar` : "Kein Verlauf vorhanden"}
          >
            ↩ Rückgängig{history.length > 0 ? ` (${history.length})` : ""}
          </button>{" "}
          {reise && !reiseLoeschenBestaetigen && (
            <button
              className="pl-edit-toggle pl-danger"
              onClick={() => setReiseLoeschenBestaetigen(true)}
            >
              Reise löschen
            </button>
          )}
          {reise && reiseLoeschenBestaetigen && (
            <span className="pl-loeschen-bestaetigen">
              „{reise.reise}“ wirklich unwiderruflich löschen?{" "}
              <button className="pl-edit-toggle pl-danger" onClick={loescheReise}>
                Ja, löschen
              </button>{" "}
              <button className="pl-edit-toggle" onClick={() => setReiseLoeschenBestaetigen(false)}>
                Abbrechen
              </button>
            </span>
          )}
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
              {/* V03-04: eigener "Fertig"-Knopf statt automatischem Verschwinden beim
                  ersten Antippen einer Menge - siehe bestaetigeNeuHinzugefuegt(). Nur
                  sichtbar, während man sich diese Ansicht wirklich anschaut. */}
              {offenFilter === "neu" && (
                <button className="pl-fertig-btn" onClick={bestaetigeNeuHinzugefuegt}>
                  Fertig
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
          {/* Seit V03-02 zunächst nur Weiter zur Teilnehmer-Auswahl (von Clemens
              ausdrücklich als erster Schritt gewünscht, 2026-08-30) - die Reise selbst
              wird erst danach in erstelleNeueReise() angelegt. */}
          <button onClick={() => setMode("neueReiseTeilnehmer")} disabled={!neuReiseName.trim()}>
            Weiter: Teilnehmer auswählen
          </button>
        </div>
      )}

      {mode === "neueReiseTeilnehmer" && data && (
        <div className="pl-newreise">
          <label>Wer ist bei „{neuReiseName.trim() || "dieser Reise"}“ dabei?</label>
          <div className="pl-teilnehmer-liste">
            {data.t05_namen.map((p) => (
              <label key={p.id} className="pl-teilnehmer-check">
                <input
                  type="checkbox"
                  checked={teilnehmerAuswahl.has(p.id)}
                  onChange={() => toggleTeilnehmer(p.id)}
                />
                {p.namen}
              </label>
            ))}
          </div>
          {/* "+ neue Person" (V03-03, von Clemens gewünscht 2026-08-31) - z.B. für einen
              Gast, der noch gar nicht in der Namensliste steht. Wird direkt angekreuzt. */}
          <label>… oder neue Person</label>
          <input
            value={neuPersonName}
            onChange={(e) => setNeuPersonName(e.target.value)}
            placeholder="Name eingeben"
          />
          <button
            onClick={() => {
              const id = erstelleNeuePerson(neuPersonName);
              if (id) {
                setTeilnehmerAuswahl((prev) => new Set(prev).add(id));
                setNeuPersonName("");
              }
            }}
            disabled={!neuPersonName.trim()}
          >
            + Person anlegen
          </button>
          <div style={{ marginTop: 12 }}>
            <button onClick={() => setMode("neueReise")}>Zurück</button>{" "}
            <button onClick={() => setMode("neueReiseAktivitaeten")} disabled={teilnehmerAuswahl.size === 0}>
              Weiter: Aktivitäten auswählen
            </button>
          </div>
        </div>
      )}

      {mode === "neueReiseAktivitaeten" && data && (
        <div className="pl-newreise">
          <label>Welche Aktivitäten sind bei „{neuReiseName.trim() || "dieser Reise"}“ dabei?</label>
          <p className="pl-save">
            „Basics" ist immer automatisch dabei. Zu jeder ausgewählten Aktivität werden die
            hinterlegten Standard-Gegenstände (siehe „Standard-Gegenstände" in der
            Werkzeugleiste) automatisch zur Liste hinzugefügt.
          </p>
          <div className="pl-teilnehmer-liste">
            {data.t02_aktivitaet
              .filter((a) => a.aktivitaet !== "Basics")
              .map((a) => (
                <label key={a.id} className="pl-teilnehmer-check">
                  <input
                    type="checkbox"
                    checked={aktivitaetenAuswahl.has(a.id)}
                    onChange={() => toggleAktivitaet(a.id)}
                  />
                  {a.aktivitaet}
                </label>
              ))}
          </div>
          <button onClick={() => setMode("neueReiseTeilnehmer")}>Zurück</button>{" "}
          <button onClick={() => erstelleNeueReise(Array.from(teilnehmerAuswahl), Array.from(aktivitaetenAuswahl))}>
            Reise anlegen &amp; Gegenstände auswählen
          </button>
        </div>
      )}

      {mode === "teilnehmerBearbeiten" && reise && data && (
        <div className="pl-newreise">
          <label>Wer ist bei „{reise.reise}“ dabei?</label>
          <div className="pl-teilnehmer-liste">
            {data.t05_namen.map((p) => (
              <label key={p.id} className="pl-teilnehmer-check">
                <input
                  type="checkbox"
                  checked={teilnehmerAuswahl.has(p.id)}
                  onChange={() => toggleTeilnehmer(p.id)}
                />
                {p.namen}
              </label>
            ))}
          </div>
          <label>… oder neue Person</label>
          <input
            value={neuPersonName}
            onChange={(e) => setNeuPersonName(e.target.value)}
            placeholder="Name eingeben"
          />
          <button
            onClick={() => {
              const id = erstelleNeuePerson(neuPersonName);
              if (id) {
                setTeilnehmerAuswahl((prev) => new Set(prev).add(id));
                setNeuPersonName("");
              }
            }}
            disabled={!neuPersonName.trim()}
          >
            + Person anlegen
          </button>
          <p className="pl-save">
            Entfernte Personen verlieren nichts – ihre Einträge bleiben erhalten und werden
            nur ausgeblendet, bis sie wieder hinzugefügt werden. Wer neu dazukommt, bekommt
            automatisch Clemens' aktuelle Gegenstände dieser Reise als Ausgangspunkt.
          </p>
          <button onClick={() => setMode("liste")}>Abbrechen</button>{" "}
          <button onClick={speichereTeilnehmer} disabled={teilnehmerAuswahl.size === 0}>
            Speichern
          </button>
        </div>
      )}

      {mode === "aktivitaetenVerwalten" && data && (
        <div className="pl-newreise">
          <label>Standard-Gegenstände je Aktivität</label>
          <p className="pl-save">
            Hier legst du fest, welche Gegenstände beim Anlegen einer neuen Reise
            automatisch vorgeschlagen werden, sobald du eine Aktivität dabei auswählst.
            Änderungen hier wirken sich nicht auf bereits angelegte Reisen aus.
          </p>
          <div className="pl-teilnehmer-liste" style={{ marginBottom: 10 }}>
            {data.t02_aktivitaet.map((a) => (
              <button
                key={a.id}
                className={"pl-edit-toggle" + (verwaltungAktivitaetId === a.id ? " active" : "")}
                onClick={() => setVerwaltungAktivitaetId(a.id)}
                style={{ marginRight: 6, marginBottom: 6 }}
              >
                {a.aktivitaet}
              </button>
            ))}
          </div>
          <label>… oder neue Aktivität</label>
          <input
            value={neuAktivitaetName}
            onChange={(e) => setNeuAktivitaetName(e.target.value)}
            placeholder="z.B. SUP"
          />
          <button onClick={erstelleNeueAktivitaet} disabled={!neuAktivitaetName.trim()}>
            + Aktivität anlegen
          </button>

          {verwaltungAktivitaetId && (() => {
            const aktivitaet = data.t02_aktivitaet.find((a) => a.id === verwaltungAktivitaetId);
            if (!aktivitaet) return null;
            const standardIds = new Set(
              data.tk02_t02_t04.filter((r) => r.id_t02 === verwaltungAktivitaetId).map((r) => r.id_t04)
            );
            const gefiltert = data.t04_gegenstand.filter(
              (g) =>
                !verwaltungSearch || g.gegenstand.toLowerCase().includes(verwaltungSearch.toLowerCase())
            );
            const byKat = new Map<string, T04Gegenstand[]>();
            for (const g of gefiltert) {
              const kat = kathegorieName(data, g.id_kathegorie);
              if (!byKat.has(kat)) byKat.set(kat, []);
              byKat.get(kat)!.push(g);
            }
            const sortiert = Array.from(byKat.entries()).sort((a, b) => a[0].localeCompare(b[0], "de"));
            return (
              <div style={{ marginTop: 16 }}>
                <label>Standard-Gegenstände für „{aktivitaet.aktivitaet}“</label>
                <input
                  value={verwaltungSearch}
                  onChange={(e) => setVerwaltungSearch(e.target.value)}
                  placeholder="Gegenstand suchen…"
                />
                {sortiert.map(([katName, items]) => (
                  <div className="pl-category" key={katName}>
                    <div className="pl-category-head">
                      <span className="pl-category-tag">{katName}</span>
                      <span className="pl-category-count">{items.length}</span>
                    </div>
                    <div className="pl-teilnehmer-liste">
                      {items
                        .slice()
                        .sort((a, b) => a.gegenstand.localeCompare(b.gegenstand, "de"))
                        .map((g) => (
                          <label key={g.id} className="pl-teilnehmer-check">
                            <input
                              type="checkbox"
                              checked={standardIds.has(g.id)}
                              onChange={() => toggleStandardGegenstand(verwaltungAktivitaetId, g.id)}
                            />
                            {g.gegenstand}
                          </label>
                        ))}
                    </div>
                  </div>
                ))}
              </div>
            );
          })()}
        </div>
      )}

      {mode === "bearbeiten" && reise && personFilter && (
        <div className="pl-body">
          {(() => {
            type Prio = "rot" | "blau" | "gruen";
            type Eintrag = { g: T04Gegenstand; prio: Prio; hint: string | null; allgemeinAktiv: boolean };
            const byKat = new Map<string, Eintrag[]>();
            // V03-04, auf Wunsch von Clemens (2026-09-03): rein informative "A"-Markierung,
            // zeigt bei jedem Gegenstand, den "Allgemein" für diese Reise schon aktiv hat -
            // unabhängig von rot/blau/grün. Bisher war das z.B. bei roten Gegenständen (ich
            // hatte es, hab's rausgenommen) gar nicht sichtbar, weil dort kein "andere
            // Person"-Hinweis berechnet wird. Keine Verwechslung mit der "A"-Spalte
            // (Ausgewählt) im Kopfbereich beabsichtigt - andere Bildschirmstelle, andere
            // Optik (Badge statt Spaltenkopf).
            const allgemeinId = data.t05_namen.find((p) => p.namen === "Allgemein")?.id ?? null;
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
              let allgemeinAktiv = false;
              if (allgemeinId && personFilter !== allgemeinId && tk03) {
                allgemeinAktiv = data.tk04_tk03_t05.some(
                  (r) => r.id_tk03 === tk03!.id && r.id_t05 === allgemeinId && r.ausgewaehlt !== -1
                );
              }
              const kat = kathegorieName(data, g.id_kathegorie);
              if (!byKat.has(kat)) byKat.set(kat, []);
              byKat.get(kat)!.push({ g, prio, hint, allgemeinAktiv });
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
                  {items.map(({ g, prio, hint, allgemeinAktiv }) => (
                    <div className={"pl-add-item pri-" + prio} key={g.id}>
                      <div className="pl-add-item-row">
                        <div className="pl-add-item-text">
                          <span className="pl-add-item-name">{g.gegenstand}</span>
                          {hint && <span className="pl-add-item-hint">{hint}</span>}
                        </div>
                        {allgemeinAktiv && (
                          <span className="pl-allgemein-badge" title="„Allgemein“ hat diesen Gegenstand bei dieser Reise schon">
                            A
                          </span>
                        )}
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
