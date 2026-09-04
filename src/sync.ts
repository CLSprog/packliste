// P03 Packliste – Sync- & Konflikterkennung (Multiuser/Offline, Phase 1+2)
//
// Grundidee (siehe Projekt-Konzeptdokument "Konzept Multiuser-Synchronisation"):
// - Jedes Gerät merkt sich den zuletzt erfolgreich synchronisierten Datenstand ("baseline").
// - Beim Speichern wird geprüft, ob sich der Stand auf dem Server seit der eigenen baseline
//   verändert hat (Inhaltsvergleich über den gesamten Dateiinhalt, siehe schemaEqual).
// - Falls ja: Zeilen-für-Zeile-Vergleich (3-Wege: baseline / eigener Stand / Server-Stand).
//   - Zeile nur lokal geändert -> eigene Änderung bleibt.
//   - Zeile nur remote geändert -> wird automatisch übernommen, aber zur Info gemeldet.
//   - Zeile auf beiden Seiten unterschiedlich geändert -> echter Konflikt, muss der Nutzer
//     pro Zeile einzeln entscheiden (nicht "alles von mir" oder "alles vom anderen Gerät").
//   - Zeile auf beiden Seiten gleich geändert -> kein Konflikt, einfach übernehmen.
//
// Dieselbe Logik gilt unverändert für den Offline-Fall: eine offline gesammelte Änderung
// ist aus Sicht dieser Funktion nichts anderes als eine "lokale Änderung", die erst beim
// Wiederverbinden mit dem dann aktuellen Server-Stand abgeglichen wird.

import type { ID, SchemaData } from "./data/schema";

export const TABLE_KEYS = [
  "t01_reise",
  "t02_aktivitaet",
  "t03_kathegorie",
  "t04_gegenstand",
  "t05_namen",
  "tk01_t01_t02",
  "tk02_t02_t04",
  "tk03_tk01_t04",
  "tk04_tk03_t05",
] as const;

export type TableKey = (typeof TABLE_KEYS)[number];

type Row = Record<string, unknown> & { id: string };

export interface RowConflict {
  table: TableKey;
  id: ID;
  /** Menschentauglicher Bezug, z.B. "Zelt – Clemens" */
  label: string;
  local: Row | null;
  remote: Row | null;
}

export interface AutoMergedChange {
  table: TableKey;
  id: ID;
  art: "hinzugefügt" | "geändert" | "entfernt";
  label: string;
}

export interface SyncResult {
  /** Datenstand nach Übernahme aller NICHT-konfliktären Änderungen. Für Zeilen mit
   *  echtem Konflikt steht hier vorläufig noch der lokale Stand, bis der Nutzer
   *  entschieden hat (siehe applyConflictResolutions). */
  merged: SchemaData;
  conflicts: RowConflict[];
  autoMerged: AutoMergedChange[];
}

// ---- Stabiler Vergleich, unabhängig von Objekt-Key-Reihenfolge ----

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const obj = value as Record<string, unknown>;
  const keys = Object.keys(obj).sort();
  return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(obj[k])}`).join(",")}}`;
}

function rowEqual(a: unknown, b: unknown): boolean {
  return stableStringify(a) === stableStringify(b);
}

/** Schneller, synchroner Inhaltsvergleich - z.B. um zu prüfen, ob es
 *  überhaupt eine lokale, noch nicht gesicherte Änderung gibt, bevor unnötig
 *  nach OneDrive geschrieben wird. */
export function schemaEqual(a: SchemaData, b: SchemaData): boolean {
  return stableStringify(a) === stableStringify(b);
}

// ---- Menschentaugliche Beschreibung einer Zeile, für Konflikt-/Info-Anzeige ----

interface NameIndex {
  gegenstand: Map<string, string>;
  person: Map<string, string>;
  reise: Map<string, string>;
  kathegorie: Map<string, string>;
  aktivitaet: Map<string, string>;
  tk03ZuGegenstand: Map<string, string>;
}

function buildNameIndex(...quellen: SchemaData[]): NameIndex {
  const idx: NameIndex = {
    gegenstand: new Map(),
    person: new Map(),
    reise: new Map(),
    kathegorie: new Map(),
    aktivitaet: new Map(),
    tk03ZuGegenstand: new Map(),
  };
  for (const d of quellen) {
    for (const g of d.t04_gegenstand) idx.gegenstand.set(g.id, g.gegenstand);
    for (const p of d.t05_namen) idx.person.set(p.id, p.namen);
    for (const r of d.t01_reise) idx.reise.set(r.id, r.reise);
    for (const k of d.t03_kathegorie) idx.kathegorie.set(k.id, k.kathegorie);
    for (const a of d.t02_aktivitaet) idx.aktivitaet.set(a.id, a.aktivitaet);
    for (const t of d.tk03_tk01_t04) idx.tk03ZuGegenstand.set(t.id, t.id_t04);
  }
  return idx;
}

function describeRow(table: TableKey, row: Row, idx: NameIndex): string {
  switch (table) {
    case "t01_reise":
      return `Reise „${row.reise as string}"`;
    case "t02_aktivitaet":
      return `Aktivität „${row.aktivitaet as string}"`;
    case "t03_kathegorie":
      return `Kategorie „${row.kathegorie as string}"`;
    case "t04_gegenstand":
      return `Gegenstand „${row.gegenstand as string}"`;
    case "t05_namen":
      return `Person „${row.namen as string}"`;
    case "tk01_t01_t02":
      return `${idx.reise.get(row.id_t01 as string) ?? "Reise"} ↔ ${idx.aktivitaet.get(row.id_t02 as string) ?? "Aktivität"}`;
    case "tk02_t02_t04":
      return `${idx.aktivitaet.get(row.id_t02 as string) ?? "Aktivität"} ↔ ${idx.gegenstand.get(row.id_t04 as string) ?? "Gegenstand"}`;
    case "tk03_tk01_t04":
      return `Gegenstand „${idx.gegenstand.get(row.id_t04 as string) ?? "?"}" auf der Reiseliste`;
    case "tk04_tk03_t05": {
      const gegenstandId = idx.tk03ZuGegenstand.get(row.id_tk03 as string);
      const gegenstandName = gegenstandId ? idx.gegenstand.get(gegenstandId) : undefined;
      const personName = idx.person.get(row.id_t05 as string);
      return `${gegenstandName ?? "Gegenstand"} – ${personName ?? "Person"}`;
    }
  }
}

function toMap(rows: Row[]): Map<string, Row> {
  return new Map(rows.map((r) => [r.id, r]));
}

/**
 * Vergleicht drei Datenstände (letzter gemeinsamer Sync-Stand, eigener aktueller Stand,
 * aktueller Server-Stand) Zeile für Zeile über alle Tabellen und liefert einen Merge-
 * Vorschlag plus die Liste der Punkte, die dem Nutzer angezeigt/zur Entscheidung
 * vorgelegt werden müssen.
 */
export function diffAndMerge(baseline: SchemaData, local: SchemaData, remote: SchemaData): SyncResult {
  const idx = buildNameIndex(baseline, local, remote);
  const conflicts: RowConflict[] = [];
  const autoMerged: AutoMergedChange[] = [];
  const mergedTables = {} as Record<TableKey, Row[]>;

  for (const table of TABLE_KEYS) {
    const baseMap = toMap(baseline[table] as unknown as Row[]);
    const localMap = toMap(local[table] as unknown as Row[]);
    const remoteMap = toMap(remote[table] as unknown as Row[]);
    const alleIds = new Set<string>([...baseMap.keys(), ...localMap.keys(), ...remoteMap.keys()]);
    const mergedRows: Row[] = [];

    for (const id of alleIds) {
      const b = baseMap.get(id) ?? null;
      const l = localMap.get(id) ?? null;
      const r = remoteMap.get(id) ?? null;
      const localChanged = !rowEqual(l, b);
      const remoteChanged = !rowEqual(r, b);

      if (!localChanged && !remoteChanged) {
        if (l) mergedRows.push(l);
        continue;
      }
      if (localChanged && !remoteChanged) {
        if (l) mergedRows.push(l);
        continue;
      }
      if (!localChanged && remoteChanged) {
        if (r) mergedRows.push(r);
        autoMerged.push({
          table,
          id,
          art: !b ? "hinzugefügt" : !r ? "entfernt" : "geändert",
          label: describeRow(table, (r ?? b)!, idx),
        });
        continue;
      }
      // beide Seiten haben diese Zeile seit der baseline verändert
      if (rowEqual(l, r)) {
        if (l) mergedRows.push(l);
        continue;
      }
      // echter Konflikt: vorläufig lokal übernehmen, bis der Nutzer entscheidet
      if (l) mergedRows.push(l);
      conflicts.push({
        table,
        id,
        label: describeRow(table, (l ?? r)!, idx),
        local: l,
        remote: r,
      });
    }
    mergedTables[table] = mergedRows;
  }

  // "Neu hinzugefügt"-Markierung ist rein informativ (siehe schema.ts) - hier reicht eine
  // einfache Vereinigung statt einer Konflikt-Abfrage: nichts geht verloren, im schlimmsten
  // Fall bleibt ein Eintrag auf einem Gerät noch kurz als "neu" markiert.
  const tk04Ids = new Set(mergedTables.tk04_tk03_t05.map((r) => r.id));
  const neuHinzugefuegt = Array.from(
    new Set([...(local.neu_hinzugefuegt ?? []), ...(remote.neu_hinzugefuegt ?? [])])
  ).filter((id) => tk04Ids.has(id));

  const merged = {
    t01_reise: mergedTables.t01_reise,
    t02_aktivitaet: mergedTables.t02_aktivitaet,
    t03_kathegorie: mergedTables.t03_kathegorie,
    t04_gegenstand: mergedTables.t04_gegenstand,
    t05_namen: mergedTables.t05_namen,
    tk01_t01_t02: mergedTables.tk01_t01_t02,
    tk02_t02_t04: mergedTables.tk02_t02_t04,
    tk03_tk01_t04: mergedTables.tk03_tk01_t04,
    tk04_tk03_t05: mergedTables.tk04_tk03_t05,
    neu_hinzugefuegt: neuHinzugefuegt,
  } as unknown as SchemaData;

  return { merged, conflicts, autoMerged };
}

/** Nachdem der Nutzer jeden Konflikt einzeln entschieden hat ("meins" oder "anderes"),
 *  wird das Ergebnis hier in den Merge-Vorschlag eingearbeitet. */
export function applyConflictResolutions(
  result: SyncResult,
  resolutions: Map<string, "local" | "remote">
): SchemaData {
  const data = result.merged;
  const tables = { ...data } as unknown as Record<string, Row[] | ID[]>;

  for (const conflict of result.conflicts) {
    const key = `${conflict.table}:${conflict.id}`;
    const choice = resolutions.get(key) ?? "local";
    const gewaehlteZeile = choice === "local" ? conflict.local : conflict.remote;
    const zeilen = (tables[conflict.table] as Row[]).filter((r) => r.id !== conflict.id);
    if (gewaehlteZeile) zeilen.push(gewaehlteZeile);
    tables[conflict.table] = zeilen;
  }

  return tables as unknown as SchemaData;
}
