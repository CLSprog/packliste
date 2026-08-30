// P03 Packliste – Aufteilung in Stammdaten-Datei + eine Datei pro Reise (Paket A,
// siehe Projekt-Konzeptdokument "Konzept Reise-Einzeldateien & Freigabe").
//
// Grundidee: Nach außen (Anzeige, Bedienung, updateData/Rückgängig in SchemaApp.tsx)
// arbeitet die App weiterhin mit EINEM zusammengeführten `SchemaData`-Objekt, genau wie
// vor Paket A - es ändert sich nichts an der Oberfläche. Nur beim Laden/Speichern wird
// dieses eine Objekt in mehrere OneDrive-Dateien auseinandergenommen bzw. wieder
// zusammengesetzt:
//   - eine zentrale Stammdaten-Datei (Aktivitäten, Kategorien, Gegenstände, Personen samt
//     ihrer Aktivität-Gegenstand-Zuordnung - alles reiseunabhängig, siehe schema.ts)
//   - je eine Datei pro Reise (die Reise selbst, ihre Aktivitäts-/Gegenstands-/Personen-
//     Zuordnungen, plus eine Momentaufnahme der dafür benötigten Stammdaten als Backup)
//
// Damit die bereits gebaute Sync-/Konfliktlogik (sync.ts, unverändert) pro Datei
// wiederverwendet werden kann, wird jede Datei für den Abgleich in die vollständige
// `SchemaData`-Form "eingebettet" (die für diese Datei nicht relevanten Tabellen bleiben
// dabei einfach leer) - siehe *AlsSchemaData()/schemaDataAls*() unten.

import type {
  ID,
  SchemaData,
  T01Reise,
  T02Aktivitaet,
  T03Kathegorie,
  T04Gegenstand,
  T05Namen,
  Tk01ReiseAktivitaet,
  Tk02AktivitaetGegenstand,
  Tk03ReiseaktivitaetGegenstand,
  Tk04GegenstandPerson,
} from "./schema";

// ---- Dateinamen ----

export const STAMMDATEN_DATEI = "P03_Packliste_Stammdaten_AI.json";

const REISE_PRAEFIX = "P03_Packliste_Reise_";
const REISE_SUFFIX = "_AI.json";

export function reiseDateiname(reiseId: ID): string {
  return `${REISE_PRAEFIX}${reiseId}${REISE_SUFFIX}`;
}

/** Liefert die Reise-ID, falls der Dateiname wie eine Reise-Datei aussieht, sonst null. */
export function reiseIdAusDateiname(name: string): ID | null {
  if (!name.startsWith(REISE_PRAEFIX) || !name.endsWith(REISE_SUFFIX)) return null;
  const id = name.slice(REISE_PRAEFIX.length, name.length - REISE_SUFFIX.length);
  return id.length > 0 ? id : null;
}

// ---- Kern-Inhalte je Datei (ohne Verlauf/Snapshot - das "Eigentliche") ----

export interface StammdatenKern {
  t02_aktivitaet: T02Aktivitaet[];
  t03_kathegorie: T03Kathegorie[];
  t04_gegenstand: T04Gegenstand[];
  t05_namen: T05Namen[];
  tk02_t02_t04: Tk02AktivitaetGegenstand[];
}

export interface ReiseKern {
  reise: T01Reise;
  tk01_t01_t02: Tk01ReiseAktivitaet[];
  tk03_tk01_t04: Tk03ReiseaktivitaetGegenstand[];
  tk04_tk03_t05: Tk04GegenstandPerson[];
  neu_hinzugefuegt: ID[];
}

// ---- Tatsächliches Datei-Format in OneDrive ----

export interface StammdatenDatei extends StammdatenKern {
  /** Rückgängig-Verlauf nur für Stammdaten-Änderungen, älteste zuerst (siehe SchemaApp.tsx). */
  verlauf: StammdatenKern[];
}

export interface ReiseDatei extends ReiseKern {
  /** Momentaufnahme der von dieser Reise benutzten Stammdaten - reines Backup/Fallback,
   *  falls die zentrale Stammdaten-Datei einmal nicht erreichbar oder verloren ist. Nie
   *  die Wahrheit, solange die zentrale Datei erreichbar ist (siehe Konzeptdokument). */
  stammdaten_snapshot: StammdatenKern;
  /** Rückgängig-Verlauf nur für Änderungen an dieser einen Reise, älteste zuerst. */
  verlauf: ReiseKern[];
}

// ---- Aufteilen: kompletter (zusammengeführter) Datenstand -> Stammdaten + je Reise ----

export function splitSchemaData(data: SchemaData): {
  stammdaten: StammdatenKern;
  reisen: Map<ID, ReiseKern>;
} {
  const stammdaten: StammdatenKern = {
    t02_aktivitaet: data.t02_aktivitaet,
    t03_kathegorie: data.t03_kathegorie,
    t04_gegenstand: data.t04_gegenstand,
    t05_namen: data.t05_namen,
    tk02_t02_t04: data.tk02_t02_t04,
  };

  const neu = new Set(data.neu_hinzugefuegt ?? []);
  const reisen = new Map<ID, ReiseKern>();
  for (const reise of data.t01_reise) {
    const tk01 = data.tk01_t01_t02.filter((r) => r.id_t01 === reise.id);
    const tk01Ids = new Set(tk01.map((r) => r.id));
    const tk03 = data.tk03_tk01_t04.filter((r) => tk01Ids.has(r.id_tk01));
    const tk03Ids = new Set(tk03.map((r) => r.id));
    const tk04 = data.tk04_tk03_t05.filter((r) => tk03Ids.has(r.id_tk03));
    const tk04Ids = new Set(tk04.map((r) => r.id));
    reisen.set(reise.id, {
      reise,
      tk01_t01_t02: tk01,
      tk03_tk01_t04: tk03,
      tk04_tk03_t05: tk04,
      neu_hinzugefuegt: Array.from(neu).filter((id) => tk04Ids.has(id)),
    });
  }
  return { stammdaten, reisen };
}

/** Backup-Momentaufnahme der von einer Reise tatsächlich benutzten Stammdaten. */
export function stammdatenSnapshotFuerReise(stammdaten: StammdatenKern, reiseKern: ReiseKern): StammdatenKern {
  const t02Ids = new Set(reiseKern.tk01_t01_t02.map((r) => r.id_t02));
  const t04Ids = new Set(reiseKern.tk03_tk01_t04.map((r) => r.id_t04));
  const t05Ids = new Set(reiseKern.tk04_tk03_t05.map((r) => r.id_t05));
  const t04 = stammdaten.t04_gegenstand.filter((g) => t04Ids.has(g.id));
  const t03Ids = new Set(t04.map((g) => g.id_kathegorie));
  return {
    t02_aktivitaet: stammdaten.t02_aktivitaet.filter((a) => t02Ids.has(a.id)),
    t03_kathegorie: stammdaten.t03_kathegorie.filter((k) => t03Ids.has(k.id)),
    t04_gegenstand: t04,
    t05_namen: stammdaten.t05_namen.filter((p) => t05Ids.has(p.id)),
    tk02_t02_t04: stammdaten.tk02_t02_t04.filter((r) => t02Ids.has(r.id_t02) && t04Ids.has(r.id_t04)),
  };
}

// ---- Zusammenführen: Stammdaten + Reisen -> ein kompletter Datenstand, exakt im
// bisherigen Format, damit der Rest der App (Anzeige, updateData, Hilfsfunktionen aus
// schema.ts) unverändert weiterläuft. ----

export function mergeSplitData(stammdaten: StammdatenKern, reisen: ReiseKern[]): SchemaData {
  // Reihenfolge der Reise-Reiter bleibt stabil (siehe `reihenfolge` in schema.ts) - ohne
  // das würde die Reihenfolge von der (nicht garantierten) Reihenfolge abhängen, in der
  // OneDrive die Dateien beim Auflisten zurückgibt.
  const sortiert = [...reisen].sort((a, b) => {
    const ra = a.reise.reihenfolge ?? Number.MAX_SAFE_INTEGER;
    const rb = b.reise.reihenfolge ?? Number.MAX_SAFE_INTEGER;
    if (ra !== rb) return ra - rb;
    return a.reise.reise.localeCompare(b.reise.reise, "de-AT");
  });
  const neu = new Set<ID>();
  for (const r of sortiert) for (const id of r.neu_hinzugefuegt) neu.add(id);
  return {
    t01_reise: sortiert.map((r) => r.reise),
    t02_aktivitaet: stammdaten.t02_aktivitaet,
    t03_kathegorie: stammdaten.t03_kathegorie,
    t04_gegenstand: stammdaten.t04_gegenstand,
    t05_namen: stammdaten.t05_namen,
    tk01_t01_t02: sortiert.flatMap((r) => r.tk01_t01_t02),
    tk02_t02_t04: stammdaten.tk02_t02_t04,
    tk03_tk01_t04: sortiert.flatMap((r) => r.tk03_tk01_t04),
    tk04_tk03_t05: sortiert.flatMap((r) => r.tk04_tk03_t05),
    neu_hinzugefuegt: Array.from(neu),
  };
}

/** Notfall-Rekonstruktion der Stammdaten aus den Backup-Momentaufnahmen aller Reise-
 *  Dateien, falls die zentrale Stammdaten-Datei einmal nicht ladbar ist. Kann Zeilen
 *  fehlen, die in keiner offenen Reise vorkommen (z.B. gerade unbenutzte Katalog-
 *  Gegenstände) - besser als ein Totalausfall, aber kein Ersatz für die echte Datei. */
export function stammdatenAusSnapshotsRekonstruieren(snapshots: StammdatenKern[]): StammdatenKern {
  function vereinigtNachId<T extends { id: string }>(listen: T[][]): T[] {
    const map = new Map<string, T>();
    for (const liste of listen) for (const zeile of liste) if (!map.has(zeile.id)) map.set(zeile.id, zeile);
    return Array.from(map.values());
  }
  return {
    t02_aktivitaet: vereinigtNachId(snapshots.map((s) => s.t02_aktivitaet)),
    t03_kathegorie: vereinigtNachId(snapshots.map((s) => s.t03_kathegorie)),
    t04_gegenstand: vereinigtNachId(snapshots.map((s) => s.t04_gegenstand)),
    t05_namen: vereinigtNachId(snapshots.map((s) => s.t05_namen)),
    tk02_t02_t04: vereinigtNachId(snapshots.map((s) => s.tk02_t02_t04)),
  };
}

// ---- Einbetten in die volle SchemaData-Form, um sync.ts (diffAndMerge/hashData/
// schemaEqual) unverändert pro Datei wiederverwenden zu können - nicht betroffene
// Tabellen bleiben schlicht leer. ----

export function stammdatenAlsSchemaData(k: StammdatenKern): SchemaData {
  return {
    t01_reise: [],
    t02_aktivitaet: k.t02_aktivitaet,
    t03_kathegorie: k.t03_kathegorie,
    t04_gegenstand: k.t04_gegenstand,
    t05_namen: k.t05_namen,
    tk01_t01_t02: [],
    tk02_t02_t04: k.tk02_t02_t04,
    tk03_tk01_t04: [],
    tk04_tk03_t05: [],
    neu_hinzugefuegt: [],
  };
}

export function schemaDataAlsStammdatenKern(d: SchemaData): StammdatenKern {
  return {
    t02_aktivitaet: d.t02_aktivitaet,
    t03_kathegorie: d.t03_kathegorie,
    t04_gegenstand: d.t04_gegenstand,
    t05_namen: d.t05_namen,
    tk02_t02_t04: d.tk02_t02_t04,
  };
}

export function reiseAlsSchemaData(k: ReiseKern): SchemaData {
  return {
    t01_reise: [k.reise],
    t02_aktivitaet: [],
    t03_kathegorie: [],
    t04_gegenstand: [],
    t05_namen: [],
    tk01_t01_t02: k.tk01_t01_t02,
    tk02_t02_t04: [],
    tk03_tk01_t04: k.tk03_tk01_t04,
    tk04_tk03_t05: k.tk04_tk03_t05,
    neu_hinzugefuegt: k.neu_hinzugefuegt,
  };
}

export function schemaDataAlsReiseKern(d: SchemaData): ReiseKern {
  return {
    reise: d.t01_reise[0],
    tk01_t01_t02: d.tk01_t01_t02,
    tk03_tk01_t04: d.tk03_tk01_t04,
    tk04_tk03_t05: d.tk04_tk03_t05,
    neu_hinzugefuegt: d.neu_hinzugefuegt ?? [],
  };
}

// ---- Rückgängig-Verlauf aufteilen/rekonstruieren ----
//
// Der bisherige, im Speicher gehaltene Rückgängig-Verlauf (`history: SchemaData[]` in
// SchemaApp.tsx, ein gemeinsamer Stapel für alle Änderungen egal welcher Reise) bleibt
// WÄHREND der laufenden Sitzung unverändert - Undo verhält sich für Clemens also genau
// wie bisher. Erst beim Speichern wird dieser eine Stapel in die Verlauf-Felder der
// jeweiligen Dateien aufgeteilt (siehe Konzeptdokument: "jede Datei bringt ihren eigenen
// Verlauf gleich mit"), und beim Laden wieder so gut wie möglich zusammengesetzt.
//
// Bewusste Vereinfachung (mit Clemens abgestimmt, 2026-08-29: der alte, große Verlauf
// war ohnehin nicht mehr gebraucht): Diese Rekonstruktion geht davon aus, dass der
// Verlauf von genau EINEM Gerät/einer Sitzung stammt. Sobald mehrere Personen wirklich
// gleichzeitig eigene Geräte-Verläufe schreiben (erst ab Paket C/Freigabe relevant),
// kann die Zusammensetzung an den Rändern ungenau werden - im schlimmsten Fall ist ein
// Rückgängig-Schritt ungenau, nie ein Datenverlust bei den eigentlichen Packlisten-Daten.

export function verlaufAufteilen(
  history: SchemaData[]
): { stammdatenVerlauf: StammdatenKern[]; reiseVerlaufMap: Map<ID, ReiseKern[]> } {
  const stammdatenVerlauf = history.map(schemaDataAlsStammdatenKern);
  const reiseVerlaufMap = new Map<ID, ReiseKern[]>();
  for (const snapshot of history) {
    const { reisen } = splitSchemaData(snapshot);
    for (const [id, kern] of reisen) {
      if (!reiseVerlaufMap.has(id)) reiseVerlaufMap.set(id, []);
      reiseVerlaufMap.get(id)!.push(kern);
    }
  }
  return { stammdatenVerlauf, reiseVerlaufMap };
}

export function historyRekonstruieren(
  stammdatenVerlauf: StammdatenKern[],
  reiseVerlaufMap: Map<ID, ReiseKern[]>,
  aktuelleReisen: ReiseKern[]
): SchemaData[] {
  const n = stammdatenVerlauf.length;
  const ergebnis: SchemaData[] = [];
  for (let i = 0; i < n; i++) {
    const reisenBeiSchritt: ReiseKern[] = [];
    for (const aktuell of aktuelleReisen) {
      const arr = reiseVerlaufMap.get(aktuell.reise.id) ?? [];
      // Rechtsbündig ausrichten: eine kürzere Liste (Reise erst kürzlich angelegt) betrifft
      // die JÜNGSTEN Schritte, nicht die ältesten. Vor ihrer Entstehung (kein Eintrag an
      // dieser Stelle) taucht die Reise in diesem Verlaufsschritt gar nicht erst auf -
      // alles andere würde ihr fälschlich schon "damals" den heutigen Stand andichten.
      const offset = n - arr.length;
      if (i >= offset) reisenBeiSchritt.push(arr[i - offset]);
    }
    ergebnis.push(mergeSplitData(stammdatenVerlauf[i], reisenBeiSchritt));
  }
  return ergebnis;
}
