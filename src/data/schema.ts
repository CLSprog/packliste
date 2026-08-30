// P03 Packliste – relationales Datenschema
// Stammtabellen (t01-t05) und Verknüpfungstabellen (tk01-tk04),
// wie mit Clemens abgestimmt. IDs sind Präfix + 5-stellige Zufallszahl,
// damit mehrere Geräte gleichzeitig neue Zeilen anlegen können, ohne
// dass IDs kollidieren.

export type ID = string;

export interface T01Reise {
  id: ID;
  reise: string;
  von: string | null; // ISO-Datum
  bis: string | null;
  notiz: string;
  // Seit der Aufteilung in Stammdaten-/Reise-Einzeldateien (Paket A, 2026-08-29) hat jede
  // Reise keinen gemeinsamen Datensatz mehr, aus dem sich eine Reihenfolge automatisch
  // ergibt (jede liegt in ihrer eigenen Datei). Damit die Reise-Reiter trotzdem in der
  // gewohnten Reihenfolge erscheinen (nicht z.B. alphabetisch nach Dateiname), wird die
  // Reihenfolge hier explizit mitgespeichert. Optional, damit ältere/unvollständige Daten
  // nicht brechen - fehlt der Wert, landet die Reise beim Sortieren ganz hinten.
  reihenfolge?: number;
  // Explizite Teilnehmerliste (V03-02, von Clemens gewünscht 2026-08-30): wer bei dieser
  // Reise überhaupt dabei ist, steuert die Personen-Reiter - unabhängig davon, ob schon
  // Gegenstände/Mengen für diese Person eingetragen sind. Wird beim Anlegen einer neuen
  // Reise als eigener Schritt abgefragt, danach über "Teilnehmer" jederzeit änderbar.
  // Optional, damit ältere Reisen ohne dieses Feld nicht brechen - siehe
  // backfillFehlendeTeilnehmer in SchemaApp.tsx (leitet es einmalig aus den tatsächlich
  // vorhandenen Personen-Zuordnungen ab, keine sichtbare Änderung für bestehende Reisen).
  teilnehmer?: ID[];
}

export interface T02Aktivitaet {
  id: ID;
  aktivitaet: string;
  notiz: string;
}

export interface T03Kathegorie {
  id: ID;
  kathegorie: string;
  notiz: string;
}

export interface T04Gegenstand {
  id: ID;
  gegenstand: string;
  id_kathegorie: ID;
  notiz: string;
}

export interface T05Namen {
  id: ID;
  namen: string;
  notiz: string;
}

// Reise <-> Aktivität (n:m) – welche Aktivitäten bei welcher Reise vorkommen
export interface Tk01ReiseAktivitaet {
  id: ID;
  id_t01: ID;
  id_t02: ID;
}

// Aktivität <-> Gegenstand (n:m) – welcher Gegenstand zu welcher Aktivität passt
export interface Tk02AktivitaetGegenstand {
  id: ID;
  id_t02: ID;
  id_t04: ID;
}

// (Reise+Aktivität) <-> Gegenstand – welche Gegenstände tatsächlich auf einer
// konkreten Reise-Liste stehen (kann von der reinen Aktivitäts-Zuordnung
// abweichen, z.B. wenn Clemens einen Gegenstand bewusst weglässt)
export interface Tk03ReiseaktivitaetGegenstand {
  id: ID;
  id_tk01: ID;
  id_t04: ID;
  notiz: string;
}

// (obige Zeile) <-> Person – wer wieviel braucht, in vier Stufen.
// NULL = noch nicht angetippt (Unterschied zu 0 = bewusst auf null gesetzt)
export interface Tk04GegenstandPerson {
  id: ID;
  id_tk03: ID;
  id_t05: ID;
  ausgewaehlt: number | null;
  hergerichtet: number | null;
  eingepackt: number | null;
  verwendet: number | null;
}

export interface SchemaData {
  t01_reise: T01Reise[];
  t02_aktivitaet: T02Aktivitaet[];
  t03_kathegorie: T03Kathegorie[];
  t04_gegenstand: T04Gegenstand[];
  t05_namen: T05Namen[];
  tk01_t01_t02: Tk01ReiseAktivitaet[];
  tk02_t02_t04: Tk02AktivitaetGegenstand[];
  tk03_tk01_t04: Tk03ReiseaktivitaetGegenstand[];
  tk04_tk03_t05: Tk04GegenstandPerson[];
  // IDs von tk04-Zeilen, die über "Liste bearbeiten" hinzugefügt/reaktiviert wurden und
  // noch nicht angetippt sind (Filter "Neu hinzugefügt"). Optional, weil ältere
  // gespeicherte Dateien dieses Feld noch nicht kennen - dann als leer behandeln.
  // Seit V01-29 fix in OneDrive gespeichert (geräte-/sitzungsübergreifend), vorher nur
  // im Arbeitsspeicher der laufenden Sitzung.
  neu_hinzugefuegt?: ID[];
}

// ---- Zufalls-ID-Erzeugung (Präfix + 5-stellige Zahl), wie abgestimmt ----
const PREFIXES = {
  t01: "R",
  t02: "A",
  t03: "K",
  t04: "G",
  t05: "N",
  tk01: "V1",
  tk02: "V2",
  tk03: "V3",
  tk04: "V4",
} as const;

export function newId(kind: keyof typeof PREFIXES, existingIds: Set<string>): string {
  const prefix = PREFIXES[kind];
  let id: string;
  do {
    const n = Math.floor(10000 + Math.random() * 90000);
    id = `${prefix}-${n}`;
  } while (existingIds.has(id));
  return id;
}

// ---- Hilfsfunktionen zum Ableiten von Anzeige-Daten aus dem relationalen Schema ----

/** Alle Gegenstände, die für eine bestimmte Reise auf der Liste stehen (über tk03) */
export function gegenstaendeFuerReise(data: SchemaData, reiseId: ID): T04Gegenstand[] {
  const tk01Ids = new Set(
    data.tk01_t01_t02.filter((r) => r.id_t01 === reiseId).map((r) => r.id)
  );
  const t04Ids = new Set(
    data.tk03_tk01_t04.filter((r) => tk01Ids.has(r.id_tk01)).map((r) => r.id_t04)
  );
  return data.t04_gegenstand.filter((g) => t04Ids.has(g.id));
}

/** Die tk03-Zeile für einen bestimmten Gegenstand innerhalb einer Reise (falls vorhanden) */
export function tk03FuerGegenstand(
  data: SchemaData,
  reiseId: ID,
  gegenstandId: ID
): Tk03ReiseaktivitaetGegenstand | undefined {
  const tk01Ids = new Set(
    data.tk01_t01_t02.filter((r) => r.id_t01 === reiseId).map((r) => r.id)
  );
  return data.tk03_tk01_t04.find(
    (r) => tk01Ids.has(r.id_tk01) && r.id_t04 === gegenstandId
  );
}

/** Alle Personen-Zuordnungen (Menge je Stufe) für eine tk03-Zeile */
export function personenFuerTk03(data: SchemaData, tk03Id: ID): Tk04GegenstandPerson[] {
  return data.tk04_tk03_t05.filter((r) => r.id_tk03 === tk03Id);
}

export function kathegorieName(data: SchemaData, id: ID): string {
  return data.t03_kathegorie.find((k) => k.id === id)?.kathegorie ?? "Unbekannt";
}

export function personName(data: SchemaData, id: ID): string {
  return data.t05_namen.find((p) => p.id === id)?.namen ?? "Unbekannt";
}
