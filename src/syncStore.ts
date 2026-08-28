// P03 Packliste – lokale Zwischenspeicherung für Offline-Betrieb (Phase 2)
//
// Zweck: Wenn die App offline weiterbenutzt wird und dabei neu geladen wird (z.B.
// Handy zwischendurch zu, App später wieder geöffnet - immer noch ohne Verbindung),
// dürfen die zwischenzeitlich gemachten Änderungen nicht verloren gehen, bis wieder
// synchronisiert werden kann. Dafür merkt sich der Browser zwei Dinge:
//  - "baseline": der letzte Stand, der nachweislich erfolgreich mit dem Server
//    abgeglichen wurde (Grundlage für den 3-Wege-Vergleich in sync.ts).
//  - "pending": der aktuelle Arbeitsstand, auch wenn er noch nicht gespeichert ist.
// Beides ist reiner Cache, kein Ersatz für OneDrive - geht der Browser-Speicher
// verloren (z.B. privates Fenster), gibt es beim nächsten Online-Laden ganz normal
// den Server-Stand.

import type { SchemaData } from "./data/schema";

const PREFIX = "p03_sync_";

function read(key: string): SchemaData | null {
  try {
    const raw = localStorage.getItem(PREFIX + key);
    return raw ? (JSON.parse(raw) as SchemaData) : null;
  } catch {
    return null; // z.B. privates Fenster ohne Speicherzugriff - dann läuft es ohne Cache
  }
}

function write(key: string, data: SchemaData) {
  try {
    localStorage.setItem(PREFIX + key, JSON.stringify(data));
  } catch {
    // Speicher voll o.ä. - Offline-Cache ist ein Komfortfeature, kein Muss
  }
}

function clear(key: string) {
  try {
    localStorage.removeItem(PREFIX + key);
  } catch {
    // ignorieren
  }
}

export function readBaseline(file: string): SchemaData | null {
  return read(`baseline_${file}`);
}
export function writeBaseline(file: string, data: SchemaData) {
  write(`baseline_${file}`, data);
}

export function readPending(file: string): SchemaData | null {
  return read(`pending_${file}`);
}
export function writePending(file: string, data: SchemaData) {
  write(`pending_${file}`, data);
}
export function clearPending(file: string) {
  clear(`pending_${file}`);
}

/** Netzwerkfehler (keine Verbindung) von echten Serverfehlern (z.B. abgelaufener
 *  Login, Berechtigungsfehler) unterscheiden - nur bei echten Netzwerkfehlern soll
 *  die App still in den Offline-Modus wechseln statt einen Fehler zu zeigen. */
export function istVerbindungsfehler(error: unknown): boolean {
  if (!navigator.onLine) return true;
  return error instanceof TypeError; // Browser werfen bei Netzwerkausfall ein TypeError aus fetch()
}
