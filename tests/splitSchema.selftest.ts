// Selbsttest für die Aufteilung in Stammdaten-/Reise-Einzeldateien (Paket A, siehe
// src/data/splitSchema.ts). Kein Teil der App, nur zur Verifikation während der
// Entwicklung. Läuft mit: npx tsx tests/splitSchema.selftest.ts
import {
  splitSchemaData,
  mergeSplitData,
  stammdatenSnapshotFuerReise,
  stammdatenAusSnapshotsRekonstruieren,
  verlaufAufteilen,
  historyRekonstruieren,
} from "../src/data/splitSchema";
import { schemaEqual } from "../src/sync";
import type { SchemaData } from "../src/data/schema";

function beispielDaten(): SchemaData {
  return {
    t01_reise: [
      { id: "R-1", reise: "Schottland 2026", von: null, bis: null, notiz: "", reihenfolge: 0 },
      { id: "R-2", reise: "China 2024", von: null, bis: null, notiz: "", reihenfolge: 1 },
    ],
    t02_aktivitaet: [
      { id: "A-1", aktivitaet: "Wandern", notiz: "" },
      { id: "A-2", aktivitaet: "Baden", notiz: "" },
    ],
    t03_kathegorie: [{ id: "K-1", kathegorie: "Kleidung", notiz: "" }],
    t04_gegenstand: [
      { id: "G-1", gegenstand: "Zelt", id_kathegorie: "K-1", notiz: "" },
      { id: "G-2", gegenstand: "Badehose", id_kathegorie: "K-1", notiz: "" },
    ],
    t05_namen: [
      { id: "N-1", namen: "Clemens", notiz: "" },
      { id: "N-2", namen: "Florian", notiz: "" },
    ],
    tk01_t01_t02: [
      { id: "V1-1", id_t01: "R-1", id_t02: "A-1" },
      { id: "V1-2", id_t01: "R-2", id_t02: "A-2" },
    ],
    tk02_t02_t04: [{ id: "V2-1", id_t02: "A-1", id_t04: "G-1" }],
    tk03_tk01_t04: [
      { id: "V3-1", id_tk01: "V1-1", id_t04: "G-1", notiz: "" },
      { id: "V3-2", id_tk01: "V1-2", id_t04: "G-2", notiz: "" },
    ],
    tk04_tk03_t05: [
      { id: "V4-1", id_tk03: "V3-1", id_t05: "N-1", ausgewaehlt: 1, hergerichtet: null, eingepackt: null, verwendet: null },
      { id: "V4-2", id_tk03: "V3-2", id_t05: "N-2", ausgewaehlt: 0, hergerichtet: null, eingepackt: null, verwendet: null },
    ],
    neu_hinzugefuegt: ["V4-2"],
  };
}

let fehler = 0;
function check(name: string, bedingung: boolean) {
  console.log(`${bedingung ? "OK  " : "FEHL"} - ${name}`);
  if (!bedingung) fehler++;
}

// Test 1: Aufteilen + Zusammenführen ergibt exakt denselben Stand (Rundreise).
{
  const original = beispielDaten();
  const { stammdaten, reisen } = splitSchemaData(original);
  const zusammengefuehrt = mergeSplitData(stammdaten, Array.from(reisen.values()));
  check("Test1: nichts geht bei Aufteilen+Zusammenführen verloren", schemaEqual(original, zusammengefuehrt));
  check("Test1: zwei Reisen erkannt", reisen.size === 2);
  check(
    "Test1: R-1 bekommt nur seine eigenen Zuordnungen",
    reisen.get("R-1")!.tk03_tk01_t04.length === 1 && reisen.get("R-1")!.tk03_tk01_t04[0].id === "V3-1"
  );
  check(
    "Test1: neu_hinzugefuegt landet bei der richtigen Reise",
    reisen.get("R-1")!.neu_hinzugefuegt.length === 0 && reisen.get("R-2")!.neu_hinzugefuegt[0] === "V4-2"
  );
}

// Test 2: Reihenfolge beim Zusammenführen bleibt stabil, auch wenn die Reisen in anderer
// Reihenfolge übergeben werden (z.B. weil OneDrive die Dateien anders auflistet).
{
  const original = beispielDaten();
  const { stammdaten, reisen } = splitSchemaData(original);
  const vertauscht = [reisen.get("R-2")!, reisen.get("R-1")!];
  const zusammengefuehrt = mergeSplitData(stammdaten, vertauscht);
  check(
    "Test2: Reihenfolge folgt `reihenfolge`, nicht der Übergabereihenfolge",
    zusammengefuehrt.t01_reise[0].id === "R-1" && zusammengefuehrt.t01_reise[1].id === "R-2"
  );
}

// Test 3: Stammdaten-Snapshot einer Reise enthält nur tatsächlich benutzte Zeilen.
{
  const original = beispielDaten();
  const { stammdaten, reisen } = splitSchemaData(original);
  const snapshot = stammdatenSnapshotFuerReise(stammdaten, reisen.get("R-1")!);
  check("Test3: Snapshot enthält nur Zelt, nicht Badehose", snapshot.t04_gegenstand.length === 1 && snapshot.t04_gegenstand[0].id === "G-1");
  check("Test3: Snapshot enthält nur die Aktivität von R-1", snapshot.t02_aktivitaet.length === 1 && snapshot.t02_aktivitaet[0].id === "A-1");
  check("Test3: Snapshot enthält nur Clemens (Person von R-1)", snapshot.t05_namen.length === 1 && snapshot.t05_namen[0].id === "N-1");
}

// Test 4: Notfall-Rekonstruktion der Stammdaten aus den Snapshots aller Reisen ergibt
// wieder alle Zeilen (Vereinigung), auch wenn keine einzelne Reise alles benutzt.
{
  const original = beispielDaten();
  const { stammdaten, reisen } = splitSchemaData(original);
  const snapshots = Array.from(reisen.values()).map((r) => stammdatenSnapshotFuerReise(stammdaten, r));
  const rekonstruiert = stammdatenAusSnapshotsRekonstruieren(snapshots);
  check("Test4: beide Gegenstände nach Rekonstruktion vorhanden", rekonstruiert.t04_gegenstand.length === 2);
  check("Test4: beide Personen nach Rekonstruktion vorhanden", rekonstruiert.t05_namen.length === 2);
}

// Test 5: Verlauf aufteilen + wieder rekonstruieren ergibt dieselbe Abfolge zurück
// (Normalfall: alle Reisen existieren über den ganzen Verlauf hinweg).
{
  const schritt1 = beispielDaten();
  const schritt2 = beispielDaten();
  schritt2.tk04_tk03_t05[0].ausgewaehlt = 2; // Clemens ändert Zelt-Menge
  const history = [schritt1, schritt2]; // älteste zuerst, wie in SchemaApp.tsx
  const { stammdatenVerlauf, reiseVerlaufMap } = verlaufAufteilen(history);
  const aktuelleReisen = Array.from(splitSchemaData(schritt2).reisen.values());
  const rekonstruiert = historyRekonstruieren(stammdatenVerlauf, reiseVerlaufMap, aktuelleReisen);
  check("Test5: gleiche Anzahl Verlaufsschritte", rekonstruiert.length === 2);
  check("Test5: Schritt 1 korrekt rekonstruiert", schemaEqual(rekonstruiert[0], schritt1));
  check("Test5: Schritt 2 korrekt rekonstruiert", schemaEqual(rekonstruiert[1], schritt2));
}

// Test 6: Eine Reise, die erst NACH ein paar Verlaufsschritten entstanden ist, hat einen
// kürzeren eigenen Verlauf - die Rekonstruktion muss das rechtsbündig (an den jüngsten
// Schritten) ausrichten, nicht an den ältesten.
{
  const schritt1 = beispielDaten();
  schritt1.t01_reise = schritt1.t01_reise.filter((r) => r.id === "R-1");
  schritt1.tk01_t01_t02 = schritt1.tk01_t01_t02.filter((r) => r.id_t01 === "R-1");
  schritt1.tk03_tk01_t04 = schritt1.tk03_tk01_t04.filter((r) => r.id_tk01 === "V1-1");
  schritt1.tk04_tk03_t05 = schritt1.tk04_tk03_t05.filter((r) => r.id_tk03 === "V3-1");
  schritt1.neu_hinzugefuegt = [];
  const schritt2 = beispielDaten(); // R-2 kommt hier neu dazu
  const history = [schritt1, schritt2];
  const { stammdatenVerlauf, reiseVerlaufMap } = verlaufAufteilen(history);
  const aktuelleReisen = Array.from(splitSchemaData(schritt2).reisen.values());
  const rekonstruiert = historyRekonstruieren(stammdatenVerlauf, reiseVerlaufMap, aktuelleReisen);
  check(
    "Test6: R-2 existiert im rekonstruierten Schritt 1 noch gar nicht (statt fälschlich mit heutigem Stand)",
    !rekonstruiert[0].t01_reise.some((r) => r.id === "R-2")
  );
  check("Test6: Schritt 2 (mit R-2) korrekt rekonstruiert", schemaEqual(rekonstruiert[1], schritt2));
}

console.log(fehler === 0 ? "\nAlle Tests bestanden." : `\n${fehler} Test(s) fehlgeschlagen.`);
process.exit(fehler === 0 ? 0 : 1);
