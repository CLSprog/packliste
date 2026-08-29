// Kleiner Selbsttest der Sync-/Konfliktlogik (sync.ts) - kein Teil der App, nur zur
// Verifikation während der Entwicklung. Läuft mit: npx tsx tests/sync.selftest.ts
import { diffAndMerge, applyConflictResolutions, schemaEqual } from "../src/sync";
import type { SchemaData } from "../src/data/schema";

function basis(): SchemaData {
  return {
    t01_reise: [{ id: "R-1", reise: "Schottland 2026", von: null, bis: null, notiz: "" }],
    t02_aktivitaet: [],
    t03_kathegorie: [{ id: "K-1", kathegorie: "Kleidung", notiz: "" }],
    t04_gegenstand: [{ id: "G-1", gegenstand: "Zelt", id_kathegorie: "K-1", notiz: "" }],
    t05_namen: [{ id: "N-1", namen: "Clemens", notiz: "" }],
    tk01_t01_t02: [],
    tk02_t02_t04: [],
    tk03_tk01_t04: [{ id: "V3-1", id_tk01: "V1-1", id_t04: "G-1", notiz: "" }],
    tk04_tk03_t05: [
      { id: "V4-1", id_tk03: "V3-1", id_t05: "N-1", ausgewaehlt: 1, hergerichtet: null, eingepackt: null, verwendet: null },
    ],
    neu_hinzugefuegt: [],
  };
}

let fehler = 0;
function check(name: string, bedingung: boolean) {
  console.log(`${bedingung ? "OK  " : "FEHL"} - ${name}`);
  if (!bedingung) fehler++;
}

// Test 1: keine Änderung irgendwo -> kein Konflikt, kein Auto-Merge
{
  const b = basis();
  const l = basis();
  const r = basis();
  const res = diffAndMerge(b, l, r);
  check("Test1: keine Konflikte", res.conflicts.length === 0);
  check("Test1: kein Auto-Merge", res.autoMerged.length === 0);
}

// Test 2: nur remote ändert eine Zeile -> Auto-Merge, kein Konflikt
{
  const b = basis();
  const l = basis();
  const r = basis();
  r.tk04_tk03_t05[0].ausgewaehlt = 2;
  const res = diffAndMerge(b, l, r);
  check("Test2: kein Konflikt", res.conflicts.length === 0);
  check("Test2: ein Auto-Merge", res.autoMerged.length === 1);
  check("Test2: Wert übernommen", res.merged.tk04_tk03_t05[0].ausgewaehlt === 2);
}

// Test 3: nur lokal ändert eine Zeile -> bleibt lokal, kein Konflikt/Auto-Merge
{
  const b = basis();
  const l = basis();
  l.tk04_tk03_t05[0].ausgewaehlt = 3;
  const r = basis();
  const res = diffAndMerge(b, l, r);
  check("Test3: kein Konflikt", res.conflicts.length === 0);
  check("Test3: kein Auto-Merge", res.autoMerged.length === 0);
  check("Test3: lokaler Wert bleibt", res.merged.tk04_tk03_t05[0].ausgewaehlt === 3);
}

// Test 4: beide ändern dieselbe Zeile unterschiedlich -> echter Konflikt
{
  const b = basis();
  const l = basis();
  l.tk04_tk03_t05[0].ausgewaehlt = 3;
  const r = basis();
  r.tk04_tk03_t05[0].ausgewaehlt = 5;
  const res = diffAndMerge(b, l, r);
  check("Test4: ein Konflikt erkannt", res.conflicts.length === 1);
  if (res.conflicts.length === 1) {
    check("Test4: lokal=3 im Konflikt", res.conflicts[0].local?.ausgewaehlt === 3);
    check("Test4: remote=5 im Konflikt", res.conflicts[0].remote?.ausgewaehlt === 5);
    check("Test4: Label lesbar", res.conflicts[0].label.includes("Zelt"));

    const resolutions = new Map<string, "local" | "remote">([[`tk04_tk03_t05:V4-1`, "remote"]]);
    const final = applyConflictResolutions(res, resolutions);
    check("Test4: Konfliktauflösung übernimmt remote", final.tk04_tk03_t05[0].ausgewaehlt === 5);
    check("Test4: genau eine Zeile übrig (keine Dopplung)", final.tk04_tk03_t05.length === 1);
  }
}

// Test 5: beide ändern dieselbe Zeile auf denselben Wert -> kein echter Konflikt
{
  const b = basis();
  const l = basis();
  l.tk04_tk03_t05[0].ausgewaehlt = 4;
  const r = basis();
  r.tk04_tk03_t05[0].ausgewaehlt = 4;
  const res = diffAndMerge(b, l, r);
  check("Test5: kein Konflikt bei gleichem Zielwert", res.conflicts.length === 0);
}

// Test 6: neue Zeile nur remote (z.B. anderes Gerät hat "Liste bearbeiten" benutzt) -> Auto-Merge, neue ID bleibt eigenständig
{
  const b = basis();
  const l = basis();
  const r = basis();
  r.t04_gegenstand.push({ id: "G-2", gegenstand: "Sandalen", id_kathegorie: "K-1", notiz: "" });
  const res = diffAndMerge(b, l, r);
  check("Test6: kein Konflikt bei neuer Zeile", res.conflicts.length === 0);
  check("Test6: Auto-Merge meldet Neuanlage", res.autoMerged.some((a) => a.art === "hinzugefügt"));
  check("Test6: neue Zeile im Merge-Ergebnis", res.merged.t04_gegenstand.some((g) => g.id === "G-2"));
}

// Test 7: beide legen unabhängig eine je eigene neue Zeile an (z.B. beide "Sandalen" neu
// angelegt, aber mit unterschiedlicher Zufalls-ID) -> keine ID-Kollision, beide bleiben,
// kein Konflikt (siehe Konzeptdokument: "fällt als zwei neue Zeilen auf, nicht als Datenverlust")
{
  const b = basis();
  const l = basis();
  l.t04_gegenstand.push({ id: "G-3", gegenstand: "Sandalen", id_kathegorie: "K-1", notiz: "" });
  const r = basis();
  r.t04_gegenstand.push({ id: "G-4", gegenstand: "Sandalen", id_kathegorie: "K-1", notiz: "" });
  const res = diffAndMerge(b, l, r);
  check("Test7: kein Konflikt", res.conflicts.length === 0);
  check("Test7: beide neuen Zeilen vorhanden", res.merged.t04_gegenstand.some((g) => g.id === "G-3") && res.merged.t04_gegenstand.some((g) => g.id === "G-4"));
}

// Test 8: schemaEqual - Grundlage für "gibt es überhaupt eine eigene Änderung zu
// speichern" (V02-02: verhindert unnötige Schreibvorgänge bei rein periodischen
// Pull-Checks, siehe SchemaApp.tsx "habenWirWasZuSpeichern").
{
  const a = basis();
  const b = basis();
  check("Test8: identische Stände gleich", schemaEqual(a, b));
  b.tk04_tk03_t05[0].ausgewaehlt = 9;
  check("Test8: geänderter Stand ungleich", !schemaEqual(a, b));
}

console.log(fehler === 0 ? "\nAlle Tests bestanden." : `\n${fehler} Test(s) fehlgeschlagen.`);
process.exit(fehler === 0 ? 0 : 1);
