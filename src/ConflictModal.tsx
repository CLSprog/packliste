import { useState } from "react";
import type { RowConflict } from "./sync";

// Anzeige-Beschriftung für Tabellenfelder, wie in der App üblich (A/H/E/V).
const FELD_LABEL: Record<string, string> = {
  reise: "Reise",
  von: "Von",
  bis: "Bis",
  notiz: "Notiz",
  aktivitaet: "Aktivität",
  kathegorie: "Kategorie",
  gegenstand: "Gegenstand",
  namen: "Name",
  ausgewaehlt: "Ausgewählt (A)",
  hergerichtet: "Hergerichtet (H)",
  eingepackt: "Eingepackt (E)",
  verwendet: "Verwendet (V)",
};

function feldWert(value: unknown): string {
  if (value === null || value === undefined) return "–";
  if (value === "") return "(leer)";
  return String(value);
}

function relevanteFelder(row: Record<string, unknown> | null): [string, unknown][] {
  if (!row) return [];
  return Object.entries(row).filter(([key]) => key !== "id" && !key.startsWith("id_"));
}

interface ConflictModalProps {
  conflicts: RowConflict[];
  onResolve: (resolutions: Map<string, "local" | "remote">) => void;
}

export default function ConflictModal({ conflicts, onResolve }: ConflictModalProps) {
  const [choices, setChoices] = useState<Map<string, "local" | "remote">>(new Map());

  const alleEntschieden = conflicts.every((c) => choices.has(`${c.table}:${c.id}`));

  function waehlen(key: string, choice: "local" | "remote") {
    setChoices((prev) => {
      const next = new Map(prev);
      next.set(key, choice);
      return next;
    });
  }

  return (
    <div className="pl-conflict-backdrop">
      <div className="pl-conflict-panel">
        <h2>Unterschiedliche Änderungen auf zwei Geräten</h2>
        <p className="pl-conflict-intro">
          Diese {conflicts.length === 1 ? "Zeile wurde" : `${conflicts.length} Zeilen wurden`} sowohl auf diesem
          als auch auf einem anderen Gerät verändert, bevor synchronisiert wurde. Bitte für jede Zeile einzeln
          auswählen, welcher Stand gelten soll.
        </p>
        <div className="pl-conflict-list">
          {conflicts.map((c) => {
            const key = `${c.table}:${c.id}`;
            const gewaehlt = choices.get(key);
            return (
              <div className="pl-conflict-row" key={key}>
                <div className="pl-conflict-label">{c.label}</div>
                <div className="pl-conflict-sides">
                  <button
                    type="button"
                    className={"pl-conflict-side" + (gewaehlt === "local" ? " chosen" : "")}
                    onClick={() => waehlen(key, "local")}
                  >
                    <div className="pl-conflict-side-title">Dein Stand (dieses Gerät)</div>
                    {c.local ? (
                      relevanteFelder(c.local).map(([field, value]) => (
                        <div className="pl-conflict-field" key={field}>
                          <span>{FELD_LABEL[field] ?? field}</span>
                          <b>{feldWert(value)}</b>
                        </div>
                      ))
                    ) : (
                      <div className="pl-conflict-field">(auf diesem Gerät gelöscht)</div>
                    )}
                  </button>
                  <button
                    type="button"
                    className={"pl-conflict-side" + (gewaehlt === "remote" ? " chosen" : "")}
                    onClick={() => waehlen(key, "remote")}
                  >
                    <div className="pl-conflict-side-title">Anderes Gerät</div>
                    {c.remote ? (
                      relevanteFelder(c.remote).map(([field, value]) => (
                        <div className="pl-conflict-field" key={field}>
                          <span>{FELD_LABEL[field] ?? field}</span>
                          <b>{feldWert(value)}</b>
                        </div>
                      ))
                    ) : (
                      <div className="pl-conflict-field">(auf dem anderen Gerät gelöscht)</div>
                    )}
                  </button>
                </div>
              </div>
            );
          })}
        </div>
        <div className="pl-conflict-actions">
          <button
            type="button"
            className="pl-conflict-submit"
            disabled={!alleEntschieden}
            onClick={() => onResolve(choices)}
          >
            {alleEntschieden ? "Übernehmen & speichern" : `Noch ${conflicts.length - choices.size} zu entscheiden`}
          </button>
        </div>
      </div>
    </div>
  );
}
