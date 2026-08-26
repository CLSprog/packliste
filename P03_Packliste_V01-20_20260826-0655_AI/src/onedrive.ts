import { getAccessToken } from "./auth";

// Die Datei liegt sichtbar im OneDrive des Nutzers unter diesem Pfad -
// nicht in einem versteckten App-Ordner, damit sie bei Bedarf auch manuell
// eingesehen, verschoben oder gesichert werden kann.
// Ab V01-19: richtiger, von Clemens vorgesehener Ordner (vorher lag die Datei
// direkt unter "P03_Packliste" auf der OneDrive-Wurzel, siehe LEGACY_FOLDER).
const FOLDER = "OneDrive_KI/ThinkTank/P03_Packliste/07_Database";
// Alter Speicherort vor der Ordner-Umstellung. Wird beim Laden nur noch als
// Fallback verwendet (falls im neuen Ordner noch nichts existiert), damit ein
// bereits laufendes Gerät seine Daten nicht verliert, wenn diese ZIP-Version
// zum ersten Mal geladen wird - siehe loadStateMitFallback in SchemaApp.tsx.
export const LEGACY_FOLDER = "P03_Packliste";
const FILE = "P03_Packliste_Zustand_AI.json";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function itemPath(file: string, folder: string = FOLDER): string {
  return `/me/drive/root:/${folder}/${file}`;
}

async function graphFetch(path: string, init: RequestInit = {}): Promise<Response> {
  const token = await getAccessToken();
  return fetch(`${GRAPH_BASE}${path}`, {
    ...init,
    headers: {
      ...(init.headers ?? {}),
      Authorization: `Bearer ${token}`,
    },
  });
}

// file ist optional, damit verschiedene Datenstände (z.B. alte App-Struktur
// vs. neues t01-tk04-Schema) nebeneinander in OneDrive existieren können,
// ohne sich gegenseitig zu überschreiben.
export async function loadState(file: string = FILE, folder: string = FOLDER): Promise<unknown | null> {
  const response = await graphFetch(`${itemPath(file, folder)}:/content`);
  if (response.status === 404) return null; // erster Start: Datei existiert noch nicht
  if (!response.ok) throw new Error(`OneDrive-Ladefehler (${response.status})`);
  return response.json();
}

export async function saveState(state: unknown, file: string = FILE, folder: string = FOLDER): Promise<void> {
  const body = JSON.stringify(state);
  const response = await graphFetch(`${itemPath(file, folder)}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body,
  });
  if (!response.ok) throw new Error(`OneDrive-Speicherfehler (${response.status})`);
}

export async function deleteState(file: string = FILE, folder: string = FOLDER): Promise<void> {
  const response = await graphFetch(itemPath(file, folder));
  if (response.status === 404) return;
  if (!response.ok) throw new Error(`OneDrive-Lesefehler (${response.status})`);
  const deleteResponse = await graphFetch(itemPath(file, folder), { method: "DELETE" });
  if (!deleteResponse.ok && deleteResponse.status !== 404) {
    throw new Error(`OneDrive-Löschfehler (${deleteResponse.status})`);
  }
}
