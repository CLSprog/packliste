import { getAccessToken } from "./auth";

// Die Datei liegt sichtbar im OneDrive des Nutzers unter diesem Pfad -
// nicht in einem versteckten App-Ordner, damit sie bei Bedarf auch manuell
// eingesehen, verschoben oder gesichert werden kann.
const FOLDER = "P03_Packliste";
const FILE = "P03_Packliste_Zustand_AI.json";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function itemPath(file: string): string {
  return `/me/drive/root:/${FOLDER}/${file}`;
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
export async function loadState(file: string = FILE): Promise<unknown | null> {
  const response = await graphFetch(`${itemPath(file)}:/content`);
  if (response.status === 404) return null; // erster Start: Datei existiert noch nicht
  if (!response.ok) throw new Error(`OneDrive-Ladefehler (${response.status})`);
  return response.json();
}

export async function saveState(state: unknown, file: string = FILE): Promise<void> {
  const body = JSON.stringify(state);
  const response = await graphFetch(`${itemPath(file)}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body,
  });
  if (!response.ok) throw new Error(`OneDrive-Speicherfehler (${response.status})`);
}

export async function deleteState(file: string = FILE): Promise<void> {
  const response = await graphFetch(itemPath(file));
  if (response.status === 404) return;
  if (!response.ok) throw new Error(`OneDrive-Lesefehler (${response.status})`);
  const deleteResponse = await graphFetch(itemPath(file), { method: "DELETE" });
  if (!deleteResponse.ok && deleteResponse.status !== 404) {
    throw new Error(`OneDrive-Löschfehler (${deleteResponse.status})`);
  }
}
