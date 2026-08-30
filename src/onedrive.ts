import { getAccessToken } from "./auth";

// Die Datei liegt sichtbar im OneDrive des Nutzers unter diesem Pfad -
// nicht in einem versteckten App-Ordner, damit sie bei Bedarf auch manuell
// eingesehen, verschoben oder gesichert werden kann.
// Einziger gültiger Speicherort (von Clemens am 2026-08-28 bestätigt).
// Ab V01-27 bewusst OHNE Fallback auf frühere/falsche Ordnernamen - Clemens
// löscht die alten Dateien manuell selbst aus OneDrive. Siehe Projektstand
// für die Historie (LEGACY_FOLDER, WRONG_FOLDER_V24 gab es bis V01-26).
const FOLDER = "_KI/ThinkTank/P03_Packliste/07_Database";

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

// Listet die Dateinamen im Datenordner auf - seit Paket A (Aufteilung in Stammdaten- +
// Reise-Einzeldateien, 2026-08-29) nötig, um beim Start herauszufinden, welche Reise-
// Dateien überhaupt existieren (die App kennt ihre Namen sonst nicht im Voraus, jede neue
// Reise bekommt ihre eigene Datei). Einfache Einzelseiten-Abfrage ohne Pagination - für
// die überschaubare Dateimenge dieses Ordners (Stammdaten + eine Handvoll Reisen)
// ausreichend; sollten es sehr viele Reisen werden, müsste hier @odata.nextLink
// nachgezogen werden.
export async function listFiles(folder: string = FOLDER): Promise<string[]> {
  const response = await graphFetch(
    `/me/drive/root:/${folder}:/children?$select=name&$top=200`
  );
  if (response.status === 404) return []; // Ordner existiert noch nicht
  if (!response.ok) throw new Error(`OneDrive-Listungsfehler (${response.status})`);
  const body = (await response.json()) as { value?: { name?: string }[] };
  return (body.value ?? []).map((item) => item.name).filter((name): name is string => !!name);
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
