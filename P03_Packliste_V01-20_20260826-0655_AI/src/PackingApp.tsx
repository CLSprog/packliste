import { useEffect, useMemo, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import catalog from "./data/catalog-data.json";
import scotland from "./data/scotland-data.json";
import { logout } from "./auth";
import { loadState as loadOneDriveState, saveState as saveOneDriveState, deleteState as deleteOneDriveState } from "./onedrive";

type View = "trips" | "create" | "prepare" | "pack" | "experience" | "categories" | "tasks" | "settings";
type Stage = "prepared" | "packed" | "used";
type Person = string;
type Priority = "Unbedingt" | "Empfohlen" | "Optional";
type CreationPhase = "selection" | "quantities" | "assignment" | "complete";
type ItemScope = "personal" | "shared";
type Quantities = { intended: number; prepared: number; packed: number; used: number };
type Assignment = { person: Person; quantities: Quantities };
type CatalogItem = {
  id: string;
  categoryId: string;
  category: string;
  name: string;
  units: string;
  notes: string;
  reviewStatus: string;
  archived: boolean;
};
type TripItem = {
  id: string;
  itemId: string;
  catalogItemId: string | null;
  name: string;
  category: string;
  unit: string;
  activity: string;
  priority: Priority;
  sourcePriority: string;
  luggage: string;
  source: string;
  notes: string;
  scope: ItemScope;
  assignments: Assignment[];
};
type Task = { id: string; area: string; label: string; priority: string; source: string; notes: string; done: boolean };
type PacklistRecord = {
  id: string;
  tripName: string;
  destination: string;
  dateFrom: string;
  dateTo: string;
  tripDates: string;
  peopleCount: number;
  people: string[];
  description: string;
  activities: string[];
  creationPhase: CreationPhase;
  tripItems: TripItem[];
  tasks: Task[];
  lastView: Exclude<View, "trips" | "categories" | "settings">;
  updatedAt: string;
};
type AppState = {
  schemaVersion: 9;
  activePacklistId: string | null;
  packlists: PacklistRecord[];
  creationPhase: CreationPhase;
  tripName: string;
  tripDates: string;
  destination: string;
  dateFrom: string;
  dateTo: string;
  peopleCount: number;
  description: string;
  activities: string[];
  people: string[];
  tripItems: TripItem[];
  tasks: Task[];
  customCatalogItems: CatalogItem[];
  catalogCategoryOverrides: Record<string, string>;
  deletedCatalogItemIds: string[];
};

type ExportRow = {
  category: string;
  item: string;
  quantity: number;
  unit: string;
  person: string;
  scope: string;
  priority: string;
  prepared: number;
  packed: number;
  used: number;
  notes: string;
};

type PendingRestore = {
  state: AppState;
  backupDate: string;
};

type PacklistDetails = {
  tripName: string;
  destination: string;
  dateFrom: string;
  dateTo: string;
  description: string;
  activities: string[];
  people: string[];
};

const APP_VERSION = "V00-12_00";

const priorities: Priority[] = ["Unbedingt", "Empfohlen", "Optional"];
const baseCatalogItems = catalog.items as CatalogItem[];
const activeBaseCatalogItems = baseCatalogItems.filter((item) => !item.archived);
const activityOptions = ["Wandern", "Baden", "Stadtbesichtigung", "Radfahren", "Camping", "Wintersport", "Sport & Fitness", "Veranstaltung"];

const scotlandItems: TripItem[] = scotland.items.map((item) => ({
  ...item,
  itemId: item.id,
  priority: item.priority as Priority,
  scope: item.assignments.length > 1 && item.assignments.every((assignment) => assignment.person !== "Alle") ? "personal" : "shared",
  assignments: item.assignments.map((assignment) => ({
    person: assignment.person === "Alle" ? "Noch nicht zugeordnet" : assignment.person,
    quantities: { intended: assignment.intended, prepared: 0, packed: 0, used: 0 },
  })),
}));

const initialScotlandRecord: PacklistRecord = {
  id: "packlist-scotland-2026",
  tripName: scotland.title,
  destination: "Schottland",
  dateFrom: "",
  dateTo: "",
  tripDates: scotland.dates,
  peopleCount: scotland.people.length,
  people: scotland.people,
  description: "Rundreise durch Schottland mit Städten, Mietauto und Wanderungen.",
  activities: ["Wandern", "Stadtbesichtigung"],
  creationPhase: "selection",
  tripItems: scotlandItems,
  tasks: scotland.tasks.map((task) => ({ ...task, done: false })),
  lastView: "create",
  updatedAt: new Date().toISOString(),
};

const initialState: AppState = {
  schemaVersion: 9,
  activePacklistId: initialScotlandRecord.id,
  packlists: [initialScotlandRecord],
  creationPhase: "selection",
  tripName: scotland.title,
  tripDates: scotland.dates,
  destination: "Schottland",
  dateFrom: "",
  dateTo: "",
  peopleCount: scotland.people.length,
  description: initialScotlandRecord.description,
  activities: initialScotlandRecord.activities,
  people: scotland.people,
  tripItems: scotlandItems,
  tasks: scotland.tasks.map((task) => ({ ...task, done: false })),
  customCatalogItems: [],
  catalogCategoryOverrides: {},
  deletedCatalogItemIds: [],
};

const blankState: AppState = {
  ...initialState,
  activePacklistId: null,
  packlists: [],
  creationPhase: "selection",
  tripName: "",
  tripDates: "",
  destination: "",
  dateFrom: "",
  dateTo: "",
  peopleCount: 1,
  description: "",
  activities: [],
  people: [],
  tripItems: [],
  tasks: [],
  customCatalogItems: [],
  catalogCategoryOverrides: {},
  deletedCatalogItemIds: [],
};

function currentAsRecord(state: AppState, lastView: View = "create"): PacklistRecord | null {
  if (!state.activePacklistId) return null;
  const safeLastView = (["create", "prepare", "pack", "experience", "tasks"] as View[]).includes(lastView)
    ? lastView as PacklistRecord["lastView"]
    : state.packlists.find((entry) => entry.id === state.activePacklistId)?.lastView ?? "create";
  return {
    id: state.activePacklistId,
    tripName: state.tripName,
    destination: state.destination,
    dateFrom: state.dateFrom,
    dateTo: state.dateTo,
    tripDates: state.tripDates,
    peopleCount: state.peopleCount,
    people: state.people,
    description: state.description,
    activities: state.activities,
    creationPhase: state.creationPhase,
    tripItems: state.tripItems,
    tasks: state.tasks,
    lastView: safeLastView,
    updatedAt: new Date().toISOString(),
  };
}

function syncCurrentRecord(state: AppState, lastView: View = "create"): AppState {
  const record = currentAsRecord(state, lastView);
  if (!record) return state;
  const exists = state.packlists.some((entry) => entry.id === record.id);
  return { ...state, packlists: exists ? state.packlists.map((entry) => entry.id === record.id ? record : entry) : [...state.packlists, record] };
}

function loadRecord(state: AppState, record: PacklistRecord): AppState {
  return {
    ...state,
    activePacklistId: record.id,
    creationPhase: record.creationPhase,
    tripName: record.tripName,
    tripDates: record.tripDates,
    destination: record.destination,
    dateFrom: record.dateFrom,
    dateTo: record.dateTo,
    peopleCount: record.peopleCount,
    description: record.description,
    activities: record.activities,
    people: record.people,
    tripItems: record.tripItems,
    tasks: record.tasks,
  };
}

function migrateState(value: unknown): AppState {
  if (!value || typeof value !== "object" || !("schemaVersion" in value)) return initialState;
  const saved = value as Partial<AppState> & { schemaVersion: number };
  if (saved.schemaVersion < 2 || !Array.isArray(saved.tripItems)) return initialState;
  const migrated: AppState = {
    ...initialState,
    ...saved,
    schemaVersion: 9,
    creationPhase: saved.creationPhase === "quantities" || saved.creationPhase === "assignment" || saved.creationPhase === "complete" ? saved.creationPhase : "selection",
    destination: typeof saved.destination === "string" ? saved.destination : "Schottland",
    dateFrom: typeof saved.dateFrom === "string" ? saved.dateFrom : "",
    dateTo: typeof saved.dateTo === "string" ? saved.dateTo : "",
    peopleCount: typeof saved.peopleCount === "number" ? saved.peopleCount : (Array.isArray(saved.people) ? saved.people.length : 3),
    description: typeof saved.description === "string" ? saved.description : "",
    activities: Array.isArray(saved.activities) ? saved.activities : ["Wandern", "Stadtbesichtigung"],
    people: Array.isArray(saved.people) ? saved.people : ["Clemens", "Florian", "Carina"],
    tripItems: saved.tripItems.map(normalizeTripItem),
    tasks: Array.isArray(saved.tasks) ? saved.tasks : initialState.tasks,
    customCatalogItems: Array.isArray(saved.customCatalogItems) ? saved.customCatalogItems : [],
    catalogCategoryOverrides: saved.catalogCategoryOverrides && typeof saved.catalogCategoryOverrides === "object" ? saved.catalogCategoryOverrides : {},
    deletedCatalogItemIds: Array.isArray(saved.deletedCatalogItemIds) ? saved.deletedCatalogItemIds : [],
  };
  if (saved.schemaVersion >= 6 && Array.isArray(saved.packlists)) {
    migrated.packlists = saved.packlists.map((record) => {
      const normalizedRecord = {
        ...record,
        people: normalizePeople(record.people, record.peopleCount),
        creationPhase: record.creationPhase === "quantities" || record.creationPhase === "assignment" || record.creationPhase === "complete" ? record.creationPhase : "selection",
        tripItems: Array.isArray(record.tripItems) ? record.tripItems.map(normalizeTripItem) : [],
      } as PacklistRecord;
      return saved.schemaVersion < 9 ? mergeScotlandAdditions(normalizedRecord) : normalizedRecord;
    });
    migrated.activePacklistId = typeof saved.activePacklistId === "string" ? saved.activePacklistId : saved.packlists[0]?.id ?? null;
    const active = migrated.packlists.find((record) => record.id === migrated.activePacklistId);
    if (active) return loadRecord(migrated, active);
    return migrated;
  }
  const legacyRecord = currentAsRecord({ ...migrated, activePacklistId: "packlist-scotland-2026", packlists: [] });
  migrated.activePacklistId = legacyRecord?.id ?? null;
  migrated.packlists = legacyRecord ? [legacyRecord] : [];
  return migrated;
}

function mergeScotlandAdditions(record: PacklistRecord): PacklistRecord {
  if (record.id !== initialScotlandRecord.id) return record;
  const existingIds = new Set(record.tripItems.flatMap((item) => [item.id, item.itemId]));
  const existingNames = new Set(record.tripItems.map((item) => item.name.toLocaleLowerCase("de")));
  const additions = scotlandItems.slice(96).filter((item) => !existingIds.has(item.id) && !existingIds.has(item.itemId) && !existingNames.has(item.name.toLocaleLowerCase("de")));
  const existingTaskIds = new Set(record.tasks.map((task) => task.id));
  const taskAdditions = initialScotlandRecord.tasks.filter((task) => !existingTaskIds.has(task.id));
  return additions.length || taskAdditions.length
    ? { ...record, tripItems: [...record.tripItems, ...additions], tasks: [...record.tasks, ...taskAdditions] }
    : record;
}

function normalizePeople(people: string[] | undefined, peopleCount: number) {
  if (Array.isArray(people) && people.length === peopleCount && !people.some((person) => /^\d+ Personen?$/.test(person))) return people;
  return Array.from({ length: Math.max(1, peopleCount || 1) }, (_, index) => `Person ${index + 1}`);
}

function normalizeTripItem(item: TripItem): TripItem {
  const sourceAssignments = Array.isArray(item.assignments) && item.assignments.length
    ? item.assignments
    : [{ person: "Noch nicht zugeordnet", quantities: { intended: 1, prepared: 0, packed: 0, used: 0 } }];
  const inferredScope: ItemScope = item.scope ?? (sourceAssignments.length > 1 && sourceAssignments.every((assignment) => assignment.person !== "Alle") ? "personal" : "shared");
  return {
    ...item,
    scope: inferredScope,
    assignments: sourceAssignments.map((assignment) => ({
      ...assignment,
      person: assignment.person === "Alle" ? "Noch nicht zugeordnet" : assignment.person,
    })),
  };
}

function Icon({ name }: { name: "list" | "prepare" | "bag" | "used" | "empty" | "grid" | "chevron" | "task" | "settings" | "help" | "search" | "sync" | "close" | "menu" | "plus" | "restore" | "download" | "upload" | "pdf" | "table" | "shield" | "trash" }) {
  const paths: Record<string, React.ReactNode> = {
    list: <><path d="M8 6h12M8 12h12M8 18h12"/><path d="m3.5 6 1 1 2-2M3.5 12l1 1 2-2M3.5 18l1 1 2-2"/></>,
    prepare: <><path d="M5 7h14v13H5z"/><path d="M8 7V4h8v3M9 12h6M9 16h4"/></>,
    bag: <><path d="M6 8h12l1 12H5L6 8Z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/></>,
    used: <><path d="M5 12.5 10 17 20 6"/><path d="M20 12v7H4V5h11"/></>,
    empty: <rect x="4.5" y="4.5" width="15" height="15" rx="2"/>,
    grid: <><rect x="4" y="4" width="6" height="6"/><rect x="14" y="4" width="6" height="6"/><rect x="4" y="14" width="6" height="6"/><rect x="14" y="14" width="6" height="6"/></>,
    chevron: <path d="m9 5 7 7-7 7"/>,
    task: <><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V2h6v2M8 10l2 2 4-4M8 17h7"/></>,
    settings: <><circle cx="12" cy="12" r="3"/><path d="M19 12a7 7 0 0 0-.1-1l2-1.6-2-3.4-2.5 1A8 8 0 0 0 15 6l-.4-2.7h-4L10 6a8 8 0 0 0-1.5.9L6 6 4 9.4 6.1 11a7 7 0 0 0 0 2L4 14.6 6 18l2.5-1A8 8 0 0 0 10 18l.5 2.7h4L15 18a8 8 0 0 0 1.5-.9L19 18l2-3.4-2.1-1.6a7 7 0 0 0 .1-1Z"/></>,
    help: <><circle cx="12" cy="12" r="10"/><path d="M9.5 9a2.7 2.7 0 1 1 3.7 2.5c-.8.4-1.2.9-1.2 1.8M12 17h.01"/></>,
    search: <><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/></>,
    sync: <><path d="M20 7h-5V2M4 17h5v5"/><path d="M18.5 9A7 7 0 0 0 6.2 6.2L4 8M5.5 15A7 7 0 0 0 17.8 17.8L20 16"/></>,
    close: <path d="m6 6 12 12M18 6 6 18"/>,
    menu: <path d="M4 7h16M4 12h16M4 17h16"/>,
    plus: <path d="M12 5v14M5 12h14"/>,
    restore: <><path d="M4 8V4h4"/><path d="M5.5 6.5A8 8 0 1 1 4 13"/></>,
    download: <><path d="M12 3v12"/><path d="m7 10 5 5 5-5M5 21h14"/></>,
    upload: <><path d="M12 16V4"/><path d="m7 9 5-5 5 5M5 21h14"/></>,
    pdf: <><path d="M6 2h8l4 4v16H6z"/><path d="M14 2v5h5M8 16h8M8 12h5"/></>,
    table: <><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M9 9v11M15 9v11"/></>,
    shield: <><path d="M12 3 20 6v6c0 5-3.4 8.1-8 9-4.6-.9-8-4-8-9V6l8-3Z"/><path d="m8.5 12 2.2 2.2 4.8-5"/></>,
    trash: <><path d="M4 7h16M9 7V4h6v3M7 7l1 14h8l1-14M10 11v6M14 11v6"/></>,
  };
  return <svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">{paths[name]}</svg>;
}

function clamp(value: number) { return Math.max(0, Math.round(Number(value) || 0)); }
function stageIsDone(item: TripItem, stage: Stage) { return item.assignments.every((assignment) => assignment.quantities[stage] >= assignment.quantities.intended); }
function assignmentSummary(item: TripItem) { return item.assignments.map((assignment) => `${assignment.person} ${assignment.quantities.intended}`).join(" · "); }
function intendedTotal(item: TripItem) { return item.assignments.reduce((sum, assignment) => sum + assignment.quantities.intended, 0); }

function safeFilename(value: string) {
  return value.trim().replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-").replace(/\s+/g, "_").replace(/_+/g, "_") || "Packliste";
}

function dateStamp() {
  const now = new Date();
  const part = (value: number) => String(value).padStart(2, "0");
  return `${now.getFullYear()}${part(now.getMonth() + 1)}${part(now.getDate())}-${part(now.getHours())}${part(now.getMinutes())}`;
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportRows(record: PacklistRecord): ExportRow[] {
  return [...record.tripItems]
    .sort((left, right) => left.category.localeCompare(right.category, "de") || left.name.localeCompare(right.name, "de"))
    .flatMap((item) => item.assignments.map((assignment) => ({
      category: item.category,
      item: item.name,
      quantity: assignment.quantities.intended,
      unit: item.unit,
      person: assignment.person,
      scope: item.scope === "personal" ? "Persönlich" : "Gemeinsam",
      priority: item.priority,
      prepared: assignment.quantities.prepared,
      packed: assignment.quantities.packed,
      used: assignment.quantities.used,
      notes: item.notes,
    })));
}

async function createPdf(record: PacklistRecord) {
  const [{ jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const rows = exportRows(record);
  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setTextColor(8, 45, 73);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text(record.tripName || "Packliste", 12, 14);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(80, 92, 98);
  doc.text(`Ziel: ${record.destination || "noch offen"}   |   Zeitraum: ${record.tripDates || "noch offen"}   |   Personen: ${record.people.join(", ") || record.peopleCount}`, 12, 20);
  doc.text("Zum Abhaken mit Stift: H = hergerichtet, E = eingepackt, G = tatsächlich gebraucht", 12, 25);

  autoTable(doc, {
    startY: 29,
    margin: { top: 15, right: 12, bottom: 13, left: 12 },
    theme: "grid",
    head: [["Kategorie", "Gegenstand", "Menge", "Für / verantwortlich", "Priorität", "H", "E", "G", "Notiz"]],
    body: rows.map((row) => [row.category, row.item, `${row.quantity} ${row.unit}`, row.person, row.priority, "", "", "", row.notes]),
    styles: { font: "helvetica", fontSize: 8, cellPadding: 1.8, lineColor: [210, 214, 207], lineWidth: 0.2, textColor: [24, 38, 45], overflow: "linebreak", valign: "middle" },
    headStyles: { fillColor: [8, 45, 73], textColor: [255, 255, 255], fontStyle: "bold", halign: "left" },
    alternateRowStyles: { fillColor: [247, 249, 244] },
    columnStyles: {
      0: { cellWidth: 31 }, 1: { cellWidth: 52 }, 2: { cellWidth: 18, halign: "center" },
      3: { cellWidth: 36 }, 4: { cellWidth: 22 }, 5: { cellWidth: 11, halign: "center" },
      6: { cellWidth: 11, halign: "center" }, 7: { cellWidth: 11, halign: "center" }, 8: { cellWidth: 76 },
    },
    didDrawCell: (data) => {
      if (data.section !== "body" || ![5, 6, 7].includes(data.column.index)) return;
      const size = 4;
      doc.setDrawColor(70, 86, 88);
      doc.rect(data.cell.x + (data.cell.width - size) / 2, data.cell.y + (data.cell.height - size) / 2, size, size);
    },
    didDrawPage: () => {
      const pageNumber = doc.getNumberOfPages();
      doc.setFontSize(7.5);
      doc.setTextColor(105, 110, 112);
      doc.text(`P03 Packliste · ${APP_VERSION} · Seite ${pageNumber}`, pageWidth - 12, doc.internal.pageSize.getHeight() - 6, { align: "right" });
    },
  });

  doc.save(`P03_${safeFilename(record.tripName)}_Packliste_${dateStamp()}.pdf`);
}

async function createExcel(record: PacklistRecord) {
  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "P03 Packliste";
  workbook.created = new Date();
  const navy = "082D49";
  const green = "39734B";
  const lightGreen = "E7EFE5";
  const borderColor = "DDD8CF";

  const overview = workbook.addWorksheet("Reise", { views: [{ showGridLines: false }] });
  overview.columns = [{ width: 24 }, { width: 72 }];
  overview.mergeCells("A1:B1");
  overview.getCell("A1").value = record.tripName || "Packliste";
  overview.getCell("A1").font = { bold: true, size: 18, color: { argb: "FFFFFFFF" } };
  overview.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${navy}` } };
  overview.getCell("A1").alignment = { vertical: "middle" };
  overview.getRow(1).height = 30;
  const overviewRows = [
    ["Ziel", record.destination || "Noch offen"], ["Zeitraum", record.tripDates || "Noch offen"],
    ["Personen", record.people.join(", ") || String(record.peopleCount)], ["Aktivitäten", record.activities.join(", ") || "Keine ausgewählt"],
    ["Beschreibung", record.description || ""], ["Gegenstände", record.tripItems.length], ["Exportiert", new Date()], ["App-Version", APP_VERSION],
  ];
  overview.addRows(overviewRows);
  overview.getColumn(1).font = { bold: true, color: { argb: `FF${green}` } };
  overview.getCell("B8").numFmt = "yyyy-mm-dd hh:mm";
  overview.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.alignment = { vertical: "top", wrapText: true };
      row.eachCell((cell) => { cell.border = { bottom: { style: "thin", color: { argb: `FF${borderColor}` } } }; });
    }
  });

  const list = workbook.addWorksheet("Packliste", { views: [{ state: "frozen", ySplit: 1, showGridLines: false }] });
  list.autoFilter = "A1:K1";
  list.columns = [
    { header: "Kategorie", key: "category", width: 24 }, { header: "Gegenstand", key: "item", width: 34 },
    { header: "Menge", key: "quantity", width: 10 }, { header: "Einheit", key: "unit", width: 11 },
    { header: "Für / verantwortlich", key: "person", width: 24 }, { header: "Art", key: "scope", width: 14 },
    { header: "Priorität", key: "priority", width: 14 }, { header: "Hergerichtet", key: "prepared", width: 14 },
    { header: "Eingepackt", key: "packed", width: 13 }, { header: "Tatsächlich gebraucht", key: "used", width: 21 },
    { header: "Notizen", key: "notes", width: 44 },
  ];
  list.addRows(exportRows(record));
  list.getRow(1).height = 28;
  list.getRow(1).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${navy}` } };
    cell.alignment = { vertical: "middle", wrapText: true };
  });
  list.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.alignment = { vertical: "top", wrapText: true };
      if (rowNumber % 2 === 0) row.eachCell((cell) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${lightGreen}` } }; });
    }
    row.eachCell((cell) => { cell.border = { bottom: { style: "thin", color: { argb: `FF${borderColor}` } } }; });
  });
  [3, 8, 9, 10].forEach((column) => { list.getColumn(column).numFmt = "0"; });

  const tasks = workbook.addWorksheet("Aufgaben", { views: [{ state: "frozen", ySplit: 1, showGridLines: false }] });
  tasks.autoFilter = "A1:E1";
  tasks.columns = [
    { header: "Bereich", key: "area", width: 24 }, { header: "Aufgabe", key: "label", width: 48 },
    { header: "Priorität", key: "priority", width: 15 }, { header: "Erledigt", key: "done", width: 12 }, { header: "Notizen", key: "notes", width: 44 },
  ];
  tasks.addRows(record.tasks.map((task) => ({ ...task, done: task.done ? "Ja" : "Nein" })));
  tasks.getRow(1).height = 28;
  tasks.getRow(1).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${navy}` } };
    cell.alignment = { vertical: "middle" };
  });
  tasks.eachRow((row, rowNumber) => { if (rowNumber > 1) row.alignment = { vertical: "top", wrapText: true }; });

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `P03_${safeFilename(record.tripName)}_Packliste_${dateStamp()}.xlsx`);
}

type Suggestion = { name: string; priority: Priority; activity: string; scope: ItemScope };
const basicSuggestions: Suggestion[] = [
  { name: "Reisepass", priority: "Unbedingt", activity: "Grundausstattung", scope: "personal" },
  { name: "E-Card", priority: "Unbedingt", activity: "Grundausstattung", scope: "personal" },
  { name: "Bankkarte + etwas Bargeld", priority: "Unbedingt", activity: "Grundausstattung", scope: "personal" },
  { name: "Persönliche Dauermedikamente", priority: "Unbedingt", activity: "Grundausstattung", scope: "personal" },
  { name: "Handy + Ladegerät", priority: "Unbedingt", activity: "Grundausstattung", scope: "personal" },
  { name: "Unterwäsche", priority: "Empfohlen", activity: "Grundausstattung", scope: "personal" },
  { name: "Socken", priority: "Empfohlen", activity: "Grundausstattung", scope: "personal" },
  { name: "T-Shirts", priority: "Empfohlen", activity: "Grundausstattung", scope: "personal" },
  { name: "Zahnbürste", priority: "Unbedingt", activity: "Grundausstattung", scope: "personal" },
  { name: "Zahnpasta", priority: "Empfohlen", activity: "Grundausstattung", scope: "shared" },
];
const activitySuggestions: Record<string, Suggestion[]> = {
  Wandern: [
    { name: "Wanderschuhe", priority: "Unbedingt", activity: "Wandern", scope: "personal" },
    { name: "Tagesrucksack", priority: "Unbedingt", activity: "Wandern", scope: "shared" },
    { name: "Regenjacke", priority: "Empfohlen", activity: "Wandern", scope: "personal" },
    { name: "Wandersocken", priority: "Empfohlen", activity: "Wandern", scope: "personal" },
  ],
  Baden: [
    { name: "Badebekleidung", priority: "Unbedingt", activity: "Baden", scope: "personal" },
    { name: "Badetuch", priority: "Empfohlen", activity: "Baden", scope: "personal" },
    { name: "Sonnencreme", priority: "Empfohlen", activity: "Baden", scope: "shared" },
  ],
  Stadtbesichtigung: [
    { name: "Bequeme Schuhe", priority: "Empfohlen", activity: "Stadtbesichtigung", scope: "personal" },
    { name: "Kleiner Tagesrucksack / Crossbody-Bag", priority: "Empfohlen", activity: "Stadtbesichtigung", scope: "shared" },
    { name: "Regenschirm", priority: "Optional", activity: "Stadtbesichtigung", scope: "shared" },
  ],
  Radfahren: [
    { name: "Radhelm", priority: "Unbedingt", activity: "Radfahren", scope: "personal" },
    { name: "Radhose", priority: "Empfohlen", activity: "Radfahren", scope: "personal" },
    { name: "Fahrradhandschuhe", priority: "Optional", activity: "Radfahren", scope: "personal" },
  ],
  Camping: [
    { name: "Zelt", priority: "Unbedingt", activity: "Camping", scope: "shared" },
    { name: "Schlafsack", priority: "Unbedingt", activity: "Camping", scope: "personal" },
    { name: "Isomatte", priority: "Empfohlen", activity: "Camping", scope: "personal" },
    { name: "Stirnlampe", priority: "Empfohlen", activity: "Camping", scope: "personal" },
  ],
  Wintersport: [
    { name: "Skihose", priority: "Unbedingt", activity: "Wintersport", scope: "personal" },
    { name: "Skibrille", priority: "Unbedingt", activity: "Wintersport", scope: "personal" },
    { name: "Handschuhe", priority: "Unbedingt", activity: "Wintersport", scope: "personal" },
  ],
  "Sport & Fitness": [
    { name: "Sportschuhe", priority: "Unbedingt", activity: "Sport & Fitness", scope: "personal" },
    { name: "Sporthose", priority: "Empfohlen", activity: "Sport & Fitness", scope: "personal" },
    { name: "Sportshirt", priority: "Empfohlen", activity: "Sport & Fitness", scope: "personal" },
  ],
  Veranstaltung: [
    { name: "Eintrittskarten", priority: "Unbedingt", activity: "Veranstaltung", scope: "personal" },
    { name: "Abendgarnitur", priority: "Optional", activity: "Veranstaltung", scope: "personal" },
  ],
};

function buildProposal(activities: string[], people: string[]): TripItem[] {
  const requests = [...basicSuggestions, ...activities.flatMap((activity) => activitySuggestions[activity] ?? [])];
  const used = new Set<string>();
  return requests.flatMap((suggestion) => {
    const exact = activeBaseCatalogItems.find((item) => item.name.toLocaleLowerCase("de") === suggestion.name.toLocaleLowerCase("de"));
    const fallback = activeBaseCatalogItems.find((item) => item.name.toLocaleLowerCase("de").includes(suggestion.name.toLocaleLowerCase("de")));
    const item = exact ?? fallback;
    if (!item || used.has(item.id)) return [];
    used.add(item.id);
    const id = `trip-${Date.now()}-${item.id}`;
    return [{
      id,
      itemId: id,
      catalogItemId: item.id,
      name: item.name,
      category: item.category,
      unit: item.units || "Stk",
      activity: suggestion.activity,
      priority: suggestion.priority,
      sourcePriority: suggestion.priority,
      luggage: "Noch offen",
      source: "Automatischer Grundvorschlag",
      notes: item.notes,
      scope: suggestion.scope,
      assignments: suggestion.scope === "personal"
        ? people.map((person) => ({ person, quantities: { intended: 1, prepared: 0, packed: 0, used: 0 } }))
        : [{ person: "Noch nicht zugeordnet", quantities: { intended: 1, prepared: 0, packed: 0, used: 0 } }],
    }];
  });
}

export default function PackingApp({ account }: { account: AccountInfo }) {
  const [state, setState] = useState<AppState>(blankState);
  const [view, setView] = useState<View>("trips");
  const [ready, setReady] = useState(false);
  const [saveStatus, setSaveStatus] = useState<"loading" | "saved" | "saving" | "offline">("loading");
  const [catalogSearch, setCatalogSearch] = useState("");
  const [catalogCategory, setCatalogCategory] = useState("Alle Kategorien");
  const [stageSearch, setStageSearch] = useState("");
  const [stageCategory, setStageCategory] = useState("Alle Kategorien");
  const [stageFilter, setStageFilter] = useState("Alle");
  const [stageSort, setStageSort] = useState<"category" | "person">("category");
  const [editingId, setEditingId] = useState<string | null>(null);
  const [catalogEditingId, setCatalogEditingId] = useState<string | null>(null);
  const [addingItem, setAddingItem] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [packMenuOpen, setPackMenuOpen] = useState(true);
  const [newPacklistOpen, setNewPacklistOpen] = useState(false);
  const [packlistEditOpen, setPacklistEditOpen] = useState(false);
  const [exportStatus, setExportStatus] = useState("");
  const [pendingRestore, setPendingRestore] = useState<PendingRestore | null>(null);
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const backupInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    loadOneDriveState()
      .then((loaded) => {
        if (loaded) {
          setState(migrateState(loaded));
        } else {
          setState(initialState); // erster Start: Schottland-Grunddaten als Startpunkt anbieten
        }
        setReady(true);
        setSaveStatus("saved");
      })
      .catch(() => { setState(blankState); setReady(false); setSaveStatus("offline"); });
  }, []);

  useEffect(() => {
    if (!ready) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      setSaveStatus("saving");
      const payload = syncCurrentRecord(state, view);
      saveOneDriveState(payload)
        .then(() => setSaveStatus("saved"))
        .catch(() => setSaveStatus("offline"));
    }, 500);
    return () => { if (saveTimer.current) clearTimeout(saveTimer.current); };
  }, [state, ready, view]);

  const effectiveCatalog = useMemo(() => [...activeBaseCatalogItems, ...state.customCatalogItems]
    .filter((item) => !state.deletedCatalogItemIds.includes(item.id))
    .map((item) => ({ ...item, category: state.catalogCategoryOverrides[item.id] ?? item.category })), [state.customCatalogItems, state.deletedCatalogItemIds, state.catalogCategoryOverrides]);
  const catalogCount = effectiveCatalog.length;
  const selectedByCatalogId = useMemo(() => new Map(state.tripItems.flatMap((item) => item.catalogItemId ? [[item.catalogItemId, item] as const] : [])), [state.tripItems]);
  const filteredCatalog = useMemo(() => effectiveCatalog.filter((item) => {
    const query = catalogSearch.trim().toLocaleLowerCase("de");
    return (catalogCategory === "Alle Kategorien" || item.category === catalogCategory) &&
      (!query || `${item.name} ${item.notes}`.toLocaleLowerCase("de").includes(query));
  }), [effectiveCatalog, catalogSearch, catalogCategory]);

  const activeStage: Stage | null = view === "prepare" ? "prepared" : view === "pack" ? "packed" : view === "experience" ? "used" : null;
  const tripCategories = useMemo(() => [...new Set(state.tripItems.map((item) => item.category))], [state.tripItems]);
  const filteredStageItems = useMemo(() => state.tripItems.filter((item) => {
    const query = stageSearch.trim().toLocaleLowerCase("de");
    const done = activeStage ? stageIsDone(item, activeStage) : false;
    const matchesStatus = stageFilter === "Alle" || (stageFilter === "Erledigt" ? done : !done);
    return matchesStatus && (stageCategory === "Alle Kategorien" || item.category === stageCategory) &&
      (!query || `${item.name} ${item.category} ${item.notes} ${item.activity}`.toLocaleLowerCase("de").includes(query));
  }), [state.tripItems, stageSearch, stageCategory, stageFilter, activeStage]);
  const stagePersonGroups = useMemo(() => {
    if (!activeStage) return [];
    const query = stageSearch.trim().toLocaleLowerCase("de");
    const entries = state.tripItems.flatMap((item) => {
      if ((stageCategory !== "Alle Kategorien" && item.category !== stageCategory) ||
        (query && !`${item.name} ${item.category} ${item.notes} ${item.activity}`.toLocaleLowerCase("de").includes(query))) return [];
      return item.assignments.flatMap((assignment) => {
        const done = assignment.quantities[activeStage] >= assignment.quantities.intended;
        const matchesStatus = stageFilter === "Alle" || (stageFilter === "Erledigt" ? done : !done);
        return matchesStatus ? [{ item, assignment }] : [];
      });
    });
    const knownPeople = [...state.people, "Noch nicht zugeordnet"];
    const allPeople = [...knownPeople, ...entries.map(({ assignment }) => assignment.person)]
      .filter((person, index, list) => person && list.indexOf(person) === index);
    return allPeople.map((person) => ({ person, entries: entries.filter((entry) => entry.assignment.person === person) }))
      .filter((group) => group.entries.length > 0);
  }, [state.tripItems, state.people, stageSearch, stageCategory, stageFilter, activeStage]);

  const editing = editingId ? state.tripItems.find((item) => item.id === editingId) ?? null : null;
  const catalogEditing = catalogEditingId ? effectiveCatalog.find((item) => item.id === catalogEditingId) ?? null : null;
  const completed = (stage: Stage) => state.tripItems.filter((item) => stageIsDone(item, stage)).length;
  const summaryStage: Stage = view === "pack" ? "packed" : view === "experience" ? "used" : "prepared";
  const summaryDone = completed(summaryStage);
  const summaryPercent = state.tripItems.length ? Math.round(summaryDone / state.tripItems.length * 100) : 0;
  const summaryLabel = summaryStage === "prepared" ? "hergerichtet" : summaryStage === "packed" ? "eingepackt" : "tatsächlich gebraucht";
  const summaryTitle = view === "create" ? "Nächster Schritt: Herrichten" : view === "prepare" ? "Herrichten" : view === "pack" ? "Einpacken" : "Tatsächlich gebraucht";
  const packlistsForDisplay = useMemo(() => {
    const current = currentAsRecord(state, view);
    if (!current) return state.packlists;
    return state.packlists.map((entry) => entry.id === current.id ? current : entry);
  }, [state, view]);

  function openPacklist(record: PacklistRecord) {
    setState((current) => loadRecord(syncCurrentRecord(current, view), record));
    setView(record.lastView || "create");
    setMenuOpen(false);
  }

  function createPacklist(values: { tripName: string; destination: string; dateFrom: string; dateTo: string; peopleCount: number; description: string; activities: string[] }) {
    const id = `packlist-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    const tripDates = values.dateFrom && values.dateTo ? `${values.dateFrom} bis ${values.dateTo}` : values.dateFrom || values.dateTo || "Zeitraum noch offen";
    const people = normalizePeople(undefined, values.peopleCount);
    const record: PacklistRecord = {
      id,
      tripName: values.tripName,
      destination: values.destination,
      dateFrom: values.dateFrom,
      dateTo: values.dateTo,
      tripDates,
      peopleCount: values.peopleCount,
      people,
      description: values.description,
      activities: values.activities,
      creationPhase: "selection",
      tripItems: buildProposal(values.activities, people),
      tasks: [],
      lastView: "create",
      updatedAt: new Date().toISOString(),
    };
    setState((current) => {
      const synced = syncCurrentRecord(current, view);
      return loadRecord({ ...synced, packlists: [...synced.packlists, record] }, record);
    });
    setReady(true);
    setNewPacklistOpen(false);
    setView("create");
  }

  function updatePacklistDetails(values: PacklistDetails) {
    setState((current) => {
      const previousPeople = current.people;
      const people = values.people.map((person) => person.trim()).filter(Boolean);
      const safePeople = people.length ? people : ["Person 1"];
      const renamedPeople = new Map(previousPeople.map((person, index) => [person, safePeople[index] ?? null]));
      const tripItems = current.tripItems.map((item) => {
        if (item.scope === "personal") {
          const template = item.assignments[0]?.quantities ?? { intended: 1, prepared: 0, packed: 0, used: 0 };
          return {
            ...item,
            assignments: safePeople.map((person, index) => {
              const previousName = previousPeople[index];
              const existing = item.assignments.find((assignment) => assignment.person === previousName)
                ?? item.assignments.find((assignment) => assignment.person === person);
              return existing
                ? { ...existing, person }
                : { person, quantities: { ...template, prepared: 0, packed: 0, used: 0 } };
            }),
          };
        }
        return {
          ...item,
          assignments: item.assignments.map((assignment) => {
            if (assignment.person === "Noch nicht zugeordnet") return assignment;
            return { ...assignment, person: renamedPeople.get(assignment.person) ?? "Noch nicht zugeordnet" };
          }),
        };
      });
      const tripDates = values.dateFrom && values.dateTo
        ? `${values.dateFrom} bis ${values.dateTo}`
        : values.dateFrom || values.dateTo || "Zeitraum noch offen";
      return {
        ...current,
        tripName: values.tripName,
        destination: values.destination,
        dateFrom: values.dateFrom,
        dateTo: values.dateTo,
        tripDates,
        description: values.description,
        activities: values.activities,
        people: safePeople,
        peopleCount: safePeople.length,
        tripItems,
      };
    });
    setPacklistEditOpen(false);
  }

  function deletePacklist(record: PacklistRecord) {
    if (!window.confirm(`Packliste „${record.tripName}“ wirklich löschen? Dieser Vorgang kann nicht rückgängig gemacht werden.`)) return;
    setState((current) => {
      const synced = syncCurrentRecord(current, view);
      const remaining = synced.packlists.filter((entry) => entry.id !== record.id);
      if (synced.activePacklistId !== record.id) return { ...synced, packlists: remaining };
      if (remaining.length) return loadRecord({ ...synced, packlists: remaining }, remaining[0]);
      return { ...synced, packlists: [], activePacklistId: null, tripName: "", tripDates: "", destination: "", dateFrom: "", dateTo: "", peopleCount: 1, description: "", activities: [], people: [], tripItems: [], tasks: [], creationPhase: "selection" };
    });
    setView("trips");
  }

  function updateTripItem(id: string, update: (item: TripItem) => TripItem) {
    setState((current) => ({ ...current, tripItems: current.tripItems.map((item) => item.id === id ? update(item) : item) }));
  }

  function updateQuantity(itemId: string, person: Person, field: keyof Quantities, value: number) {
    updateTripItem(itemId, (current) => ({
      ...current,
      assignments: current.assignments.map((assignment) => assignment.person === person
        ? { ...assignment, quantities: { ...assignment.quantities, [field]: clamp(value) } }
        : assignment),
    }));
  }

  function updateIntendedTotal(itemId: string, value: number) {
    const total = Math.max(1, clamp(value));
    updateTripItem(itemId, (current) => {
      const assignments = current.assignments.length ? current.assignments : [{ person: "Noch nicht zugeordnet", quantities: { intended: 1, prepared: 0, packed: 0, used: 0 } }];
      const base = Math.floor(total / assignments.length);
      const remainder = total % assignments.length;
      return {
        ...current,
        assignments: assignments.map((assignment, index) => ({
          ...assignment,
          quantities: { ...assignment.quantities, intended: base + (index < remainder ? 1 : 0) },
        })),
      };
    });
  }

  function choosePriority(itemId: string, priority: Priority) {
    const selected = selectedByCatalogId.get(itemId);
    if (selected) {
      if (selected.priority === priority) {
        removeCatalogItem(itemId);
        return;
      }
      updateTripItem(selected.id, (item) => ({ ...item, priority }));
      return;
    }
    const item = effectiveCatalog.find((entry) => entry.id === itemId);
    if (!item) return;
    const id = `trip-custom-${item.id}`;
    setState((current) => ({ ...current, tripItems: [...current.tripItems, {
      id,
      itemId: id,
      catalogItemId: item.id,
      name: item.name,
      category: item.category,
      unit: item.units || "Stk",
      activity: "",
      priority,
      sourcePriority: priority,
      luggage: "Noch offen",
      source: "Maximalliste",
      notes: item.notes,
      scope: "shared",
      assignments: [{ person: "Noch nicht zugeordnet", quantities: { intended: 1, prepared: 0, packed: 0, used: 0 } }],
    }] }));
  }

  function removeCatalogItem(catalogItemId: string) {
    setState((current) => ({ ...current, tripItems: current.tripItems.filter((item) => item.catalogItemId !== catalogItemId) }));
  }

  function addNewCatalogItem(values: { name: string; category: string; unit: string; notes: string; priority: Priority }) {
    const id = `CUSTOM-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    const catalogItem: CatalogItem = { id, categoryId: "CUSTOM", category: values.category, name: values.name, units: values.unit, notes: values.notes, reviewStatus: "Eigene Ergänzung", archived: false };
    const tripItem: TripItem = {
      id: `trip-${id}`, itemId: `trip-${id}`, catalogItemId: id, name: values.name, category: values.category,
      unit: values.unit || "Stk", activity: "", priority: values.priority, sourcePriority: values.priority,
      luggage: "Noch offen", source: "Eigene Ergänzung", notes: values.notes,
      scope: "shared",
      assignments: [{ person: "Noch nicht zugeordnet", quantities: { intended: 1, prepared: 0, packed: 0, used: 0 } }],
    };
    setState((current) => ({ ...current, customCatalogItems: [...current.customCatalogItems, catalogItem], tripItems: [...current.tripItems, tripItem] }));
    setAddingItem(false);
    setEditingId(tripItem.id);
  }

  function deleteFromCatalog(item: CatalogItem) {
    if (!window.confirm(`„${item.name}“ wirklich aus der Maximalliste löschen? Der Gegenstand kann wiederhergestellt werden.`)) return;
    setState((current) => ({
      ...current,
      deletedCatalogItemIds: [...new Set([...current.deletedCatalogItemIds, item.id])],
      tripItems: current.tripItems.filter((entry) => entry.catalogItemId !== item.id),
    }));
    setCatalogEditingId(null);
  }

  function restoreDeletedCatalogItems() {
    setState((current) => ({ ...current, deletedCatalogItemIds: [] }));
  }

  function setCatalogItemCategory(itemId: string, category: string) {
    setState((current) => ({
      ...current,
      catalogCategoryOverrides: { ...current.catalogCategoryOverrides, [itemId]: category },
      tripItems: current.tripItems.map((item) => item.catalogItemId === itemId ? { ...item, category } : item),
    }));
  }

  function toggleStage(item: TripItem, stage: Stage) {
    const done = stageIsDone(item, stage);
    updateTripItem(item.id, (current) => ({
      ...current,
      assignments: current.assignments.map((assignment) => ({
        ...assignment,
        quantities: { ...assignment.quantities, [stage]: done ? 0 : assignment.quantities.intended },
      })),
    }));
  }

  function toggleAssignmentStage(itemId: string, person: Person, stage: Stage) {
    updateTripItem(itemId, (current) => ({
      ...current,
      assignments: current.assignments.map((assignment) => assignment.person === person
        ? { ...assignment, quantities: { ...assignment.quantities, [stage]: assignment.quantities[stage] >= assignment.quantities.intended ? 0 : assignment.quantities.intended } }
        : assignment),
    }));
  }

  function setItemScope(itemId: string, scope: ItemScope) {
    setState((current) => ({
      ...current,
      tripItems: current.tripItems.map((item) => {
        if (item.id !== itemId || item.scope === scope) return item;
        const total = Math.max(1, intendedTotal(item));
        if (scope === "shared") {
          return { ...item, scope, assignments: [{ person: "Noch nicht zugeordnet", quantities: { intended: total, prepared: 0, packed: 0, used: 0 } }] };
        }
        const people = current.people.length ? current.people : ["Person 1"];
        const base = Math.max(1, Math.floor(total / people.length));
        const remainder = Math.max(0, total - base * people.length);
        return {
          ...item,
          scope,
          assignments: people.map((person, index) => ({ person, quantities: { intended: base + (index < remainder ? 1 : 0), prepared: 0, packed: 0, used: 0 } })),
        };
      }),
    }));
  }

  function setSharedCarrier(itemId: string, person: string) {
    updateTripItem(itemId, (item) => ({ ...item, assignments: item.assignments.map((assignment, index) => index === 0 ? { ...assignment, person } : assignment) }));
  }

  function renamePerson(index: number, nextName: string) {
    setState((current) => {
      const previous = current.people[index];
      const name = nextName;
      const people = current.people.map((person, personIndex) => personIndex === index ? name : person);
      const tripItems = current.tripItems.map((item) => ({
        ...item,
        assignments: item.assignments.map((assignment) => assignment.person === previous ? { ...assignment, person: name } : assignment),
      }));
      return { ...current, people, tripItems };
    });
  }

  function downloadBackup() {
    const completeState = syncCurrentRecord(state, view);
    const payload = {
      format: "P03_PACKLIST_BACKUP",
      formatVersion: 1,
      appVersion: APP_VERSION,
      exportedAt: new Date().toISOString(),
      data: completeState,
    };
    downloadBlob(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" }), `P03_Packliste_Datensicherung_${dateStamp()}_AI.json`);
    setExportStatus(`Vollständige Datensicherung mit ${completeState.packlists.length} Packlisten wurde heruntergeladen.`);
  }

  async function restoreBackup(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as { format?: string; data?: unknown; exportedAt?: string } | AppState;
      const candidate = "format" in parsed && parsed.format === "P03_PACKLIST_BACKUP" ? parsed.data : parsed;
      if (!candidate || typeof candidate !== "object" || !("schemaVersion" in candidate) || !("packlists" in candidate) || !Array.isArray((candidate as AppState).packlists)) {
        throw new Error("Diese Datei ist keine gültige P03-Datensicherung.");
      }
      const restored = migrateState(candidate);
      const backupDate = "exportedAt" in parsed && parsed.exportedAt ? new Date(parsed.exportedAt).toLocaleString("de-AT") : "unbekannt";
      setPendingRestore({ state: restored, backupDate });
    } catch (error) {
      setExportStatus(error instanceof Error ? error.message : "Die Sicherung konnte nicht gelesen werden.");
    }
  }

  function applyRestoredBackup() {
    if (!pendingRestore) return;
    const restored = pendingRestore.state;
    setState(restored);
    setReady(true);
    setView(restored.activePacklistId ? "trips" : "settings");
    setExportStatus(`Sicherung erfolgreich wiederhergestellt: ${restored.packlists.length} ${restored.packlists.length === 1 ? "Packliste wurde" : "Packlisten wurden"} übernommen.`);
    setPendingRestore(null);
  }

  async function downloadPdf() {
    const record = currentAsRecord(state, view);
    if (!record) return;
    setExportStatus("PDF wird erstellt …");
    try {
      await createPdf(record);
      setExportStatus(`PDF für „${record.tripName}“ wurde heruntergeladen.`);
    } catch {
      setExportStatus("Das PDF konnte nicht erstellt werden.");
    }
  }

  async function downloadExcel() {
    const record = currentAsRecord(state, view);
    if (!record) return;
    setExportStatus("Excel-Datei wird erstellt …");
    try {
      await createExcel(record);
      setExportStatus(`Excel-Datei für „${record.tripName}“ wurde heruntergeladen.`);
    } catch {
      setExportStatus("Die Excel-Datei konnte nicht erstellt werden.");
    }
  }

  async function deleteAllStoredData() {
    setExportStatus("Gespeicherte Daten werden gelöscht …");
    try {
      await deleteOneDriveState();
      setReady(false);
      setState(blankState);
      setView("settings");
      setDeleteDialogOpen(false);
      setExportStatus("Alle in der App gespeicherten Packlisten und persönlichen Angaben wurden gelöscht.");
    } catch {
      setDeleteDialogOpen(false);
      setExportStatus("Die gespeicherten Daten konnten nicht gelöscht werden.");
    }
  }

  const workflowNav = [
    { id: "create" as const, label: "Liste erstellen", short: "Liste", icon: "list" as const },
    { id: "prepare" as const, label: "Herrichten", short: "Herrichten", icon: "prepare" as const },
    { id: "pack" as const, label: "Einpacken", short: "Einpacken", icon: "used" as const },
    { id: "experience" as const, label: "Tatsächlich gebraucht", short: "Gebraucht", icon: "empty" as const },
  ];
  const workflowActive = workflowNav.some((entry) => entry.id === view);
  const hasActivePacklist = Boolean(state.activePacklistId);

  const stageCopy = activeStage ? {
    prepared: { step: "Schritt 2", title: "Herrichten", description: "Hake alles ab, was bereits bereitliegt.", action: "hergerichtet" },
    packed: { step: "Schritt 3", title: "Einpacken", description: "Hake erst beim tatsächlichen Einpacken ab.", action: "eingepackt" },
    used: { step: "Schritt 4 · nach der Reise", title: "Tatsächlich gebraucht", description: "Damit wird die nächste Packliste besser.", action: "gebraucht" },
  }[activeStage] : null;

  return (
    <main className="app-shell">
      <aside className={`sidebar ${menuOpen ? "open" : ""}`}>
        <button className="brand" onClick={() => { setView("trips"); setMenuOpen(false); }}><span className="brand-mark"><Icon name="list" /></span><strong>P03 Packliste</strong></button>
        <nav aria-label="Hauptnavigation">
          <button className={`nav-parent ${workflowActive || view === "trips" ? "active" : ""}`} aria-expanded={packMenuOpen} onClick={() => setPackMenuOpen((value) => !value)}><Icon name="bag" /><span>Packliste</span><span className={`nav-chevron ${packMenuOpen ? "open" : ""}`}><Icon name="chevron" /></span></button>
          {packMenuOpen && <div className="nav-submenu" aria-label="Packliste Arbeitsschritte">
            <button className={view === "trips" ? "active" : ""} onClick={() => { setView("trips"); setMenuOpen(false); }}><Icon name="list" /><span>Packliste auswählen</span></button>
            {workflowNav.map((entry) => <button key={entry.id} disabled={!hasActivePacklist} className={view === entry.id ? "active" : ""} onClick={() => { setView(entry.id); setMenuOpen(false); }}><Icon name={entry.icon} /><span>{entry.label}</span></button>)}
          </div>}
          <button className={view === "categories" ? "active" : ""} onClick={() => { setView("categories"); setMenuOpen(false); }}><Icon name="grid" /><span>Kategorien</span></button>
          <button disabled={!hasActivePacklist} className={view === "tasks" ? "active" : ""} onClick={() => { setView("tasks"); setMenuOpen(false); }}><Icon name="task" /><span>Vor der Reise</span></button>
          <button className={view === "settings" ? "active" : ""} onClick={() => { setView("settings"); setMenuOpen(false); }}><Icon name="settings" /><span>Daten &amp; Einstellungen</span></button>
        </nav>
        <button className="help-button"><Icon name="help" /><span>Hilfe</span></button>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <button className="menu-button" aria-label="Menü öffnen" onClick={() => setMenuOpen((value) => !value)}><Icon name="menu" /></button>
          <div className="topbar-title"><p className="eyebrow">{view === "trips" ? "Verwaltung" : view === "settings" ? "Sicherung und Ausgabe" : "Aktuelle Reise"}</p><h1>{view === "trips" ? "Packliste auswählen" : view === "settings" ? "Daten & Einstellungen" : state.tripName}</h1></div>
          <div className="topbar-actions">
            {hasActivePacklist && !["trips", "categories", "settings"].includes(view) && <button className="secondary compact-action topbar-edit" onClick={() => setPacklistEditOpen(true)}>Packliste bearbeiten</button>}
            <div className={`save-status ${saveStatus}`}><Icon name="sync" /><span>{saveStatus === "loading" ? "Lädt …" : saveStatus === "saving" ? "Speichert …" : saveStatus === "saved" ? "Arbeitsstand gespeichert" : "Nicht gespeichert"}</span></div>
          </div>
        </header>

        {view === "trips" && <section className="card page-card packlist-picker">
          <div className="section-heading picker-heading"><div><p className="eyebrow">Ihre Reisen</p><h2>Packliste auswählen</h2><p className="section-note">Eine vorhandene Liste am letzten Stand fortsetzen oder eine neue Reise anlegen.</p></div><button className="primary compact-action" onClick={() => setNewPacklistOpen(true)}><Icon name="plus" /> Neue Packliste</button></div>
          <div className="packlist-grid">{packlistsForDisplay.map((record) => <article className={`packlist-card ${record.id === state.activePacklistId ? "current" : ""}`} key={record.id}>
            <div className="packlist-card-head"><div><span className="packlist-status">{record.id === state.activePacklistId ? "Zuletzt geöffnet" : "Gespeichert"}</span><h3>{record.tripName}</h3></div><button className="packlist-delete" aria-label={`${record.tripName} löschen`} onClick={() => deletePacklist(record)}>Löschen</button></div>
            <dl><div><dt>Ziel</dt><dd>{record.destination || "Noch offen"}</dd></div><div><dt>Zeitraum</dt><dd>{record.tripDates || "Noch offen"}</dd></div><div><dt>Personen</dt><dd>{record.peopleCount}</dd></div><div><dt>Stand</dt><dd>{record.lastView === "create" ? "Liste erstellen" : record.lastView === "prepare" ? "Herrichten" : record.lastView === "pack" ? "Einpacken" : record.lastView === "experience" ? "Tatsächlich gebraucht" : "Vor der Reise"}</dd></div></dl>
            {record.activities.length > 0 && <div className="activity-chips">{record.activities.map((activity) => <span key={activity}>{activity}</span>)}</div>}
            <div className="packlist-card-footer"><span>{record.tripItems.length} Gegenstände</span><button className="primary compact-action" onClick={() => openPacklist(record)}>Öffnen und fortsetzen</button></div>
          </article>)}</div>
          {packlistsForDisplay.length === 0 && <div className="empty-state picker-empty"><strong>Noch keine Packliste vorhanden.</strong><span>Legen Sie Ihre erste Reise an; danach erstellt die App einen Grundvorschlag.</span></div>}
        </section>}

        {["create", "prepare", "pack", "experience"].includes(view) && <section className="workflow-overview card">
          <div className="progress-ring workflow-ring" style={{ "--progress": `${summaryPercent * 3.6}deg` } as React.CSSProperties}><span>{summaryPercent}<small>%</small></span></div>
          <div className="progress-copy workflow-progress-copy"><p className="eyebrow">{view === "create" ? "Stand nach der Listenerstellung" : `Schritt ${view === "prepare" ? 2 : view === "pack" ? 3 : 4}`}</p><h2>{summaryTitle}</h2><span className="mobile-progress-percent">{summaryPercent}%</span><p><strong>{summaryDone} von {state.tripItems.length}</strong> Gegenständen vollständig {summaryLabel}</p><div className="bar"><span style={{ width: `${summaryPercent}%` }} /></div></div>
          <div className="overview-trip-data"><p>Reise</p><strong>{state.tripName}</strong><span>{state.tripDates}</span><p>Ausgewählt</p><strong>{state.tripItems.length} Gegenstände</strong></div>
        </section>}

        {view === "create" && <section className="card page-card creation-page">
          <div className="section-heading creation-heading"><div><p className="eyebrow">Schritt 1</p><h2>{state.creationPhase === "selection" ? "Liste schnell erstellen" : state.creationPhase === "quantities" ? "Mengen festlegen" : "Personen zuordnen"}</h2><p className="section-note">{state.creationPhase === "selection" ? "Pro Gegenstand das passende Auswahlfeld markieren. Ein gesetztes Häkchen nochmals anklicken, um es zu entfernen." : state.creationPhase === "quantities" ? "Nur die ausgewählten Gegenstände: geplante Menge prüfen und bei Bedarf ändern." : "Persönliche Dinge gelten für die Reisenden einzeln. Gemeinsame Dinge werden einer verantwortlichen Person zugeordnet."}</p></div><div className="heading-actions"><button className="secondary compact-action" onClick={() => setAddingItem(true)}><Icon name="plus" /> Neuer Gegenstand</button>{state.deletedCatalogItemIds.length > 0 && <button className="restore-button" onClick={restoreDeletedCatalogItems}><Icon name="restore" /> {state.deletedCatalogItemIds.length} wiederherstellen</button>}<div className="selected-count"><strong>{state.tripItems.length}</strong><span>ausgewählt</span></div></div></div>
          <div className="creation-phase-tabs" role="tablist" aria-label="Teilabschnitte von Liste erstellen">
            <button role="tab" aria-selected={state.creationPhase === "selection"} onClick={() => setState((current) => ({ ...current, creationPhase: "selection" }))}>Schnellauswahl</button>
            <button role="tab" aria-selected={state.creationPhase === "quantities"} disabled={state.tripItems.length === 0} onClick={() => setState((current) => ({ ...current, creationPhase: "quantities" }))}>Mengen festlegen</button>
            <button role="tab" aria-selected={state.creationPhase === "assignment" || state.creationPhase === "complete"} disabled={state.tripItems.length === 0} onClick={() => setState((current) => ({ ...current, creationPhase: "assignment" }))}>Personen zuordnen</button>
          </div>
          {state.creationPhase === "selection" ? <>
            <div className="filters creation-filters">
              <label className="search"><Icon name="search" /><input value={catalogSearch} onChange={(event) => setCatalogSearch(event.target.value)} placeholder="Gegenstand suchen …" /></label>
              <select aria-label="Kategorie wählen" value={catalogCategory} onChange={(event) => setCatalogCategory(event.target.value)}><option>Alle Kategorien</option>{catalog.categories.map((entry) => <option key={entry.id}>{entry.name}</option>)}</select>
            </div>
            <p className="result-count">{filteredCatalog.length} von {catalogCount} Gegenständen · {state.tripItems.length} für diese Reise ausgewählt</p>
            <div className="selection-table">
              <div className="selection-header"><span>Gegenstand</span>{priorities.map((priority) => <span key={priority}>{priority}</span>)}</div>
              <div className="selection-list">{filteredCatalog.map((item) => {
                const selected = selectedByCatalogId.get(item.id);
                return <div className={`selection-row ${selected ? "selected" : ""}`} key={item.id}>
                  <button className="selection-item" onClick={() => selected ? setEditingId(selected.id) : setCatalogEditingId(item.id)}><strong>{item.name}</strong><small>{item.category}{item.units ? ` · ${item.units}` : ""}</small></button>
                  {priorities.map((priority) => <button key={priority} className={`priority-check ${priority.toLowerCase()} ${selected?.priority === priority ? "checked" : ""}`} aria-label={`${item.name}: ${priority}`} aria-pressed={selected?.priority === priority} onClick={() => choosePriority(item.id, priority)}><span className="priority-mobile-label">{priority}</span><span className="priority-checkmark">{selected?.priority === priority ? "✓" : ""}</span></button>)}
                </div>;
              })}</div>
            </div>
            {filteredCatalog.length === 0 && <p className="empty-state">Für diese Auswahl wurden keine Gegenstände gefunden.</p>}
            <div className="creation-footer"><span>{state.tripItems.length} {state.tripItems.length === 1 ? "Gegenstand ausgewählt" : "Gegenstände ausgewählt"}</span><button className="primary compact-action" disabled={state.tripItems.length === 0} onClick={() => setState((current) => ({ ...current, creationPhase: "quantities" }))}>Auswahl fertig</button></div>
          </> : state.creationPhase === "quantities" ? <>
            <div className="quantity-list">{state.tripItems.map((item) => {
              const total = intendedTotal(item);
              return <div className="quantity-row" key={item.id}>
                <button className="quantity-item" onClick={() => setEditingId(item.id)}><strong>{item.name}</strong><small>{item.category} · <span className={`priority-label ${item.priority.toLowerCase()}`}>{item.priority}</span>{item.assignments.length > 1 ? ` · ${assignmentSummary(item)}` : ""}</small></button>
                <div className="quantity-control" aria-label={`Geplante Menge für ${item.name}`}>
                  <button aria-label={`${item.name}: eins weniger`} onClick={() => updateIntendedTotal(item.id, total - 1)}>−</button>
                  <input aria-label={`${item.name}: geplante Menge`} type="number" min="1" inputMode="numeric" value={total} onChange={(event) => updateIntendedTotal(item.id, Number(event.target.value))} />
                  <button aria-label={`${item.name}: eins mehr`} onClick={() => updateIntendedTotal(item.id, total + 1)}>+</button>
                </div>
                <button className="remove-quantity" aria-label={`${item.name} aus der Reise entfernen`} onClick={() => setState((current) => ({ ...current, tripItems: current.tripItems.filter((entry) => entry.id !== item.id) }))}>×</button>
              </div>;
            })}</div>
            {state.tripItems.length === 0 && <p className="empty-state">Noch keine Gegenstände ausgewählt.</p>}
            <div className="creation-footer"><button className="secondary compact-action" onClick={() => setState((current) => ({ ...current, creationPhase: "selection" }))}>Zurück zur Auswahl</button><button className="primary compact-action" disabled={state.tripItems.length === 0} onClick={() => setState((current) => ({ ...current, creationPhase: "assignment" }))}>Weiter zu Personen</button></div>
          </> : <>
            <div className="people-editor">
              <div><strong>Reisende</strong><span>Die Namen werden bei Herrichten und Einpacken als eigene Listen verwendet.</span></div>
              <div className="people-name-grid">{state.people.map((person, index) => <label key={index}><span>Person {index + 1}</span><input value={person} onChange={(event) => renamePerson(index, event.target.value)} onBlur={(event) => { if (!event.target.value.trim()) renamePerson(index, `Person ${index + 1}`); }} /></label>)}</div>
            </div>
            <div className="assignment-table">
              <div className="assignment-header"><span>Gegenstand</span><span>Art</span><span>Zuordnung</span></div>
              {state.tripItems.map((item) => <div className="assignment-row" key={item.id}>
                <button className="assignment-item" onClick={() => setEditingId(item.id)}><strong>{item.name}</strong><small>{item.category} · Menge {intendedTotal(item)}</small></button>
                <div className="scope-toggle" role="group" aria-label={`${item.name}: Art der Zuordnung`}>
                  <button className={item.scope === "personal" ? "active" : ""} aria-pressed={item.scope === "personal"} onClick={() => setItemScope(item.id, "personal")}>Persönlich</button>
                  <button className={item.scope === "shared" ? "active" : ""} aria-pressed={item.scope === "shared"} onClick={() => setItemScope(item.id, "shared")}>Gemeinsam</button>
                </div>
                {item.scope === "shared" ? <label className="carrier-select"><span className="sr-only">Verantwortliche Person für {item.name}</span><select value={item.assignments[0]?.person ?? "Noch nicht zugeordnet"} onChange={(event) => setSharedCarrier(item.id, event.target.value)}><option>Noch nicht zugeordnet</option>{state.people.map((person) => <option key={person}>{person}</option>)}</select></label> : <button className="personal-summary" onClick={() => setEditingId(item.id)}>{assignmentSummary(item)}</button>}
              </div>)}
            </div>
            <div className="creation-footer"><button className="secondary compact-action" onClick={() => setState((current) => ({ ...current, creationPhase: "quantities" }))}>Zurück zu Mengen</button><button className="primary compact-action" disabled={state.tripItems.length === 0} onClick={() => { setState((current) => ({ ...current, creationPhase: "complete" })); setView("prepare"); }}>Zuordnung fertig</button></div>
          </>}
        </section>}

        {activeStage && stageCopy && <section className="card page-card stage-page">
          <div className="section-heading"><div><p className="eyebrow">{stageCopy.step}</p><h2>{stageCopy.title}</h2><p className="section-note">{stageCopy.description}</p></div><div className="selected-count"><strong>{completed(activeStage)}/{state.tripItems.length}</strong><span>{stageCopy.action}</span></div></div>
          <div className="stage-progress"><span style={{ width: `${state.tripItems.length ? completed(activeStage) / state.tripItems.length * 100 : 0}%` }} /></div>
          <div className="filters trip-filters">
            <label className="search"><Icon name="search" /><input value={stageSearch} onChange={(event) => setStageSearch(event.target.value)} placeholder={`Beim ${stageCopy.title.toLocaleLowerCase("de")} suchen …`} /></label>
            <select aria-label="Ansicht sortieren" value={stageSort} onChange={(event) => setStageSort(event.target.value as "category" | "person")}><option value="category">Nach Kategorien</option><option value="person">Nach Personen</option></select>
            <select value={stageCategory} onChange={(event) => setStageCategory(event.target.value)}><option>Alle Kategorien</option>{tripCategories.map((entry) => <option key={entry}>{entry}</option>)}</select>
            <select value={stageFilter} onChange={(event) => setStageFilter(event.target.value)}><option>Alle</option><option>Offen</option><option>Erledigt</option></select>
          </div>
          {stageSort === "category" ? <>
            <p className="result-count">{filteredStageItems.length} von {state.tripItems.length} ausgewählten Gegenständen werden angezeigt</p>
            <div className="stage-list">{filteredStageItems.map((item) => <StageRow key={item.id} item={item} stage={activeStage} onToggle={() => toggleStage(item, activeStage)} onEdit={() => setEditingId(item.id)} />)}</div>
            {filteredStageItems.length === 0 && <p className="empty-state">In dieser Auswahl ist nichts mehr offen.</p>}
          </> : <>
            <p className="result-count">Persönliche und übernommene gemeinsame Gegenstände sind nach Reisenden gruppiert.</p>
            <div className="person-stage-groups">{stagePersonGroups.map((group) => <section className={`person-stage-group ${group.person === "Noch nicht zugeordnet" ? "unassigned" : ""}`} key={group.person}>
              <div className="person-group-heading"><strong>{group.person}</strong><span>{group.entries.length} {group.entries.length === 1 ? "Gegenstand" : "Gegenstände"}</span></div>
              <div className="stage-list">{group.entries.map(({ item, assignment }) => <StageAssignmentRow key={`${item.id}-${assignment.person}`} item={item} assignment={assignment} stage={activeStage} onToggle={() => toggleAssignmentStage(item.id, assignment.person, activeStage)} onEdit={() => setEditingId(item.id)} />)}</div>
            </section>)}</div>
            {stagePersonGroups.length === 0 && <p className="empty-state">In dieser Auswahl ist nichts mehr offen.</p>}
          </>}
        </section>}

        {view === "categories" && <section className="card page-card category-page">
          <div className="section-heading creation-heading"><div><p className="eyebrow">Gegenstandsbasis</p><h2>Kategorien verwalten</h2><p className="section-note">Diese Zuordnung gilt unabhängig von einer einzelnen Reise.</p></div><div className="heading-actions"><button className="secondary compact-action" onClick={() => setAddingItem(true)}><Icon name="plus" /> Neuer Gegenstand</button>{state.deletedCatalogItemIds.length > 0 && <button className="restore-button" onClick={restoreDeletedCatalogItems}><Icon name="restore" /> {state.deletedCatalogItemIds.length} wiederherstellen</button>}</div></div>
          <div className="filters creation-filters">
            <label className="search"><Icon name="search" /><input value={catalogSearch} onChange={(event) => setCatalogSearch(event.target.value)} placeholder="Gegenstand suchen …" /></label>
            <select aria-label="Kategorie wählen" value={catalogCategory} onChange={(event) => setCatalogCategory(event.target.value)}><option>Alle Kategorien</option>{catalog.categories.map((entry) => <option key={entry.id}>{entry.name}</option>)}</select>
          </div>
          <p className="result-count">{filteredCatalog.length} von {catalogCount} Gegenständen</p>
          <div className="catalog-list">{filteredCatalog.map((item) => <div className="catalog-row" key={item.id}><div><strong><span className="category-dot" />{item.name}</strong><p>{item.category}{item.units ? ` · ${item.units}` : ""}</p></div><button onClick={() => setCatalogEditingId(item.id)}>Kategorie zuordnen</button></div>)}</div>
          {filteredCatalog.length === 0 && <p className="empty-state">Für diese Auswahl wurden keine Gegenstände gefunden.</p>}
        </section>}

        {view === "tasks" && <section className="card page-card narrow">
          <div className="section-heading"><div><p className="eyebrow">Getrennt von Packgegenständen</p><h2>Vor der Reise</h2></div><span className="counter">{state.tasks.filter((task) => !task.done).length} offen</span></div>
          <div className="task-list large">{state.tasks.map((task) => <TaskRow key={task.id} task={task} onToggle={() => setState((current) => ({ ...current, tasks: current.tasks.map((entry) => entry.id === task.id ? { ...entry, done: !entry.done } : entry) }))} showArea />)}</div>
        </section>}

        {view === "settings" && <section className="settings-layout">
          <section className="card page-card export-page">
            <div className="section-heading export-heading"><div><p className="eyebrow">Alle Daten</p><h2>Datenexport und Sicherung</h2><p className="section-note">Sicherung für die Wiederherstellung oder lesbare Dateien für Papier, Excel und Versand.</p></div></div>
            <div className="export-grid">
              <article className="export-card"><span className="export-icon"><Icon name="download" /></span><div><h3>Datensicherung</h3><p>Speichert alle Packlisten, Arbeitsstände und eigenen Ergänzungen in einer Datei.</p></div><button className="primary compact-action" onClick={downloadBackup}>Sicherung herunterladen</button></article>
              <article className="export-card"><span className="export-icon"><Icon name="upload" /></span><div><h3>Sicherung wiederherstellen</h3><p>Liest eine frühere P03-Sicherung ein. Erst nach Ihrer Bestätigung werden die aktuellen Daten ersetzt.</p></div><button className="secondary compact-action" onClick={() => backupInputRef.current?.click()}>Sicherung auswählen</button><input ref={backupInputRef} className="sr-only" type="file" accept="application/json,.json" onChange={restoreBackup} /></article>
              <article className={`export-card ${!hasActivePacklist ? "disabled" : ""}`}><span className="export-icon"><Icon name="pdf" /></span><div><h3>Packliste als PDF</h3><p>A4-Querformat mit freien Kästchen für Herrichten, Einpacken und Gebrauch - geeignet zum Ausdrucken und Abhaken mit Stift.</p></div><button disabled={!hasActivePacklist} className="primary compact-action" onClick={downloadPdf}>PDF herunterladen</button></article>
              <article className={`export-card ${!hasActivePacklist ? "disabled" : ""}`}><span className="export-icon"><Icon name="table" /></span><div><h3>Packliste als Excel</h3><p>Enthält Reiseübersicht, vollständige Packliste nach Personen und die Aufgaben vor der Reise.</p></div><button disabled={!hasActivePacklist} className="primary compact-action" onClick={downloadExcel}>Excel herunterladen</button></article>
            </div>
            {exportStatus && <div className="export-status" role="status">{exportStatus}</div>}
            <div className="notice"><strong>Die Dateien werden in den Download-Ordner des Browsers gespeichert.</strong><p>Von dort können Sie sie in Ihren gewünschten OneDrive-Ordner verschieben, ausdrucken oder verschicken. GitHub wird dafür nicht benötigt.</p></div>
          </section>

          <section className="card page-card privacy-page">
            <div className="section-heading export-heading"><div><p className="eyebrow">Technische Übersicht</p><h2>Datenschutz &amp; Datenverwaltung</h2><p className="section-note">Welche Daten die Packliste benötigt und wie Sie diese selbst kontrollieren.</p></div><span className="export-icon"><Icon name="shield" /></span></div>
            <div className="privacy-grid">
              <article className="privacy-fact"><strong>Gespeicherte Daten</strong><p>Packlisten, Reiseziele und -daten, Namen der Mitreisenden, Gegenstände, Mengen, Aufgaben, eigene Ergänzungen und Bearbeitungsstände.</p></article>
              <article className="privacy-fact"><strong>Speicherort</strong><p>Alle Daten liegen ausschließlich in Ihrem eigenen OneDrive im Ordner „P03_Packliste“ – nicht auf einem Server dieser App.</p></article>
              <article className="privacy-fact"><strong>Zweck und Dauer</strong><p>Die Daten werden für Speicherung und geräteübergreifende Synchronisierung verwendet und bleiben gespeichert, bis Sie sie hier oder direkt in OneDrive löschen.</p></article>
              <article className="privacy-fact"><strong>Keine zusätzlichen Tracker</strong><p>Die Packlisten-App selbst verwendet keine eigenen Werbe-, Analyse- oder Trackingdienste und übermittelt keine Packlistendaten an weitere Apps.</p></article>
            </div>
            <div className="privacy-advice"><strong>Angemeldet als {account.username}</strong><p><button type="button" className="secondary compact-action" onClick={() => logout()}>Abmelden</button></p></div>
            <div className="privacy-advice"><strong>Bitte keine besonders sensiblen Angaben eintragen.</strong><p>Für die Packliste genügen Bezeichnungen wie „Reisepass“ oder „Medikamente“. Passnummern, Gesundheitsdaten oder andere vertrauliche Inhalte gehören nicht in Notizfelder.</p></div>
            <div className="danger-zone"><div><strong>Alle gespeicherten App-Daten löschen</strong><p>Löscht Ihren vollständigen Datenbestand aus der Packlisten-App. Zuvor heruntergeladene Dateien auf Ihrem Gerät oder in OneDrive werden dadurch nicht gelöscht.</p></div><button className="danger-action" onClick={() => setDeleteDialogOpen(true)}><Icon name="trash" /> Alle Daten löschen</button></div>
            <p className="privacy-note">Diese Seite beschreibt die technische Datenverarbeitung der Packlisten-App. Eine formelle Datenschutzerklärung benötigt zusätzlich die Angaben des späteren Betreibers, insbesondere Name, Kontakt, Rechtsgrundlage und Hostinginformationen.</p>
          </section>

          {hasActivePacklist && <section className="card page-card narrow settings-page trip-settings-card">
            <div><p className="eyebrow">Aktuelle Reise</p><h2>Einstellungen</h2></div>
            <label>Reisename<input value={state.tripName} onChange={(event) => setState((current) => ({ ...current, tripName: event.target.value }))} /></label>
            <label>Reisedaten<input value={state.tripDates} onChange={(event) => setState((current) => ({ ...current, tripDates: event.target.value }))} /></label>
            <div><span className="field-label">Personen</span><div className="people">{state.people.map((person) => <span key={person}>{person}</span>)}</div></div>
          </section>}
        </section>}
      </section>

      <nav className="mobile-nav" aria-label="Arbeitsschritte">{workflowNav.map((entry) => <button key={entry.id} disabled={!hasActivePacklist} className={view === entry.id ? "active" : ""} onClick={() => setView(entry.id)}><Icon name={entry.icon} /><span>{entry.short}</span></button>)}</nav>

      {newPacklistOpen && <NewPacklistPanel onClose={() => setNewPacklistOpen(false)} onCreate={createPacklist} />}
      {packlistEditOpen && <PacklistEditPanel state={state} onClose={() => setPacklistEditOpen(false)} onSave={updatePacklistDetails} />}
      {editing && <DetailPanel item={editing} onClose={() => setEditingId(null)} onUpdateItem={(update) => updateTripItem(editing.id, update)} onUpdateQuantity={(person, field, value) => updateQuantity(editing.id, person, field, value)} onRemove={() => { setState((current) => ({ ...current, tripItems: current.tripItems.filter((entry) => entry.id !== editing.id) })); setEditingId(null); }} />}
      {addingItem && <NewItemPanel categories={catalog.categories.map((entry) => entry.name)} onClose={() => setAddingItem(false)} onAdd={addNewCatalogItem} />}
      {catalogEditing && <CatalogItemPanel item={catalogEditing} categories={catalog.categories.map((entry) => entry.name)} selected={selectedByCatalogId.has(catalogEditing.id)} onClose={() => setCatalogEditingId(null)} onAdd={(priority) => { choosePriority(catalogEditing.id, priority); setCatalogEditingId(null); }} onCategoryChange={(category) => setCatalogItemCategory(catalogEditing.id, category)} onDelete={() => deleteFromCatalog(catalogEditing)} />}
      {pendingRestore && <RestoreBackupDialog backupDate={pendingRestore.backupDate} packlistCount={pendingRestore.state.packlists.length} onCancel={() => setPendingRestore(null)} onConfirm={applyRestoredBackup} />}
      {deleteDialogOpen && <DeleteAllDataDialog onCancel={() => setDeleteDialogOpen(false)} onConfirm={deleteAllStoredData} />}
    </main>
  );
}

function DeleteAllDataDialog({ onCancel, onConfirm }: { onCancel: () => void; onConfirm: () => void }) {
  return <div className="modal-backdrop centered-backdrop restore-confirm-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onCancel(); }}>
    <section className="restore-confirm-panel delete-confirm-panel" role="dialog" aria-modal="true" aria-labelledby="delete-confirm-title">
      <span className="export-icon danger-icon"><Icon name="trash" /></span>
      <div><p className="eyebrow">Endgültige Löschung</p><h2 id="delete-confirm-title">Alle gespeicherten Daten löschen?</h2></div>
      <p>Damit werden alle Packlisten, Personen, Reiseangaben, eigenen Gegenstände und Arbeitsstände dieses Kontos aus der App entfernt.</p>
      <div className="restore-warning"><strong>Dieser Vorgang kann nicht rückgängig gemacht werden.</strong><span>Laden Sie vorher eine Datensicherung herunter, wenn Sie den Stand später wiederherstellen möchten.</span></div>
      <div className="panel-actions panel-actions-end"><button type="button" className="secondary" onClick={onCancel}>Abbrechen</button><button type="button" className="danger-confirm" onClick={onConfirm}>Endgültig löschen</button></div>
    </section>
  </div>;
}

function RestoreBackupDialog({ backupDate, packlistCount, onCancel, onConfirm }: { backupDate: string; packlistCount: number; onCancel: () => void; onConfirm: () => void }) {
  return <div className="modal-backdrop centered-backdrop restore-confirm-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onCancel(); }}>
    <section className="restore-confirm-panel" role="dialog" aria-modal="true" aria-labelledby="restore-confirm-title">
      <span className="export-icon"><Icon name="restore" /></span>
      <div><p className="eyebrow">Sicherheitsabfrage</p><h2 id="restore-confirm-title">Sicherung wiederherstellen?</h2></div>
      <p>Die Sicherung vom <strong>{backupDate}</strong> enthält <strong>{packlistCount} {packlistCount === 1 ? "Packliste" : "Packlisten"}</strong>.</p>
      <div className="restore-warning"><strong>Der derzeitige Arbeitsstand wird vollständig ersetzt.</strong><span>Brechen Sie ab, wenn Sie vorher noch eine aktuelle Sicherung herunterladen möchten.</span></div>
      <div className="panel-actions panel-actions-end"><button type="button" className="secondary" onClick={onCancel}>Abbrechen</button><button type="button" className="primary" onClick={onConfirm}>Sicherung wiederherstellen</button></div>
    </section>
  </div>;
}

function NewPacklistPanel({ onClose, onCreate }: { onClose: () => void; onCreate: (values: { tripName: string; destination: string; dateFrom: string; dateTo: string; peopleCount: number; description: string; activities: string[] }) => void }) {
  const [tripName, setTripName] = useState("");
  const [destination, setDestination] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [peopleCount, setPeopleCount] = useState(1);
  const [description, setDescription] = useState("");
  const [activities, setActivities] = useState<string[]>([]);
  const [activityPickerOpen, setActivityPickerOpen] = useState(false);
  function toggleActivity(activity: string) {
    setActivities((current) => current.includes(activity) ? current.filter((entry) => entry !== activity) : [...current, activity]);
  }
  return <div className="modal-backdrop centered-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><form className="new-packlist-panel" role="dialog" aria-modal="true" aria-labelledby="new-packlist-title" onSubmit={(event) => { event.preventDefault(); if (!tripName.trim() || !destination.trim()) return; onCreate({ tripName: tripName.trim(), destination: destination.trim(), dateFrom, dateTo, peopleCount: Math.max(1, peopleCount), description: description.trim(), activities }); }}>
    <div className="new-packlist-header"><div><p className="eyebrow">Neue Reise</p><h2 id="new-packlist-title">Packliste neu erstellen</h2><p>Die Angaben bilden die Grundlage für den ersten Vorschlag.</p></div><button type="button" className="close-button" onClick={onClose}><Icon name="close" /></button></div>
    <div className="new-packlist-fields">
      <label className="wide-field">Name der Packliste<input autoFocus required value={tripName} onChange={(event) => setTripName(event.target.value)} placeholder="z. B. Schottland 2027" /></label>
      <label className="wide-field">Reiseziel<input required value={destination} onChange={(event) => setDestination(event.target.value)} placeholder="Land, Region oder Ort" /></label>
      <label>Von<input type="date" value={dateFrom} onChange={(event) => setDateFrom(event.target.value)} /></label>
      <label>Bis<input type="date" min={dateFrom || undefined} value={dateTo} onChange={(event) => setDateTo(event.target.value)} /></label>
      <label className="wide-field people-field">Anzahl der Personen<input type="number" min="1" max="30" inputMode="numeric" value={peopleCount} onChange={(event) => setPeopleCount(Math.max(1, Number(event.target.value) || 1))} /></label>
      <label className="wide-field">Reise kurz beschreiben<textarea value={description} onChange={(event) => setDescription(event.target.value)} placeholder="Zum Beispiel: Flugreise, danach Mietauto, Hotels und B&B, nur kleiner Koffer …" /></label>
    </div>
    <div className="activity-picker">
      <button type="button" className={`activity-picker-button ${activityPickerOpen ? "open" : ""}`} aria-expanded={activityPickerOpen} onClick={() => setActivityPickerOpen((value) => !value)}><span><strong>Aktivitäten auswählen</strong><small>{activities.length ? `${activities.length} ausgewählt` : "Noch keine ausgewählt"}</small></span><span className="activity-picker-chevron"><Icon name="chevron" /></span></button>
      {activityPickerOpen && <div className="activity-option-list">{activityOptions.map((activity) => <label key={activity}><span>{activity}</span><input type="checkbox" checked={activities.includes(activity)} onChange={() => toggleActivity(activity)} /></label>)}</div>}
      {activities.length > 0 && <div className="activity-chips selected-activities">{activities.map((activity) => <span key={activity}>{activity}<button type="button" aria-label={`${activity} entfernen`} onClick={() => toggleActivity(activity)}>×</button></span>)}</div>}
    </div>
    <div className="proposal-note"><strong>Lokaler Grundvorschlag</strong><p>Die App wählt Grundausstattung und passende Gegenstände zu den Aktivitäten vor. Danach können alle drei Prioritäten geändert und weitere Dinge ergänzt werden.</p></div>
    <div className="panel-actions panel-actions-end"><button type="button" className="secondary" onClick={onClose}>Abbrechen</button><button className="primary" type="submit">Vorschlag erstellen</button></div>
  </form></div>;
}

function PacklistEditPanel({ state, onClose, onSave }: { state: AppState; onClose: () => void; onSave: (values: PacklistDetails) => void }) {
  const [tripName, setTripName] = useState(state.tripName);
  const [destination, setDestination] = useState(state.destination);
  const [dateFrom, setDateFrom] = useState(state.dateFrom);
  const [dateTo, setDateTo] = useState(state.dateTo);
  const [description, setDescription] = useState(state.description);
  const [activities, setActivities] = useState<string[]>(state.activities);
  const [people, setPeople] = useState<string[]>(state.people.length ? state.people : ["Person 1"]);
  const [activityPickerOpen, setActivityPickerOpen] = useState(false);

  function toggleActivity(activity: string) {
    setActivities((current) => current.includes(activity) ? current.filter((entry) => entry !== activity) : [...current, activity]);
  }

  function updatePerson(index: number, value: string) {
    setPeople((current) => current.map((person, personIndex) => personIndex === index ? value : person));
  }

  function removePerson(index: number) {
    setPeople((current) => current.length === 1 ? current : current.filter((_, personIndex) => personIndex !== index));
  }

  return <div className="modal-backdrop centered-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><form className="new-packlist-panel edit-packlist-panel" role="dialog" aria-modal="true" aria-labelledby="edit-packlist-title" onSubmit={(event) => {
    event.preventDefault();
    const cleanedPeople = people.map((person, index) => person.trim() || `Person ${index + 1}`);
    if (!tripName.trim() || !destination.trim()) return;
    onSave({ tripName: tripName.trim(), destination: destination.trim(), dateFrom, dateTo, description: description.trim(), activities, people: cleanedPeople });
  }}>
    <div className="new-packlist-header"><div><p className="eyebrow">Grundangaben ändern</p><h2 id="edit-packlist-title">Packliste bearbeiten</h2><p>Hier ändern Sie die Reise selbst – einschließlich aller mitreisenden Personen.</p></div><button type="button" className="close-button" onClick={onClose}><Icon name="close" /></button></div>
    <div className="new-packlist-fields">
      <label className="wide-field">Name der Packliste<input autoFocus required value={tripName} onChange={(event) => setTripName(event.target.value)} /></label>
      <label className="wide-field">Reiseziel<input required value={destination} onChange={(event) => setDestination(event.target.value)} /></label>
      <label>Von<input type="date" value={dateFrom} onChange={(event) => setDateFrom(event.target.value)} /></label>
      <label>Bis<input type="date" min={dateFrom || undefined} value={dateTo} onChange={(event) => setDateTo(event.target.value)} /></label>
      <label className="wide-field">Reise kurz beschreiben<textarea value={description} onChange={(event) => setDescription(event.target.value)} /></label>
    </div>
    <section className="edit-people-section" aria-labelledby="edit-people-heading">
      <div className="edit-section-heading"><div><strong id="edit-people-heading">Mitreisende Personen</strong><span>Person hinzufügen, Namen ändern oder eine Person entfernen.</span></div><button type="button" className="secondary compact-action" onClick={() => setPeople((current) => [...current, `Person ${current.length + 1}`])}><Icon name="plus" /> Person hinzufügen</button></div>
      <div className="edit-people-list">{people.map((person, index) => <div className="edit-person-row" key={index}><label><span>Person {index + 1}</span><input value={person} onChange={(event) => updatePerson(index, event.target.value)} /></label><button type="button" className="remove-person" disabled={people.length === 1} aria-label={`Person ${index + 1} entfernen`} onClick={() => removePerson(index)}>Entfernen</button></div>)}</div>
    </section>
    <div className="activity-picker">
      <button type="button" className={`activity-picker-button ${activityPickerOpen ? "open" : ""}`} aria-expanded={activityPickerOpen} onClick={() => setActivityPickerOpen((value) => !value)}><span><strong>Aktivitäten bearbeiten</strong><small>{activities.length ? `${activities.length} ausgewählt` : "Noch keine ausgewählt"}</small></span><span className="activity-picker-chevron"><Icon name="chevron" /></span></button>
      {activityPickerOpen && <div className="activity-option-list">{activityOptions.map((activity) => <label key={activity}><span>{activity}</span><input type="checkbox" checked={activities.includes(activity)} onChange={() => toggleActivity(activity)} /></label>)}</div>}
      {activities.length > 0 && <div className="activity-chips selected-activities">{activities.map((activity) => <span key={activity}>{activity}<button type="button" aria-label={`${activity} entfernen`} onClick={() => toggleActivity(activity)}>×</button></span>)}</div>}
    </div>
    <div className="proposal-note"><strong>Die bestehende Packliste bleibt erhalten.</strong><p>Neue Personen werden bei persönlichen Gegenständen ergänzt. Wird eine zuständige Person entfernt, bleibt der gemeinsame Gegenstand erhalten und wird als noch nicht zugeordnet markiert.</p></div>
    <div className="panel-actions panel-actions-end"><button type="button" className="secondary" onClick={onClose}>Abbrechen</button><button className="primary" type="submit">Änderungen speichern</button></div>
  </form></div>;
}

function NewItemPanel({ categories, onClose, onAdd }: { categories: string[]; onClose: () => void; onAdd: (values: { name: string; category: string; unit: string; notes: string; priority: Priority }) => void }) {
  const [name, setName] = useState("");
  const [category, setCategory] = useState(categories[0] ?? "Sonstiges");
  const [unit, setUnit] = useState("Stk");
  const [notes, setNotes] = useState("");
  const [priority, setPriority] = useState<Priority>("Empfohlen");
  return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><form className="detail-panel" role="dialog" aria-modal="true" onSubmit={(event) => { event.preventDefault(); if (name.trim()) onAdd({ name: name.trim(), category, unit: unit.trim() || "Stk", notes: notes.trim(), priority }); }}>
    <button type="button" className="close-button" onClick={onClose}><Icon name="close" /></button><p className="eyebrow">Maximalliste ergänzen</p><h2>Neuer Gegenstand</h2><p className="item-note">Der neue Gegenstand wird gespeichert und gleich für diese Reise ausgewählt.</p>
    <div className="detail-fields new-item-fields"><label>Bezeichnung<input autoFocus required value={name} onChange={(event) => setName(event.target.value)} placeholder="z. B. Fernglas" /></label><label>Kategorie<select value={category} onChange={(event) => setCategory(event.target.value)}>{categories.map((entry) => <option key={entry}>{entry}</option>)}</select></label><label>Einheit<input value={unit} onChange={(event) => setUnit(event.target.value)} /></label><label>Priorität<select value={priority} onChange={(event) => setPriority(event.target.value as Priority)}>{priorities.map((entry) => <option key={entry}>{entry}</option>)}</select></label><label>Notiz<textarea value={notes} onChange={(event) => setNotes(event.target.value)} placeholder="Optional" /></label></div>
    <div className="panel-actions panel-actions-end"><button type="button" className="secondary" onClick={onClose}>Abbrechen</button><button className="primary" type="submit">Gegenstand anlegen</button></div>
  </form></div>;
}

function CatalogItemPanel({ item, categories, selected, onClose, onAdd, onCategoryChange, onDelete }: { item: CatalogItem; categories: string[]; selected: boolean; onClose: () => void; onAdd: (priority: Priority) => void; onCategoryChange: (category: string) => void; onDelete: () => void }) {
  return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className="detail-panel" role="dialog" aria-modal="true">
    <button className="close-button" onClick={onClose}><Icon name="close" /></button><p className="eyebrow">{item.category}</p><h2>{item.name}</h2>{item.notes && <p className="item-note">{item.notes}</p>}
    <div className="detail-fields new-item-fields"><label>Kategorie<select value={item.category} onChange={(event) => onCategoryChange(event.target.value)}>{categories.map((entry) => <option key={entry}>{entry}</option>)}</select></label></div>
    <div className="source-note"><p><strong>Einheit:</strong> {item.units || "Stk"}</p><p><strong>Eintrag:</strong> {item.reviewStatus}</p></div>
    {!selected && <div className="catalog-priority-actions"><p className="field-label">Für diese Reise auswählen:</p>{priorities.map((priority) => <button key={priority} className={`priority-choice ${priority.toLowerCase()}`} onClick={() => onAdd(priority)}>{priority}</button>)}</div>}
    {selected && <div className="notice"><strong>Dieser Gegenstand ist bereits für die Reise ausgewählt.</strong></div>}
    <div className="panel-actions"><button className="danger" onClick={onDelete}>Aus Maximalliste löschen</button><button className="primary" onClick={onClose}>Fertig</button></div>
  </section></div>;
}

function StageRow({ item, stage, onToggle, onEdit }: { item: TripItem; stage: Stage; onToggle: () => void; onEdit: () => void }) {
  const done = stageIsDone(item, stage);
  return <div className={`stage-row ${done ? "done" : ""}`}>
    <button className="stage-item" onClick={onEdit}><strong>{item.name}</strong><small>{item.category} · {assignmentSummary(item)}</small></button>
    <span className={`priority ${item.priority.toLowerCase()}`}>{item.priority}</span>
    <button className={`stage-check ${done ? "checked" : ""}`} aria-label={`${item.name} ${done ? "wieder als offen markieren" : "abhaken"}`} aria-pressed={done} onClick={onToggle}>{done ? "✓" : ""}</button>
  </div>;
}

function StageAssignmentRow({ item, assignment, stage, onToggle, onEdit }: { item: TripItem; assignment: Assignment; stage: Stage; onToggle: () => void; onEdit: () => void }) {
  const done = assignment.quantities[stage] >= assignment.quantities.intended;
  return <div className={`stage-row assignment-stage-row ${done ? "done" : ""}`}>
    <button className="stage-item" onClick={onEdit}><strong>{item.name}</strong><small>{item.category} · {item.scope === "personal" ? "Persönlich" : "Gemeinsam übernommen"} · Menge {assignment.quantities.intended}</small></button>
    <span className={`priority ${item.priority.toLowerCase()}`}>{item.priority}</span>
    <button className={`stage-check ${done ? "checked" : ""}`} aria-label={`${item.name} für ${assignment.person} ${done ? "wieder als offen markieren" : "abhaken"}`} aria-pressed={done} onClick={onToggle}>{done ? "✓" : ""}</button>
  </div>;
}

function TaskRow({ task, onToggle, showArea = false }: { task: Task; onToggle: () => void; showArea?: boolean }) {
  return <label className={task.done ? "done" : ""}><input type="checkbox" checked={task.done} onChange={onToggle}/><span><strong>{task.label}</strong>{showArea && <small>{task.area} · {task.priority}</small>}</span></label>;
}

function DetailPanel({ item, onClose, onUpdateItem, onUpdateQuantity, onRemove }: {
  item: TripItem;
  onClose: () => void;
  onUpdateItem: (update: (item: TripItem) => TripItem) => void;
  onUpdateQuantity: (person: Person, field: keyof Quantities, value: number) => void;
  onRemove: () => void;
}) {
  const labels: Record<keyof Quantities, string> = { intended: "Vorgesehen", prepared: "Hergerichtet", packed: "Eingepackt", used: "Tatsächlich gebraucht" };
  return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className="detail-panel" role="dialog" aria-modal="true" aria-labelledby="detail-title">
    <button className="close-button" onClick={onClose}><Icon name="close" /></button><p className="eyebrow">{item.category}</p><h2 id="detail-title">{item.name}</h2>
    {item.notes && <p className="item-note">{item.notes}</p>}
    <div className="assignment-list">{item.assignments.map((assignment) => <section className="assignment-card" key={assignment.person}><h3>{assignment.person}</h3><div className="quantity-grid">{(["intended", "prepared", "packed", "used"] as const).map((field) => <label key={field}><span>{labels[field]}</span><input type="number" min="0" value={assignment.quantities[field]} onChange={(event) => onUpdateQuantity(assignment.person, field, Number(event.target.value))}/></label>)}</div></section>)}</div>
    <div className="detail-fields"><label>Priorität<select value={item.priority} onChange={(event) => onUpdateItem((current) => ({ ...current, priority: event.target.value as Priority }))}>{priorities.map((priority) => <option key={priority}>{priority}</option>)}</select></label><label>Gepäckbereich<input value={item.luggage} onChange={(event) => onUpdateItem((current) => ({ ...current, luggage: event.target.value }))}/></label></div>
    {(item.activity || item.source) && <div className="source-note">{item.activity && <p><strong>Aktivität:</strong> {item.activity}</p>}{item.source && <p><strong>Herkunft:</strong> {item.source}</p>}</div>}
    <div className="panel-actions"><button className="danger" onClick={onRemove}>Aus Reise entfernen</button><button className="primary" onClick={onClose}>Fertig</button></div>
  </section></div>;
}
