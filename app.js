import { CONFIG } from "./config.js";

const $ = (id) => document.getElementById(id);
const status = (m) => ($("status").textContent = m);

let templateArrayBuffer = null;     // originele template bytes
let workbook = null;               // XLSX workbook
let items = [];                    // {Id, Item, Waarde}

function nowArubaTimestamp() {
  // Aruba is UTC-4; in browser nemen we lokale tijd van user (meestal Aruba)
  // Format: yyyy-MM-dd_HH-mm-ss
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const yyyy = d.getFullYear();
  const MM = pad(d.getMonth() + 1);
  const dd = pad(d.getDate());
  const hh = pad(d.getHours());
  const mm = pad(d.getMinutes());
  const ss = pad(d.getSeconds());
  return `${yyyy}-${MM}-${dd}_${hh}-${mm}-${ss}`;
}

async function downloadTemplate() {
  status("Template downloaden…");
  const r = await fetch(CONFIG.TEMPLATE_DOWNLOAD_URL, { cache: "no-store" });
  if (!r.ok) throw new Error("Template download mislukt (controleer download-link).");
  templateArrayBuffer = await r.arrayBuffer();

  workbook = XLSX.read(templateArrayBuffer, { type: "array" });
  if (!workbook.Sheets[CONFIG.SHEET_ITEMS]) {
    throw new Error(`Tabblad '${CONFIG.SHEET_ITEMS}' niet gevonden in Excel.`);
  }

  const wsItems = workbook.Sheets[CONFIG.SHEET_ITEMS];

  // Lees items als JSON (kolomnamen uit header row)
  const rows = XLSX.utils.sheet_to_json(wsItems, { defval: "", raw: true });

  // Verwacht kolommen: Id, Item, Waarde
  items = rows
    .filter(r => String(r.Id || "").trim() !== "")
    .map(r => ({
      Id: String(r.Id).trim(),
      Item: String(r.Item || "").trim(),
      Waarde: r.Waarde
    }));

  if (items.length === 0) throw new Error("Geen items gevonden (controleer Items-tabblad).");

  renderItems();
  status(`✅ Template geladen (${items.length} items).`);
}

function renderItems() {
  const host = $("items");
  host.innerHTML = "";

  for (const it of items) {
    const row = document.createElement("div");
    row.className = "itemRow";

    const label = document.createElement("label");
    label.textContent = it.Item || it.Id;

    const input = document.createElement("input");
    input.type = "checkbox";
    input.dataset.id = it.Id;

    row.appendChild(label);
    row.appendChild(input);
    host.appendChild(row);
  }
}

function collectAnswersMap() {
  const checks = [...document.querySelectorAll("#items input[type=checkbox]")];
  const map = new Map();
  for (const ch of checks) map.set(ch.dataset.id, ch.checked);
  return map;
}

function setCell(ws, a1, value) {
  ws[a1] = { t: "s", v: String(value ?? "") };
}

function writeBackToWorkbook() {
  if (!workbook) throw new Error("Geen workbook geladen.");

  const answers = collectAnswersMap();
  const filledBy = $("filledBy").value || "Onbekend";
  const remark = $("remark").value || "";
  const ts = nowArubaTimestamp();

  // 1) Schrijf metadata naar tabblad Checklist (optioneel)
  const wsChecklist = workbook.Sheets[CONFIG.SHEET_CHECKLIST];
  if (wsChecklist) {
    setCell(wsChecklist, "B1", CONFIG.CHECKLIST_TITLE);
    setCell(wsChecklist, "B2", CONFIG.VEHICLE_TYPE);
    setCell(wsChecklist, "B3", filledBy);
    setCell(wsChecklist, "B4", ts);
    setCell(wsChecklist, "B5", remark);
  }

  // 2) Schrijf Waarde kolom terug in Items sheet
  const wsItems = workbook.Sheets[CONFIG.SHEET_ITEMS];

  // We herschrijven de hele sheet vanuit JSON om simpel en stabiel te blijven
  const newRows = items.map(it => ({
    Id: it.Id,
    Item: it.Item,
    Waarde: answers.get(it.Id) === true ? "JA" : "NEE"
  }));

  const newWs = XLSX.utils.json_to_sheet(newRows, { header: ["Id", "Item", "Waarde"] });

  // Vervang sheet
  workbook.Sheets[CONFIG.SHEET_ITEMS] = newWs;
}

function downloadFilledExcel() {
  if (!templateArrayBuffer || !workbook) throw new Error("Laad eerst het template.");

  writeBackToWorkbook();

  const ts = nowArubaTimestamp();
  const filledBy = ($("filledBy").value || "Onbekend").replace(/\s+/g, "");
  const fileName = `${CONFIG.VEHICLE_TYPE}_${CONFIG.CHECKLIST_ID}_${ts}_${filledBy}.xlsx`;

  XLSX.writeFile(workbook, fileName);
  status(`✅ Gedownload: ${fileName} (upload dit naar OneDrive/02_ingevuld).`);
}

function init() {
  $("viewLink").href = CONFIG.TEMPLATE_VIEW_URL;
  $("viewLink").textContent = "open";

  $("reloadBtn").addEventListener("click", async () => {
    try { await downloadTemplate(); }
    catch (e) { status("❌ " + e.message); }
  });

  $("downloadBtn").addEventListener("click", () => {
    try { downloadFilledExcel(); }
    catch (e) { status("❌ " + e.message); }
  });

  // Auto-load bij start
  downloadTemplate().catch(e => status("❌ " + e.message));
}

init();
