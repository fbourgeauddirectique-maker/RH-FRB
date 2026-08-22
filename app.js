import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";
import { initializeApp } from "https://www.gstatic.com/firebasejs/12.17.1/firebase-app.js";
import { getAuth, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/12.17.1/firebase-auth.js";
import { getFirestore, collection, query, where, getDocs } from "https://www.gstatic.com/firebasejs/12.17.1/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyCTz-YS4aaOS81XvDUFx3tiBISh0V7oUHo",
  authDomain: "rh-frb.firebaseapp.com",
  projectId: "rh-frb",
  storageBucket: "rh-frb.firebasestorage.app",
  messagingSenderId: "1069409309284",
  appId: "1:1069409309284:web:cac99c3a2457eef6a58571"
};

const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);
const $ = (id) => document.getElementById(id);
const state = { user: null, days: [] };

function toMinutes(value) { if (!value) return null; const [h, m] = value.split(":").map(Number); return h * 60 + m; }
function totalMinutes(day) { const morning = toMinutes(day.morningEnd) - toMinutes(day.morningStart); const afternoon = toMinutes(day.dayEnd) - toMinutes(day.afternoonStart); return [morning, afternoon].filter(Number.isFinite).reduce((a, b) => a + b, 0); }
function weekInfo(value) { const d = new Date(`${value}T12:00:00`); const day = d.getDay() || 7; d.setDate(d.getDate() + 4 - day); const year = d.getFullYear(); const first = new Date(year, 0, 1); const week = Math.ceil((((d - first) / 86400000) + first.getDay() + 1) / 7); return { year, week }; }
function dateLabel(value) { return value ? new Date(`${value}T12:00:00`).toLocaleDateString("fr-FR") : ""; }
function show(text) { const e = $("message"); e.textContent = text; e.hidden = false; }

function render() {
  const list = $("history-list");
  list.innerHTML = state.days.length ? state.days.map((day) => `<div class="history-row"><span><strong>${dateLabel(day.date)}</strong><br>${day.morningStart || "--:--"} / ${day.morningEnd || "--:--"} · ${day.afternoonStart || "--:--"} / ${day.dayEnd || "--:--"}<br>${day.mission || "Sans mission"} · ${day.zone || "Zone non renseignée"}</span><strong>${Math.floor(totalMinutes(day) / 60)}h${String(totalMinutes(day) % 60).padStart(2, "0")}</strong></div>`).join("") : "<p>Aucune journée enregistrée.</p>";
}

async function loadDays() {
  if (!state.user) return;
  const snapshot = await getDocs(query(collection(db, "timesheets"), where("uid", "==", state.user.uid)));
  state.days = snapshot.docs.map((item) => ({ id: item.id, ...item.data() })).sort((a, b) => String(a.date).localeCompare(String(b.date)));
  render();
}

function exportXlsx() {
  if (!state.days.length) { show("Aucune journée à exporter."); return; }
  const rows = state.days.map((day) => { const week = weekInfo(day.date); const minutes = totalMinutes(day); return {
    "Semaine": `S${String(week.week).padStart(2, "0")}`,
    "Date": day.date,
    "Jour": new Date(`${day.date}T12:00:00`).toLocaleDateString("fr-FR", { weekday: "long" }),
    "Début matin": day.morningStart || "",
    "Fin matin": day.morningEnd || "",
    "Début AM": day.afternoonStart || "",
    "Fin journée": day.dayEnd || "",
    "Total heures": `${Math.floor(minutes / 60)}:${String(minutes % 60).padStart(2, "0")}`,
    "Mission": day.mission || "",
    "Zone": day.zone || "",
    "Équipe": day.team || "",
    "Voiture": day.car || "",
    "Affectation": day.assignment || ""
  }; });
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet["!cols"] = [{ wch: 10 }, { wch: 14 }, { wch: 15 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 24 }, { wch: 24 }, { wch: 18 }, { wch: 18 }, { wch: 24 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Pointages");
  XLSX.writeFile(workbook, `releve-heures-${new Date().toISOString().slice(0, 10)}.xlsx`);
  show("Export XLSX généré.");
}

function init() { $("export-xlsx").addEventListener("click", exportXlsx); onAuthStateChanged(auth, async (user) => { state.user = user; $("auth-status").textContent = user ? user.email : "Non connecté"; if (user) await loadDays(); }); }
init();
