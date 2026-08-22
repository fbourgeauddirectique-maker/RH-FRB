import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";
import { initializeApp } from "https://www.gstatic.com/firebasejs/12.17.1/firebase-app.js";
import { getAuth, signInAnonymously, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/12.17.1/firebase-auth.js";
import { getFirestore, collection, query, where, getDocs, doc, setDoc, deleteDoc, serverTimestamp } from "https://www.gstatic.com/firebasejs/12.17.1/firebase-firestore.js";

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
const timeFields = ["morningStart", "morningEnd", "afternoonStart", "dayEnd"];
let currentUser = null;
let entries = [];
let current = { date: "", mission: "", zone: "", team: "", car: "", assignment: "", morningStart: "", morningEnd: "", afternoonStart: "", dayEnd: "" };

function todayISO() { const d = new Date(); const offset = d.getTimezoneOffset(); return new Date(d.getTime() - offset * 60000).toISOString().slice(0, 10); }
function minutes(value) { if (!value) return null; const [hours, mins] = value.split(":").map(Number); return hours * 60 + mins; }
function totalMinutes(item) { const morning = minutes(item.morningEnd) - minutes(item.morningStart); const afternoon = minutes(item.dayEnd) - minutes(item.afternoonStart); return [morning, afternoon].filter(Number.isFinite).reduce((sum, value) => sum + value, 0); }
function formatMinutes(value) { if (!Number.isFinite(value) || value < 0) return "00:00"; return `${String(Math.floor(value / 60)).padStart(2, "0")}:${String(Math.round(value % 60)).padStart(2, "0")}`; }
function isoWeek(value) { const date = new Date(`${value}T12:00:00`); const day = date.getDay() || 7; date.setDate(date.getDate() + 4 - day); const year = date.getFullYear(); const first = new Date(year, 0, 1); return Math.ceil((((date - first) / 86400000) + first.getDay() + 1) / 7); }
function dateLabel(value) { return value ? new Date(`${value}T12:00:00`).toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }) : ""; }
function message(id, text, type = "") { const element = $(id); if (!element) return; element.textContent = text; element.className = `message ${type}`; element.hidden = false; }
function authMessage(text, type = "") { const element = $("auth-message"); if (element) { element.textContent = text; element.className = `message ${type}`; } }
function showView(name) { document.querySelectorAll("[data-view]").forEach((view) => view.classList.toggle("active", view.dataset.view === name)); document.querySelectorAll("[data-nav]").forEach((button) => button.classList.toggle("active", button.dataset.nav === name)); }

function updateInterface() {
  timeFields.forEach((field) => { $(`${field}-view`).textContent = current[field] || "Non pointé"; document.querySelector(`[data-field="${field}"]`).classList.toggle("done", Boolean(current[field])); });
  $("date").value = current.date;
  $("total").textContent = formatMinutes(totalMinutes(current));
  $("mission-view").textContent = current.mission.trim() || "Non renseignée";
  $("zone-view").textContent = current.zone.trim() || "Non renseignée";
  $("team-view").textContent = current.team.trim() || "Non renseignée";
  $("car-view").textContent = current.car.trim() || "Non renseignée";
  $("assignment-view").textContent = current.assignment.trim() || "Non renseignée";
  const status = $("status"); status.textContent = current.dayEnd ? "🔵 Journée complète" : current.morningStart ? "🟢 Journée en cours" : "⚪ En attente";
  renderStats();
}
function fillParameters() { ["mission", "zone", "team", "car", "assignment"].forEach((field) => { $(field).value = current[field] || ""; }); }
function renderStats() { const week = entries.filter((entry) => entry.date && isoWeek(entry.date) === isoWeek(current.date)); const weekTotal = week.reduce((sum, entry) => sum + totalMinutes(entry), 0); $("day-week-label").textContent = `Semaine S${String(isoWeek(current.date)).padStart(2, "0")}`; $("day-week-total").textContent = formatMinutes(weekTotal); $("stats-week-label").textContent = `Semaine S${String(isoWeek(current.date)).padStart(2, "0")}`; $("stats-week-total").textContent = formatMinutes(weekTotal); $("stats-days").textContent = entries.length; const all = entries.reduce((sum, entry) => sum + totalMinutes(entry), 0); $("stats-all").textContent = formatMinutes(all); $("stats-average").textContent = entries.length ? formatMinutes(all / entries.length) : "00:00"; }
function renderHistory() { $("history-list").innerHTML = entries.length ? entries.map((entry) => `<div class="history-card"><strong>S${String(isoWeek(entry.date)).padStart(2, "0")} · ${dateLabel(entry.date)}</strong><div>${entry.morningStart || "--:--"} / ${entry.morningEnd || "--:--"} · ${entry.afternoonStart || "--:--"} / ${entry.dayEnd || "--:--"}</div><div>${entry.mission || "Sans mission"} · ${entry.zone || "Zone non renseignée"} · ${entry.team || "Équipe non renseignée"} · ${entry.car || "Voiture non renseignée"} · ${entry.assignment || "Affectation non renseignée"}</div><strong>${formatMinutes(totalMinutes(entry))}</strong></div>`).join("") : '<p class="empty">Aucune journée enregistrée.</p>'; }
async function loadEntries() { if (!currentUser) return; try { const snapshot = await getDocs(query(collection(db, "timesheets"), where("uid", "==", currentUser.uid))); entries = snapshot.docs.map((item) => ({ id: item.id, ...item.data() })).sort((a, b) => String(b.date).localeCompare(String(a.date))); const todayEntry = entries.find((entry) => entry.date === current.date); if (todayEntry) { current = { ...current, ...todayEntry }; fillParameters(); } renderHistory(); updateInterface(); } catch (error) { message("message", `Erreur lecture : ${error.code}\n${error.message}`, "error"); } }
async function saveDay() { if (!currentUser) return message("message", "Session anonyme non disponible.", "error"); try { const id = `${currentUser.uid}_${current.date}`; const payload = { ...current, uid: currentUser.uid, authMode: "anonymous", totalMinutes: totalMinutes(current), week: isoWeek(current.date), updatedAt: serverTimestamp() }; await setDoc(doc(db, "timesheets", id), payload, { merge: true }); await loadEntries(); showView("historique"); message("message", "Journée enregistrée.", "ok"); } catch (error) { message("message", `${error.code}\n${error.message}`, "error"); } }
function editTime(field) { const overlay = document.createElement("div"); overlay.style.cssText = "position:fixed;inset:0;z-index:100;background:#000b;display:grid;place-items:center;padding:20px"; overlay.innerHTML = `<div class="panel"><h2>Modifier l’heure</h2><input id="time-value" type="time" value="${current[field] || ""}"><button id="time-ok">Valider</button><button id="time-cancel">Annuler</button></div>`; document.body.append(overlay); overlay.querySelector("#time-ok").onclick = () => { const value = overlay.querySelector("#time-value").value; if (!value) return; current[field] = value; overlay.remove(); updateInterface(); message("message", "Horaire modifié.", "ok"); }; overlay.querySelector("#time-cancel").onclick = () => overlay.remove(); }
function exportXlsx() { if (!entries.length) return message("message", "Aucune journée à exporter.", "error"); const rows = entries.map((entry) => ({ Semaine: `S${String(isoWeek(entry.date)).padStart(2, "0")}`, Date: entry.date, Jour: new Date(`${entry.date}T12:00:00`).toLocaleDateString("fr-FR", { weekday: "long" }), "Début matin": entry.morningStart || "", "Fin matin": entry.morningEnd || "", "Début AM": entry.afternoonStart || "", "Fin journée": entry.dayEnd || "", "Total heures": formatMinutes(totalMinutes(entry)), Mission: entry.mission || "", Zone: entry.zone || "", Équipe: entry.team || "", Voiture: entry.car || "", Affectation: entry.assignment || "" })); const sheet = XLSX.utils.json_to_sheet(rows); const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, sheet, "Pointages"); XLSX.writeFile(workbook, `releve-heures-${todayISO()}.xlsx`); }
async function startAnonymousSession() { try { authMessage("Connexion automatique en cours…"); const credential = await signInAnonymously(auth); authMessage(`Session automatique active · ${credential.user.uid}`, "ok"); } catch (error) { authMessage(`${error.code}\n${error.message}`, "error"); } }

function init() {
  current.date = todayISO();
  updateInterface();
  document.querySelectorAll("[data-nav]").forEach((button) => button.onclick = () => showView(button.dataset.nav));
  $("edit").onclick = () => showView("parametres");
  $("save").onclick = saveDay;
  $("export-xlsx").onclick = exportXlsx;
  $("duplicate").onclick = () => { if (entries.length) { current = { ...entries[0], date: current.date }; fillParameters(); updateInterface(); } };
  document.querySelectorAll("[data-field]").forEach((button) => button.onclick = () => editTime(button.dataset.field));
  [["date", "date"], ["mission", "mission"], ["zone", "zone"], ["team", "team"], ["car", "car"], ["assignment", "assignment"]].forEach(([elementId, field]) => $(elementId).oninput = (event) => { current[field] = event.target.value; updateInterface(); });
  const times = ["08:00", "08:30", "12:00", "13:00", "17:00", "18:00"]; $("quick").innerHTML = times.map((time) => `<button type="button" data-quick="${time}">${time}</button>`).join(""); let selected = null; document.querySelectorAll("[data-quick]").forEach((button) => button.onclick = () => { selected = selected || timeFields.find((field) => !current[field]) || "morningStart"; current[selected] = button.dataset.quick; updateInterface(); });
  onAuthStateChanged(auth, async (user) => { currentUser = user; if (user) { authMessage(`Session automatique active · ${user.uid}`, "ok"); $("logout").classList.remove("hidden"); await loadEntries(); } else { $("logout").classList.add("hidden"); await startAnonymousSession(); } });
  $("logout").onclick = () => signOut(auth);
}
init();
