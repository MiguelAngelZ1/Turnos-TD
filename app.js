console.log('--- Turnos TD v1.2 (Commit: Glass+Bugfix) ---');
/**
 * Turnos TD - App de administración de turnos
 * Regla: turno viernes = mismo turno sábado y domingo
 * Reasignación automática al cambiar personal
 */

const STORAGE_KEY = 'turnos_td_data';

// Estado global
let state = {
  personnel: [],           // { id, name, active }
  shifts: {},             // { "2025-3": { "1": "uuid", "2": "uuid", ... } }
  currentYear: new Date().getFullYear(),
  currentMonth: new Date().getMonth() + 1,
  editingPersonId: null
};

// Utilidad: ID único
function uid() {
  return 'id_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 9);
}

// Clave del mes actual
function monthKey(year, month) {
  return `${year}-${month}`;
}

// Cargar desde localStorage
function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      state.personnel = parsed.personnel || [];
      state.shifts = parsed.shifts || {};
      if (parsed.currentYear) state.currentYear = parsed.currentYear;
      if (parsed.currentMonth) state.currentMonth = parsed.currentMonth;
    }
  } catch (e) {
    console.warn('Error loading state', e);
  }
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      personnel: state.personnel,
      shifts: state.shifts,
      currentYear: state.currentYear,
      currentMonth: state.currentMonth
    }));
  } catch (e) {
    console.warn('Error saving state', e);
  }
}

// --- Toasts (feedback inmediato)
const TOAST_DURATION_MS = 2800;

function showToast(message, type = 'default') {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const toast = document.createElement('div');
  toast.className = 'toast' + (type === 'success' ? ' toast-success' : type === 'error' ? ' toast-error' : '');
  toast.setAttribute('role', 'status');
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => {
    toast.remove();
  }, TOAST_DURATION_MS);
}

// Obtener personal activo
function getActivePersonnel() {
  return state.personnel.filter(p => p.active);
}

// Obtener asignaciones del mes
function getMonthShifts() {
  const key = monthKey(state.currentYear, state.currentMonth);
  if (!state.shifts[key]) state.shifts[key] = {};
  return state.shifts[key];
}

// Día de la semana (0=domingo ... 6=sábado)
function getDayOfWeek(year, month, day) {
  const d = new Date(year, month - 1, day);
  return d.getDay();
}

function isWeekend(year, month, day) {
  const dow = getDayOfWeek(year, month, day);
  return dow === 0 || dow === 6;
}

// Viernes = 5 en getDay()
function isFriday(year, month, day) {
  return getDayOfWeek(year, month, day) === 5;
}

// Días del mes
function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

// Al asignar el viernes, rellenar sábado y domingo con la misma persona
function applyFridayRule(year, month, day, personId) {
  const key = monthKey(year, month);
  if (!state.shifts[key]) state.shifts[key] = {};
  const shifts = state.shifts[key];
  shifts[String(day)] = personId;
  if (isFriday(year, month, day)) {
    shifts[String(day + 1)] = personId; // sábado
    shifts[String(day + 2)] = personId; // domingo
  }
}

// Al quitar asignación de un viernes, quitar también sábado y domingo
function clearFridayWeekend(year, month, day) {
  const key = monthKey(year, month);
  const shifts = state.shifts[key] || {};
  delete shifts[String(day)];
  if (isFriday(year, month, day)) {
    delete shifts[String(day + 1)];
    delete shifts[String(day + 2)];
  }
}

// Reasignación inteligente: cuando se elimina o desactiva una persona,
// redistribuir sus turnos entre el personal activo (quien menos tiene primero)
function reassignShiftsFromPerson(personId) {
  const active = getActivePersonnel().filter(p => p.id !== personId);
  if (active.length === 0) return;

  const key = monthKey(state.currentYear, state.currentMonth);
  const shifts = state.shifts[key] || {};
  const daysToReassign = [];

  for (const [dayStr, assignedId] of Object.entries(shifts)) {
    if (assignedId === personId) daysToReassign.push(parseInt(dayStr, 10));
  }

  // Ordenar días para procesar viernes primero (así al asignar viernes se rellenan sáb/dom)
  daysToReassign.sort((a, b) => a - b);

  for (const day of daysToReassign) {
    const countPerPerson = {};
    active.forEach(p => { countPerPerson[p.id] = 0; });
    for (const d of Object.keys(shifts)) {
      const id = shifts[d];
      if (countPerPerson[id] !== undefined) countPerPerson[id]++;
    }
    const sorted = active.slice().sort((a, b) => countPerPerson[a.id] - countPerPerson[b.id]);
    const chosen = sorted[0];
    if (chosen) {
      clearFridayWeekend(state.currentYear, state.currentMonth, day);
      applyFridayRule(state.currentYear, state.currentMonth, day, chosen.id);
    }
  }
  saveState();
}

// Redistribuir todos los turnos del mes de forma equitativa entre activos
function redistributeAllShifts() {
  const active = getActivePersonnel();
  if (active.length === 0) return;

  const key = monthKey(state.currentYear, state.currentMonth);
  const totalDays = daysInMonth(state.currentYear, state.currentMonth);
  const newShifts = {};
  let idx = 0;

  for (let day = 1; day <= totalDays; day++) {
    if (newShifts[String(day)] !== undefined) continue; // ya asignado (sáb/dom por viernes)
    const person = active[idx % active.length];
    newShifts[String(day)] = person.id;
    if (isFriday(state.currentYear, state.currentMonth, day)) {
      newShifts[String(day + 1)] = person.id;
      newShifts[String(day + 2)] = person.id;
    }
    idx++;
  }

  state.shifts[key] = newShifts;
  saveState();
}

// Obtener nombre de persona por id
function getPersonName(id) {
  const p = state.personnel.find(x => x.id === id);
  return p ? p.name : '—';
}

// --- UI: Navegación
function initNavigation() {
  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.dataset.view;
      document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
      const panel = document.getElementById('view-' + view);
      if (panel) panel.classList.add('active');

      if (view === 'turnos') renderCalendar();
      if (view === 'personal') renderPersonList();
      if (view === 'estadisticas') renderStats();
    });
  });
}

// --- UI: Mes
function setMonthLabel() {
  const d = new Date(state.currentYear, state.currentMonth - 1, 1);
  const label = d.toLocaleDateString('es-AR', { month: 'long', year: 'numeric' });
  const cap = label.charAt(0).toUpperCase() + label.slice(1);
  document.getElementById('monthLabel').textContent = cap;
  document.getElementById('currentMonthTitle').textContent = cap;
}

function initMonthNav() {
  document.getElementById('prevMonth').addEventListener('click', () => {
    state.currentMonth--;
    if (state.currentMonth < 1) {
      state.currentMonth = 12;
      state.currentYear--;
    }
    setMonthLabel();
    renderCalendar();
    saveState();
  });
  document.getElementById('nextMonth').addEventListener('click', () => {
    state.currentMonth++;
    if (state.currentMonth > 12) {
      state.currentMonth = 1;
      state.currentYear++;
    }
    setMonthLabel();
    renderCalendar();
    saveState();
  });
}

// --- UI: Calendario
function renderCalendar() {
  const grid = document.getElementById('calendarGrid');
  const totalDays = daysInMonth(state.currentYear, state.currentMonth);
  const shifts = getMonthShifts();

  const weekDays = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
  grid.innerHTML = '';

  weekDays.forEach(d => {
    const h = document.createElement('div');
    h.className = 'cal-day-header';
    h.textContent = d;
    grid.appendChild(h);
  });

  const firstDow = getDayOfWeek(state.currentYear, state.currentMonth, 1);
  const emptyStart = firstDow;
  for (let i = 0; i < emptyStart; i++) {
    const cell = document.createElement('div');
    cell.className = 'cal-cell empty';
    grid.appendChild(cell);
  }

  for (let day = 1; day <= totalDays; day++) {
    const cell = document.createElement('div');
    const weekend = isWeekend(state.currentYear, state.currentMonth, day);
    cell.className = 'cal-cell ' + (weekend ? 'weekend' : 'weekday');
    cell.dataset.day = day;

    const dayNum = document.createElement('span');
    dayNum.className = 'day-num';
    dayNum.textContent = day;
    cell.appendChild(dayNum);

    const personId = shifts[String(day)];
    const nameSpan = document.createElement('span');
    nameSpan.className = 'person-name';
    nameSpan.textContent = personId ? getPersonName(personId) : 'Sin asignar';
    cell.appendChild(nameSpan);

    cell.addEventListener('click', () => openAssignModal(day));
    grid.appendChild(cell);
  }

  saveState();
}

// --- Modal asignar
let assignDayPending = null;

function openAssignModal(day) {
  assignDayPending = day;
  const d = new Date(state.currentYear, state.currentMonth - 1, day);
  const label = d.toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long' });
  const cap = label.charAt(0).toUpperCase() + label.slice(1);
  document.getElementById('modalAssignDay').textContent = cap;

  const select = document.getElementById('assignPersonSelect');
  select.innerHTML = '';
  const active = getActivePersonnel();
  const currentId = getMonthShifts()[String(day)];

  const optEmpty = document.createElement('option');
  optEmpty.value = '';
  optEmpty.textContent = '— Sin asignar —';
  select.appendChild(optEmpty);

  active.forEach(p => {
    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = p.name;
    if (p.id === currentId) opt.selected = true;
    select.appendChild(opt);
  });

  document.getElementById('modalAssign').showModal();
  setTimeout(() => document.getElementById('assignPersonSelect').focus(), 0);
}

function confirmAssign() {
  if (assignDayPending == null) return;
  const select = document.getElementById('assignPersonSelect');
  const personId = select.value || null;
  const day = assignDayPending;

  clearFridayWeekend(state.currentYear, state.currentMonth, day);
  if (personId) applyFridayRule(state.currentYear, state.currentMonth, day, personId);

  assignDayPending = null;
  document.getElementById('modalAssign').close();
  renderCalendar();
  saveState();
  showToast('Turno asignado', 'success');
}

// --- Personal: lista y modal
function renderPersonList() {
  const list = document.getElementById('personList');
  list.innerHTML = '';

  state.personnel.forEach(p => {
    const li = document.createElement('li');
    li.className = 'person-item' + (p.active ? '' : ' inactive');
    li.innerHTML = `
      <div>
        <span class="name">${escapeHtml(p.name)}</span>
        <span class="badge">${p.active ? 'Activo' : 'Inactivo'}</span>
      </div>
      <div class="person-actions">
        <button type="button" class="btn-icon-sm btn-edit" data-edit="${p.id}" aria-label="Editar"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button>
        <button type="button" class="btn-icon-sm btn-delete" data-delete="${p.id}" aria-label="Eliminar"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg></button>
      </div>
    `;
    list.appendChild(li);
  });

  list.querySelectorAll('[data-edit]').forEach(btn => {
    btn.addEventListener('click', () => openPersonModal(btn.dataset.edit));
  });
  list.querySelectorAll('[data-delete]').forEach(btn => {
    btn.addEventListener('click', () => openConfirmDeleteModal(btn.dataset.delete));
  });
}

function escapeHtml(s) {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function openPersonModal(personId = null) {
  state.editingPersonId = personId;
  const modal = document.getElementById('modalPerson');
  const title = document.getElementById('modalPersonTitle');
  const nameInput = document.getElementById('personName');
  const activeGroup = document.getElementById('editActiveGroup');
  const activeCheck = document.getElementById('personActive');

  if (personId) {
    title.textContent = 'Editar personal';
    const p = state.personnel.find(x => x.id === personId);
    nameInput.value = p ? p.name : '';
    activeCheck.checked = p ? p.active : true;
    activeGroup.hidden = false;
  } else {
    title.textContent = 'Agregar personal';
    nameInput.value = '';
    activeCheck.checked = true;
    activeGroup.hidden = true;
  }
  modal.showModal();
  setTimeout(() => nameInput.focus(), 0);
}

function savePerson() {
  const name = document.getElementById('personName').value.trim();
  if (!name) return;

  if (state.editingPersonId) {
    const p = state.personnel.find(x => x.id === state.editingPersonId);
    if (p) {
      const wasActive = p.active;
      p.name = name;
      p.active = document.getElementById('personActive').checked;
      if (wasActive && !p.active) reassignShiftsFromPerson(p.id);
    }
  } else {
    const newPerson = { id: uid(), name, active: true };
    state.personnel.push(newPerson);
    redistributeAllShifts();
  }

  document.getElementById('modalPerson').close();
  state.editingPersonId = null;
  renderPersonList();
  renderCalendar();
  saveState();
  showToast('Guardado', 'success');
}

let deletePersonIdPending = null;

function openConfirmDeleteModal(personId) {
  const person = state.personnel.find(p => p.id === personId);
  const name = person ? person.name : 'esta persona';
  deletePersonIdPending = personId;
  document.getElementById('modalConfirmDeleteMessage').textContent =
    `¿Eliminar a ${escapeHtml(name)}? Se reasignarán sus turnos al resto del personal.`;
  document.getElementById('modalConfirmDelete').showModal();
  setTimeout(() => document.querySelector('[data-close="modalConfirmDelete"]').focus(), 0);
}

function confirmDeletePerson() {
  if (deletePersonIdPending == null) return;
  const personId = deletePersonIdPending;
  deletePersonIdPending = null;
  document.getElementById('modalConfirmDelete').close();
  reassignShiftsFromPerson(personId);
  state.personnel = state.personnel.filter(p => p.id !== personId);
  renderPersonList();
  renderCalendar();
  saveState();
  showToast('Persona eliminada', 'success');
}

function deletePerson(personId) {
  reassignShiftsFromPerson(personId);
  state.personnel = state.personnel.filter(p => p.id !== personId);
  renderPersonList();
  renderCalendar();
  saveState();
}

// --- Estadísticas: total, normales, fines de semana
function getStatsForMonth() {
  const key = monthKey(state.currentYear, state.currentMonth);
  const shifts = state.shifts[key] || {};
  const totalDays = daysInMonth(state.currentYear, state.currentMonth);
  const stats = {};

  state.personnel.forEach(p => {
    stats[p.id] = { name: p.name, total: 0, weekdays: 0, weekends: 0 };
  });

  for (let day = 1; day <= totalDays; day++) {
    const personId = shifts[String(day)];
    if (!personId || !stats[personId]) continue;
    stats[personId].total++;
    if (isWeekend(state.currentYear, state.currentMonth, day)) {
      stats[personId].weekends++;
    } else {
      stats[personId].weekdays++;
    }
  }

  return Object.entries(stats).map(([id, data]) => ({ id, ...data }));
}

function renderStats() {
  const container = document.getElementById('statsCards');
  const stats = getStatsForMonth();
  container.innerHTML = '';

  stats
    .filter(s => s.total > 0 || state.personnel.find(p => p.id === s.id))
    .sort((a, b) => b.total - a.total)
    .forEach(s => {
      const card = document.createElement('div');
      card.className = 'stat-card';
      card.innerHTML = `
        <h4>${escapeHtml(s.name)}</h4>
        <div class="stat-row"><span>Total turnos</span><span>${s.total}</span></div>
        <div class="stat-row"><span>Días laborales</span><span>${s.weekdays}</span></div>
        <div class="stat-row"><span>Fines de semana</span><span>${s.weekends}</span></div>
      `;
      container.appendChild(card);
    });
}

function doDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

// Abreviaturas día de la semana (Dom=0 ... Sáb=6)
const WEEKDAY_ABBR = ['DO', 'LU', 'MA', 'MI', 'JU', 'VI', 'SA'];

// --- Exportar/Compartir Excel (formato GCIG: matriz + resumen)
let pendingShare = null; // { blob, fileName, shareText } para el modal "Planilla lista"

/** Genera el Excel del mes y devuelve { blob, fileName, shareText }. Lanza si falla. */
async function generateExcelBlob() {
  if (typeof ExcelJS === 'undefined') {
    throw new Error('Librería de Excel no cargada');
  }
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Turnos TD';
  const sheet = workbook.addWorksheet('Turnos del mes', { views: [{ state: 'frozen', ySplit: 3 }] });

  const key = monthKey(state.currentYear, state.currentMonth);
  const shifts = state.shifts[key] || {};
  const totalDays = daysInMonth(state.currentYear, state.currentMonth);
  const activePersonnel = getActivePersonnel();
  const stats = getStatsForMonth();
  const statsById = {};
  stats.forEach(s => { statsById[s.id] = s.total; });

  const monthLabel = new Date(state.currentYear, state.currentMonth - 1, 1).toLocaleDateString('es-AR', { month: 'long', year: 'numeric' });
  const monthUpper = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
  const titleText = `TURNO DE OPERADORES GCIG COMODORO RIVADAVIA ${monthUpper.toUpperCase()} DEL ${state.currentYear}`;

  const borderThin = { style: 'thin' };
  const borderThick = { style: 'medium' };
  const weekendFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
  const centerAlign = { horizontal: 'center', vertical: 'middle' };
  const bold = { bold: true };

  // --- Fila 1: Título (fusionado)
  sheet.mergeCells(1, 1, 1, totalDays + 1);
  const titleCell = sheet.getCell(1, 1);
  titleCell.value = titleText;
  titleCell.font = bold;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  sheet.getRow(1).height = 28;

  // --- Filas 2 y 3: Encabezados de días (día de semana + número)
  for (let day = 1; day <= totalDays; day++) {
    const col = day + 1;
    const dow = getDayOfWeek(state.currentYear, state.currentMonth, day);
    const weekend = isWeekend(state.currentYear, state.currentMonth, day);
    const dayStyle = weekend ? { fill: weekendFill, font: bold, alignment: centerAlign, border: { top: borderThin, left: borderThin, bottom: borderThin, right: borderThin } } : { font: bold, alignment: centerAlign, border: { top: borderThin, left: borderThin, bottom: borderThin, right: borderThin } };
    sheet.getCell(2, col).value = WEEKDAY_ABBR[dow];
    sheet.getCell(2, col).style = dayStyle;
    sheet.getCell(3, col).value = String(day).padStart(2, '0');
    sheet.getCell(3, col).style = dayStyle;
  }
  sheet.getCell(2, 1).value = 'Grado Apellido y Nombre';
  sheet.getCell(2, 1).style = { font: bold, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: borderThin, left: borderThick, bottom: borderThin, right: borderThin }, fill: headerFill };
  sheet.getCell(3, 1).value = '';
  sheet.getCell(3, 1).style = { border: { top: borderThin, left: borderThick, bottom: borderThin, right: borderThin }, fill: headerFill };
  sheet.getRow(2).height = 20;
  sheet.getRow(3).height = 20;

  // --- Filas de operadores: una fila por persona, "T" en el día asignado
  activePersonnel.forEach((person, idx) => {
    const rowNum = 4 + idx;
    sheet.getCell(rowNum, 1).value = person.name.toUpperCase();
    sheet.getCell(rowNum, 1).style = { font: bold, border: { top: borderThin, left: borderThick, bottom: borderThin, right: borderThin } };
    for (let day = 1; day <= totalDays; day++) {
      const col = day + 1;
      const hasShift = shifts[String(day)] === person.id;
      const weekend = isWeekend(state.currentYear, state.currentMonth, day);
      const cell = sheet.getCell(rowNum, col);
      cell.value = hasShift ? 'T' : '';
      cell.alignment = centerAlign;
      cell.border = { top: borderThin, left: borderThin, bottom: borderThin, right: borderThin };
      if (weekend) cell.fill = weekendFill;
    }
  });

  // Borde exterior de la tabla principal (derecha e inferior)
  for (let r = 2; r <= 3 + activePersonnel.length; r++) {
    sheet.getCell(r, totalDays + 1).border = { ...sheet.getCell(r, totalDays + 1).border, right: borderThick };
  }
  for (let c = 1; c <= totalDays + 1; c++) {
    const lastRow = 3 + activePersonnel.length;
    sheet.getCell(lastRow, c).border = { ...sheet.getCell(lastRow, c).border, bottom: borderThick };
  }

  // --- Tabla de resumen: OPERADORES | T (total turnos)
  const summaryStartRow = 5 + activePersonnel.length;
  sheet.getCell(summaryStartRow, 1).value = 'OPERADORES';
  sheet.getCell(summaryStartRow, 1).style = { font: bold, border: { top: borderThick, left: borderThick, bottom: borderThin, right: borderThin }, fill: headerFill };
  sheet.getCell(summaryStartRow, 2).value = 'T';
  sheet.getCell(summaryStartRow, 2).style = { font: bold, alignment: centerAlign, border: { top: borderThick, left: borderThin, bottom: borderThin, right: borderThick }, fill: headerFill };
  sheet.getRow(summaryStartRow).height = 20;

  activePersonnel.forEach((person, idx) => {
    const r = summaryStartRow + 1 + idx;
    const total = statsById[person.id] || 0;
    sheet.getCell(r, 1).value = person.name.toUpperCase();
    sheet.getCell(r, 1).style = { border: { top: borderThin, left: borderThick, bottom: borderThin, right: borderThin } };
    sheet.getCell(r, 2).value = total > 0 ? total : '';
    sheet.getCell(r, 2).style = { alignment: centerAlign, border: { top: borderThin, left: borderThin, bottom: borderThin, right: borderThick } };
  });
  const lastSummaryRow = summaryStartRow + activePersonnel.length;
  sheet.getCell(lastSummaryRow, 1).border = { ...sheet.getCell(lastSummaryRow, 1).border, bottom: borderThick };
  sheet.getCell(lastSummaryRow, 2).border = { ...sheet.getCell(lastSummaryRow, 2).border, bottom: borderThick };

  // Anchos de columna
  sheet.getColumn(1).width = 28;
  for (let d = 1; d <= totalDays; d++) sheet.getColumn(d + 1).width = 5;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const fileName = `Turnos_${state.currentYear}_${String(state.currentMonth).padStart(2, '0')}.xlsx`;
  const monthLabelShare = new Date(state.currentYear, state.currentMonth - 1, 1).toLocaleDateString('es-AR', { month: 'long', year: 'numeric' });
  const capShare = monthLabelShare.charAt(0).toUpperCase() + monthLabelShare.slice(1);
  const shareText = `Planilla de turnos - ${capShare}. Revisa el día que te toca.`;
  return { blob, fileName, shareText };
}

async function downloadExcel() {
  const btn = document.getElementById('btnDownloadExcel');
  if (btn) { btn.disabled = true; btn.textContent = 'Generando…'; }
  try {
    const { blob, fileName } = await generateExcelBlob();
    doDownload(blob, fileName);
    showToast('Descargado', 'success');
  } catch (err) {
    console.error(err);
    showToast('Error al generar el Excel.', 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Descargar Excel'; }
  }
}

async function shareExcel() {
  const btn = document.getElementById('btnShareExcel');
  const btnText = document.getElementById('btnShareExcelText');
  if (btn) btn.disabled = true;
  if (btnText) btnText.textContent = 'Generando…';

  try {
    // Generar el blob primero. Esto puede tardar un poco y rompería el "User Gesture"
    // si intentamos llamar a navigator.share directamente después.
    const data = await generateExcelBlob();
    pendingShare = data;

    // Mostramos el modal de "Planilla lista". 
    // El usuario tocará "Compartir" en el modal, lo que disparará un gesto fresco y válido.
    document.getElementById('modalShareReady').showModal();
  } catch (err) {
    console.error(err);
    showToast('Error al generar el Excel.', 'error');
  } finally {
    if (btn) btn.disabled = false;
    if (btnText) btnText.textContent = 'Compartir';
  }
}

// --- Modal "Planilla lista": Compartir (con gesto de usuario) o Descargar
function doShareFromModal() {
  if (!pendingShare) return;
  const data = pendingShare; // Referencia local para evitar race conditions
  pendingShare = null; // Lo limpiamos inmediatamente en el estado global
  document.getElementById('modalShareReady').close();

  if (!navigator.share) {
    doDownload(data.blob, data.fileName);
    return;
  }

  const file = new File([data.blob], data.fileName, { type: data.blob.type });
  navigator.share({
    files: [file],
    title: 'Turnos TD',
    text: data.shareText
  }).then(() => {
    showToast('Compartido', 'success');
  }).catch((e) => {
    if (e.name !== 'AbortError') {
      console.error('Share failed, falling back to download:', e);
      doDownload(data.blob, data.fileName);
    }
  });
}

document.getElementById('btnShareNow').addEventListener('click', doShareFromModal);

document.getElementById('modalShareReady').addEventListener('close', () => {
  pendingShare = null;
  const btn = document.getElementById('btnShareExcel');
  const btnText = document.getElementById('btnShareExcelText');
  if (btn) btn.disabled = false;
  if (btnText) btnText.textContent = 'Compartir';
});

// --- Cerrar modales
document.querySelectorAll('[data-close]').forEach(btn => {
  btn.addEventListener('click', () => {
    const id = btn.dataset.close;
    document.getElementById(id).close();
  });
});

document.getElementById('confirmAssign').addEventListener('click', confirmAssign);
document.getElementById('savePerson').addEventListener('click', savePerson);
document.getElementById('confirmDeletePerson').addEventListener('click', confirmDeletePerson);
document.getElementById('btnAddPerson').addEventListener('click', () => openPersonModal(null));
document.getElementById('btnDownloadExcel').addEventListener('click', downloadExcel);
document.getElementById('btnShareExcel').addEventListener('click', shareExcel);

// Inicio
loadState();
setMonthLabel();
initNavigation();
initMonthNav();
renderCalendar();
renderPersonList();
