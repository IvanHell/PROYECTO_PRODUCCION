/* ===================================================================
   PANEL DE PRODUCCIÓN — MÁQUINAS VALLE (v2)
   Todos los JavaScript en un solo archivo
   Incluye: Máquinas, Manuales, Paro, Tarimas, Contenedores + Excel
   =================================================================== */

/* ======================= CONFIGURACIÓN ======================= */

const MACHINES = [
    { id: 'FLEXO', label: 'Flexo', piezasMin: 26, ajuste: 30, flexo: true, aliases: ['FLEXO'] },
    { id: 'ROLADORA', label: 'Roladora I', piezasMin: 10, ajuste: 15, aliases: ['ROLADORA', 'ROLADORA I'] },
    { id: 'RANURADORA', label: 'Ranuradora', piezasMin: 37, ajuste: 35, aliases: ['RANURADORA'] },
    { id: 'T_ROTATIVA', label: 'Troqueladora Rotativa', piezasMin: 35, ajuste: 50, aliases: ['T. ROTATIVA', 'T . ROTATIVA', 'TROQUELADORA ROTATIVA'] },
    { id: 'CAIMAN', label: 'Caimán', piezasMin: 5, ajuste: 15, aliases: ['CAIMAN', 'CAIMÁN'] },
    { id: 'PEGADORA', label: 'Pegadora', piezasMin: 80, ajuste: 30, aliases: ['PEGADORA', 'Pegadora'] },
    { id: 'T_PLANA', label: 'Troqueladora Plana', piezasMin: 35, ajuste: 50, aliases: ['T. PLANA', 'TROQ. PLANA', 'TROQUELADORA PLANA'] },
    { id: 'ARMADORA_REJILLAS', label: 'Máquina Armadora (Armadora de Rejillas)', piezasMin: 26, ajuste: 30, aliases: ['ARMADORA DE REJILLAS', 'MAQUINA ARMADORA', 'MÁQUINA ARMADORA'] },
];

const OPERACIONES = ['DESPIQUE', 'ARMADO', 'CORTE', 'GRAPADO', 'PEGADO', 'ENSAMBLE', 'DOBLADO', 'CONTEO', 'SELLADO', 'AMARRADO', 'TARIMA'];

const DIAS_ABR = ['D', 'L', 'M', 'MI', 'J', 'V', 'S'];
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

/* ======================= CACHE CLIENTE ======================= */
const SHEETS_CLIENT_CACHE = {};
const CACHE_TTL_MS = 2 * 60 * 1000; // 2 minutos

/* ======================= HELPERS DE CONFIGURACIÓN ======================= */

function machineById(id) {
    return MACHINES.find(m => m.id === id);
}

function normalizeMachine(raw) {
    if (!raw) return null;
    const clean = String(raw).trim().toUpperCase().replace(/\s+/g, ' ');
    for (const m of MACHINES) {
        if (m.aliases.some(a => a.toUpperCase() === clean)) return m.id;
    }
    return null;
}

function defaultTiempoDisponible(machineId, dateStr) { // sirve para crear la constante de tiempo "defaul: 570min"
    const m = machineById(machineId);
    const d = new Date(dateStr + 'T00:00:00');
    const dow = d.getDay();
    //if (m && m.flexo) return null;
    if (dow === 0 || dow === 6) return 0;
    if (dow === 2) return 600;
    return 570;
}

function pad2(n) {
    return String(n).padStart(2, '0');
}

function toISODate(d) {
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

// Formatea una fecha interna (YYYY-MM-DD) al formato definitivo de
// visualización DD/MM/YYYY (el que usamos en Excel/Sheets/México).
function formatDateDisplay(isoDate) {
    if (!isoDate || typeof isoDate !== 'string') return '';
    const [y, m, d] = isoDate.split('-');
    if (!y || !m || !d) return isoDate;
    return `${d}/${m}/${y}`;
}

// Parser único de fechas: acepta objetos Date, números de serie de Excel,
// texto ISO (YYYY-MM-DD) y texto DD/MM/YYYY o DD/MM/YY.
// REGLA DEFINITIVA: cualquier fecha con "/" ambigua (como "05/06/2026")
// se interpreta SIEMPRE como DD/MM/YYYY, nunca como MM/DD/YYYY.
// Internamente todo se sigue guardando como YYYY-MM-DD (necesario para
// poder ordenar y comparar fechas como texto sin errores).
function parseFlexibleDate(v) {
    if (!v && v !== 0) return '';
    if (v instanceof Date) return isNaN(v) ? '' : toISODate(v);

    if (typeof v === 'number') {
        const d = XLSX.SSF.parse_date_code(v);
        return d ? `${d.y}-${pad2(d.m)}-${pad2(d.d)}` : '';
    }

    const s = String(v).trim();
    if (!s) return '';

    // Ya viene en formato interno ISO
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

    // DD/MM/YYYY o DD/MM/YY — el estándar definitivo para fechas con "/"
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
        let [, d, mo, y] = m;
        if (y.length === 2) y = (Number(y) < 50 ? '20' : '19') + y;
        return `${y}-${pad2(mo)}-${pad2(d)}`;
    }

    // Último recurso (formatos no previstos)
    const d2 = new Date(s);
    return isNaN(d2) ? '' : toISODate(d2);
}

// Se mantiene el nombre anterior por compatibilidad con el resto del código
function parseExcelDate(v) {
    return parseFlexibleDate(v);
}

function uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/* ======================= STORAGE HELPERS ======================= */
/*
 * Detecta automáticamente dónde está corriendo el panel:
 * - Dentro de una conversación de Claude (vista previa de artefacto):
 *   usa window.storage, que Claude provee.
 * - Abierto como archivo local (file://) o subido a un servidor propio:
 *   window.storage NO existe ahí, así que usa localStorage del navegador
 *   (nativo, no depende de Claude, funciona igual en cualquier lado).
 * Un mismo prefijo evita chocar con otras cosas guardadas en el navegador.
 */
const LS_PREFIX = 'panel_produccion_';
const usaClaudeStorage = typeof window.storage !== 'undefined' && window.storage !== null;

async function loadArr(key) {
    try {
        if (usaClaudeStorage) {
            const r = await window.storage.get(key, false);
            return r && r.value ? JSON.parse(r.value) : [];
        }
        const raw = localStorage.getItem(LS_PREFIX + key);
        return raw ? JSON.parse(raw) : [];
    } catch (e) {
        console.error('storage error (loadArr)', e);
        return [];
    }
}

async function saveArr(key, arr) {
    try {
        if (usaClaudeStorage) {
            await window.storage.set(key, JSON.stringify(arr), false);
        } else {
            localStorage.setItem(LS_PREFIX + key, JSON.stringify(arr));
        }
        return true;
    } catch (e) {
        console.error('storage error (saveArr)', e);
        return false;
    }
}

async function loadObj(key) {
    try {
        if (usaClaudeStorage) {
            const r = await window.storage.get(key, false);
            return r && r.value ? JSON.parse(r.value) : {};
        }
        const raw = localStorage.getItem(LS_PREFIX + key);
        return raw ? JSON.parse(raw) : {};
    } catch (e) {
        console.error('storage error (loadObj)', e);
        return {};
    }
}

async function saveObj(key, obj) {
    try {
        if (usaClaudeStorage) {
            await window.storage.set(key, JSON.stringify(obj), false);
        } else {
            localStorage.setItem(LS_PREFIX + key, JSON.stringify(obj));
        }
        return true;
    } catch (e) {
        console.error('storage error (saveObj)', e);
        return false;
    }
}

/* ======================= ESTADO GLOBAL ======================= */

const DB = {
    maquinas: [],
    manuales: [],
    paro: [],
    tarimas: [],
    contenedores: [],
    tiempoOverrides: {}
};

let SHEETS_CONFIG = { url: '', enabled: false };

async function loadAll() {
    DB.maquinas = await loadArr('entries-maquinas');
    DB.manuales = await loadArr('entries-manuales');
    DB.paro = await loadArr('entries-paro');
    DB.tarimas = await loadArr('entries-tarimas');
    DB.contenedores = await loadArr('entries-contenedores');
    DB.tiempoOverrides = await loadObj('config-tiempo-overrides');
    SHEETS_CONFIG = Object.assign({ url: '', enabled: false }, await loadObj('config-sheets-sync'));
}

/* ======================= UTILIDADES ======================= */

function val(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : '';
}

function numval(id) {
    const el = document.getElementById(id);
    return el ? (Number(el.value) || 0) : 0;
}

function setStatus(id, msg, isError = false) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = msg;
    el.style.color = isError ? '#dc2626' : '#16a34a';
    setTimeout(() => {
        el.textContent = '';
    }, 4000);
}

function getToday() {
    return new Date().toISOString().slice(0, 10);
}

function getCurrentMonth() {
    return new Date().getMonth() + 1;
}

function getDaysInMonth(year, month) {
    return new Date(year, month, 0).getDate();
}

function getDayOfWeek(dateStr) {
    return new Date(dateStr + 'T00:00:00').getDay();
}

function parseDate(dateStr) {
    return new Date(dateStr + 'T00:00:00');
}

function filterByMonthYear(arr, month, year) {
    return arr.filter(e => {
        const d = parseDate(e.fecha);
        return d.getFullYear() === year && (d.getMonth() + 1) === month;
    });
}

function sum(arr, prop) {
    return arr.reduce((s, e) => s + (Number(e[prop]) || 0), 0);
}

function opsSummary(ops) {
    if (!ops) return '';
    return OPERACIONES.filter(op => ops[op]).join(', ');
}

/* ======================= UI HELPERS ======================= */

function fillMachineSelect(sel) {
    const select = document.getElementById(sel);
    if (!select) return;
    select.innerHTML = '<option value="">Selecciona...</option>' +
        MACHINES.map(m => `<option value="${m.id}">${m.label}</option>`).join('');
}

function fillMonthSelect(sel) {
    const select = document.getElementById(sel);
    if (!select) return;
    select.innerHTML = MESES.map((m, i) => `<option value="${i + 1}">${m}</option>`).join('');
}

function buildChecklist() {
    const el = document.getElementById('m-checklist');
    if (!el) return;
    el.innerHTML = OPERACIONES.map(op =>
        `<label class="check-item"><input type="checkbox" class="m-op" value="${op}" onchange="window.updateSubProc()"> ${op}</label>`
    ).join('');
}

function updateSubProc() {
    const n = document.querySelectorAll('.m-op:checked').length;
    const el = document.getElementById('m-subproc');
    if (el) el.value = n;
}
window.updateSubProc = updateSubProc;

/* ================================================================
   OPERADORES DINÁMICOS (Enter o botón "+" agrega otro nombre;
   el RRHH se calcula solo contando cuántos nombres hay capturados)
   ================================================================ */

function makeOperadorRow(containerId, rrhhId, primero) {
    const row = document.createElement('div');
    row.className = 'operador-row';

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'operador-item';
    input.placeholder = primero ? 'Nombre del operador' : 'Otro operador';
    input.addEventListener('input', () => updateOperadoresRRHH(containerId, rrhhId));
    input.addEventListener('keydown', (ev) => {
        if (ev.key === 'Enter') {
            ev.preventDefault();
            if (input.value.trim()) addOperadorRow(containerId, rrhhId);
        }
    });

    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'op-btn op-add';
    addBtn.textContent = '+';
    addBtn.title = 'Agregar otro operador';
    addBtn.onclick = () => addOperadorRow(containerId, rrhhId);

    row.appendChild(input);
    row.appendChild(addBtn);

    if (!primero) {
        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'op-btn op-remove';
        removeBtn.textContent = '×';
        removeBtn.title = 'Quitar este operador';
        removeBtn.onclick = () => {
            row.remove();
            updateOperadoresRRHH(containerId, rrhhId);
        };
        row.appendChild(removeBtn);
    }

    return { row, input };
}

function initOperadorList(containerId, rrhhId) {
    const cont = document.getElementById(containerId);
    if (!cont) return;
    cont.innerHTML = '';
    const { row, input } = makeOperadorRow(containerId, rrhhId, true);
    cont.appendChild(row);
    updateOperadoresRRHH(containerId, rrhhId);
}

function addOperadorRow(containerId, rrhhId) {
    const cont = document.getElementById(containerId);
    if (!cont) return;
    const { row, input } = makeOperadorRow(containerId, rrhhId, false);
    cont.appendChild(row);
    input.focus();
    updateOperadoresRRHH(containerId, rrhhId);
}

function updateOperadoresRRHH(containerId, rrhhId) {
    const cont = document.getElementById(containerId);
    const rrhhEl = document.getElementById(rrhhId);
    if (!cont || !rrhhEl) return;
    const n = [...cont.querySelectorAll('.operador-item')].filter(i => i.value.trim()).length;
    rrhhEl.value = n;
}

function getOperadoresString(containerId) {
    const cont = document.getElementById(containerId);
    if (!cont) return '';
    return [...cont.querySelectorAll('.operador-item')]
        .map(i => i.value.trim())
        .filter(x => x)
        .join(', ');
}

function resetOperadorList(containerId, rrhhId) {
    initOperadorList(containerId, rrhhId);
}

function setupTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
            const target = document.getElementById('panel-' + btn.dataset.tab);
            if (target) target.classList.add('active');
        });
    });

    document.querySelectorAll('.subtabs').forEach(group => {
        group.querySelectorAll('.subtab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                group.querySelectorAll('.subtab-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                const parent = group.parentElement;
                parent.querySelectorAll('.subpanel').forEach(p => p.classList.remove('active'));
                const target = document.getElementById(btn.dataset.sub);
                if (target) target.classList.add('active');
            });
        });
    });
}

function setDateInputs() {
    const today = toISODate(new Date());
    ['f-fecha', 'm-fecha', 'p-fecha', 't-fecha', 'c-fecha'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = today;
    });
}

function setMonthSelects() {
    const month = getCurrentMonth();
    ['r-mes', 'rk-mes', 'rm-mes', 'rt-mes', 'rc-mes', 'rtop-mes', 'rnp-mes'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = month;
    });
}

/* ======================= HERENCIA DE HORARIOS ======================= */

const LAST_END_KEY = LS_PREFIX + 'ultima-hora-fin-';

function guardarUltimaHoraFin(contexto, horaFin, fecha) {
    try {
        localStorage.setItem(LAST_END_KEY + contexto, JSON.stringify({ hora: horaFin, fecha }));
    } catch(e) {}
}

function obtenerUltimaHoraFin(contexto, fechaActual) {
    try {
        const raw = localStorage.getItem(LAST_END_KEY + contexto);
        if (!raw) return null;
        const obj = JSON.parse(raw);
        return obj.fecha === fechaActual ? obj.hora : null;
    } catch(e) { return null; }
}

function minutosDesdeMedianoche(hora) {
    const [h, m] = hora.split(':').map(Number);
    return h * 60 + m;
}

// Precarga la hora de inicio heredada y, si es necesario, resetea el fin a 17:00
function aplicarHoraHerencia(inputInicioId, inputFinId, contexto, fechaActual) {
    const ultima = obtenerUltimaHoraFin(contexto, fechaActual);
    if (!ultima) return;
    const elInicio = document.getElementById(inputInicioId);
    const elFin   = document.getElementById(inputFinId);
    if (!elInicio) return;
    
    elInicio.value = ultima;
    
    // Si el fin queda antes o igual al inicio, lo reseteamos a 17:00 para evitar 0 min
    if (elFin && minutosDesdeMedianoche(elFin.value) <= minutosDesdeMedianoche(ultima)) {
        elFin.value = '17:00';
    }
    
    if (inputInicioId === 'f-hora-inicio') calcularTiempoMaquina();
    if (inputInicioId === 'm-hora-inicio') calcularTiempoManual();
}

/* ================================================================
   EXCEL IMPORT / EXPORT
   ================================================================ */

function downloadExcel(dataArrayOfObjects, filename, sheetName) {
    const ws = XLSX.utils.json_to_sheet(dataArrayOfObjects);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, (sheetName || 'Hoja1').slice(0, 31));
    XLSX.writeFile(wb, filename);
}

/* ================================================================
   RENDERIZADO DE TABLAS (CAPTURA)
   ================================================================ */

function renderMaquinasTable() {
    const t = document.getElementById('tbl-maquinas');
    const hoy = new Date();
    const rows = filterByMonthYear(DB.maquinas, hoy.getMonth() + 1, hoy.getFullYear())
        .sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Máquina</th><th>Producto</th><th>NP</th><th>H.V.</th><th>Operador</th><th>RRHH</th><th>Cant. Prod.</th><th>Tiempo</th><th>T. Disp.</th><th>Comentarios</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${machineById(e.maquina)?.label || e.maquina}</td>
                <td class="left">${e.producto || ''}</td>
                <td>${e.np || ''}</td>
                <td>${e.hv || ''}</td>
                <td class="left">${e.operador || ''}</td>
                <td>${e.rrhh || 0}</td>
                <td>${e.cant || 0}</td>
                <td>${e.tiempo || 0}</td>
                <td>${e.tiempoDisp || ''}</td>
                <td class="left">${e.comentarios || ''}</td>
                <td><button class="btn small danger" onclick="window.deleteMaquinaEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="12" class="empty">Sin registros este mes</td></tr>`;
}

function renderManualesTable() {
    const t = document.getElementById('tbl-manuales');
    const hoy = new Date();
    const rows = filterByMonthYear(DB.manuales, hoy.getMonth() + 1, hoy.getFullYear())
        .sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Producto</th><th>NP</th><th>H.V.</th><th>Operador</th><th>RRHH</th><th>Cant.</th><th>Tiempo</th><th>Operaciones</th><th>Sub Proc</th><th>Comentarios</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${e.producto || ''}</td>
                <td>${e.np || ''}</td>
                <td>${e.hv || ''}</td>
                <td class="left">${e.operador || ''}</td>
                <td>${e.rrhh || 0}</td>
                <td>${e.cant || 0}</td>
                <td>${e.tiempo || 0}</td>
                <td class="left">${opsSummary(e.ops)}</td>
                <td>${e.subproc || 0}</td>
                <td class="left">${e.comentarios || ''}</td>
                <td><button class="btn small danger" onclick="window.deleteManualEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="12" class="empty">Sin registros este mes</td></tr>`;
}

function renderParoTable() {
    const t = document.getElementById('tbl-paro');
    const hoy = new Date();
    const rows = filterByMonthYear(DB.paro, hoy.getMonth() + 1, hoy.getFullYear())
        .sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Máquina</th><th>Minutos</th><th>Motivo</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${machineById(e.maquina)?.label || e.maquina}</td>
                <td>${e.minutos || 0}</td>
                <td class="left">${e.motivo || ''}</td>
                <td><button class="btn small danger" onclick="window.deleteParoEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="5" class="empty">Sin registros este mes</td></tr>`;
}

function renderTarimasTable() {
    const t = document.getElementById('tbl-tarimas');
    const hoy = new Date();
    const rows = filterByMonthYear(DB.tarimas, hoy.getMonth() + 1, hoy.getFullYear())
        .sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Persona</th><th>Tipo</th><th>Cantidad</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${e.persona || ''}</td>
                <td class="left">${e.tipo || ''}</td>
                <td>${e.cant || 0}</td>
                <td><button class="btn small danger" onclick="window.deleteTarimaEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="5" class="empty">Sin registros este mes</td></tr>`;
}

function renderContenedoresTable() {
    const t = document.getElementById('tbl-contenedores');
    const hoy = new Date();
    const rows = filterByMonthYear(DB.contenedores, hoy.getMonth() + 1, hoy.getFullYear())
        .sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Turno</th><th>Personas</th><th>Contenedores</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${e.turno || ''}</td>
                <td>${e.personas || 0}</td>
                <td>${e.cant || 0}</td>
                <td><button class="btn small danger" onclick="window.deleteContenedorEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="5" class="empty">Sin registros este mes</td></tr>`;
}

/* ================================================================
   CAPTURA DE DATOS
   ================================================================ */

// ===== MÁQUINAS =====

async function addMaquinaEntry() {
    const fecha = val('f-fecha');
    const maquina = val('f-maquina');
    if (!fecha || !maquina) {
        setStatus('status-maquinas', 'Fecha y máquina son obligatorias.', true);
        return;
    }
    // Verificar si ya existe un tiempo disponible para esta máquina+fecha
    const existente = DB.maquinas.find(e => e.maquina === maquina && e.fecha === fecha && e.tiempoDisp > 0);
    let tiempoDisp = numval('f-tiempo-disp');

    if (existente) {
        // Ya hay un tiempo disponible registrado para este día: usar el existente
        tiempoDisp = existente.tiempoDisp;
    } else if (tiempoDisp > 0) {
        // Primer registro del día: guardar como override para el resumen
        const overrideKey = maquina + '_' + fecha;
        DB.tiempoOverrides[overrideKey] = tiempoDisp;
        await saveObj('config-tiempo-overrides', DB.tiempoOverrides);
    }

    const nuevo = {
        id: uid(),
        fecha,
        maquina,
        producto: val('f-producto'),
        np: val('f-np'),
        operador: getOperadoresString('f-operadores-list'),
        rrhh: numval('f-rrhh'),
        cant: numval('f-cant'),
        tiempo: numval('f-tiempo'),
        tiempoDisp: tiempoDisp,
        hv: val('f-hv'),
        comentarios: val('f-comentarios')
    };
    DB.maquinas.push(nuevo);
    await saveArr('entries-maquinas', DB.maquinas);

    pushToSheets('MAQUINAS', mapMaquinaRow(nuevo));
    // Clear rápido: conserva operador(es) y RRHH para capturas seguidas de la misma máquina
    ['f-producto', 'f-np', 'f-cant', 'f-tiempo', 'f-comentarios', 'f-hv']
        .forEach(id => document.getElementById(id).value = '');

    // >>> HERENCIA DE HORARIO <<<
    const finGuardado = val('f-hora-fin');
    guardarUltimaHoraFin(maquina, finGuardado, fecha);
    aplicarHoraHerencia('f-hora-inicio', 'f-hora-fin', maquina, fecha);
    // -----------------------------

    // Si acabamos de registrar el primer tiempo disponible del día, bloquear el campo
    if (!existente && tiempoDisp > 0) {
        updateTiempoDispField();
    }
    setStatus('status-maquinas', 'Guardado ✓ — siguiente captura desde las ' + finGuardado);
    renderMaquinasTable();
}

async function deleteMaquinaEntry(id) {
    DB.maquinas = DB.maquinas.filter(e => e.id !== id);
    await saveArr('entries-maquinas', DB.maquinas);
    await saveObj('config-tiempo-overrides', DB.tiempoOverrides);
    renderMaquinasTable();
}

async function handleImportMaquinas(file) {
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { defval: null });
    let imported = 0,
        ignored = 0;
    json.forEach(row => {
        const fecha = parseExcelDate(row['FECHA']); 
        const maquinaId = normalizeMachine(row['MAQUINA']);
        if (!fecha || !maquinaId) {
            if (row['FECHA']) ignored++;
            return;
        }
        const tiempoDispImport = Number(row['TIEMPO DISP']) || 0;
        DB.maquinas.push({
            id: uid(),
            fecha,
            maquina: maquinaId,
            producto: row['PRODUCTO'] || '',
            np: String(row['NP'] ?? ''),
            hv: String(row['H.V.'] ?? ''),
            operador: row['OPERADOR'] || '',
            rrhh: Number(row['RRHH']) || 0,
            cant: Number(row['CANT PROD']) || 0,
            tiempo: Number(row['TIEMPO']) || 0,
            tiempoDisp: tiempoDispImport,
            comentarios: row['COMENTARIOS'] || ''
        });
        // Si el Excel trae tiempo disponible, guardar como override
        if (tiempoDispImport > 0) {
            const overrideKey = maquinaId + '_' + fecha;
            DB.tiempoOverrides[overrideKey] = tiempoDispImport;
        }
        imported++;
    });
    await saveArr('entries-maquinas', DB.maquinas);
    await saveObj('config-tiempo-overrides', DB.tiempoOverrides);
    renderMaquinasTable();
    document.getElementById('import-maquinas').value = '';
    setStatus('status-maquinas', `Importados ${imported} registros (${ignored} ignorados).`);
}

function exportMaquinas() {
    const data = DB.maquinas.map(e => ({
        FECHA: formatDateDisplay(e.fecha),
        MAQUINA: machineById(e.maquina)?.label || e.maquina,
        PRODUCTO: e.producto,
        NP: e.np,
        'H.V.': e.hv || '',
        OPERADOR: e.operador,
        RRHH: e.rrhh,
        'CANT PROD': e.cant,
        TIEMPO: e.tiempo,
        'TIEMPO DISP': e.tiempoDisp || '',
        COMENTARIOS: e.comentarios
    }));
    downloadExcel(data, 'produccion_maquinas.xlsx', 'PRODUCCION');
}

// ===== MANUALES =====

async function addManualEntry() {
    const fecha = val('m-fecha');
    if (!fecha) {
        setStatus('status-manuales', 'La fecha es obligatoria.', true);
        return;
    }
    // Asegurar que el tiempo esté calculado antes de guardar
    calcularTiempoManual();
    const ops = {};
    OPERACIONES.forEach(op => {
        const cb = document.querySelector(`.m-op[value="${op}"]`);
        ops[op] = cb && cb.checked ? 1 : 0;
    });
    const subproc = Object.values(ops).reduce((a, b) => a + b, 0);
    const nuevo = {
        id: uid(),
        fecha,
        producto: val('m-producto'),
        np: val('m-np'),
        hv: val('m-hv'),
        operador: getOperadoresString('m-operadores-list'),
        rrhh: numval('m-rrhh'),
        cant: numval('m-cant'),
        tiempo: numval('m-tiempo'),
        ops,
        subproc,
        comentarios: val('m-comentarios')
    };
    DB.manuales.push(nuevo);
    await saveArr('entries-manuales', DB.manuales);
    pushToSheets('MANUALES', mapManualRow(nuevo));
    ['m-producto', 'm-np', 'm-cant', 'm-tiempo', 'm-comentarios']
        .forEach(id => document.getElementById(id).value = '');
    document.querySelectorAll('.m-op').forEach(cb => cb.checked = false);
    document.getElementById('m-subproc').value = '';
    
    // >>> HERENCIA DE HORARIO <<<
    const finGuardado = val('m-hora-fin');
    guardarUltimaHoraFin('manual', finGuardado, fecha);
    aplicarHoraHerencia('m-hora-inicio', 'm-hora-fin', 'manual', fecha);
    // -----------------------------
    
    setStatus('status-manuales', 'Guardado ✓ — siguiente captura desde las ' + finGuardado);
    renderManualesTable();
}

async function deleteManualEntry(id) {
    DB.manuales = DB.manuales.filter(e => e.id !== id);
    await saveArr('entries-manuales', DB.manuales);
    renderManualesTable();
}

async function handleImportManuales(file) {
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { defval: null });
    let imported = 0;
    json.forEach(row => {
        const fecha = parseExcelDate(row['FECHA']);
        if (!fecha) return;
        const ops = {};
        OPERACIONES.forEach(op => { ops[op] = (row[op] == 1 || row[op] === '1') ? 1 : 0; });
        const subproc = Object.values(ops).reduce((a, b) => a + b, 0);
        DB.manuales.push({
            id: uid(),
            fecha,
            producto: row['PRODUCTO'] || '',
            np: String(row['NP'] ?? ''),
            hv: String(row['H.V.'] ?? ''),
            operador: row['OPERADOR'] || '',
            rrhh: Number(row['RRHH']) || 0,
            cant: Number(row['CANT PROD']) || 0,
            tiempo: Number(row['TIEMPO']) || 0,
            ops,
            subproc,
            comentarios: row['COMENTARIOS'] || ''
        });
        imported++;
    });
    await saveArr('entries-manuales', DB.manuales);
    renderManualesTable();
    document.getElementById('import-manuales').value = '';
    setStatus('status-manuales', `Importados ${imported} registros.`);
}

function exportManuales() {
    const data = DB.manuales.map(e => {
        const row = {
            FECHA: formatDateDisplay(e.fecha),
            PRODUCTO: e.producto,
            NP: e.np,
            'H.V.': e.hv || '',
            OPERADOR: e.operador,
            RRHH: e.rrhh,
            'CANT PROD': e.cant,
            TIEMPO: e.tiempo
        };
        OPERACIONES.forEach(op => row[op] = (e.ops && e.ops[op]) ? 1 : '');
        row['SUB PROC'] = e.subproc || 0;
        row['COMENTARIOS'] = e.comentarios;
        return row;
    });
    downloadExcel(data, 'produccion_manuales.xlsx', 'PRODUCCION');
}

// ===== PARO =====

async function addParoEntry() {
    const fecha = val('p-fecha');
    const maquina = val('p-maquina');
    if (!fecha || !maquina) {
        setStatus('status-paro', 'Fecha y máquina son obligatorias.', true);
        return;
    }
    const nuevo = {
        id: uid(),
        fecha,
        maquina,
        minutos: numval('p-minutos'),
        motivo: val('p-motivo')
    };
    DB.paro.push(nuevo);
    await saveArr('entries-paro', DB.paro);
    pushToSheets('PARO', mapParoRow(nuevo));
    ['p-minutos', 'p-motivo'].forEach(id => document.getElementById(id).value = '');
    setStatus('status-paro', 'Guardado ✓');
    renderParoTable();
}

async function deleteParoEntry(id) {
    DB.paro = DB.paro.filter(e => e.id !== id);
    await saveArr('entries-paro', DB.paro);
    renderParoTable();
}

let lastFrecuenciaParosRows = [];

async function renderFrecuenciaParos(forceRefresh = false) {
    const mes = parseInt(val('rp-mes'), 10);
    const anio = parseInt(val('rp-anio'), 10);

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-paros', resultado.fuente, resultado.cacheado);

    const rows = (datosSheets.PARO || [])
        .map(convertirFilaParo)
        .filter(Boolean);

    // Agrupar por motivo
    const porMotivo = {};
    rows.forEach(e => {
        const motivo = (e.motivo || 'SIN MOTIVO').trim().toUpperCase();
        if (!porMotivo[motivo]) porMotivo[motivo] = { veces: 0, minutos: 0 };
        porMotivo[motivo].veces += 1;
        porMotivo[motivo].minutos += Number(e.minutos) || 0;
    });

    // Convertir a array y ordenar por minutos (descendente)
    let entries = Object.entries(porMotivo).map(([motivo, data]) => ({
        motivo, veces: data.veces, minutos: data.minutos
    })).sort((a, b) => b.minutos - a.minutos);

    const totalMinutos = entries.reduce((s, e) => s + e.minutos, 0);
    if (!totalMinutos) {
        document.getElementById('chart-paros-container').innerHTML = 
            '<div class="empty">Sin registros de paro este mes</div>';
        return;
    }

    // Calcular % acumulado
    let acum = 0;
    entries = entries.map(e => {
        acum += e.minutos;
        return { ...e, pct: Math.round((e.minutos / totalMinutos) * 100), pctAcum: Math.round((acum / totalMinutos) * 100) };
    });

    // Guardar para exportar
    lastFrecuenciaParosRows = entries.map(e => ({
        MOTIVO: e.motivo,
        VECES: e.veces,
        MINUTOS: e.minutos,
        '% DEL TOTAL': e.pct,
        '% ACUMULADO': e.pctAcum
    }));

    const max = entries[0].minutos;
    const limite80 = entries.findIndex(e => e.pctAcum > 80); // Último índice que forma el 80%

    // Renderizar
    const cont = document.getElementById('chart-paros-container');
    cont.innerHTML = `
        <div style="display:grid; grid-template-columns: 180px 1fr 70px 60px 60px; gap:8px; align-items:center; font-size:12px; font-weight:bold; color:var(--sub); margin-bottom:6px; padding:0 4px;">
            <div>Motivo</div>
            <div style="text-align:center;">Minutos totales (% acumulado)</div>
            <div style="text-align:center;">Veces</div>
            <div style="text-align:center;">% Ind.</div>
            <div style="text-align:center;">% Acum.</div>
        </div>
        ${entries.map((e, i) => {
            const barPct = max ? Math.round((e.minutos / max) * 100) : 0;
            const esPareto = i <= limite80; // Los que forman el 80%
            return `
            <div style="display:grid; grid-template-columns: 180px 1fr 70px 60px 60px; gap:8px; align-items:center; margin-bottom:10px; padding:8px; border-radius:8px; ${esPareto ? 'background:#fef3c7; border:1px solid #fcd34d;' : ''}">
                <div style="font-size:12.5px; font-weight:bold; word-break:break-word;" title="${e.motivo}">${e.motivo}</div>
                <div>
                    <div style="display:flex; align-items:center; gap:8px;">
                        <div class="top5-bar-bg" style="flex:1; height:24px;">
                            <div class="top5-bar-fill" style="width:${barPct}%; ${esPareto ? 'background:linear-gradient(90deg, #d97706, #f59e0b);' : 'background:linear-gradient(90deg, var(--accent), #60a5fa);'}"></div>
                        </div>
                        <span style="font-size:11px; color:var(--sub); white-space:nowrap;">${e.minutos} min</span>
                    </div>
                    ${esPareto && i === limite80 ? '<div style="font-size:10px; color:#92400e; margin-top:2px;">▲ Aquí se concentra el 80% del tiempo perdido</div>' : ''}
                </div>
                <div style="text-align:center; font-size:12px;">${e.veces}</div>
                <div style="text-align:center; font-size:12px; font-weight:bold;">${e.pct}%</div>
                <div style="text-align:center; font-size:12px; font-weight:bold; ${e.pctAcum >= 80 ? 'color:var(--bad);' : ''}">${e.pctAcum}%</div>
            </div>
            `;
        }).join('')}
        <div style="margin-top:12px; padding:10px; background:#f8f9fb; border-radius:8px; font-size:12px; color:var(--sub);">
            <strong>Total minutos de paro:</strong> ${totalMinutos} min &nbsp;|&nbsp; 
            <strong>Motivos distintos:</strong> ${entries.length} &nbsp;|&nbsp; 
            <strong>Regla 80/20:</strong> Los primeros ${limite80 + 1} motivos causan el 80% del tiempo perdido.
        </div>
    `;
}

function exportFrecuenciaParos() {
    if (!lastFrecuenciaParosRows.length) { alert('Genera el Pareto primero.'); return; }
    downloadExcel(lastFrecuenciaParosRows, `pareto_paros_${val('rp-mes')}_${val('rp-anio')}.xlsx`, 'PAROS_PARETO');
}



// ===== TARIMAS =====

async function addTarimaEntry() {
    const fecha = val('t-fecha');
    if (!fecha) {
        setStatus('status-tarimas', 'La fecha es obligatoria.', true);
        return;
    }
    const nuevo = {
        id: uid(),
        fecha,
        persona: val('t-persona'),
        tipo: val('t-tipo'),
        cant: numval('t-cant')
    };
    DB.tarimas.push(nuevo);
    await saveArr('entries-tarimas', DB.tarimas);
    pushToSheets('TARIMAS', mapTarimaRow(nuevo));
    ['t-persona', 't-cant'].forEach(id => document.getElementById(id).value = '');
    setStatus('status-tarimas', 'Guardado ✓');
    renderTarimasTable();
}

async function deleteTarimaEntry(id) {
    DB.tarimas = DB.tarimas.filter(e => e.id !== id);
    await saveArr('entries-tarimas', DB.tarimas);
    renderTarimasTable();
}

// ===== CONTENEDORES =====

async function addContenedorEntry() {
    const fecha = val('c-fecha');
    if (!fecha) {
        setStatus('status-contenedores', 'La fecha es obligatoria.', true);
        return;
    }
    const nuevo = {
        id: uid(),
        fecha,
        turno: val('c-turno'),
        personas: numval('c-personas'),
        cant: numval('c-cant')
    };
    DB.contenedores.push(nuevo);
    await saveArr('entries-contenedores', DB.contenedores);
    pushToSheets('CONTENEDORES', mapContenedorRow(nuevo));
    ['c-personas', 'c-cant'].forEach(id => document.getElementById(id).value = '');
    setStatus('status-contenedores', 'Guardado ✓');
    renderContenedoresTable();
}

async function deleteContenedorEntry(id) {
    DB.contenedores = DB.contenedores.filter(e => e.id !== id);
    await saveArr('entries-contenedores', DB.contenedores);
    renderContenedoresTable();
}

/* ================================================================
   RESÚMENES
   ================================================================ */

let lastResumenMaquinaRows = [];
let lastResumenKitRows = [];
let lastResumenManualesRows = [];
let lastResumenTarimasRows = [];
let lastResumenContenedoresRows = [];

// ===== RESUMEN: MÁQUINAS =====

async function renderResumenMaquina(forceRefresh = false) {
    const maquina = val('r-maquina');
    const mes = parseInt(val('r-mes'), 10);
    const anio = parseInt(val('r-anio'), 10);
    if (!maquina) return;
    
    const maquinaLabel = machineById(maquina)?.label || maquina;
    const mesNombre = MESES[mes - 1];
    const tituloEl = document.getElementById('titulo-resumen-maquina');
    if (tituloEl) {
        tituloEl.textContent = `${maquinaLabel} - ${mesNombre.toUpperCase()} ${anio}`;
        tituloEl.style.opacity = '1';
    }

    const tbody = document.querySelector('#tbl-resumen-maquina tbody');
    if (tbody) {
        tbody.innerHTML = `<tr><td colspan="10" class="empty"><div class="spinner-inline"></div> Consultando datos…</td></tr>`;
    }
    const kpiEl = document.getElementById('kpi-maquina');
    if (kpiEl) {
        kpiEl.innerHTML = `<div class="kpi" style="grid-column:1/-1;text-align:center;color:var(--sub);"><div class="num">⏳</div><div class="lbl">Consultando Google Sheets…</div></div>`;
    }

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-maquinas', resultado.fuente, resultado.cacheado);

    const datosMaquinasDelMes = (datosSheets.MAQUINAS || [])
        .map(convertirFilaMaquina)
        .filter(Boolean);
    const datosParoDelMes = (datosSheets.PARO || [])
        .map(convertirFilaParo)
        .filter(Boolean);

    const diasEnMes = getDaysInMonth(anio, mes);
    const cfg = machineById(maquina);

    let totalProd = 0, totalParo = 0, sumPct = 0, nPct = 0;
    const rowsHtml = [];
    lastResumenMaquinaRows = [];

    for (let d = 1; d <= diasEnMes; d++) {
        const fecha = `${anio}-${pad2(mes)}-${pad2(d)}`;
        const dow = getDayOfWeek(fecha);
        const entriesDia = datosMaquinasDelMes.filter(e => e.fecha === fecha && e.maquina === maquina);
        const produccion = sum(entriesDia, 'cant');
        const partidas = new Set(entriesDia.map(e => (e.np || '').trim().toUpperCase()).filter(x => x)).size ||
            (entriesDia.length ? 1 : 0);

        const overrideKey = maquina + '_' + fecha;
        let tiempoDisp = DB.tiempoOverrides[overrideKey];

        // Si no hay override manual, buscar si algún registro de producción capturó tiempo disponible
        if (tiempoDisp === undefined) {
            const regConTiempoDisp = entriesDia.find(e => e.tiempoDisp > 0);
            if (regConTiempoDisp) {
                tiempoDisp = regConTiempoDisp.tiempoDisp;
            }
        }

        // Si aún no hay nada, usar el default por día de la semana
        if (tiempoDisp === undefined) {
            const def = defaultTiempoDisponible(maquina, fecha);
            tiempoDisp = def === null ? '' : def;
        }
        const tiempoNum = Number(tiempoDisp) || 0;

        const real = tiempoNum > 0 ? produccion / tiempoNum : 0;
        const esperado = tiempoNum > 0 ?
            (cfg.piezasMin * tiempoNum) / (tiempoNum + (partidas * cfg.ajuste)) :
            0;
        const pct = esperado > 0 ? real / esperado : 0;

        const paroDia = datosParoDelMes.filter(e => e.fecha === fecha && e.maquina === maquina);
        const paroMin = sum(paroDia, 'minutos');
        const motivos = paroDia.map(e => e.motivo).filter(x => x).join('; ');
        const operadoresDia = [...new Set(entriesDia.map(e => e.operador).filter(x => x))].join(', ');
        const fechaTitle = operadoresDia ? operadoresDia : 'Sin operador capturado';

        if (tiempoNum > 0) {
            totalProd += produccion;
            totalParo += paroMin;
            sumPct += pct;
            nPct++;
        }

        const pctPill = tiempoNum === 0 ? '' :
            `<span class="pill ${pct >= 0.9 ? 'good' : pct >= 0.7 ? 'warn' : 'bad'}">${(pct * 100).toFixed(0)}%</span>`;

        rowsHtml.push(`<tr>
            <td>${DIAS_ABR[dow]}</td>
            <td class="op-tooltip" data-operadores="${fechaTitle}">${formatDateDisplay(fecha)}</td>
            <td><input type="number" style="width:70px" value="${tiempoDisp}" onchange="window.updateTiempoOverride('${maquina}','${fecha}',this.value)"></td>
            <td>${produccion}</td>
            <td>${partidas}</td>
            <td>${real ? real.toFixed(2) : '-'}</td>
            <td>${esperado ? esperado.toFixed(2) : '-'}</td>
            <td>${pctPill}</td>
            <td>${paroMin || ''}</td>
            <td class="left">${motivos}</td>
        </tr>`);

        lastResumenMaquinaRows.push({
            DIA: DIAS_ABR[dow],
            FECHA: formatDateDisplay(fecha),
            'TIEMPO DISPONIBLE': tiempoNum,
            PRODUCCION: produccion,
            PARTIDAS: partidas,
            REAL: Number(real.toFixed(2)),
            ESPERADO: Number(esperado.toFixed(2)),
            '%': Math.round(pct * 100),
            'PARO (MIN)': paroMin,
            MOTIVOS: motivos
        });
    }

    const t = document.getElementById('tbl-resumen-maquina');
    t.querySelector('thead').innerHTML =
        `<tr><th>Día</th><th>Fecha</th><th>Tiempo Disp. (min)</th><th>Producción</th><th>Partidas</th><th>Real (u/min)</th><th>Esperado (u/min)</th><th>%</th><th>Paro (min)</th><th>Motivo(s)</th></tr>`;
    t.querySelector('tbody').innerHTML = rowsHtml.join('');

    const pctProm = nPct ? (sumPct / nPct * 100).toFixed(0) : '-';
    document.getElementById('kpi-maquina').innerHTML = `
        <div class="kpi"><div class="num">${totalProd.toLocaleString()}</div><div class="lbl">Producción total del mes</div></div>
        <div class="kpi"><div class="num">${pctProm}${nPct ? '%' : ''}</div><div class="lbl">Cumplimiento promedio</div></div>
        <div class="kpi"><div class="num">${totalParo}</div><div class="lbl">Minutos de paro en el mes</div></div>
    `;

    renderTop5Maquina(maquina, datosMaquinasDelMes);
}


async function updateTiempoOverride(maquina, fecha, value) {
    const key = maquina + '_' + fecha;
    const numVal = value === '' ? '' : Number(value);
    if (value === '') {
        delete DB.tiempoOverrides[key];
    } else {
        DB.tiempoOverrides[key] = numVal;
    }
    await saveObj('config-tiempo-overrides', DB.tiempoOverrides);

    // Sincronizar con Google Sheets
    const rowOverride = {
        FECHA: formatDateDisplay(fecha),
        MAQUINA: machineById(maquina)?.label || maquina,
        OVERRIDE: numVal
    };
    pushToSheets('TIEMPO_OVERRIDES', rowOverride);

    renderResumenMaquina();
}

function exportResumenMaquina() {
    if (!lastResumenMaquinaRows.length) { alert('Genera el resumen primero.'); return; }
    const maquina = machineById(val('r-maquina'))?.label || 'maquina';
    downloadExcel(lastResumenMaquinaRows,
        `resumen_${maquina.replace(/\s+/g, '_')}_${val('r-mes')}_${val('r-anio')}.xlsx`,
        'RESUMEN');
}

// ===== RESUMEN: KIT (derivado de Manuales) =====

async function renderResumenKit(forceRefresh = false) {
    const mes = parseInt(val('rk-mes'), 10);
    const anio = parseInt(val('rk-anio'), 10);

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-kit', resultado.fuente, resultado.cacheado);

    // SOLO datos de Sheets
    const rows = (datosSheets.MANUALES || [])
        .map(convertirFilaManual)
        .filter(e => e && (e.producto || '').trim().toUpperCase() === 'KIT')
        .sort((a, b) => a.fecha.localeCompare(b.fecha));

    const total = sum(rows, 'cant');
    lastResumenKitRows = rows.map(e => ({
        FECHA: formatDateDisplay(e.fecha),
        PERSONAL: e.operador,
        'NP KIT': e.np,
        CANTIDAD: e.cant
    }));

    const t = document.getElementById('tbl-resumen-kit');
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Personal</th><th>NP KIT</th><th>Cantidad</th></tr>`;
    t.querySelector('tbody').innerHTML = (rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${e.operador || ''}</td>
                <td class="left">${e.np || ''}</td>
                <td>${e.cant || 0}</td>
            </tr>
        `).join('') :
        `<tr><td colspan="4" class="empty">Sin registros este mes</td></tr>`) +
        `<tr class="totals"><td colspan="3">Total del mes</td><td>${total}</td></tr>`;
}

function exportResumenKit() {
    if (!lastResumenKitRows.length) { alert('Genera el resumen primero.'); return; }
    downloadExcel(lastResumenKitRows, `resumen_kit_${val('rk-mes')}_${val('rk-anio')}.xlsx`, 'ENSAMBLE KITS');
}

// ===== RESUMEN: MANUALES (general) =====

async function renderResumenManuales(forceRefresh = false) {
    const mes = parseInt(val('rm-mes'), 10);
    const anio = parseInt(val('rm-anio'), 10);
    const filtroProd = val('rm-producto').toUpperCase();

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-manuales', resultado.fuente, resultado.cacheado);

    // SOLO datos de Sheets
    const rows = (datosSheets.MANUALES || [])
        .map(convertirFilaManual)
        .filter(e => {
            if (!e) return false;
            if (filtroProd && (e.producto || '').toUpperCase() !== filtroProd) return false;
            return true;
        })
        .sort((a, b) => a.fecha.localeCompare(b.fecha));

    lastResumenManualesRows = rows.map(e => ({
        FECHA: formatDateDisplay(e.fecha),
        PRODUCTO: e.producto,
        NP: e.np,
        'H.V.': e.hv || '',
        OPERADOR: e.operador,
        RRHH: e.rrhh,
        'CANT PROD': e.cant,
        OPERACIONES: opsSummary(e.ops),
        'SUB PROC': e.subproc
    }));

    const totalCant = sum(rows, 'cant');
    const t = document.getElementById('tbl-resumen-manuales');
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Producto</th><th>NP</th><th>H.V.</th><th>Operador</th><th>RRHH</th><th>Cant. Prod.</th><th>Operaciones</th><th>Sub Proc</th></tr>`;
    t.querySelector('tbody').innerHTML = (rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${e.producto || ''}</td>
                <td>${e.np || ''}</td>
                <td>${e.hv || ''}</td>
                <td class="left">${e.operador || ''}</td>
                <td>${e.rrhh || 0}</td>
                <td>${e.cant || 0}</td>
                <td class="left">${opsSummary(e.ops)}</td>
                <td>${e.subproc || 0}</td>
            </tr>
        `).join('') :
        `<tr><td colspan="9" class="empty">Sin registros este mes</td></tr>`) +
        `<tr class="totals"><td colspan="6">Total</td><td>${totalCant}</td><td colspan="2"></td></tr>`;
}

function exportResumenManuales() {
    if (!lastResumenManualesRows.length) { alert('Genera el resumen primero.'); return; }
    downloadExcel(lastResumenManualesRows,
        `resumen_manuales_${val('rm-mes')}_${val('rm-anio')}.xlsx`,
        'MANUALES');
}

// ===== RESUMEN: TARIMAS =====

async function renderResumenTarimas(forceRefresh = false) {
    const mes = parseInt(val('rt-mes'), 10);
    const anio = parseInt(val('rt-anio'), 10);

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-tarimas', resultado.fuente, resultado.cacheado);

    // SOLO datos de Sheets
    const rows = (datosSheets.TARIMAS || [])
        .map(convertirFilaTarima)
        .filter(Boolean);

    const porPersona = {};
    rows.forEach(e => {
        const key = (e.persona || 'Sin nombre') + ' | ' + (e.tipo || '');
        porPersona[key] = (porPersona[key] || 0) + (Number(e.cant) || 0);
    });

    const entries = Object.entries(porPersona);
    lastResumenTarimasRows = entries.map(([k, v]) => {
        const [persona, tipo] = k.split(' | ');
        return { PERSONA: persona, TIPO: tipo, TOTAL: v };
    });

    const t = document.getElementById('tbl-resumen-tarimas');
    t.querySelector('thead').innerHTML =
        `<tr><th>Persona</th><th>Tipo</th><th>Total del mes</th></tr>`;
    t.querySelector('tbody').innerHTML = entries.length ?
        entries.map(([k, v]) => {
            const [persona, tipo] = k.split(' | ');
            return `<tr><td class="left">${persona}</td><td class="left">${tipo}</td><td>${v}</td></tr>`;
        }).join('') :
        `<tr><td colspan="3" class="empty">Sin registros este mes</td></tr>`;
}
function exportResumenTarimas() {
    if (!lastResumenTarimasRows.length) { alert('Genera el resumen primero.'); return; }
    downloadExcel(lastResumenTarimasRows,
        `resumen_tarimas_${val('rt-mes')}_${val('rt-anio')}.xlsx`,
        'TARIMAS');
}

// ===== RESUMEN: CONTENEDORES =====

async function renderResumenContenedores(forceRefresh = false) {
    const mes = parseInt(val('rc-mes'), 10);
    const anio = parseInt(val('rc-anio'), 10);

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-contenedores', resultado.fuente, resultado.cacheado);

    // SOLO datos de Sheets
    const rows = (datosSheets.CONTENEDORES || [])
        .map(convertirFilaContenedor)
        .filter(Boolean);

    const porTurno = {};
    rows.forEach(e => {
        const key = e.turno || 'Sin turno';
        if (!porTurno[key]) porTurno[key] = { personas: 0, cant: 0 };
        porTurno[key].personas += Number(e.personas) || 0;
        porTurno[key].cant += Number(e.cant) || 0;
    });

    const entries = Object.entries(porTurno);
    lastResumenContenedoresRows = entries.map(([k, v]) => ({
        TURNO: k,
        PERSONAS: v.personas,
        CONTENEDORES: v.cant,
        'CONTENEDORES/PERSONA-DIA': v.personas ? Number((v.cant / v.personas).toFixed(2)) : 0
    }));

    const t = document.getElementById('tbl-resumen-contenedores');
    t.querySelector('thead').innerHTML =
        `<tr><th>Turno</th><th>Personas (suma)</th><th>Contenedores (total)</th><th>Contenedores / persona-día</th></tr>`;
    t.querySelector('tbody').innerHTML = entries.length ?
        entries.map(([k, v]) => `
            <tr>
                <td class="left">${k}</td>
                <td>${v.personas}</td>
                <td>${v.cant}</td>
                <td>${v.personas ? (v.cant / v.personas).toFixed(2) : '-'}</td>
            </tr>
        `).join('') :
        `<tr><td colspan="4" class="empty">Sin registros este mes</td></tr>`;
}

function exportResumenContenedores() {
    if (!lastResumenContenedoresRows.length) { alert('Genera el resumen primero.'); return; }
    downloadExcel(lastResumenContenedoresRows,
        `resumen_contenedores_${val('rc-mes')}_${val('rc-anio')}.xlsx`,
        'CONTENEDORES');
}

// ===== RESUMEN: TOP 5 PRODUCTOS (reutilizable) =====

// Dibuja barras horizontales a partir de pares [etiqueta, valor] ya calculados
function drawBars(containerId, pares, opts) {
    const cont = document.getElementById(containerId);
    if (!cont) return;
    opts = opts || {};
    const limite = opts.limite || 5;
    const sufijo = opts.sufijo || '';

    const top = pares.sort((a, b) => b[1] - a[1]).slice(0, limite);

    if (!top.length) {
        cont.innerHTML = `<div class="empty">Sin datos capturados este mes</div>`;
        return;
    }

    const max = top[0][1];
    cont.innerHTML = top.map(([etiqueta, valor]) => {
        const pct = max ? Math.round((valor / max) * 100) : 0;
        return `
            <div class="top5-row">
                <div class="top5-label" title="${etiqueta}">${etiqueta}</div>
                <div class="top5-bar-bg"><div class="top5-bar-fill" style="width:${pct}%"></div></div>
                <div class="top5-value">${valor.toLocaleString()}${sufijo}</div>
            </div>
        `;
    }).join('');
}

// Mantiene el nombre anterior por compatibilidad: agrupa por producto sumando cantidad
function drawTop5Bars(containerId, rows) {
    const porProducto = {};
    rows.forEach(e => {
        const nombre = (e.producto || 'SIN PRODUCTO').trim().toUpperCase();
        porProducto[nombre] = (porProducto[nombre] || 0) + (Number(e.cant) || 0);
    });
    drawBars(containerId, Object.entries(porProducto), { limite: 5 });
}

async function renderTop5(forceRefresh = false) {
    const mes = parseInt(val('rtop-mes'), 10);
    const anio = parseInt(val('rtop-anio'), 10);
    const fuente = val('rtop-fuente') || 'ambos';

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-top5', resultado.fuente, resultado.cacheado);
    if (!datosSheets) {
        document.getElementById('chart-top5').innerHTML = '<div class="empty">No se pudo conectar con Google Sheets ni hay datos locales.</div>';
        return;
    }

    let registros = [];
    if (fuente === 'ambos' || fuente === 'maquinas') {
        registros = registros.concat((datosSheets.MAQUINAS || [])
            .map(convertirFilaMaquina).filter(Boolean));
    }
    if (fuente === 'ambos' || fuente === 'manuales') {
        registros = registros.concat((datosSheets.MANUALES || [])
            .map(convertirFilaManual).filter(Boolean));
    }

    drawTop5Bars('chart-top5', registros);
}

// Top 5 de la máquina seleccionada en "Resumen por Máquina" — recibe los
// datos del mes que renderResumenMaquina ya trajo y mezcló (no vuelve a
// consultar Sheets por separado)
function renderTop5Maquina(maquina, datosMaquinasDelMes) {
    const delMaquina = datosMaquinasDelMes.filter(e => e.maquina === maquina);
    drawTop5Bars('chart-top5-maquina', delMaquina);
}

// ===== RESUMEN: FRECUENCIA DE NÚMEROS DE PARTE =====

async function renderFrecuenciaNP(forceRefresh = false) {
    const mes = parseInt(val('rnp-mes'), 10);
    const anio = parseInt(val('rnp-anio'), 10);
    const maquina = val('rnp-maquina');
    const fuente = val('rnp-fuente') || 'ambos';

    const resultado = await obtenerDatosResumen(mes, anio, forceRefresh);
    const datosSheets = resultado.datos;
    mostrarFuenteDatos('res-frecnp', resultado.fuente, resultado.cacheado);
    if (!datosSheets) {
        document.getElementById('chart-frecnp').innerHTML = '<div class="empty">No se pudo conectar con Google Sheets ni hay datos locales.</div>';
        return;
    }

    let registros = [];
    if (fuente === 'ambos' || fuente === 'maquinas') {
        registros = registros.concat((datosSheets.MAQUINAS || [])
            .map(convertirFilaMaquina).filter(Boolean));
    }
    if (fuente === 'ambos' || fuente === 'manuales') {
        registros = registros.concat((datosSheets.MANUALES || [])
            .map(convertirFilaManual).filter(Boolean));
    }

    if (maquina) registros = registros.filter(e => e.maquina === maquina);

    const porNP = {};
    registros.forEach(e => {
        const np = (e.np || '').trim();
        if (!np) return;
        const label = `${np}${e.producto ? ' (' + e.producto.trim().toUpperCase() + ')' : ''}`;
        porNP[label] = (porNP[label] || 0) + 1;
    });

    drawBars('chart-frecnp', Object.entries(porNP), { limite: 8, sufijo: ' veces' });
}


// ================================================================
// LISTENER PARA CAMBIO DE MÁQUINA 
// ================================================================

function onMaquinaChange() {
    const selector = document.getElementById('f-maquina');
    if (!selector) return;
    const maquinaActual = selector.value;
    if (!maquinaActual) return;

    // Clear completo...
    ['f-producto', 'f-np', 'f-cant', 'f-tiempo', 'f-comentarios', 'f-hv']
        .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
    
    // ... (todo lo demás que ya tienes: tiempo disp, operadores, etc.) ...

    // >>> HERENCIA DE HORARIO <<<
    const hoy = getToday();
    aplicarHoraHerencia('f-hora-inicio', 'f-hora-fin', maquinaActual, hoy);
    // -----------------------------

    const maquina = machineById(maquinaActual);
    if (maquina) {
        setStatus('status-maquinas', `🔄 Campos limpiados para ${maquina.label}`, false);
    }
}
window.onMaquinaChange = onMaquinaChange;

function updateTiempoDispField() {
    const maquina = val('f-maquina');
    const fecha = val('f-fecha');
    const el = document.getElementById('f-tiempo-disp');
    const msgEl = document.getElementById('f-tiempo-disp-msg');
    if (!el || !maquina || !fecha) return;

    // Verificar si ya existe un tiempo disponible capturado para esta máquina+fecha
    const existente = DB.maquinas.find(e => e.maquina === maquina && e.fecha === fecha && e.tiempoDisp > 0);

    if (existente) {
        // Ya está bloqueado: mostrar el valor capturado, no editable
        el.value = existente.tiempoDisp;
        el.readOnly = true;
        el.classList.add('bloqueado');
        if (msgEl) {
            msgEl.style.display = 'flex';
            msgEl.className = 'tiempo-disp-msg bloqueado';
            msgEl.innerHTML = `🔒 Tiempo disponible ya registrado para ${machineById(maquina)?.label || maquina} <strong>${existente.tiempoDisp} min</strong>`;
            //msgEl.innerHTML = `🔒 Tiempo disponible ya registrado para ${machineById(maquina)?.label || maquina} — <strong>${existente.tiempoDisp} min</strong>. Para cambiarlo, borra el primer registro de este día.`;
        }
    } else {
        // No está bloqueado: precargar default si está vacío
        el.readOnly = false;
        el.classList.remove('bloqueado');
        if (msgEl) {
            msgEl.style.display = 'none';
            msgEl.innerHTML = '';
        }
        if (el.value === '') {
            const def = defaultTiempoDisponible(maquina, fecha);
            el.value = def === null ? '' : def;
        }
    }
}
window.updateTiempoDispField = updateTiempoDispField;

function onTiempoDispManualChange() {
    const maquina = val('f-maquina');
    const fecha = val('f-fecha');
    const el = document.getElementById('f-tiempo-disp');
    if (!el || el.readOnly || !maquina || !fecha) return;

    // Cuando el usuario edita manualmente el tiempo disponible, guardarlo como override
    // para que el resumen lo use, pero solo hasta que se guarde el primer registro
    const valNum = Number(el.value);
    if (valNum > 0) {
        const overrideKey = maquina + '_' + fecha;
        DB.tiempoOverrides[overrideKey] = valNum;
        saveObj('config-tiempo-overrides', DB.tiempoOverrides);
    }
}
window.onTiempoDispManualChange = onTiempoDispManualChange;

function minutosEntreHoras(horaInicio, horaFin) {
    if (!horaInicio || !horaFin) return 0;
    const [h1, m1] = horaInicio.split(':').map(Number);
    const [h2, m2] = horaFin.split(':').map(Number);
    let min = (h2 * 60 + m2) - (h1 * 60 + m1);
    if (min < 0) min += 24 * 60; // Cruza medianoche
    return min;
}

function calcularTiempoMaquina() {
    const inicio = val('f-hora-inicio');
    const fin = val('f-hora-fin');
    const min = minutosEntreHoras(inicio, fin);
    const resEl = document.getElementById('f-tiempo-res');
    const hiddenEl = document.getElementById('f-tiempo');
    if (resEl) resEl.textContent = `= ${min} min`;
    if (hiddenEl) hiddenEl.value = min;
}
window.calcularTiempoMaquina = calcularTiempoMaquina;

function calcularTiempoManual() {
    const inicio = val('m-hora-inicio');
    const fin = val('m-hora-fin');
    const min = minutosEntreHoras(inicio, fin);
    const resEl = document.getElementById('m-tiempo-res');
    const hiddenEl = document.getElementById('m-tiempo');
    if (resEl) resEl.textContent = `= ${min} min`;
    if (hiddenEl) hiddenEl.value = min;
}
window.calcularTiempoManual = calcularTiempoManual;


/* ================================================================
   INICIALIZACIÓN (BOOT)
   ================================================================ */

/* ================================================================
   SINCRONIZACIÓN CON GOOGLE SHEETS (vía Google Apps Script Web App)
   Es de una sola vía: el panel manda datos, no los lee de vuelta.
   Se usa fetch en modo 'no-cors' porque Apps Script no responde con
   encabezados CORS, así que no hay confirmación visible de éxito.
   ================================================================ */

function mapMaquinaRow(e) {
    return {
        FECHA: formatDateDisplay(e.fecha),
        MAQUINA: machineById(e.maquina)?.label || e.maquina,
        PRODUCTO: e.producto,
        NP: e.np,
        'H.V.': e.hv || '',
        OPERADOR: e.operador,
        RRHH: e.rrhh,
        'CANT PROD': e.cant,
        TIEMPO: e.tiempo,
        'TIEMPO DISP': e.tiempoDisp || '',
        COMENTARIOS: e.comentarios
    };
}
function mapManualRow(e) {
    const row = {
        FECHA: formatDateDisplay(e.fecha), PRODUCTO: e.producto, NP: e.np, 'H.V.': e.hv || '',
        OPERADOR: e.operador,
        RRHH: e.rrhh, 'CANT PROD': e.cant, TIEMPO: e.tiempo
    };
    OPERACIONES.forEach(op => row[op] = (e.ops && e.ops[op]) ? 1 : '');
    row['SUB PROC'] = e.subproc || 0;
    row['COMENTARIOS'] = e.comentarios;
    return row;
}
function mapParoRow(e) {
    return { FECHA: formatDateDisplay(e.fecha), MAQUINA: machineById(e.maquina)?.label || e.maquina, MINUTOS: e.minutos, MOTIVO: e.motivo };
}
function mapTarimaRow(e) {
    return { FECHA: formatDateDisplay(e.fecha), PERSONA: e.persona, TIPO: e.tipo, CANTIDAD: e.cant };
}
function mapContenedorRow(e) {
    return { FECHA: formatDateDisplay(e.fecha), TURNO: e.turno, PERSONAS: e.personas, CONTENEDORES: e.cant };
}

function pushToSheets(sheetName, row) {
    if (!SHEETS_CONFIG.enabled || !SHEETS_CONFIG.url) return Promise.resolve();
    return fetch(SHEETS_CONFIG.url, {
        method: 'POST',
        mode: 'no-cors', // Apps Script no manda headers CORS; así evitamos el bloqueo del navegador
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // evita el preflight OPTIONS que Apps Script no soporta
        body: JSON.stringify({ sheet: sheetName, row })
    }).catch(err => console.error('Error enviando a Google Sheets:', err));
}

async function guardarConfigSheets() {
    const url = val('cfg-webhook-url');
    const enabled = document.getElementById('cfg-webhook-enabled').checked;
    SHEETS_CONFIG = { url, enabled };
    await saveObj('config-sheets-sync', SHEETS_CONFIG);
    setStatus('status-config', 'Configuración guardada ✓');
}

function probarConexionSheets() {
    const url = val('cfg-webhook-url');
    if (!url) { setStatus('status-config', 'Pega primero la URL del webhook.', true); return; }
    fetch(url, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
            sheet: 'PRUEBA',
            row: { FECHA: formatDateDisplay(toISODate(new Date())), MENSAJE: 'Conexión de prueba desde el Panel de Producción' }
        })
    }).then(() => {
        setStatus('status-config', 'Prueba enviada — revisa la pestaña "PRUEBA" en tu Google Sheet en unos segundos.');
    }).catch(() => {
        setStatus('status-config', 'Error de red al enviar la prueba.', true);
    });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function reenviarTodoASheets() {
    if (!val('cfg-webhook-url')) { setStatus('status-reenvio', 'Configura y guarda primero la URL.', true); return; }
    // Fuerza el envío aunque el toggle esté apagado, ya que el usuario lo pidió explícitamente aquí
    const urlActiva = val('cfg-webhook-url');
    const configTemporal = { url: urlActiva, enabled: true };
    const configOriginal = SHEETS_CONFIG;
    SHEETS_CONFIG = configTemporal;

    const tareas = [
        ...DB.maquinas.map(e => ['MAQUINAS', mapMaquinaRow(e)]),
        ...DB.manuales.map(e => ['MANUALES', mapManualRow(e)]),
        ...DB.paro.map(e => ['PARO', mapParoRow(e)]),
        ...DB.tarimas.map(e => ['TARIMAS', mapTarimaRow(e)]),
        ...DB.contenedores.map(e => ['CONTENEDORES', mapContenedorRow(e)]),
        ...Object.entries(DB.tiempoOverrides).map(([key, val]) => {
            const [maquinaId, fecha] = key.split('_');
            return ['TIEMPO_OVERRIDES', {
                FECHA: formatDateDisplay(fecha),
                MAQUINA: machineById(maquinaId)?.label || maquinaId,
                OVERRIDE: val
            }];
        })
    ];

    if (!tareas.length) {
        setStatus('status-reenvio', 'No hay registros capturados todavía.', true);
        SHEETS_CONFIG = configOriginal;
        return;
    }

    setStatus('status-reenvio', `Enviando ${tareas.length} registros, no cierres esta pestaña...`);
    for (let i = 0; i < tareas.length; i++) {
        await pushToSheets(tareas[i][0], tareas[i][1]); // espera a que termine antes de mandar el siguiente
        if ((i + 1) % 10 === 0 || i === tareas.length - 1) {
            setStatus('status-reenvio', `Enviando... ${i + 1} de ${tareas.length}`);
        }
        await sleep(150); // pequeño respiro extra para no saturar el webhook
    }
    SHEETS_CONFIG = configOriginal;
    setStatus('status-reenvio', `Listo — se reenviaron ${tareas.length} registros.`);
}

window.guardarConfigSheets = guardarConfigSheets;
window.probarConexionSheets = probarConexionSheets;
window.reenviarTodoASheets = reenviarTodoASheets;

/* ================================================================
   IMPORTAR (LEER) DESDE GOOGLE SHEETS
   A diferencia del push (POST en modo no-cors, "a ciegas"), aquí SÍ
   podemos leer la respuesta: los doGet de Apps Script sí permiten
   fetch normal desde el navegador.
   Solo se agregan filas que no existan ya localmente (comparando
   por los mismos campos que se guardan en el Sheet), para no duplicar.
   ================================================================ */

// Convierte un objeto a una "llave" comparable: texto en mayúsculas
// sin espacios sobrantes, números normalizados, y fechas normalizadas
// a su valor real (sin importar si el texto dice "08/06/2026",
// "2026-06-08" u otra variante) — para poder comparar un registro
// local contra uno leído de Sheets sin falsos negativos.
function buildRowKey(row) {
    return Object.keys(row).sort().map(k => {
        const v = row[k];
        if (k === 'FECHA') return parseFlexibleDate(v) || '';
        if (v === null || v === undefined || v === '') return '';
        const n = Number(v);
        return !isNaN(n) && v !== '' ? String(n) : String(v).trim().toUpperCase();
    }).join('|');
}

// Se mantiene el nombre anterior por compatibilidad; usa el mismo parser único
function parseSheetDate(v) {
    return parseFlexibleDate(v);
}

// Convierte una fila leída de Sheets (formato de columnas del Sheet) a la
// forma interna de cada entrada. Devuelven null cuando la fila no se puede
// usar (ej. nombre de máquina no reconocido) para que quien llama la ignore.

function convertirFilaMaquina(row) {
    const maquinaId = normalizeMachine(row.MAQUINA);
    if (!maquinaId) return null;
    const tiempoDispSheet = Number(row['TIEMPO DISP']) || 0;
    const fecha = parseSheetDate(row.FECHA);
    // Si el Sheet trae tiempo disponible, guardarlo como override
    if (tiempoDispSheet > 0 && fecha) {
        const overrideKey = maquinaId + '_' + fecha;
        DB.tiempoOverrides[overrideKey] = tiempoDispSheet;
    }
    return {
        id: uid(), fecha: fecha, maquina: maquinaId,
        producto: row.PRODUCTO || '', np: String(row.NP ?? ''), hv: String(row['H.V.'] ?? ''),
        operador: row.OPERADOR || '',
        rrhh: Number(row.RRHH) || 0, cant: Number(row['CANT PROD']) || 0,
        tiempo: Number(row.TIEMPO) || 0, tiempoDisp: tiempoDispSheet,
        comentarios: row.COMENTARIOS || ''
    };
}
function convertirFilaManual(row) {
    const ops = {};
    OPERACIONES.forEach(op => { ops[op] = (row[op] == 1 || row[op] === '1') ? 1 : 0; });
    const subproc = Object.values(ops).reduce((a, b) => a + b, 0);
    return {
        id: uid(), fecha: parseSheetDate(row.FECHA), producto: row.PRODUCTO || '',
        np: String(row.NP ?? ''), hv: String(row['H.V.'] ?? ''),
        operador: row.OPERADOR || '', rrhh: Number(row.RRHH) || 0,
        cant: Number(row['CANT PROD']) || 0, tiempo: Number(row.TIEMPO) || 0,
        ops, subproc, comentarios: row.COMENTARIOS || ''
    };
}
function convertirFilaParo(row) {
    const maquinaId = normalizeMachine(row.MAQUINA);
    if (!maquinaId) return null;
    return {
        id: uid(), fecha: parseSheetDate(row.FECHA), maquina: maquinaId,
        minutos: Number(row.MINUTOS) || 0, motivo: row.MOTIVO || ''
    };
}
function convertirFilaTarima(row) {
    return {
        id: uid(), fecha: parseSheetDate(row.FECHA), persona: row.PERSONA || '',
        tipo: row.TIPO || '', cant: Number(row.CANTIDAD) || 0
    };
}
function convertirFilaContenedor(row) {
    return {
        id: uid(), fecha: parseSheetDate(row.FECHA), turno: row.TURNO || '',
        personas: Number(row.PERSONAS) || 0, cant: Number(row.CONTENEDORES) || 0
    };
}

// Junta un arreglo local con filas de Sheets, agregando solo lo que no
// exista ya (comparando por buildRowKey). No modifica el arreglo local
// que le pasaron — regresa uno nuevo, salvo que se pida lo contrario.
function mergeSinDuplicar(localArr, sheetRows, mapLocalFn, convertFn) {
    if (!Array.isArray(sheetRows) || !sheetRows.length) return localArr;
    const existentes = new Set(localArr.map(e => buildRowKey(mapLocalFn(e))));
    const extra = [];
    sheetRows.forEach(row => {
        const key = buildRowKey(row);
        if (existentes.has(key)) return;
        const convertido = convertFn(row);
        if (!convertido) return;
        existentes.add(key);
        extra.push(convertido);
    });
    return localArr.concat(extra);
}

// Pide datos a Apps Script en formato JSONP en vez de fetch(): Chrome
// bloquea con CORB la lectura de las respuestas normales de Apps Script
// sin importar dónde esté hospedado el panel (ni con https:// se libra),
// pero SÍ permite cargar <script> entre orígenes distintos — por eso el
// truco funciona. Se resuelve con null si tarda demasiado o falla.
let jsonpContador = 0;
function fetchSheetsJSONP(url, timeoutMs = 10000) {
    return new Promise(resolve => {
        const callbackName = `panelSheetsCallback_${jsonpContador++}_${Date.now()}`;
        const script = document.createElement('script');
        let terminado = false;

        const limpiar = () => {
            delete window[callbackName];
            script.remove();
            clearTimeout(timer);
        };

        window[callbackName] = (datos) => {
            if (terminado) return;
            terminado = true;
            limpiar();
            resolve(datos);
        };

        const timer = setTimeout(() => {
            if (terminado) return;
            terminado = true;
            limpiar();
            console.error('JSONP a Google Sheets: se agotó el tiempo de espera.');
            resolve(null);
        }, timeoutMs);

        script.onerror = () => {
            if (terminado) return;
            terminado = true;
            limpiar();
            console.error('JSONP a Google Sheets: no se pudo cargar (revisa la URL o el despliegue).');
            resolve(null);
        };

        const sep = url.includes('?') ? '&' : '?';
        script.src = `${url}${sep}callback=${callbackName}`;
        document.head.appendChild(script);
    });
}

// Consulta Sheets filtrado a un mes/año específico (el propio Apps Script
// filtra, así no se descarga el historial completo cada vez). Devuelve
// null si la sincronización está apagada, no hay URL, o falla la consulta
// (en ese caso los resúmenes simplemente siguen con los datos locales).
async function fetchSheetsMonth(mes, anio, forceRefresh = false) {
    if (!SHEETS_CONFIG.enabled || !SHEETS_CONFIG.url) return null;
    
    const cacheKey = `${mes}_${anio}`;
    const now = Date.now();
    
    if (!forceRefresh && SHEETS_CLIENT_CACHE[cacheKey] && (now - SHEETS_CLIENT_CACHE[cacheKey].ts < CACHE_TTL_MS)) {
        return SHEETS_CLIENT_CACHE[cacheKey].data;
    }
    
    let lastError = null;
    for (let attempt = 0; attempt < 2; attempt++) {
        try {
            const sep = SHEETS_CONFIG.url.includes('?') ? '&' : '?';
            const data = await fetchSheetsJSONP(`${SHEETS_CONFIG.url}${sep}mes=${mes}&anio=${anio}`, 15000);
            if (data) {
                SHEETS_CLIENT_CACHE[cacheKey] = { data, ts: now };
                return data;
            }
        } catch (e) {
            lastError = e;
        }
        if (attempt < 1) await sleep(1000);
    }
    console.error('fetchSheetsMonth falló después de 2 intentos:', lastError);
    return null;
}

async function obtenerDatosResumen(mes, anio, forceRefresh = false) {
    const sheets = await fetchSheetsMonth(mes, anio, forceRefresh);
    if (sheets) {
        // Aplicar overrides de tiempo desde Sheets sobre los locales
        const overridesSheets = (sheets.TIEMPO_OVERRIDES || []);
        let overridesNuevos = 0;
        overridesSheets.forEach(row => {
            const maquinaId = normalizeMachine(row.MAQUINA);
            const fecha = parseSheetDate(row.FECHA);
            const overrideVal = row.OVERRIDE !== undefined && row.OVERRIDE !== '' ? Number(row.OVERRIDE) : undefined;
            if (maquinaId && fecha && overrideVal !== undefined && !isNaN(overrideVal)) {
                const key = maquinaId + '_' + fecha;
                // Solo actualizar si el valor de Sheets es diferente al local
                if (DB.tiempoOverrides[key] !== overrideVal) {
                    DB.tiempoOverrides[key] = overrideVal;
                    overridesNuevos++;
                }
            }
        });
        if (overridesNuevos > 0) {
            saveObj('config-tiempo-overrides', DB.tiempoOverrides);
        }

        const cacheKey = `${mes}_${anio}`;
        const cacheado = !forceRefresh && SHEETS_CLIENT_CACHE[cacheKey] && (Date.now() - SHEETS_CLIENT_CACHE[cacheKey].ts < CACHE_TTL_MS);
        return { fuente: 'sheets', datos: sheets, cacheado };
    }
    
    const filtrar = (arr) => arr.filter(e => {
        const d = parseDate(e.fecha);
        return d.getFullYear() === anio && (d.getMonth() + 1) === mes;
    });
    
    return {
        fuente: 'local',
        datos: {
            MAQUINAS: filtrar(DB.maquinas).map(e => ({
                FECHA: formatDateDisplay(e.fecha),
                MAQUINA: machineById(e.maquina)?.label || e.maquina,
                PRODUCTO: e.producto || '',
                NP: e.np || '',
                OPERADOR: e.operador || '',
                RRHH: e.rrhh || 0,
                'CANT PROD': e.cant || 0,
                TIEMPO: e.tiempo || 0,
                COMENTARIOS: e.comentarios || ''
            })),
            MANUALES: filtrar(DB.manuales).map(e => {
                const row = {
                    FECHA: formatDateDisplay(e.fecha),
                    PRODUCTO: e.producto || '',
                    NP: e.np || '',
                    OPERADOR: e.operador || '',
                    RRHH: e.rrhh || 0,
                    'CANT PROD': e.cant || 0,
                    TIEMPO: e.tiempo || 0
                };
                OPERACIONES.forEach(op => row[op] = (e.ops && e.ops[op]) ? 1 : '');
                row['SUB PROC'] = e.subproc || 0;
                row['COMENTARIOS'] = e.comentarios || '';
                return row;
            }),
            PARO: filtrar(DB.paro).map(e => ({
                FECHA: formatDateDisplay(e.fecha),
                MAQUINA: machineById(e.maquina)?.label || e.maquina,
                MINUTOS: e.minutos || 0,
                MOTIVO: e.motivo || ''
            })),
            TARIMAS: filtrar(DB.tarimas).map(e => ({
                FECHA: formatDateDisplay(e.fecha),
                PERSONA: e.persona || '',
                TIPO: e.tipo || '',
                CANTIDAD: e.cant || 0
            })),
            CONTENEDORES: filtrar(DB.contenedores).map(e => ({
                FECHA: formatDateDisplay(e.fecha),
                TURNO: e.turno || '',
                PERSONAS: e.personas || 0,
                CONTENEDORES: e.cant || 0
            }))
        },
        cacheado: false
    };
}

function mostrarFuenteDatos(panelId, fuente, cacheado) {
    const panel = document.getElementById(panelId);
    if (!panel) return;
    let badge = panel.querySelector('.fuente-badge');
    if (!badge) {
        badge = document.createElement('div');
        badge.className = 'fuente-badge';
        const firstCard = panel.querySelector('.card');
        if (firstCard) {
            panel.insertBefore(badge, firstCard);
        } else {
            panel.prepend(badge);
        }
    }
    if (fuente === 'sheets') {
        badge.innerHTML = cacheado 
            ? '✅ Datos de Google Sheets <span class="fuente-sub">(en caché — <a href="#" onclick="window.recargarResumen(\'' + panelId + '\');return false;">↻ recargar</a>)</span>' 
            : '✅ Datos de Google Sheets <span class="fuente-sub">(<a href="#" onclick="window.recargarResumen(\'' + panelId + '\');return false;">↻ recargar</a>)</span>';
        badge.className = 'fuente-badge fuente-sheets';
    } else {
        badge.innerHTML = '⚠️ Datos locales <span class="fuente-sub">(Sheets no disponible — <a href="#" onclick="window.recargarResumen(\'' + panelId + '\');return false;">↻ reintentar</a>)</span>';
        badge.className = 'fuente-badge fuente-local';
    }
}

window.recargarResumen = function(panelId) {
    switch(panelId) {
        case 'res-maquinas': renderResumenMaquina(true); break;
        case 'res-kit': renderResumenKit(true); break;
        case 'res-manuales': renderResumenManuales(true); break;
        case 'res-tarimas': renderResumenTarimas(true); break;
        case 'res-contenedores': renderResumenContenedores(true); break;
        case 'res-top5': renderTop5(true); break;
        case 'res-frecnp': renderFrecuenciaNP(true); break;
    }
};

async function importarDesdeSheets() {
    const url = val('cfg-webhook-url');
    if (!url) { setStatus('status-importar-sheets', 'Configura y guarda primero la URL.', true); return; }

    // Regla del panel: el almacenamiento local NUNCA guarda meses que no
    // sean el actual. Los meses anteriores viven solo en Sheets y se
    // consultan al vuelo desde cada Resumen (sin guardarse aquí). Por eso
    // este botón, aunque diga "Importar", solo trae el mes en curso —
    // pensado como respaldo/recuperación si algo se perdió localmente.
    const hoy = new Date();
    const mesActual = hoy.getMonth() + 1;
    const anioActual = hoy.getFullYear();

    setStatus('status-importar-sheets', `Consultando el mes actual en Google Sheets...`);
    const sep = url.includes('?') ? '&' : '?';
    const datos = await fetchSheetsJSONP(`${url}${sep}mes=${mesActual}&anio=${anioActual}`);
    if (!datos) {
        setStatus('status-importar-sheets', 'No se pudo leer el Sheet. Revisa que el Apps Script esté implementado (Nueva versión) con el código más reciente.', true);
        return;
    }

    const conteos = [
        ['maquinas', 'entries-maquinas', 'MAQUINAS', mapMaquinaRow, convertirFilaMaquina],
        ['manuales', 'entries-manuales', 'MANUALES', mapManualRow, convertirFilaManual],
        ['paro', 'entries-paro', 'PARO', mapParoRow, convertirFilaParo],
        ['tarimas', 'entries-tarimas', 'TARIMAS', mapTarimaRow, convertirFilaTarima],
        ['contenedores', 'entries-contenedores', 'CONTENEDORES', mapContenedorRow, convertirFilaContenedor]
    ];

    // Importar overrides de tiempo desde Sheets
    const overridesSheets = datos['TIEMPO_OVERRIDES'] || [];
    let overridesImportados = 0;
    overridesSheets.forEach(row => {
        const maquinaId = normalizeMachine(row.MAQUINA);
        const fecha = parseSheetDate(row.FECHA);
        const overrideVal = row.OVERRIDE !== undefined && row.OVERRIDE !== '' ? Number(row.OVERRIDE) : undefined;
        if (maquinaId && fecha && overrideVal !== undefined && !isNaN(overrideVal)) {
            const key = maquinaId + '_' + fecha;
            if (DB.tiempoOverrides[key] !== overrideVal) {
                DB.tiempoOverrides[key] = overrideVal;
                overridesImportados++;
            }
        }
    });
    if (overridesImportados > 0) {
        await saveObj('config-tiempo-overrides', DB.tiempoOverrides);
    }

    let totalNuevos = 0;
    for (const [dbKey, storageKey, sheetKey, mapFn, convertFn] of conteos) {
        const antes = DB[dbKey].length;
        DB[dbKey] = mergeSinDuplicar(DB[dbKey], datos[sheetKey], mapFn, convertFn);
        totalNuevos += DB[dbKey].length - antes;
        await saveArr(storageKey, DB[dbKey]);
    }

    renderMaquinasTable();
    renderManualesTable();
    renderParoTable();
    renderTarimasTable();
    renderContenedoresTable();

    setStatus('status-importar-sheets', `Listo — ${totalNuevos} registros nuevos importados del mes actual.`);
}
window.importarDesdeSheets = importarDesdeSheets;

async function boot() {

    // Configurar UI
    fillMachineSelect('f-maquina');
    fillMachineSelect('p-maquina');
    fillMachineSelect('r-maquina');
    fillMachineSelect('rnp-maquina');
    fillMonthSelect('r-mes');
    fillMonthSelect('rk-mes');
    fillMonthSelect('rm-mes');
    fillMonthSelect('rt-mes');
    fillMonthSelect('rc-mes');
    fillMonthSelect('rtop-mes');
    fillMonthSelect('rnp-mes');
    fillMonthSelect('rp-mes');

    buildChecklist();
    initOperadorList('f-operadores-list', 'f-rrhh');
    initOperadorList('m-operadores-list', 'm-rrhh');
    setDateInputs();
    setMonthSelects();
    setupTabs();

    // Cargar datos
    await loadAll();

    // Precargar herencia de horarios del día actual
    const hoy = getToday();
    const maqInicial = document.getElementById('f-maquina').value;
    if (maqInicial) aplicarHoraHerencia('f-hora-inicio', 'f-hora-fin', maqInicial, hoy);
    aplicarHoraHerencia('m-hora-inicio', 'm-hora-fin', 'manual', hoy);

    // Precargar configuración de Google Sheets en el formulario
    //aqui se precargan los datos 
    document.getElementById('cfg-webhook-url').value = SHEETS_CONFIG.url || 'https://script.google.com/macros/s/AKfycbxkND7SULKytuYTQog2DDx6XSQ8XtUxaJhii0PLqrTQ5uZSOc5ZsRWClfuQImcCnJCg/exec'; //esta url es la que se puede remplazar recuerda dejar ('';)
    document.getElementById('cfg-webhook-enabled').checked = !!SHEETS_CONFIG.enabled;

    // Renderizar tablas
    renderMaquinasTable();
    renderManualesTable();
    renderParoTable();
    renderTarimasTable();
    renderContenedoresTable();
}

/* ================================================================
   EXPONER FUNCIONES AL ÁMBITO GLOBAL (para onclick en HTML)
   ================================================================ */

// Captura - Máquinas
window.addMaquinaEntry = addMaquinaEntry;
window.deleteMaquinaEntry = deleteMaquinaEntry;
window.handleImportMaquinas = handleImportMaquinas;
window.exportMaquinas = exportMaquinas;

// Captura - Manuales
window.addManualEntry = addManualEntry;
window.deleteManualEntry = deleteManualEntry;
window.handleImportManuales = handleImportManuales;
window.exportManuales = exportManuales;

// Captura - Paro
window.addParoEntry = addParoEntry;
window.deleteParoEntry = deleteParoEntry;

// Captura - Tarimas
window.addTarimaEntry = addTarimaEntry;
window.deleteTarimaEntry = deleteTarimaEntry;

// Captura - Contenedores
window.addContenedorEntry = addContenedorEntry;
window.deleteContenedorEntry = deleteContenedorEntry;

// Resúmenes
window.renderResumenMaquina = renderResumenMaquina;
window.updateTiempoOverride = updateTiempoOverride;
window.exportResumenMaquina = exportResumenMaquina;

window.renderResumenKit = renderResumenKit;
window.exportResumenKit = exportResumenKit;

window.renderResumenManuales = renderResumenManuales;
window.exportResumenManuales = exportResumenManuales;

window.renderResumenTarimas = renderResumenTarimas;
window.exportResumenTarimas = exportResumenTarimas;

window.renderResumenContenedores = renderResumenContenedores;
window.exportResumenContenedores = exportResumenContenedores;

window.renderTop5 = renderTop5;
window.renderFrecuenciaNP = renderFrecuenciaNP;

window.renderFrecuenciaParos = renderFrecuenciaParos;
window.exportFrecuenciaParos = exportFrecuenciaParos;

// UI
window.updateSubProc = updateSubProc;

/* ================================================================
   INICIAR APLICACIÓN
   ================================================================ */

boot();

/* ================================================================
   FIN DEL ARCHIVO
   ================================================================ */