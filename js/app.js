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
    { id: 'PEGADORA', label: 'Semiautomática (Pegadora)', piezasMin: 80, ajuste: 30, aliases: ['PEGADORA', 'SEMIAUTOMATICA', 'SEMIAUTOMÁTICA'] },
    { id: 'T_PLANA', label: 'Troqueladora Plana', piezasMin: 35, ajuste: 50, aliases: ['T. PLANA', 'TROQ. PLANA', 'TROQUELADORA PLANA'] },
    { id: 'ARMADORA_REJILLAS', label: 'Máquina Armadora (Armadora de Rejillas)', piezasMin: 26, ajuste: 30, aliases: ['ARMADORA DE REJILLAS', 'MAQUINA ARMADORA', 'MÁQUINA ARMADORA'] },
];

const OPERACIONES = ['DESPIQUE', 'ARMADO', 'CORTE', 'GRAPADO', 'PEGADO', 'ENSAMBLE', 'DOBLADO', 'CONTEO', 'SELLADO', 'AMARRADO', 'TARIMA'];

const DIAS_ABR = ['D', 'L', 'M', 'MI', 'J', 'V', 'S'];
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

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

async function loadArr(key) {
    try {
        const r = await window.storage.get(key, false);
        return r && r.value ? JSON.parse(r.value) : [];
    } catch (e) {
        return [];
    }
}

async function saveArr(key, arr) {
    try {
        await window.storage.set(key, JSON.stringify(arr), false);
        return true;
    } catch (e) {
        console.error('storage error', e);
        return false;
    }
}

async function loadObj(key) {
    try {
        const r = await window.storage.get(key, false);
        return r && r.value ? JSON.parse(r.value) : {};
    } catch (e) {
        return {};
    }
}

async function saveObj(key, obj) {
    try {
        await window.storage.set(key, JSON.stringify(obj), false);
        return true;
    } catch (e) {
        console.error('storage error', e);
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
        `<tr><th>Fecha</th><th>Máquina</th><th>Producto</th><th>NP</th><th>Operador</th><th>RRHH</th><th>Cant. Prod.</th><th>Tiempo</th><th>Comentarios</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${machineById(e.maquina)?.label || e.maquina}</td>
                <td class="left">${e.producto || ''}</td>
                <td>${e.np || ''}</td>
                <td class="left">${e.operador || ''}</td>
                <td>${e.rrhh || 0}</td>
                <td>${e.cant || 0}</td>
                <td>${e.tiempo || 0}</td>
                <td class="left">${e.comentarios || ''}</td>
                <td><button class="btn small danger" onclick="window.deleteMaquinaEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="10" class="empty">Sin registros este mes</td></tr>`;
}

function renderManualesTable() {
    const t = document.getElementById('tbl-manuales');
    const hoy = new Date();
    const rows = filterByMonthYear(DB.manuales, hoy.getMonth() + 1, hoy.getFullYear())
        .sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Producto</th><th>NP</th><th>Operador</th><th>RRHH</th><th>Cant.</th><th>Tiempo</th><th>Operaciones</th><th>Sub Proc</th><th>Comentarios</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${e.producto || ''}</td>
                <td>${e.np || ''}</td>
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
        `<tr><td colspan="11" class="empty">Sin registros este mes</td></tr>`;
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
        comentarios: val('f-comentarios')
    };
    DB.maquinas.push(nuevo);
    await saveArr('entries-maquinas', DB.maquinas);
    pushToSheets('MAQUINAS', mapMaquinaRow(nuevo));
    // Clear rápido: conserva operador(es) y RRHH para capturas seguidas de la misma máquina
    ['f-producto', 'f-np', 'f-cant', 'f-tiempo', 'f-comentarios']
        .forEach(id => document.getElementById(id).value = '');
    setStatus('status-maquinas', 'Guardado ✓');
    renderMaquinasTable();
}

async function deleteMaquinaEntry(id) {
    DB.maquinas = DB.maquinas.filter(e => e.id !== id);
    await saveArr('entries-maquinas', DB.maquinas);
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
        DB.maquinas.push({
            id: uid(),
            fecha,
            maquina: maquinaId,
            producto: row['PRODUCTO'] || '',
            np: String(row['NP'] ?? ''),
            operador: row['OPERADOR'] || '',
            rrhh: Number(row['RRHH']) || 0,
            cant: Number(row['CANT PROD']) || 0,
            tiempo: Number(row['TIEMPO']) || 0,
            comentarios: row['COMENTARIOS'] || ''
        });
        imported++;
    });
    await saveArr('entries-maquinas', DB.maquinas);
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
        OPERADOR: e.operador,
        RRHH: e.rrhh,
        'CANT PROD': e.cant,
        TIEMPO: e.tiempo,
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
    setStatus('status-manuales', 'Guardado ✓');
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

async function renderResumenMaquina() {
    const maquina = val('r-maquina');
    const mes = parseInt(val('r-mes'), 10);
    const anio = parseInt(val('r-anio'), 10);
    if (!maquina) return;

    const datosSheets = await fetchSheetsMonth(mes, anio);
    const datosMaquinasDelMes = mergeSinDuplicar(
        filterByMonthYear(DB.maquinas, mes, anio), datosSheets?.MAQUINAS, mapMaquinaRow, convertirFilaMaquina
    );
    const datosParoDelMes = mergeSinDuplicar(
        filterByMonthYear(DB.paro, mes, anio), datosSheets?.PARO, mapParoRow, convertirFilaParo
    );

    const diasEnMes = getDaysInMonth(anio, mes);
    const cfg = machineById(maquina);

    let totalProd = 0,
        totalParo = 0,
        sumPct = 0,
        nPct = 0;
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
    if (value === '') {
        delete DB.tiempoOverrides[key];
    } else {
        DB.tiempoOverrides[key] = Number(value);
    }
    await saveObj('config-tiempo-overrides', DB.tiempoOverrides);
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

async function renderResumenKit() {
    const mes = parseInt(val('rk-mes'), 10);
    const anio = parseInt(val('rk-anio'), 10);

    const datosSheets = await fetchSheetsMonth(mes, anio);
    const manualesDelMes = mergeSinDuplicar(
        filterByMonthYear(DB.manuales, mes, anio), datosSheets?.MANUALES, mapManualRow, convertirFilaManual
    );
    const rows = manualesDelMes
        .filter(e => (e.producto || '').trim().toUpperCase() === 'KIT')
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

async function renderResumenManuales() {
    const mes = parseInt(val('rm-mes'), 10);
    const anio = parseInt(val('rm-anio'), 10);
    const filtroProd = val('rm-producto').toUpperCase();

    const datosSheets = await fetchSheetsMonth(mes, anio);
    const manualesDelMes = mergeSinDuplicar(
        filterByMonthYear(DB.manuales, mes, anio), datosSheets?.MANUALES, mapManualRow, convertirFilaManual
    );
    const rows = manualesDelMes.filter(e => {
        if (filtroProd && (e.producto || '').toUpperCase() !== filtroProd) return false;
        return true;
    }).sort((a, b) => a.fecha.localeCompare(b.fecha));

    lastResumenManualesRows = rows.map(e => ({
        FECHA: formatDateDisplay(e.fecha),
        PRODUCTO: e.producto,
        NP: e.np,
        OPERADOR: e.operador,
        RRHH: e.rrhh,
        'CANT PROD': e.cant,
        OPERACIONES: opsSummary(e.ops),
        'SUB PROC': e.subproc
    }));

    const totalCant = sum(rows, 'cant');
    const t = document.getElementById('tbl-resumen-manuales');
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Producto</th><th>NP</th><th>Operador</th><th>RRHH</th><th>Cant. Prod.</th><th>Operaciones</th><th>Sub Proc</th></tr>`;
    t.querySelector('tbody').innerHTML = (rows.length ?
        rows.map(e => `
            <tr>
                <td>${formatDateDisplay(e.fecha)}</td>
                <td class="left">${e.producto || ''}</td>
                <td>${e.np || ''}</td>
                <td class="left">${e.operador || ''}</td>
                <td>${e.rrhh || 0}</td>
                <td>${e.cant || 0}</td>
                <td class="left">${opsSummary(e.ops)}</td>
                <td>${e.subproc || 0}</td>
            </tr>
        `).join('') :
        `<tr><td colspan="8" class="empty">Sin registros este mes</td></tr>`) +
        `<tr class="totals"><td colspan="5">Total</td><td>${totalCant}</td><td colspan="2"></td></tr>`;
}

function exportResumenManuales() {
    if (!lastResumenManualesRows.length) { alert('Genera el resumen primero.'); return; }
    downloadExcel(lastResumenManualesRows,
        `resumen_manuales_${val('rm-mes')}_${val('rm-anio')}.xlsx`,
        'MANUALES');
}

// ===== RESUMEN: TARIMAS =====

async function renderResumenTarimas() {
    const mes = parseInt(val('rt-mes'), 10);
    const anio = parseInt(val('rt-anio'), 10);

    const datosSheets = await fetchSheetsMonth(mes, anio);
    const rows = mergeSinDuplicar(
        filterByMonthYear(DB.tarimas, mes, anio), datosSheets?.TARIMAS, mapTarimaRow, convertirFilaTarima
    );

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

async function renderResumenContenedores() {
    const mes = parseInt(val('rc-mes'), 10);
    const anio = parseInt(val('rc-anio'), 10);

    const datosSheets = await fetchSheetsMonth(mes, anio);
    const rows = mergeSinDuplicar(
        filterByMonthYear(DB.contenedores, mes, anio), datosSheets?.CONTENEDORES, mapContenedorRow, convertirFilaContenedor
    );

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

async function renderTop5() {
    const mes = parseInt(val('rtop-mes'), 10);
    const anio = parseInt(val('rtop-anio'), 10);
    const fuente = val('rtop-fuente') || 'ambos';

    const datosSheets = await fetchSheetsMonth(mes, anio);
    let registros = [];
    if (fuente === 'ambos' || fuente === 'maquinas') {
        registros = registros.concat(mergeSinDuplicar(
            filterByMonthYear(DB.maquinas, mes, anio), datosSheets?.MAQUINAS, mapMaquinaRow, convertirFilaMaquina
        ));
    }
    if (fuente === 'ambos' || fuente === 'manuales') {
        registros = registros.concat(mergeSinDuplicar(
            filterByMonthYear(DB.manuales, mes, anio), datosSheets?.MANUALES, mapManualRow, convertirFilaManual
        ));
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

async function renderFrecuenciaNP() {
    const mes = parseInt(val('rnp-mes'), 10);
    const anio = parseInt(val('rnp-anio'), 10);
    const maquina = val('rnp-maquina');
    const fuente = val('rnp-fuente') || 'ambos';

    const datosSheets = await fetchSheetsMonth(mes, anio);
    let registros = [];
    if (fuente === 'ambos' || fuente === 'maquinas') {
        registros = registros.concat(mergeSinDuplicar(
            filterByMonthYear(DB.maquinas, mes, anio), datosSheets?.MAQUINAS, mapMaquinaRow, convertirFilaMaquina
        ));
    }
    if (fuente === 'ambos' || fuente === 'manuales') {
        registros = registros.concat(mergeSinDuplicar(
            filterByMonthYear(DB.manuales, mes, anio), datosSheets?.MANUALES, mapManualRow, convertirFilaManual
        ));
    }

    if (maquina) registros = registros.filter(e => e.maquina === maquina);

    const porNP = {};
    registros.forEach(e => {
        const np = (e.np || '').trim();
        if (!np) return; // los registros sin NP capturado no cuentan para esta gráfica
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

    // Clear completo: esta es una entrada nueva de máquina, a diferencia del
    // clear parcial de addMaquinaEntry (que conserva operador/RRHH para capturas seguidas)
    ['f-producto', 'f-np', 'f-cant', 'f-tiempo', 'f-comentarios']
        .forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = '';
        });
    resetOperadorList('f-operadores-list', 'f-rrhh');

    const maquina = machineById(maquinaActual);
    if (maquina) {
        setStatus('status-maquinas', `🔄 Campos limpiados para ${maquina.label}`, false);
    }
}
window.onMaquinaChange = onMaquinaChange;


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
        OPERADOR: e.operador,
        RRHH: e.rrhh,
        'CANT PROD': e.cant,
        TIEMPO: e.tiempo,
        COMENTARIOS: e.comentarios
    };
}
function mapManualRow(e) {
    const row = {
        FECHA: formatDateDisplay(e.fecha), PRODUCTO: e.producto, NP: e.np, OPERADOR: e.operador,
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
        ...DB.contenedores.map(e => ['CONTENEDORES', mapContenedorRow(e)])
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
    return {
        id: uid(), fecha: parseSheetDate(row.FECHA), maquina: maquinaId,
        producto: row.PRODUCTO || '', np: String(row.NP ?? ''), operador: row.OPERADOR || '',
        rrhh: Number(row.RRHH) || 0, cant: Number(row['CANT PROD']) || 0,
        tiempo: Number(row.TIEMPO) || 0, comentarios: row.COMENTARIOS || ''
    };
}
function convertirFilaManual(row) {
    const ops = {};
    OPERACIONES.forEach(op => { ops[op] = (row[op] == 1 || row[op] === '1') ? 1 : 0; });
    const subproc = Object.values(ops).reduce((a, b) => a + b, 0);
    return {
        id: uid(), fecha: parseSheetDate(row.FECHA), producto: row.PRODUCTO || '',
        np: String(row.NP ?? ''), operador: row.OPERADOR || '', rrhh: Number(row.RRHH) || 0,
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

// Consulta Sheets filtrado a un mes/año específico (el propio Apps Script
// filtra, así no se descarga el historial completo cada vez). Devuelve
// null si la sincronización está apagada, no hay URL, o falla la consulta
// (en ese caso los resúmenes simplemente siguen con los datos locales).
async function fetchSheetsMonth(mes, anio) {
    if (!SHEETS_CONFIG.enabled || !SHEETS_CONFIG.url) return null;
    try {
        const sep = SHEETS_CONFIG.url.includes('?') ? '&' : '?';
        const resp = await fetch(`${SHEETS_CONFIG.url}${sep}mes=${mes}&anio=${anio}`, { method: 'GET' });
        return await resp.json();
    } catch (err) {
        console.error('No se pudo consultar Google Sheets para el resumen:', err);
        return null;
    }
}

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
    let datos;
    try {
        const sep = url.includes('?') ? '&' : '?';
        const resp = await fetch(`${url}${sep}mes=${mesActual}&anio=${anioActual}`, { method: 'GET' });
        datos = await resp.json();
    } catch (err) {
        setStatus('status-importar-sheets', 'No se pudo leer el Sheet. ¿Actualizaste el Apps Script a la versión con doGet?', true);
        console.error(err);
        return;
    }

    const conteos = [
        ['maquinas', 'entries-maquinas', 'MAQUINAS', mapMaquinaRow, convertirFilaMaquina],
        ['manuales', 'entries-manuales', 'MANUALES', mapManualRow, convertirFilaManual],
        ['paro', 'entries-paro', 'PARO', mapParoRow, convertirFilaParo],
        ['tarimas', 'entries-tarimas', 'TARIMAS', mapTarimaRow, convertirFilaTarima],
        ['contenedores', 'entries-contenedores', 'CONTENEDORES', mapContenedorRow, convertirFilaContenedor]
    ];

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

    buildChecklist();
    initOperadorList('f-operadores-list', 'f-rrhh');
    initOperadorList('m-operadores-list', 'm-rrhh');
    setDateInputs();
    setMonthSelects();
    setupTabs();

    // Cargar datos
    await loadAll();

    // Precargar configuración de Google Sheets en el formulario
    document.getElementById('cfg-webhook-url').value = SHEETS_CONFIG.url || '';
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

// UI
window.updateSubProc = updateSubProc;

/* ================================================================
   INICIAR APLICACIÓN
   ================================================================ */

boot();

/* ================================================================
   FIN DEL ARCHIVO
   ================================================================ */