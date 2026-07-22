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

function parseExcelDate(v) {
    if (v instanceof Date) return toISODate(v);
    if (typeof v === 'number') {
        const d = XLSX.SSF.parse_date_code(v);
        return d ? `${d.y}-${pad2(d.m)}-${pad2(d.d)}` : '';
    }
    if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
    return '';
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

async function loadAll() {
    DB.maquinas = await loadArr('entries-maquinas');
    DB.manuales = await loadArr('entries-manuales');
    DB.paro = await loadArr('entries-paro');
    DB.tarimas = await loadArr('entries-tarimas');
    DB.contenedores = await loadArr('entries-contenedores');
    DB.tiempoOverrides = await loadObj('config-tiempo-overrides');
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
    ['r-mes', 'rk-mes', 'rm-mes', 'rt-mes', 'rc-mes', 'rtop-mes'].forEach(id => {
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
    const rows = [...DB.maquinas].sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Máquina</th><th>Producto</th><th>NP</th><th>Operador</th><th>RRHH</th><th>Cant. Prod.</th><th>Tiempo</th><th>Comentarios</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${e.fecha}</td>
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
        `<tr><td colspan="10" class="empty">Sin registros aún</td></tr>`;
}

function renderManualesTable() {
    const t = document.getElementById('tbl-manuales');
    const rows = [...DB.manuales].sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Producto</th><th>NP</th><th>Operador</th><th>RRHH</th><th>Cant.</th><th>Tiempo</th><th>Operaciones</th><th>Sub Proc</th><th>Comentarios</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${e.fecha}</td>
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
        `<tr><td colspan="11" class="empty">Sin registros aún</td></tr>`;
}

function renderParoTable() {
    const t = document.getElementById('tbl-paro');
    const rows = [...DB.paro].sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Máquina</th><th>Minutos</th><th>Motivo</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${e.fecha}</td>
                <td class="left">${machineById(e.maquina)?.label || e.maquina}</td>
                <td>${e.minutos || 0}</td>
                <td class="left">${e.motivo || ''}</td>
                <td><button class="btn small danger" onclick="window.deleteParoEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="5" class="empty">Sin registros aún</td></tr>`;
}

function renderTarimasTable() {
    const t = document.getElementById('tbl-tarimas');
    const rows = [...DB.tarimas].sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Persona</th><th>Tipo</th><th>Cantidad</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${e.fecha}</td>
                <td class="left">${e.persona || ''}</td>
                <td class="left">${e.tipo || ''}</td>
                <td>${e.cant || 0}</td>
                <td><button class="btn small danger" onclick="window.deleteTarimaEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="5" class="empty">Sin registros aún</td></tr>`;
}

function renderContenedoresTable() {
    const t = document.getElementById('tbl-contenedores');
    const rows = [...DB.contenedores].sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Turno</th><th>Personas</th><th>Contenedores</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${e.fecha}</td>
                <td class="left">${e.turno || ''}</td>
                <td>${e.personas || 0}</td>
                <td>${e.cant || 0}</td>
                <td><button class="btn small danger" onclick="window.deleteContenedorEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="5" class="empty">Sin registros aún</td></tr>`;
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
    DB.maquinas.push({
        id: uid(),
        fecha,
        maquina,
        producto: val('f-producto'),
        np: val('f-np'),
        operador: val('f-operador'),
        rrhh: numval('f-rrhh'),
        cant: numval('f-cant'),
        tiempo: numval('f-tiempo'),
        comentarios: val('f-comentarios')
    });
    await saveArr('entries-maquinas', DB.maquinas); 
    ['f-producto', 'f-np', 'f-operador', 'f-rrhh' , 'f-cant', 'f-tiempo', 'f-comentarios'] 
        .forEach(id => document.getElementById(id).value = ''); //BORRA LOS CAMPOS NO GUARDADOS
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
        FECHA: e.fecha,
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
    DB.manuales.push({
        id: uid(),
        fecha,
        producto: val('m-producto'),
        np: val('m-np'),
        operador: val('m-operador'),
        rrhh: numval('m-rrhh'),
        cant: numval('m-cant'),
        tiempo: numval('m-tiempo'),
        ops,
        subproc,
        comentarios: val('m-comentarios')
    });
    await saveArr('entries-manuales', DB.manuales);
    ['m-producto', 'm-np', 'm-operador', 'm-rrhh', 'm-cant', 'm-tiempo', 'm-comentarios']
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
            FECHA: e.fecha,
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
    DB.paro.push({
        id: uid(),
        fecha,
        maquina,
        minutos: numval('p-minutos'),
        motivo: val('p-motivo')
    });
    await saveArr('entries-paro', DB.paro);
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
    DB.tarimas.push({
        id: uid(),
        fecha,
        persona: val('t-persona'),
        tipo: val('t-tipo'),
        cant: numval('t-cant')
    });
    await saveArr('entries-tarimas', DB.tarimas);
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
    DB.contenedores.push({
        id: uid(),
        fecha,
        turno: val('c-turno'),
        personas: numval('c-personas'),
        cant: numval('c-cant')
    });
    await saveArr('entries-contenedores', DB.contenedores);
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
        const entriesDia = DB.maquinas.filter(e => e.fecha === fecha && e.maquina === maquina);
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

        const paroDia = DB.paro.filter(e => e.fecha === fecha && e.maquina === maquina);
        const paroMin = sum(paroDia, 'minutos');
        const motivos = paroDia.map(e => e.motivo).filter(x => x).join('; ');
        const operadoresDia = [...new Set(entriesDia.map(e => e.operador).filter(x => x))].join(', ');
        const fechaTitle = operadoresDia ? `Operador(es): ${operadoresDia}` : 'Sin operador capturado';

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
            <td title="${fechaTitle}" style="cursor:help;border-bottom:1px dotted var(--sub);">${fecha}</td>
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
            FECHA: fecha,
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

function renderResumenKit() {
    const mes = parseInt(val('rk-mes'), 10);
    const anio = parseInt(val('rk-anio'), 10);
    const rows = DB.manuales.filter(e => {
        if ((e.producto || '').trim().toUpperCase() !== 'KIT') return false;
        const d = parseDate(e.fecha);
        return d.getFullYear() === anio && (d.getMonth() + 1) === mes;
    }).sort((a, b) => a.fecha.localeCompare(b.fecha));

    const total = sum(rows, 'cant');
    lastResumenKitRows = rows.map(e => ({
        FECHA: e.fecha,
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
                <td>${e.fecha}</td>
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

function renderResumenManuales() {
    const mes = parseInt(val('rm-mes'), 10);
    const anio = parseInt(val('rm-anio'), 10);
    const filtroProd = val('rm-producto').toUpperCase();

    const rows = DB.manuales.filter(e => {
        const d = parseDate(e.fecha);
        if (d.getFullYear() !== anio || (d.getMonth() + 1) !== mes) return false;
        if (filtroProd && (e.producto || '').toUpperCase() !== filtroProd) return false;
        return true;
    }).sort((a, b) => a.fecha.localeCompare(b.fecha));

    lastResumenManualesRows = rows.map(e => ({
        FECHA: e.fecha,
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
                <td>${e.fecha}</td>
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

function renderResumenTarimas() {
    const mes = parseInt(val('rt-mes'), 10);
    const anio = parseInt(val('rt-anio'), 10);
    const rows = filterByMonthYear(DB.tarimas, mes, anio);

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

function renderResumenContenedores() {
    const mes = parseInt(val('rc-mes'), 10);
    const anio = parseInt(val('rc-anio'), 10);
    const rows = filterByMonthYear(DB.contenedores, mes, anio);

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

// ===== RESUMEN: TOP 5 PRODUCTOS =====

function renderTop5() {
    const mes = parseInt(val('rtop-mes'), 10);
    const anio = parseInt(val('rtop-anio'), 10);
    const fuente = val('rtop-fuente') || 'ambos';

    let registros = [];
    if (fuente === 'ambos' || fuente === 'maquinas') registros = registros.concat(DB.maquinas);
    if (fuente === 'ambos' || fuente === 'manuales') registros = registros.concat(DB.manuales);

    const delMes = filterByMonthYear(registros, mes, anio);

    const porProducto = {};
    delMes.forEach(e => {
        const nombre = (e.producto || 'SIN PRODUCTO').trim().toUpperCase();
        porProducto[nombre] = (porProducto[nombre] || 0) + (Number(e.cant) || 0);
    });

    const top5 = Object.entries(porProducto)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    const cont = document.getElementById('chart-top5');
    if (!top5.length) {
        cont.innerHTML = `<div class="empty">Sin datos capturados este mes</div>`;
        return;
    }

    const max = top5[0][1];
    cont.innerHTML = top5.map(([producto, cant]) => {
        const pct = max ? Math.round((cant / max) * 100) : 0;
        return `
            <div class="top5-row">
                <div class="top5-label" title="${producto}">${producto}</div>
                <div class="top5-bar-bg"><div class="top5-bar-fill" style="width:${pct}%"></div></div>
                <div class="top5-value">${cant.toLocaleString()}</div>
            </div>
        `;
    }).join('');
}

// ================================================================
// LISTENER PARA CAMBIO DE MÁQUINA 
// ================================================================

/*function onMaquinaChange() {
    const selector = document.getElementById('f-maquina');
    if (!selector) return;
    const maquinaActual = selector.value;
    if (!maquinaActual) return;

    // Limpia los mismos campos que se limpian al guardar un registro
    // (deja fecha, máquina, operador y RRHH tal como quedaron, igual que addMaquinaEntry)
    ['f-producto', 'f-np', 'f-cant', 'f-tiempo', 'f-comentarios']
        .forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = '';
        });

    const maquina = machineById(maquinaActual);
    if (maquina) {
        setStatus('status-maquinas', `🔄 Campos limpiados para ${maquina.label}`, false);
    }
}
window.onMaquinaChange = onMaquinaChange;*/


/* ================================================================
   INICIALIZACIÓN (BOOT)
   ================================================================ */

async function boot() {
    // Configurar UI
    fillMachineSelect('f-maquina');
    fillMachineSelect('p-maquina');
    fillMachineSelect('r-maquina');
    fillMonthSelect('r-mes');
    fillMonthSelect('rk-mes');
    fillMonthSelect('rm-mes');
    fillMonthSelect('rt-mes');
    fillMonthSelect('rc-mes');
    fillMonthSelect('rtop-mes');

    buildChecklist();
    setDateInputs();
    setMonthSelects();
    setupTabs();

    // Cargar datos
    await loadAll();

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

// UI
window.updateSubProc = updateSubProc;

/* ================================================================
   INICIAR APLICACIÓN
   ================================================================ */

boot();

/* ================================================================
   FIN DEL ARCHIVO
   ================================================================ */