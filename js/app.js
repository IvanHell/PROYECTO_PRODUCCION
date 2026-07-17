/* ===================================================================
   PANEL DE PRODUCCIÓN — MÁQUINAS VALLE
   Todos los JavaScript en un solo archivo
   =================================================================== */

/* ======================= CONFIGURACIÓN ======================= */

const MACHINES = [
    { id: 'FLEXO', label: 'Flexo', piezasMin: 26, ajuste: 30, flexo: true },
    { id: 'ROLADORA', label: 'Roladora I', piezasMin: 10, ajuste: 15 },
    { id: 'RANURADORA', label: 'Ranuradora', piezasMin: 37, ajuste: 35 },
    { id: 'T_ROTATIVA', label: 'Troqueladora Rotativa', piezasMin: 35, ajuste: 50 },
    { id: 'CAIMAN', label: 'Caimán', piezasMin: 5, ajuste: 15 },
    { id: 'PEGADORA', label: 'Semiautomática (Pegadora)', piezasMin: 80, ajuste: 30 },
    { id: 'T_PLANA', label: 'Troqueladora Plana', piezasMin: 35, ajuste: 50 },
    { id: 'ARMADORA_REJILLAS', label: 'Máquina Armadora (Armadora de Rejillas)', piezasMin: 26, ajuste: 30 },
];

const DIAS = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
const DIAS_ABR = ['D', 'L', 'M', 'MI', 'J', 'V', 'S'];
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

/* ======================= HELPERS DE CONFIGURACIÓN ======================= */

function machineById(id) {
    return MACHINES.find(m => m.id === id);
}

function defaultTiempoDisponible(machineId, dateStr) {
    const m = machineById(machineId);
    const d = new Date(dateStr + 'T00:00:00');
    const dow = d.getDay();
    if (m && m.flexo) return null;
    if (dow === 0 || dow === 6) return 0;
    if (dow === 2) return 600;
    return 570;
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
    kit: [],
    paro: [],
    tarimas: [],
    contenedores: [],
    tiempoOverrides: {}
};

async function loadAll() {
    DB.maquinas = await loadArr('entries-maquinas');
    DB.kit = await loadArr('entries-kit');
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
    }, 3000);
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
    const today = getToday();
    ['f-fecha', 'k-fecha', 'p-fecha', 't-fecha', 'c-fecha'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = today;
    });
}

function setMonthSelects() {
    const month = getCurrentMonth();
    ['r-mes', 'rk-mes', 'rt-mes', 'rc-mes'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = month;
    });
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
                <td><button class="btn small danger" onclick="deleteMaquinaEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="10" class="empty">Sin registros aún</td></tr>`;
}

function renderKitTable() {
    const t = document.getElementById('tbl-kit');
    const rows = [...DB.kit].sort((a, b) => b.fecha.localeCompare(a.fecha));
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Personal</th><th>NP KIT</th><th>Cantidad</th><th></th></tr>`;
    t.querySelector('tbody').innerHTML = rows.length ?
        rows.map(e => `
            <tr>
                <td>${e.fecha}</td>
                <td class="left">${e.personal || ''}</td>
                <td class="left">${e.np || ''}</td>
                <td>${e.cant || 0}</td>
                <td><button class="btn small danger" onclick="deleteKitEntry('${e.id}')">Borrar</button></td>
            </tr>
        `).join('') :
        `<tr><td colspan="5" class="empty">Sin registros aún</td></tr>`;
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
                <td><button class="btn small danger" onclick="deleteParoEntry('${e.id}')">Borrar</button></td>
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
                <td><button class="btn small danger" onclick="deleteTarimaEntry('${e.id}')">Borrar</button></td>
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
                <td><button class="btn small danger" onclick="deleteContenedorEntry('${e.id}')">Borrar</button></td>
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
    ['f-producto', 'f-np', 'f-operador', 'f-rrhh', 'f-cant', 'f-tiempo', 'f-comentarios']
        .forEach(id => document.getElementById(id).value = '');
    setStatus('status-maquinas', 'Guardado ✓');
    renderMaquinasTable();
}

async function deleteMaquinaEntry(id) {
    DB.maquinas = DB.maquinas.filter(e => e.id !== id);
    await saveArr('entries-maquinas', DB.maquinas);
    renderMaquinasTable();
}

// ===== KIT =====

async function addKitEntry() {
    const fecha = val('k-fecha');
    if (!fecha) {
        setStatus('status-kit', 'La fecha es obligatoria.', true);
        return;
    }
    DB.kit.push({
        id: uid(),
        fecha,
        personal: val('k-personal'),
        np: val('k-np'),
        cant: numval('k-cant')
    });
    await saveArr('entries-kit', DB.kit);
    ['k-personal', 'k-np', 'k-cant'].forEach(id => document.getElementById(id).value = '');
    setStatus('status-kit', 'Guardado ✓');
    renderKitTable();
}

async function deleteKitEntry(id) {
    DB.kit = DB.kit.filter(e => e.id !== id);
    await saveArr('entries-kit', DB.kit);
    renderKitTable();
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

    for (let d = 1; d <= diasEnMes; d++) {
        const fecha = `${anio}-${String(mes).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
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
            <td>${fecha}</td>
            <td><input type="number" style="width:70px" value="${tiempoDisp}" onchange="updateTiempoOverride('${maquina}','${fecha}',this.value)"></td>
            <td>${produccion}</td>
            <td>${partidas}</td>
            <td>${real ? real.toFixed(2) : '-'}</td>
            <td>${esperado ? esperado.toFixed(2) : '-'}</td>
            <td>${pctPill}</td>
            <td>${paroMin || ''}</td>
            <td class="left">${motivos}</td>
        </tr>`);
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

// ===== RESUMEN: KIT =====

function renderResumenKit() {
    const mes = parseInt(val('rk-mes'), 10);
    const anio = parseInt(val('rk-anio'), 10);
    const rows = filterByMonthYear(DB.kit, mes, anio).sort((a, b) => a.fecha.localeCompare(b.fecha));
    const total = sum(rows, 'cant');

    const t = document.getElementById('tbl-resumen-kit');
    t.querySelector('thead').innerHTML =
        `<tr><th>Fecha</th><th>Personal</th><th>NP KIT</th><th>Cantidad</th></tr>`;
    t.querySelector('tbody').innerHTML = (rows.length ?
        rows.map(e => `
            <tr>
                <td>${e.fecha}</td>
                <td class="left">${e.personal || ''}</td>
                <td class="left">${e.np || ''}</td>
                <td>${e.cant || 0}</td>
            </tr>
        `).join('') :
        `<tr><td colspan="4" class="empty">Sin registros este mes</td></tr>`) +
        `<tr class="totals"><td colspan="3">Total del mes</td><td>${total}</td></tr>`;
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

    const t = document.getElementById('tbl-resumen-tarimas');
    t.querySelector('thead').innerHTML =
        `<tr><th>Persona</th><th>Tipo</th><th>Total del mes</th></tr>`;
    const entries = Object.entries(porPersona);
    t.querySelector('tbody').innerHTML = entries.length ?
        entries.map(([k, v]) => {
            const [persona, tipo] = k.split(' | ');
            return `<tr><td class="left">${persona}</td><td class="left">${tipo}</td><td>${v}</td></tr>`;
        }).join('') :
        `<tr><td colspan="3" class="empty">Sin registros este mes</td></tr>`;
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

    const t = document.getElementById('tbl-resumen-contenedores');
    t.querySelector('thead').innerHTML =
        `<tr><th>Turno</th><th>Personas (suma)</th><th>Contenedores (total)</th><th>Contenedores / persona-día</th></tr>`;
    const entries = Object.entries(porTurno);
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
    fillMonthSelect('rt-mes');
    fillMonthSelect('rc-mes');

    setDateInputs();
    setMonthSelects();
    setupTabs();

    // Cargar datos
    await loadAll();

    // Renderizar tablas
    renderMaquinasTable();
    renderKitTable();
    renderParoTable();
    renderTarimasTable();
    renderContenedoresTable();
}

// ================================================================
// EXPONER FUNCIONES AL ÁMBITO GLOBAL (para onclick en HTML)
// ================================================================

window.addMaquinaEntry = addMaquinaEntry;
window.deleteMaquinaEntry = deleteMaquinaEntry;
window.addKitEntry = addKitEntry;
window.deleteKitEntry = deleteKitEntry;
window.addParoEntry = addParoEntry;
window.deleteParoEntry = deleteParoEntry;
window.addTarimaEntry = addTarimaEntry;
window.deleteTarimaEntry = deleteTarimaEntry;
window.addContenedorEntry = addContenedorEntry;
window.deleteContenedorEntry = deleteContenedorEntry;

window.renderResumenMaquina = renderResumenMaquina;
window.updateTiempoOverride = updateTiempoOverride;
window.renderResumenKit = renderResumenKit;
window.renderResumenTarimas = renderResumenTarimas;
window.renderResumenContenedores = renderResumenContenedores;

// ================================================================
// INICIAR APLICACIÓN
// ================================================================

boot();

/* ================================================================
   FIN DEL ARCHIVO
   ================================================================ */