let allData = [];
let filteredData = [];
let __barChartData = null;

const COLUMN_ALIASES = {
    referencia: ['referencia', 'ref', 'sku', 'código', 'codigo', 'code', 'item', 'producto', 'id', 'cód', 'cod'],
    color: ['color', 'cod_color', 'codcolor', 'desc_color', 'desccolor', 'nombre_color', 'col'],
    estado: ['estado', 'status', 'state', 'condición', 'condicion'],
    stock: ['stock', 'inventario', 'existencia', 'existencias', 'cantidad', 'qty', 'quantity', 'inv actual', 'inv_actual', 'saldo'],
    cierre: ['cierre', 'closing'],
    agotados: ['agotados', 'agotado', 'agot', 'faltante', 'faltantes', 'shortage', 'stockout'],
    descripcion: ['descripción', 'descripcion', 'description', 'producto', 'nombre', 'name', 'detalle'],
    categoria: ['categoría', 'categoria', 'category', 'linea', 'line', 'departamento', 'subcategoria'],
    es_vmi: ['es vmi', 'vmi', 'tipo vmi', 'tipo'],
    talla: ['talla', 'tamaño', 'size', 'medida'],
    tendencia: ['tendencia', 'trend', 'tend'],
    cumplimiento: ['cumplimiento', 'compliance', 'cumpl'],
    tend_dia: ['tend dia', 'tendencia dia', 'tendencia día', 'daily trend', 'tendencia diaria']
};

function findColumnLoose(aliases, columns) {
    const lowerCols = columns.map(c => c.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''));
    for (const col of lowerCols) {
        for (const alias of aliases) {
            if (col.includes(alias)) return columns[lowerCols.indexOf(col)];
        }
    }
    return null;
}

function findColumn(aliases, columns) {
    const lowerCols = columns.map(c => c.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''));
    for (const alias of aliases) {
        const idx = lowerCols.indexOf(alias);
        if (idx !== -1) return columns[idx];
    }
    return null;
}

function normalize(s) {
    return String(s ?? '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function getStockValue(row, stockCol) {
    const v = row[stockCol];
    if (v === undefined || v === null || v === '') return NaN;
    const n = Number(String(v).replace(/[^0-9.,\-]/g, '').replace(',', '.'));
    return isNaN(n) ? NaN : n;
}

function getEstado(stock) {
    if (stock === 0) return 'AGOTADO';
    if (stock <= 5) return 'CRÍTICO';
    if (stock <= 20) return 'BAJO STOCK';
    return 'DISPONIBLE';
}

const ESTADO_MAP = {
    'd': 'DESCONTINUADO',
    'e': 'ESTABLECIDO',
    'ee': 'ESTRATEGIA ESTABLECIDO',
    'n': 'NUEVO',
    'ne': 'NUEVO ESTABLECIDO',
    'np': 'NO PORTAFOLIO',
    'r': 'RETIRADO'
};

function getEstadoFull(code) {
    return ESTADO_MAP[normalize(code)] || code || '';
}

function getEstadoColor(estado) {
    switch (estado) {
        case 'AGOTADO': return { bg: '#ba1a1a', text: '#fff' };
        case 'CRÍTICO': return { bg: '#F59E0B', text: '#fff' };
        case 'BAJO STOCK': return { bg: '#facc15', text: '#0b1c30' };
        case 'D':
        case 'DESCONTINUADO': return { bg: '#6b7280', text: '#fff' };
        case 'E':
        case 'ESTABLECIDO': return { bg: '#10b981', text: '#fff' };
        case 'EE':
        case 'ESTRATEGIA ESTABLECIDO': return { bg: '#3b82f6', text: '#fff' };
        case 'N':
        case 'NUEVO': return { bg: '#8b5cf6', text: '#fff' };
        case 'NE':
        case 'NUEVO ESTABLECIDO': return { bg: '#06b6d4', text: '#fff' };
        case 'NP':
        case 'NO PORTAFOLIO': return { bg: '#f97316', text: '#fff' };
        case 'R':
        case 'RETIRADO': return { bg: '#ef4444', text: '#fff' };
        default: return { bg: '#10b981', text: '#fff' };
    }
}

function getPrioridad(stock) {
    if (stock === 0) return { label: 'INMEDIATA', color: '#ba1a1a' };
    if (stock <= 5) return { label: 'ALTA', color: '#F59E0B' };
    if (stock <= 20) return { label: 'MEDIA', color: '#737780' };
    return { label: 'NORMAL', color: '#10b981' };
}

// ─── FILE UPLOAD ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('file-input');
    const uploadBtn = document.getElementById('upload-btn');
    const loadingMsg = document.getElementById('loading-msg');

    uploadBtn.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (!file) return;
        loadingMsg.classList.remove('hidden');
        uploadBtn.disabled = true;

        if (typeof XLSX === 'undefined') {
            alert("La librería XLSX no se cargó. Verifica tu conexión a internet.");
            loadingMsg.classList.add('hidden');
            uploadBtn.disabled = false;
            return;
        }

        const reader = new FileReader();
        reader.onload = function (event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                if (!workbook.SheetNames.includes('INFORME DE VENTAS')) {
                    alert("Error: No se encontró la hoja 'INFORME DE VENTAS'. Verifica que el nombre sea exacto.");
                    loadingMsg.classList.add('hidden');
                    uploadBtn.disabled = false;
                    return;
                }

                const worksheet = workbook.Sheets['INFORME DE VENTAS'];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

                if (!jsonData || jsonData.length === 0) {
                    alert("El archivo está vacío.");
                    loadingMsg.classList.add('hidden');
                    uploadBtn.disabled = false;
                    return;
                }

                allData = jsonData;
                document.getElementById('drop-zone').classList.add('hidden');

                const now = new Date();
                const label = `Última actualización: ${now.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' })}`;
                document.getElementById('last-updated').textContent = label;

                lsSave('main-data', { data: allData, columns: Object.keys(allData[0]), label, timestamp: Date.now() });

                buildDashboard();
                loadingMsg.classList.add('hidden');
                uploadBtn.disabled = false;

            } catch (err) {
                console.error(err);
                alert("Error al leer el archivo: " + err.message);
                loadingMsg.classList.add('hidden');
                uploadBtn.disabled = false;
            }
        };
        reader.onerror = function () {
            alert("No se pudo leer el archivo. Verifica que no esté dañado o abierto en otro programa.");
            loadingMsg.classList.add('hidden');
            uploadBtn.disabled = false;
        };
        reader.readAsArrayBuffer(file);
    });
});

// ─── BUILD DASHBOARD ──────────────────────────────────────────
function buildDashboard() {
    if (!allData || allData.length === 0) return;

    const columns = Object.keys(allData[0]);

    const colRef = findColumn(COLUMN_ALIASES.referencia, columns);
    const colColor = findColumn(COLUMN_ALIASES.color, columns);
    const colEstado = findColumn(COLUMN_ALIASES.estado, columns);
    const colStock = findColumn(COLUMN_ALIASES.stock, columns);
    let colCierre = findColumn(COLUMN_ALIASES.cierre, columns);
    if (!colCierre) colCierre = findColumnLoose(COLUMN_ALIASES.cierre, columns);
    const colAgotados = findColumn(COLUMN_ALIASES.agotados, columns) || findColumnLoose(COLUMN_ALIASES.agotados, columns);
    const colDesc = findColumn(COLUMN_ALIASES.descripcion, columns);
    const colCategoria = findColumn(COLUMN_ALIASES.categoria, columns);

    console.log('Columnas detectadas:', { colRef, colColor, colEstado, colStock, colCierre, colAgotados, colDesc, colCategoria });
    console.log('Todas las columnas del Excel:', columns);

    allData.forEach(row => {
        row.__stock = getStockValue(row, colStock);
        row.__cierre = getStockValue(row, colCierre || colStock);
        row.__agotados = colAgotados ? getStockValue(row, colAgotados) : NaN;
        row.__estado = getEstado(row.__stock);
        row.__prioridad = getPrioridad(row.__stock);
    });

    populateFilters(columns, colRef, colColor);
    applyFilters(columns, colRef, colColor, colStock, colAgotados, colDesc, colCategoria);
}

// ─── POPULATE FILTERS ─────────────────────────────────────────
function populateFilters(columns, colRef, colColor) {
    setSelectOptions('filter-color', allData, colColor);

    const refInput = document.getElementById('filter-ref');
    const refDropdown = document.getElementById('ref-dropdown');
    let selectedIndex = -1;
    let currentSuggestions = [];

    function getAllRefs() {
        if (!colRef) return [];
        return [...new Set(allData.map(r => String(r[colRef] ?? '').trim()).filter(Boolean))].sort();
    }

    function updateColorSelect(refValue) {
        let subset = allData;
        if (refValue) {
            subset = subset.filter(r => normalize(r[colRef]).includes(normalize(refValue)));
        }
        const cc = findColumn(COLUMN_ALIASES.color, Object.keys(allData[0]));
        setSelectOptions('filter-color', subset, cc, null);
    }

    function renderDropdown(match) {
        const allRefs = getAllRefs();
        if (!match) {
            refDropdown.classList.add('hidden');
            refDropdown.innerHTML = '';
            currentSuggestions = [];
            return;
        }
        const nm = normalize(match);
        currentSuggestions = allRefs.filter(r => normalize(r).includes(nm));
        if (!currentSuggestions.length) {
            refDropdown.classList.add('hidden');
            refDropdown.innerHTML = '';
            return;
        }
        refDropdown.innerHTML = currentSuggestions.map((r, i) =>
            `<div class="px-3 py-2 cursor-pointer hover:bg-surface-container-low text-body-md ${i === 0 ? 'bg-surface-container-low' : ''}" data-index="${i}">${r}</div>`
        ).join('');
        refDropdown.classList.remove('hidden');
        selectedIndex = 0;
        highlightItem(0);
    }

    function highlightItem(idx) {
        refDropdown.querySelectorAll('div').forEach((el, i) => {
            el.classList.toggle('bg-surface-container-low', i === idx);
        });
    }

    function selectRef(value) {
        refInput.value = value;
        refDropdown.classList.add('hidden');
        refDropdown.innerHTML = '';
        selectedIndex = -1;
        updateColorSelect(value);
        applyCurrentFilters();
    }

    function applyCurrentFilters() {
        const cols = Object.keys(allData[0]);
        const cr = findColumn(COLUMN_ALIASES.referencia, cols);
        const cc = findColumn(COLUMN_ALIASES.color, cols);
        const cs = findColumn(COLUMN_ALIASES.stock, cols);
        const ca = findColumn(COLUMN_ALIASES.agotados, cols);
        const cd = findColumn(COLUMN_ALIASES.descripcion, cols);
        const ccat = findColumn(COLUMN_ALIASES.categoria, cols);
        applyFilters(cols, cr, cc, cs, ca, cd, ccat);
    }

    let debounceTimer;

    refInput.addEventListener('input', () => {
        const val = refInput.value.trim();
        renderDropdown(val);
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            updateColorSelect(val);
            applyCurrentFilters();
        }, 300);
    });

    refInput.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            if (currentSuggestions.length) {
                selectedIndex = Math.min(selectedIndex + 1, currentSuggestions.length - 1);
                highlightItem(selectedIndex);
            }
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            if (currentSuggestions.length) {
                selectedIndex = Math.max(selectedIndex - 1, 0);
                highlightItem(selectedIndex);
            }
        } else if (e.key === 'Enter') {
            e.preventDefault();
            if (selectedIndex >= 0 && currentSuggestions[selectedIndex]) {
                selectRef(currentSuggestions[selectedIndex]);
            }
        } else if (e.key === 'Escape') {
            refDropdown.classList.add('hidden');
            refDropdown.innerHTML = '';
            selectedIndex = -1;
        }
    });

    refInput.addEventListener('blur', () => {
        setTimeout(() => {
            refDropdown.classList.add('hidden');
            refDropdown.innerHTML = '';
            selectedIndex = -1;
        }, 200);
    });

    refDropdown.addEventListener('mousedown', (e) => {
        const div = e.target.closest('div[data-index]');
        if (div) {
            e.preventDefault();
            selectRef(currentSuggestions[parseInt(div.dataset.index)]);
        }
    });

    document.getElementById('filter-color').addEventListener('change', applyCurrentFilters);

    document.getElementById('btn-clear-filters').addEventListener('click', () => {
        document.getElementById('filter-ref').value = '';
        document.getElementById('filter-color').value = '__all__';
        renderDropdown('');
        applyCurrentFilters();
    });

    document.getElementById('global-search').addEventListener('input', applyCurrentFilters);
}

function setSelectOptions(selectId, data, col, keepVal) {
    const sel = document.getElementById(selectId);
    if (!col || !data.length) return;
    const values = [...new Set(data.map(r => String(r[col] ?? '').trim()).filter(Boolean))].sort();
    const current = sel.value;
    sel.innerHTML = `<option value="__all__">${sel.options[0].textContent}</option>`;
    values.forEach(v => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        sel.appendChild(opt);
    });
    if (keepVal && values.includes(keepVal)) sel.value = keepVal;
    else if (values.includes(current)) sel.value = current;
    else sel.value = '__all__';
}

// ─── APPLY FILTERS ────────────────────────────────────────────
function applyFilters(columns, colRef, colColor, colStock, colAgotados, colDesc, colCategoria) {
    const cols = Object.keys(allData[0]);
    let colCierre = findColumn(COLUMN_ALIASES.cierre, cols);
    if (!colCierre) colCierre = findColumnLoose(COLUMN_ALIASES.cierre, cols);

    const vRef = document.getElementById('filter-ref').value.trim();
    const vColor = document.getElementById('filter-color').value;
    const vSearch = normalize(document.getElementById('global-search').value);

    filteredData = allData.filter(row => {
        if (vRef && colRef && !normalize(row[colRef]).includes(normalize(vRef))) return false;
        if (vColor !== '__all__' && colColor && normalize(row[colColor]) !== normalize(vColor)) return false;
        if (vSearch) {
            const haystack = columns.map(c => normalize(row[c])).join(' ');
            if (!haystack.includes(vSearch)) return false;
        }
        return true;
    });

    renderKPIs(columns, colRef, colStock, colCierre, colColor, vColor, colAgotados);
    renderAlerts(columns, colRef, colDesc, colStock, colColor, vRef, vColor);
    renderTable(columns, colRef, colDesc, colStock);
    renderBottomItems(columns, colRef, colStock);
    renderCharts(columns, colCategoria, colStock, colColor);
    updateInsights(columns, colCategoria, colStock);
    renderRevisionTable(colRef, colColor, colDesc, colStock, colCierre, colAgotados, findColumn(COLUMN_ALIASES.estado, cols));
    renderRevisionDashboard(colRef, colColor, colStock, colCierre, colAgotados, cols);
}

// ─── KPIs ─────────────────────────────────────────────────────
function renderKPIs(columns, colRef, colStock, colCierre, colColor, vColor, colAgotados) {
    const _colColor = colColor || findColumn(COLUMN_ALIASES.color, columns) || findColumnLoose(COLUMN_ALIASES.color, columns);
    const _colRef = colRef || findColumn(COLUMN_ALIASES.referencia, columns) || findColumnLoose(COLUMN_ALIASES.referencia, columns);
    const refColorSet = new Set();
    filteredData.forEach(r => {
        const ref = _colRef ? r[_colRef] : '';
        const color = _colColor ? r[_colColor] : '';
        if (ref || color) refColorSet.add(ref + '|' + color);
    });
    const totalRefs = refColorSet.size;
    const stockValues = filteredData.map(r => r.__stock).filter(v => !isNaN(v));
    const agotadosCount = stockValues.filter(v => v <= 0).length;
    const agotadosValues = colAgotados
        ? filteredData.map(r => r.__agotados).filter(v => !isNaN(v))
        : stockValues.filter(v => v <= 0);
    const totalAgotados = agotadosValues.reduce((a, b) => a + b, 0);
    const criticos = stockValues.filter(v => v > 0 && v <= 5).length;
    const promedio = stockValues.length ? (stockValues.reduce((a, b) => a + b, 0) / stockValues.length) : 0;
    const tasa = totalRefs ? ((agotadosCount / totalRefs) * 100) : 0;

    const totalInventario = stockValues.reduce((a, b) => a + b, 0);

    const cierreGroups = {};
    filteredData.forEach(row => {
        const ref = colRef ? normalize(row[colRef]) : '';
        const color = colColor ? normalize(row[colColor]) : '';
        const key = ref + '|' + color;
        const v = row.__cierre;
        if (!isNaN(v)) {
            cierreGroups[key] = (cierreGroups[key] || 0) + v;
        }
    });
    console.log('cierreGroups:', cierreGroups);
    const totalCierreGrupos = Object.values(cierreGroups).reduce((a, b) => a + b, 0);
    const totalCierreSimple = stockValues.reduce((a, b) => a + b, 0);
    console.log('totalCierre (agrupado):', totalCierreGrupos, '| totalCierre (simple):', totalCierreSimple);
    const totalCierre = totalCierreGrupos;

    document.getElementById('kpi-total-refs').textContent = (() => {
        const knownKeys = new Set(['__stock', '__agotados', '__estado', '__prioridad', '__cierre']);
        let cCols = columns.filter(c => {
            const norm = c.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
            return /c\d+/.test(norm) && !knownKeys.has(c);
        });
        cCols.sort((a, b) => parseInt(String(a).match(/\d+/)) - parseInt(String(b).match(/\d+/)));
        const colRange = columns.slice(15, 20);
        cCols = cCols.filter(c => colRange.includes(c));
        const total = cCols.reduce((sum, col) => {
            return sum + filteredData.reduce((s, row) => {
                const v = Number(String(row[col] ?? '').replace(/[^0-9.,\-]/g, '').replace(',', '.'));
                return s + (isNaN(v) ? 0 : v);
            }, 0);
        }, 0);
        return Math.round(total).toLocaleString();
    })();
    document.getElementById('kpi-total-agotados').textContent = Math.round(totalAgotados).toLocaleString();
    document.getElementById('kpi-tasa').textContent = tasa.toFixed(2) + '%';
    document.getElementById('kpi-cierre').textContent = Math.round(totalInventario).toLocaleString();
    document.getElementById('kpi-promedio').textContent = Math.round(totalCierre).toLocaleString();
    const cols = Object.keys(allData[0]);
    const colTendencia = findColumn(COLUMN_ALIASES.tendencia, cols) || findColumnLoose(COLUMN_ALIASES.tendencia, cols);
    const totalTendencia = colTendencia
        ? filteredData.reduce((s, row) => {
            const v = Number(String(row[colTendencia] ?? '').replace(/[^0-9.,\-]/g, '').replace(',', '.'));
            return s + (isNaN(v) ? 0 : v);
        }, 0)
        : 0;
    document.getElementById('kpi-registros').textContent = Math.round(totalTendencia).toLocaleString();
    document.getElementById('alert-count').textContent = agotadosCount;
    document.getElementById('table-count').textContent = filteredData.length + ' registros';
}

// ─── ALERTS ───────────────────────────────────────────────────
function renderAlerts(columns, colRef, colDesc, colStock, colColor, vRef, vColor) {
    const summary = document.getElementById('alerts-summary');
    const body = document.getElementById('alerts-body');
    const empty = document.getElementById('alerts-empty');
    const countBadge = document.getElementById('alert-count');

    const refLabel = (vRef && vRef !== '__all__') ? vRef : '—';
    const colorLabel = (vColor && vColor !== '__all__') ? vColor : '—';
    const totalAg = document.getElementById('kpi-total-agotados').textContent;
    const totalInv = document.getElementById('kpi-cierre').textContent;
    const totalCie = document.getElementById('kpi-promedio').textContent;

    // KPI cards
    const kpis = [
        { label: 'Ref', value: refLabel },
        { label: 'Color', value: colorLabel },
        { label: 'Total Agotados', value: totalAg },
        { label: 'Total Inventario', value: totalInv },
        { label: 'Total de Cierre', value: totalCie }
    ];
    summary.innerHTML = kpis.map(k =>
        `<div class="bg-[#f8f9fa] border border-gray-200 rounded-lg px-2.5 py-1.5 text-center flex-shrink-0">
            <div class="text-sm font-semibold text-primary leading-tight">${k.value}</div>
            <div class="text-[10px] text-gray-500 mt-0.5 whitespace-nowrap">${k.label}</div>
        </div>`
    ).join('');

    const alerts = filteredData
        .filter(r => r.__stock === 0 || r.__stock <= 5)
        .sort((a, b) => a.__stock - b.__stock)
        .slice(0, 10);

    // Badge color
    const n = alerts.length;
    countBadge.textContent = n;
    countBadge.style.background = n === 0 ? '#d1fae5' : n <= 5 ? '#fed7aa' : '#fecaca';
    countBadge.style.color = n === 0 ? '#065f46' : n <= 5 ? '#9a3412' : '#991b1b';

    if (!alerts.length) {
        body.innerHTML = '';
        body.classList.add('hidden');
        empty.classList.remove('hidden');
        return;
    }

    empty.classList.add('hidden');
    body.classList.remove('hidden');
    body.innerHTML = alerts.map(row => {
        const ref = colRef ? (row[colRef] ?? '—') : '—';
        const desc = colDesc ? (row[colDesc] ?? '') : '';
        const stock = row.__stock;
        const prioridad = getPrioridad(stock);
        const barColor = prioridad.label === 'INMEDIATA' ? '#dc2626' :
                         prioridad.label === 'ALTA' ? '#f97316' : '#9ca3af';
        const badgeBg = prioridad.label === 'INMEDIATA' ? '#fef2f2' :
                        prioridad.label === 'ALTA' ? '#fff7ed' : '#f3f4f6';
        const badgeText = prioridad.label === 'INMEDIATA' ? '#991b1b' :
                          prioridad.label === 'ALTA' ? '#9a3412' : '#4b5563';

        return `<div class="flex items-center gap-2 px-3 py-2 hover:bg-gray-50 transition-colors relative">
            <div class="w-1 self-stretch rounded-full flex-shrink-0" style="background:${barColor}"></div>
            <div class="flex-1 min-w-0 overflow-hidden">
                <div class="font-semibold text-primary text-sm truncate" title="${ref}">${ref}</div>
                ${desc ? `<div class="text-[10px] text-gray-400 truncate" title="${desc}">${desc}</div>` : ''}
            </div>
            <span class="font-data-mono text-xs font-semibold text-gray-700 flex-shrink-0 ml-auto">${stock}</span>
            <span class="px-2 py-0.5 rounded-full text-[10px] font-bold flex-shrink-0 whitespace-nowrap" style="background:${badgeBg};color:${badgeText}">${prioridad.label}</span>
        </div>`;
    }).join('');
}

// ─── TABLE ────────────────────────────────────────────────────
function renderTable(columns, colRef, colDesc, colStock) {
    const thead = document.getElementById('table-head');
    const tbody = document.getElementById('table-body');

    const displayCols = columns.filter(c => !c.startsWith('__'));
    thead.innerHTML = displayCols.map(c =>
        `<th class="px-md py-4 whitespace-nowrap">${c}</th>`
    ).join('') + `<th class="px-md py-4 text-center">Estado</th><th class="px-md py-4 text-right">Prioridad</th>`;

    tbody.innerHTML = filteredData.slice(0, 500).map(row => {
        const est = getEstadoColor(row.__estado);
        const pri = row.__prioridad;
        const cells = displayCols.map(c =>
            `<td class="px-md py-3 whitespace-nowrap">${row[c] ?? ''}</td>`
        ).join('');
        return `<tr class="hover:bg-surface-container transition-colors">
            ${cells}
            <td class="px-md py-3 text-center">
                <span class="px-3 py-1 rounded font-bold text-[11px]" style="background:${est.bg};color:${est.text}">${row.__estado}</span>
            </td>
            <td class="px-md py-3 text-right font-bold" style="color:${pri.color}">${pri.label}</td>
        </tr>`;
    }).join('') || '<tr><td colspan="99" class="text-center p-4 text-outline">Sin resultados</td></tr>';
}

// ─── BOTTOM ITEMS ─────────────────────────────────────────────
function renderBottomItems(columns, colRef, colStock) {
    const container = document.getElementById('bottom-items');
    const items = [...filteredData]
        .filter(r => !isNaN(r.__stock))
        .sort((a, b) => a.__stock - b.__stock)
        .slice(0, 20);

    container.innerHTML = items.map(row => {
        const ref = colRef ? (row[colRef] ?? '—') : '—';
        const s = row.__stock;
        const borderColor = s === 0 ? '#ba1a1a' : s <= 5 ? '#F59E0B' : s <= 20 ? '#facc15' : '#10b981';
        return `<div class="flex items-center justify-between p-3 bg-surface-container-low rounded border-l-4" style="border-color:${borderColor}">
            <span class="font-data-mono text-xs truncate mr-1">${ref}</span>
            <span class="font-bold text-xs" style="color:${borderColor}">${s}</span>
        </div>`;
    }).join('') || '<div class="col-span-4 text-center text-outline">Sin datos</div>';
}

// ─── REVISION TABLE ─────────────────────────────────────────
function renderRevisionTable(colRef, colColor, colDesc, colStock, colCierre, colAgotados, colEstado) {
    const cols = Object.keys(allData[0]);
    const colEsVmi = findColumn(COLUMN_ALIASES.es_vmi, cols) || findColumnLoose(COLUMN_ALIASES.es_vmi, cols);
    const colTalla = findColumn(COLUMN_ALIASES.talla, cols) || findColumnLoose(COLUMN_ALIASES.talla, cols);
    const colTendencia = findColumn(COLUMN_ALIASES.tendencia, cols) || findColumnLoose(COLUMN_ALIASES.tendencia, cols);
    const colCumplimiento = findColumn(COLUMN_ALIASES.cumplimiento, cols) || findColumnLoose(COLUMN_ALIASES.cumplimiento, cols);
    const colTendDia = findColumn(COLUMN_ALIASES.tend_dia, cols) || findColumnLoose(COLUMN_ALIASES.tend_dia, cols);
    const colPryC8 = cols.slice(15, 20).find(c => /c\d+/i.test(c));
    const colDiasInv = cols.find(c => /dias.*inventario|días.*inventario|inv.*dia/i.test(c));

    function parseNum(v) {
        const n = Number(String(v ?? '').replace(/[^0-9.\-]/g, '').replace(',', '.'));
        return isNaN(n) ? 0 : n;
    }

    // Filter only rows that belong to the 3 export groups
    const revRows = [];
    filteredData.forEach(r => {
        const tend = colTendencia ? parseNum(r[colTendencia]) : 0;
        const cierre = colCierre ? parseNum(r[colCierre]) : 0;
        let grupo = '';
        if (tend > 0 && cierre < 0) grupo = 'Tend+ Cierre-';
        else if (tend < 0 && cierre > 0) grupo = 'Tend- Cierre+';
        else if (tend < 0 && cierre < 0) grupo = 'Tend- Cierre-';
        if (grupo) revRows.push({ row: r, grupo });
    });

    const tbody = document.getElementById('rev-table-body');
    document.getElementById('rev-table-count').textContent = revRows.length + ' registros';

    tbody.innerHTML = revRows.map(({ row, grupo }) => {
        const ref = colRef ? (row[colRef] ?? '') : '';
        const color = colColor ? (row[colColor] ?? '') : '';
        const desc = colDesc ? (row[colDesc] ?? '') : '';
        const stock = !isNaN(row.__stock) ? row.__stock : '';
        const cierre = !isNaN(row.__cierre) ? row.__cierre : '';
        const agot = colAgotados ? (!isNaN(row.__agotados) ? row.__agotados : '') : (stock <= 0 ? Math.abs(stock) : 0);
        const estadoCode = colEstado ? (row[colEstado] ?? '') : '';
        const estadoFull = getEstadoFull(estadoCode);
        const esVmi = colEsVmi ? (row[colEsVmi] ?? '') : '';
        const talla = colTalla ? (row[colTalla] ?? '') : '';
        const tendencia = colTendencia ? (row[colTendencia] ?? '') : '';
        const cumplimiento = colCumplimiento ? (row[colCumplimiento] ?? '') : '';
        const tendDia = colTendDia ? (row[colTendDia] ?? '') : '';
        const pryC8 = colPryC8 ? (row[colPryC8] ?? '') : '';
        const diasInv = colDiasInv ? (row[colDiasInv] ?? '') : '';

        const estColor = getEstadoColor(estadoCode);
        const grupoColor = grupo === 'Tend+ Cierre-' ? '#ef4444' : grupo === 'Tend- Cierre+' ? '#f59e0b' : '#6b7280';

        return `<tr class="hover:bg-surface-container transition-colors">
            <td class="px-md py-3 whitespace-nowrap">${esVmi}</td>
            <td class="px-md py-3 whitespace-nowrap text-center">
                <span class="px-2 py-1 rounded font-bold text-[11px]" style="background:${estColor.bg};color:${estColor.text}">${estadoFull}</span>
            </td>
            <td class="px-md py-3 whitespace-nowrap font-semibold text-primary">${ref}</td>
            <td class="px-md py-3 whitespace-nowrap">${color}</td>
            <td class="px-md py-3 whitespace-nowrap">${talla}</td>
            <td class="px-md py-3 whitespace-nowrap">${desc}</td>
            <td class="px-md py-3 whitespace-nowrap">${tendencia}</td>
            <td class="px-md py-3 whitespace-nowrap">${pryC8}</td>
            <td class="px-md py-3 whitespace-nowrap">${cumplimiento}</td>
            <td class="px-md py-3 whitespace-nowrap text-right font-data-mono">${agot}</td>
            <td class="px-md py-3 whitespace-nowrap text-right font-data-mono">${stock}</td>
            <td class="px-md py-3 whitespace-nowrap text-right font-data-mono">${cierre}</td>
            <td class="px-md py-3 whitespace-nowrap text-right">${diasInv}</td>
            <td class="px-md py-3 whitespace-nowrap text-right">${tendDia}</td>
            <td class="px-md py-3 whitespace-nowrap text-center"><span class="px-2 py-0.5 rounded text-[10px] font-bold text-white" style="background:${grupoColor}">${grupo}</span></td>
        </tr>`;
    }).join('') || '<tr><td colspan="15" class="text-center p-4 text-outline">Sin resultados</td></tr>';
}

// ─── REVISION DASHBOARD ──────────────────────────────────────
function renderRevisionDashboard(colRef, colColor, colStock, colCierre, colAgotados, cols) {
    window.__revDashData = { colRef, colColor, colStock, colCierre, colAgotados, cols };
    const colTendencia = findColumn(COLUMN_ALIASES.tendencia, cols) || findColumnLoose(COLUMN_ALIASES.tendencia, cols);

    function parseNum(v) {
        const n = Number(String(v ?? '').replace(/[^0-9.\-]/g, '').replace(',', '.'));
        return isNaN(n) ? 0 : n;
    }

    const groups = { tpc: [], tcp: [], tcc: [] };
    filteredData.forEach(r => {
        const tend = colTendencia ? parseNum(r[colTendencia]) : 0;
        const cierre = colCierre ? parseNum(r[colCierre]) : 0;
        if (tend > 0 && cierre < 0) groups.tpc.push(r);
        else if (tend < 0 && cierre > 0) groups.tcp.push(r);
        else if (tend < 0 && cierre < 0) groups.tcc.push(r);
    });

    function sumAgot(rows) {
        return rows.reduce((s, r) => {
            const v = !isNaN(r.__agotados) ? r.__agotados : (r.__stock <= 0 ? Math.abs(r.__stock) : 0);
            return s + (isNaN(v) ? 0 : v);
        }, 0);
    }

    function sumCierre(rows) {
        return rows.reduce((s, r) => s + (isNaN(r.__cierre) ? 0 : r.__cierre), 0);
    }

    const kpis = {
        tpc: { count: groups.tpc.length, agot: sumAgot(groups.tpc), cierre: sumCierre(groups.tpc) },
        tcp: { count: groups.tcp.length, agot: sumAgot(groups.tcp), cierre: sumCierre(groups.tcp) },
        tcc: { count: groups.tcc.length, agot: sumAgot(groups.tcc), cierre: sumCierre(groups.tcc) },
    };

    Object.keys(kpis).forEach(k => {
        document.getElementById(`rev-kpi-${k}-count`).textContent = kpis[k].count;
        document.getElementById(`rev-kpi-${k}-agot`).textContent = Math.round(kpis[k].agot).toLocaleString();
        document.getElementById(`rev-kpi-${k}-cierre`).textContent = Math.round(kpis[k].cierre).toLocaleString();
    });

    // Draw cierre comparison chart
    const ctxC = document.getElementById('rev-chart-cierre').getContext('2d');
    const cierreVals = [kpis.tpc.cierre, kpis.tcp.cierre, kpis.tcc.cierre];
    drawRevBarChart(ctxC, ['Tend+\nCierre-', 'Tend-\nCierre+', 'Tend-\nCierre-'], cierreVals, '#ba1a1a');

    // Draw agotados comparison chart
    const ctxA = document.getElementById('rev-chart-agot').getContext('2d');
    const agotVals = [kpis.tpc.agot, kpis.tcp.agot, kpis.tcc.agot];
    drawRevBarChart(ctxA, ['Tend+\nCierre-', 'Tend-\nCierre+', 'Tend-\nCierre-'], agotVals, '#006399');
}

function drawRevBarChart(ctx, labels, values, color) {
    const rect = ctx.canvas.parentElement.getBoundingClientRect();
    const W = rect.width || 400;
    const H = rect.height || 250;
    ctx.canvas.width = W * 2;
    ctx.canvas.height = H * 2;
    ctx.scale(2, 2);
    ctx.clearRect(0, 0, W, H);

    const max = Math.max(...values.map(Math.abs), 1);
    const pad = { top: 20, bottom: 60, left: 10, right: 10 };
    const chartW = W - pad.left - pad.right;
    const chartH = H - pad.top - pad.bottom;
    const gap = chartW / labels.length;

    ctx.font = '11px IBM Plex Sans, sans-serif';
    ctx.textAlign = 'center';

    const barW = Math.max(Math.min(gap * 0.6, 100), 20);

    values.forEach((v, i) => {
        const x = pad.left + gap * i + gap / 2;
        const barH = chartH * Math.abs(v) / max;
        const y = v >= 0 ? pad.top + chartH - barH : pad.top + chartH;

        ctx.fillStyle = v < 0 ? '#ef4444' : '#22c55e';
        ctx.fillRect(x - barW / 2, y, barW, Math.max(barH, 1));

        ctx.fillStyle = '#0b1c30';
        ctx.font = 'bold 11px IBM Plex Sans, sans-serif';
        ctx.textBaseline = 'bottom';
        ctx.fillText(Math.round(v).toLocaleString(), x, y - 4);

        ctx.fillStyle = '#737780';
        ctx.font = '10px IBM Plex Sans, sans-serif';
        ctx.textBaseline = 'top';
        ctx.fillText(labels[i], x, pad.top + chartH + 8);
    });

    ctx.strokeStyle = '#d1d5db';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(pad.left, pad.top + chartH);
    ctx.lineTo(W - pad.right, pad.top + chartH);
    ctx.stroke();
}
function renderCharts(columns, colCategoria, colStock, colColor) {
    const ctxCierre = document.getElementById('chartCierre').getContext('2d');
    const ctxBar = document.getElementById('chartBars').getContext('2d');

    // ── Horizontal bar chart: Cierre by Ref+Color (rows with "Total" in COLOR) ──
    const _colCierre = findColumn(COLUMN_ALIASES.cierre, columns) || findColumnLoose(COLUMN_ALIASES.cierre, columns);
    const _colColor = findColumn(COLUMN_ALIASES.color, columns) || findColumnLoose(COLUMN_ALIASES.color, columns);
    const _colRef = findColumn(COLUMN_ALIASES.referencia, columns) || findColumnLoose(COLUMN_ALIASES.referencia, columns);

    let cierreRows = [];
    if (_colCierre && _colColor && _colRef) {
        cierreRows = filteredData
            .filter(r => {
                const colorVal = String(r[_colColor] ?? '');
                return /total/i.test(colorVal);
            })
            .map(r => {
                const ref = r[_colRef] ?? '';
                const color = String(r[_colColor] ?? '').replace(/^Total\s*/i, '');
                const val = Number(String(r[_colCierre] ?? '').replace(/[^0-9.\-]/g, ''));
                return { label: ref + ' / ' + color, value: isNaN(val) ? 0 : val };
            })
            .filter(r => r.value !== 0)
            .sort((a, b) => Math.abs(b.value) - Math.abs(a.value));
    }

    drawCierreChart(ctxCierre, cierreRows);
    if (!ctxCierre.canvas._tooltipSetup) {
        setupCierreTooltip(ctxCierre.canvas);
        ctxCierre.canvas._tooltipSetup = true;
    }

    // ── Bars: columns 16-20 (c9, c10, c11, …) ──
    const knownKeys = new Set(['__stock', '__agotados', '__estado', '__prioridad', '__cierre']);
    let cCols = columns.filter(c => {
        const norm = c.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        return /c\d+/.test(norm) && !knownKeys.has(c);
    });
    cCols.sort((a, b) => parseInt(String(a).match(/\d+/)) - parseInt(String(b).match(/\d+/)));
    // Keep only columns that exist in positions 16-20 (1-indexed) of original columns array
    const colRange = columns.slice(15, 20);
    cCols = cCols.filter(c => colRange.includes(c));
    const barLabels = cCols.map(c => 'PRY ' + c.toUpperCase());
    const barValues = cCols.map(col => {
        return filteredData.reduce((sum, row) => {
            const v = Number(String(row[col] ?? '').replace(/[^0-9.,\-]/g, '').replace(',', '.'));
            return sum + (isNaN(v) ? 0 : v);
        }, 0);
    });

    __barChartData = { ctx: ctxBar, labels: barLabels, values: barValues };

    const barCanvas = ctxBar.canvas;
    if (barCanvas._resizeObserver) barCanvas._resizeObserver.disconnect();
    const ro = new ResizeObserver(() => drawBarChart(ctxBar, barLabels, barValues));
    ro.observe(barCanvas.parentElement);
    barCanvas._resizeObserver = ro;
    drawBarChart(ctxBar, barLabels, barValues);

    // ── Cierre chart resize ──
    const cierreCanvas = ctxCierre.canvas;
    if (cierreCanvas._resizeObserver) cierreCanvas._resizeObserver.disconnect();
    const ro2 = new ResizeObserver(() => drawCierreChart(ctxCierre, cierreRows));
    ro2.observe(cierreCanvas.parentElement);
    cierreCanvas._resizeObserver = ro2;

}

// ─── INSIGHTS ─────────────────────────────────────────────────
function updateInsights(columns, colCategoria, colStock) {
    const agotados = filteredData.filter(r => r.__stock === 0).length;
    const total = filteredData.filter(r => !isNaN(r.__stock)).length;
    const tasa = total ? ((agotados / total) * 100) : 0;

    const insightEl = document.getElementById('insight-text');
    const recoEl = document.getElementById('recomendacion-text');

    let topCat = '';
    if (colCategoria && filteredData.length) {
        const catAgotados = {};
        filteredData.forEach(r => {
            if (r.__stock === 0) {
                const c = r[colCategoria] ?? 'Sin categoría';
                catAgotados[c] = (catAgotados[c] || 0) + 1;
            }
        });
        const entries = Object.entries(catAgotados).sort((a, b) => b[1] - a[1]);
        if (entries.length) topCat = entries[0][0];
    }

    if (total === 0) {
        insightEl.textContent = 'Cargue un archivo de datos para ver análisis inteligente.';
        recoEl.textContent = 'Los datos mostrarán recomendaciones automáticas.';
    } else {
        insightEl.textContent = tasa > 5
            ? `Se detecta una tasa de agotados del ${tasa.toFixed(1)}%. ${topCat ? `La categoría "${topCat}" es la más afectada.` : ''}`
            : `La tasa de agotados es del ${tasa.toFixed(1)}%. El inventario se mantiene estable.`;
        recoEl.textContent = agotados > 0
            ? `Priorizar reposición de ${agotados} referencias agotadas para evitar pérdida de ventas.`
            : `No hay referencias agotadas. Mantener el monitoreo constante.`;
    }
}

// ─── HORIZONTAL BAR: Cierre by Ref+Color ─────────────────────
function drawCierreChart(ctx, data) {
    const dpr = window.devicePixelRatio || 1;
    const rect = ctx.canvas.getBoundingClientRect();
    const W = rect.width;
    const H = rect.height;
    if (!W || !H) return;
    ctx.canvas.width = W * dpr;
    ctx.canvas.height = H * dpr;
    ctx.scale(dpr, dpr);
    ctx.clearRect(0, 0, W, H);

    if (!data.length) {
        ctx.fillStyle = '#9ca3af';
        ctx.font = '14px IBM Plex Sans, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('Sin datos de cierre con Total', W / 2, H / 2);
        return;
    }

    const rowH = Math.max(48, Math.min(56, Math.floor((H - 80) / data.length)));
    const padding = { top: 12, right: 90, bottom: 32, left: 190 };
    const chartH = data.length * rowH;
    const chartW = W - padding.left - padding.right;
    const needH = padding.top + chartH + padding.bottom;

    const values = data.map(d => d.value);
    const maxAbs = Math.max(...values.map(Math.abs), 1);
    const tickCount = 4;
    const mag = Math.pow(10, Math.max(0, Math.floor(Math.log10(maxAbs))));
    const tickStep = Math.max(Math.ceil(maxAbs / tickCount / mag) * mag, mag);
    const axisMax = Math.max(tickStep * tickCount, maxAbs);
    const scaleX = chartW / (axisMax * 2);
    const zeroX = padding.left + chartW / 2;

    const bars = [];
    const barColorNeg = '#E57373';
    const barColorPos = '#66BB6A';
    const barHoverNeg = '#EF5350';
    const barHoverPos = '#43A047';

    // Horizontal grid lines
    ctx.strokeStyle = '#f3f4f6';
    ctx.lineWidth = 1;
    for (let i = 0; i <= data.length; i++) {
        const y = padding.top + i * rowH;
        ctx.beginPath();
        ctx.moveTo(padding.left, y);
        ctx.lineTo(padding.left + chartW, y);
        ctx.stroke();
    }

    // X-axis ticks
    ctx.fillStyle = '#9ca3af';
    ctx.font = '400 11px IBM Plex Mono, monospace';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    for (let i = -tickCount; i <= tickCount; i++) {
        const x = zeroX + i * tickStep * scaleX;
        if (x < padding.left - 5 || x > padding.left + chartW + 5) continue;
        ctx.beginPath();
        ctx.moveTo(x, padding.top + chartH);
        ctx.lineTo(x, padding.top + chartH + 6);
        ctx.strokeStyle = '#d1d5db';
        ctx.stroke();
        const label = Math.abs(i * tickStep);
        if (label > 0) ctx.fillText(label.toLocaleString(), x, padding.top + chartH + 8);
    }
    ctx.fillStyle = '#6b7280';
    ctx.font = '500 11px IBM Plex Mono, monospace';
    ctx.fillText('0', zeroX, padding.top + chartH + 8);

    // Center line
    ctx.strokeStyle = '#cbd5e1';
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(zeroX, padding.top);
    ctx.lineTo(zeroX, padding.top + chartH);
    ctx.stroke();

    const hoverIdx = ctx.canvas._cierreHoverIdx;

    for (let i = 0; i < data.length; i++) {
        const y = padding.top + i * rowH;
        const val = data[i].value;
        const barW = Math.abs(val) * scaleX;
        const isNeg = val < 0;
        const barX = isNeg ? zeroX - barW : zeroX;
        const isHover = i === hoverIdx;
        const barColor = isNeg ? (isHover ? barHoverNeg : barColorNeg) : (isHover ? barHoverPos : barColorPos);

        // Zebra
        if (i % 2 === 0) {
            ctx.fillStyle = '#fafafa';
            ctx.fillRect(padding.left, y, chartW, rowH);
        }

        // Hover glow
        if (isHover) {
            ctx.fillStyle = 'rgba(59,130,246,0.06)';
            ctx.fillRect(padding.left, y, chartW, rowH);
        }

        // Bar with rounded ends
        const barY = y + 8;
        const barH = rowH - 16;
        const bW = Math.max(barW, 4);
        ctx.fillStyle = barColor;

        if (isNeg) {
            const r = Math.min(6, bW / 2, barH / 2);
            ctx.beginPath();
            ctx.moveTo(zeroX, barY);
            ctx.lineTo(barX + r, barY);
            ctx.quadraticCurveTo(barX, barY, barX, barY + r);
            ctx.lineTo(barX, barY + barH - r);
            ctx.quadraticCurveTo(barX, barY + barH, barX + r, barY + barH);
            ctx.lineTo(zeroX, barY + barH);
            ctx.closePath();
            ctx.fill();
        } else {
            const r = Math.min(6, bW / 2, barH / 2);
            ctx.beginPath();
            ctx.moveTo(zeroX, barY);
            ctx.lineTo(barX + bW - r, barY);
            ctx.quadraticCurveTo(barX + bW, barY, barX + bW, barY + r);
            ctx.lineTo(barX + bW, barY + barH - r);
            ctx.quadraticCurveTo(barX + bW, barY + barH, barX + bW - r, barY + barH);
            ctx.lineTo(zeroX, barY + barH);
            ctx.closePath();
            ctx.fill();
        }

        bars.push({ x: barX, y: barY, w: bW, h: barH, idx: i, val, isNeg });

        // Label (Ref / Color) — right-aligned
        ctx.fillStyle = isHover ? '#111827' : '#374151';
        ctx.font = isHover ? '600 12px IBM Plex Sans, sans-serif' : '500 11px IBM Plex Sans, sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        const labelMaxW = padding.left - 16;
        let label = data[i].label;
        while (ctx.measureText(label).width > labelMaxW && label.length > 6) {
            label = label.slice(0, -4) + '…';
        }
        ctx.fillText(label, padding.left - 12, y + rowH / 2);

        // Value at bar end
        ctx.fillStyle = isHover ? '#111827' : '#374151';
        ctx.font = isHover ? '600 12px IBM Plex Mono, monospace' : '500 11px IBM Plex Mono, monospace';
        ctx.textBaseline = 'middle';
        const valStr = (isNeg ? '' : '+') + val.toLocaleString();
        if (isNeg) {
            ctx.textAlign = 'right';
            ctx.fillText(valStr, zeroX - bW - 8, y + rowH / 2);
        } else {
            ctx.textAlign = 'left';
            ctx.fillText(valStr, zeroX + bW + 8, y + rowH / 2);
        }
    }

    ctx.canvas._cierreData = data;
    ctx.canvas._cierreBars = bars;
    ctx.canvas._cierrePadding = padding;
    ctx.canvas._cierreRowH = rowH;
    ctx.canvas._cierreZeroX = zeroX;
    ctx.canvas._cierreScaleX = scaleX;
    ctx.canvas._cierreAxisMax = axisMax;
}

function setupCierreTooltip(canvas) {
    const tip = document.getElementById('chartCierreTooltip');
    if (!tip) return;

    function hide() { tip.classList.add('hidden'); }

    function getHoverIndex(e) {
        const rect2 = canvas.getBoundingClientRect();
        const mx = e.clientX - rect2.left;
        const my = e.clientY - rect2.top;
        const bars = canvas._cierreBars;
        if (!bars) return -1;
        for (const b of bars) {
            if (mx >= b.x && mx <= b.x + b.w && my >= b.y && my <= b.y + b.h) return b.idx;
        }
        return -1;
    }

    let rafId = null;

    canvas.addEventListener('mousemove', (e) => {
        if (rafId) cancelAnimationFrame(rafId);
        rafId = requestAnimationFrame(() => {
            const idx = getHoverIndex(e);
            const prev = canvas._cierreHoverIdx;
            if (idx !== prev) {
                canvas._cierreHoverIdx = idx;
                const data = canvas._cierreData;
                if (data) {
                    const ctx = canvas.getContext('2d');
                    drawCierreChart(ctx, data);
                }
            }
            if (idx >= 0) {
                const d = canvas._cierreData[idx];
                tip.textContent = `${d.label}  |  ${d.value < 0 ? '' : '+'}${d.value.toLocaleString()}`;
                tip.classList.remove('hidden');
                tip.style.left = (e.clientX + 14) + 'px';
                tip.style.top = (e.clientY - 24) + 'px';
            } else {
                hide();
            }
            rafId = null;
        });
    });

    canvas.addEventListener('mouseleave', () => {
        if (rafId) cancelAnimationFrame(rafId);
        rafId = null;
        canvas._cierreHoverIdx = -1;
        const data = canvas._cierreData;
        if (data) {
            const ctx = canvas.getContext('2d');
            drawCierreChart(ctx, data);
        }
        hide();
    });
}

// ─── CANVAS BAR CHART ─────────────────────────────────────────
function drawBarChart(ctx, labels, values) {
    const dpr = window.devicePixelRatio || 1;
    const rect = ctx.canvas.getBoundingClientRect();
    const W = rect.width;
    const H = rect.height;
    if (!W || !H) return;
    ctx.canvas.width = W * dpr;
    ctx.canvas.height = H * dpr;
    ctx.scale(dpr, dpr);

    ctx.clearRect(0, 0, W, H);

    if (!labels.length) {
        ctx.fillStyle = '#9ca3af';
        ctx.font = '13px IBM Plex Sans, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('Sin datos de campañas', W / 2, H / 2);
        return;
    }

    const padding = { top: 28, right: 16, bottom: 56, left: 52 };
    const chartW = W - padding.left - padding.right;
    const chartH = H - padding.top - padding.bottom;

    const maxVal = Math.max(...values, 1);
    const tickCount = 5;
    const mag = Math.pow(10, Math.max(0, Math.floor(Math.log10(maxVal))));
    const maxTick = Math.ceil(maxVal / mag) * mag;
    const tickStep = Math.max(Math.ceil(maxTick / tickCount), 1);
    const yMax = tickStep * tickCount;

    // Y-axis ticks & grid
    ctx.strokeStyle = '#e5e7eb';
    ctx.lineWidth = 1;
    for (let i = 0; i <= tickCount; i++) {
        const yVal = i * tickStep;
        const y = padding.top + chartH - (yVal / yMax) * chartH;
        ctx.beginPath();
        ctx.moveTo(padding.left, y);
        ctx.lineTo(W - padding.right, y);
        ctx.stroke();
        ctx.fillStyle = '#6b7280';
        ctx.font = '400 10px IBM Plex Mono, monospace';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(yVal.toLocaleString(), padding.left - 8, y);
    }

    const barCount = labels.length;
    const barGap = Math.max(4, chartW * 0.08 / barCount);
    const barWidth = Math.min((chartW - barGap * (barCount + 1)) / barCount, 80);
    const colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];

    // Check if X labels overlap — if so, rotate
    ctx.font = '500 11px IBM Plex Sans, sans-serif';
    const maxLabelW = barWidth + barGap;
    let needsRotate = false;
    for (let i = 0; i < barCount; i++) {
        if (ctx.measureText(labels[i]).width > maxLabelW) { needsRotate = true; break; }
    }

    for (let i = 0; i < barCount; i++) {
        const x = padding.left + barGap + i * (barWidth + barGap);
        const barH = (values[i] / yMax) * chartH;
        const y = padding.top + chartH - barH;

        // Bar
        ctx.fillStyle = colors[i % colors.length];
        ctx.fillRect(x, y, barWidth, barH);

        // Value on top — 4px gap
        ctx.fillStyle = '#111827';
        ctx.font = '700 18px IBM Plex Mono, monospace';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(values[i].toLocaleString(), x + barWidth / 2, y - 4);

        // X label
        const labelX = x + barWidth / 2;
        const labelY = padding.top + chartH + 8;
        ctx.fillStyle = '#374151';
        ctx.font = '500 11px IBM Plex Sans, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        if (needsRotate) {
            ctx.save();
            ctx.translate(labelX, labelY);
            ctx.rotate(-Math.PI / 4);
            ctx.fillText(labels[i], 0, 0);
            ctx.restore();
        } else {
            ctx.fillText(labels[i], labelX, labelY);
        }
    }
}

function redrawBarChart() {
    if (__barChartData) {
        drawBarChart(__barChartData.ctx, __barChartData.labels, __barChartData.values);
    }
}

function redrawCierreChart() {
    const canvas = document.getElementById('chartCierre');
    if (canvas && canvas._cierreData) {
        canvas._cierreHoverIdx = -1;
        const ctx = canvas.getContext('2d');
        drawCierreChart(ctx, canvas._cierreData);
    }
}

// ─── EXPORT BUTTONS ────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('btn-export-excel').addEventListener('click', () => {
        if (!filteredData.length) { alert('No hay datos para exportar.'); return; }

        const cols = Object.keys(allData[0]);
        const colCierre = findColumn(COLUMN_ALIASES.cierre, cols) || findColumnLoose(COLUMN_ALIASES.cierre, cols);
        const colTendencia = findColumn(COLUMN_ALIASES.tendencia, cols) || findColumnLoose(COLUMN_ALIASES.tendencia, cols);
        const colRef = findColumn(COLUMN_ALIASES.referencia, cols) || findColumnLoose(COLUMN_ALIASES.referencia, cols);
        const colColor = findColumn(COLUMN_ALIASES.color, cols) || findColumnLoose(COLUMN_ALIASES.color, cols);
        const colStock = findColumn(COLUMN_ALIASES.stock, cols) || findColumnLoose(COLUMN_ALIASES.stock, cols);
        const colAgot = findColumn(COLUMN_ALIASES.agotados, cols) || findColumnLoose(COLUMN_ALIASES.agotados, cols);
        const colDesc = findColumn(COLUMN_ALIASES.descripcion, cols) || findColumnLoose(COLUMN_ALIASES.descripcion, cols);
        const colEsVmi = findColumn(COLUMN_ALIASES.es_vmi, cols) || findColumnLoose(COLUMN_ALIASES.es_vmi, cols);
        const colTalla = findColumn(COLUMN_ALIASES.talla, cols) || findColumnLoose(COLUMN_ALIASES.talla, cols);
        const colCumplimiento = findColumn(COLUMN_ALIASES.cumplimiento, cols) || findColumnLoose(COLUMN_ALIASES.cumplimiento, cols);
        const colTendDia = findColumn(COLUMN_ALIASES.tend_dia, cols) || findColumnLoose(COLUMN_ALIASES.tend_dia, cols);
        const colMap = {};
        const setCol = (userName, detected) => { if (detected) colMap[userName] = detected; };
        setCol('ES VMI', colEsVmi || cols.find(c => /es vmi|vmi|tipo vmi/i.test(c)));
        setCol('ESTADO', cols.find(c => /estado|status/i.test(c)));
        setCol('REF', colRef);
        setCol('COLOR', colColor);
        setCol('TALLA', colTalla);
        setCol('DESCRIPCION', colDesc);
        setCol('TENDENCIA', colTendencia);
        // PRY C8: first c* column in positions 16-20
        setCol('PRY C8', cols.slice(15, 20).find(c => /c\d+/i.test(c)));
        setCol('CUMPLIMIENTO', colCumplimiento);
        setCol('AGOT', colAgot);
        setCol('INVENTARIO', colStock);
        setCol('CIERRE', colCierre);
        setCol('DIAS INVENTARIO', cols.find(c => /dias.*inventario|días.*inventario|inv.*dia/i.test(c)));
        setCol('TEND DIA', colTendDia);

        function parseNum(v) {
            const n = Number(String(v ?? '').replace(/[^0-9.,\-]/g, '').replace(',', '.'));
            return isNaN(n) ? 0 : n;
        }

        function buildTable(rows, label) {
            const keys = Object.keys(colMap);
            const h = keys.map(c => `<th>${c}</th>`).join('');
            const b = rows.map(r => {
                const cells = keys.map(c => {
                    const raw = r[colMap[c]] ?? '';
                    const isCierre = c === 'CIERRE';
                    return isCierre ? `<td style="color:red">${raw}</td>` : `<td>${raw}</td>`;
                }).join('');
                return `<tr>${cells}</tr>`;
            }).join('');
            return `<h2>${label}</h2><table border="1" cellpadding="4" cellspacing="0">${h}${b}</table><br>`;
        }

        const tPosCNeg = [], tNegCPos = [], tNegCNeg = [];
        filteredData.forEach(r => {
            const tend = colTendencia ? parseNum(r[colTendencia]) : 0;
            const cierre = colCierre ? parseNum(r[colCierre]) : 0;
            if (tend > 0 && cierre < 0) tPosCNeg.push(r);
            else if (tend < 0 && cierre > 0) tNegCPos.push(r);
            else if (tend < 0 && cierre < 0) tNegCNeg.push(r);
        });

        let html = '<html><meta charset="utf-8"><body style="font-family:sans-serif">';
        if (tPosCNeg.length) html += buildTable(tPosCNeg, 'Tend+ Cierre-');
        if (tNegCPos.length) html += buildTable(tNegCPos, 'Tend- Cierre+');
        if (tNegCNeg.length) html += buildTable(tNegCNeg, 'Tend- Cierre-');
        if (tPosCNeg.length + tNegCPos.length + tNegCNeg.length === 0) {
            html += buildTable(filteredData, 'Datos');
        }
        html += '</body></html>';

        const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'inventario_agotados_campana.xls';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });

    document.getElementById('btn-export-pdf').addEventListener('click', () => {
        if (!filteredData.length) { alert('No hay datos para exportar.'); return; }
        window.print();
    });
});
