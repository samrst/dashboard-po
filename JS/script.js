let rawData = [], filteredData = [], charts = {};
const $ = id => document.getElementById(id);
const filters = ['unidade', 'curso', 'modalidade', 'categoria'];
const keysMap = { unidade: 'Unidade operacional (Escola)', curso: 'Curso', modalidade: 'Modalidade', categoria: 'Categoria' };

// Event Listeners
if ($('file-input')) $('file-input').onchange = handleFile;
filters.forEach(f => { if ($(`filter-${f}`)) $(`filter-${f}`).onchange = applyFilters; });

const pick = (d, keys) => {
    for (const k of keys) if (d[k] != null && d[k] !== '') {
        const v = String(d[k]).trim();
        if (v.startsWith('#')) return 0;
        return parseInt(d[k]) || 0;
    }
    return 0;
};

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    if ($('loading-overlay')) $('loading-overlay').querySelector('p').textContent = 'Sincronizando...';

    const reader = new FileReader();
    reader.onload = ev => {
        try {
            const workbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
            const sheetName = workbook.SheetNames.find(n => n.includes('Worksheet')) || workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Ler como array de arrays (header: 1)
            const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            // Encontrar a linha de cabeçalhos (contém "Unidade operacional (Escola)")
            let headerRowIdx = -1;
            for (let i = 0; i < Math.min(allRows.length, 5); i++) {
                const row = allRows[i];
                if (row && row.some(cell => cell && String(cell).trim() === 'Unidade operacional (Escola)')) {
                    headerRowIdx = i;
                    break;
                }
            }

            if (headerRowIdx === -1) {
                throw new Error('Não foi possível encontrar a linha de cabeçalhos na planilha. Verifique se a planilha contém a coluna "Unidade operacional (Escola)".');
            }

            const normalizeKey = value => String(value ?? '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .replace(/\u00A0/g, ' ')
                .replace(/[\r\n]+/g, ' ')
                .replace(/\s*\/\s*/g, '/')
                .replace(/\s+/g, ' ')
                .trim()
                .toLowerCase();

            const headers = allRows[headerRowIdx];
            const colIndex = {};
            const normColIndex = {};
            headers.forEach((h, i) => {
                if (h != null) {
                    const rawStr = String(h).trim();
                    colIndex[rawStr] = i;
                    normColIndex[normalizeKey(h)] = i;
                }
            });

            const findCol = candidates => {
                for (const c of candidates) {
                    if (colIndex[c] !== undefined) return colIndex[c];
                    const n = normalizeKey(c);
                    if (normColIndex[n] !== undefined) return normColIndex[n];
                }
                return undefined;
            };

            // Mapeamento dos índices das colunas oficiais
            const idx = {
                unidade: findCol(['Unidade operacional (Escola)', 'Unidade operacional', 'Unidade', 'Escola']),
                curso: findCol(['Curso', 'Nome do Curso']),
                modalidade: findCol(['Modalidade']),
                categoria: findCol(['Categoria']),
                alunos: findCol(['Alunos', 'Total de Alunos', 'Total Alunos']),
                aptos: findCol([
                    'Alunos Aptos/Agendados',
                    'Alunos Aptos / Agendados',
                    'Aptos/Agendados',
                    'Aptos / Agendados',
                    'Alunos Aptos',
                    'Aptos'
                ]),
                inaptos: findCol([
                    'Alunos Inaptos/Não Agendados',
                    'Alunos Inaptos / Não Agendados',
                    'Alunos Inaptos/Nao Agendados',
                    'Alunos Inaptos / Nao Agendados',
                    'Inaptos/Não Agendados',
                    'Inaptos / Não Agendados',
                    'Inaptos/Nao Agendados',
                    'Inaptos / Nao Agendados',
                    'Alunos Inaptos',
                    'Inaptos'
                ]),
                presentes: findCol(['Presentes', 'Presente']),
                ausentes: findCol(['Ausentes', 'Ausente'])
            };

            // Validar colunas obrigatórias
            const missingCols = [];
            if (idx.unidade === undefined) missingCols.push('Unidade operacional (Escola)');
            if (idx.aptos === undefined) missingCols.push('Alunos Aptos/Agendados');
            if (idx.inaptos === undefined) missingCols.push('Alunos Inaptos/Não Agendados');

            if (missingCols.length > 0) {
                throw new Error(`Não foi possível processar a planilha.\n\nColuna(s) obrigatória(s) não encontrada(s):\n- ${missingCols.join('\n- ')}`);
            }

            const parseCellNum = value => {
                if (value === null || value === undefined) return 0;
                if (typeof value === 'number') return Math.round(value) || 0;
                const text = String(value).trim();
                if (text === '' || text === '-' || text.startsWith('#')) return 0;
                const normalized = text.replace(/\./g, '').replace(',', '.');
                const num = parseInt(normalized, 10);
                return Number.isFinite(num) ? num : 0;
            };

            // Processar as linhas de dados (pular cabeçalhos)
            rawData = [];
            for (let i = headerRowIdx + 1; i < allRows.length; i++) {
                const row = allRows[i];
                if (!row) continue;
                const unidade = row[idx.unidade];
                // Pular linhas vazias ou de total
                if (!unidade || String(unidade).trim() === '' || String(unidade).toLowerCase() === 'total') continue;

                rawData.push({
                    'Unidade operacional (Escola)': unidade,
                    'Curso': (idx.curso !== undefined && row[idx.curso] != null) ? String(row[idx.curso]).trim() : '-',
                    'Modalidade': (idx.modalidade !== undefined && row[idx.modalidade] != null) ? String(row[idx.modalidade]).trim() : '-',
                    'Categoria': (idx.categoria !== undefined && row[idx.categoria] != null && String(row[idx.categoria]).trim() !== '') ? String(row[idx.categoria]).trim() : '-',
                    '_total_alunos': parseCellNum(row[idx.alunos]),
                    '_alunos_aptos': parseCellNum(row[idx.aptos]),
                    '_alunos_inaptos': parseCellNum(row[idx.inaptos]),
                    '_presentes': parseCellNum(row[idx.presentes]),
                    '_ausentes': parseCellNum(row[idx.ausentes])
                });
            }

            if (rawData.length === 0) {
                throw new Error('Nenhum dado encontrado na planilha. Verifique se há dados a partir da linha abaixo dos cabeçalhos.');
            }

            initDashboard();
        } catch (err) {
            alert('Erro ao importar: ' + err.message);
        }
        finally { e.target.value = ''; }
    };
    reader.readAsArrayBuffer(file);
}

function initDashboard() {
    if ($('loading-overlay')) $('loading-overlay').style.display = 'none';
    if ($('main-content')) $('main-content').style.display = 'block';
    populateFilters();
    applyFilters();
}

function populateFilters() {
    filters.forEach(f => {
        const key = keysMap[f], sel = $(`filter-${f}`);
        if (!sel) return;
        const vals = [...new Set(rawData.map(d => d[key]))].filter(Boolean).sort();
        const current = sel.value;
        sel.innerHTML = '<option value="ALL">Todas</option>' + vals.map(v => `<option value="${v}">${v}</option>`).join('');
        if (vals.includes(current)) sel.value = current;
    });
}

function applyFilters() {
    const vals = Object.fromEntries(filters.map(f => [f, $(`filter-${f}`)?.value || 'ALL']));
    filteredData = rawData.filter(d => filters.every(f => vals[f] === 'ALL' || d[keysMap[f]] === vals[f]));
    updateKPIs(); updateCharts(); updateTable();
}

function updateKPIs() {
    const sum = k => filteredData.reduce((a, b) => a + (Number(b[k]) || 0), 0);
    const kpis = {
        objetivo: '_total_alunos',
        homologados: '_alunos_aptos',
        total: '_alunos_inaptos',
        pendentes: '_presentes',
        confirmadas: '_ausentes'
    };
    const vals = Object.fromEntries(Object.entries(kpis).map(([k, v]) => [k, sum(v)]));

    Object.entries(vals).forEach(([k, v]) => { if ($(`kpi-${k}`)) $(`kpi-${k}`).textContent = v.toLocaleString('pt-BR'); });

    const pcts = {
        'homologados-pct': [vals.homologados, vals.objetivo, 'dos Alunos'],
        'total-pct': [vals.total, vals.objetivo, 'dos Alunos'],
        'pendentes-pct': [vals.pendentes, vals.homologados, 'dos Aptos/Agendados'],
        'confirmadas-pct': [vals.confirmadas, vals.homologados, 'dos Aptos/Agendados']
    };
    Object.entries(pcts).forEach(([id, [v, total, msg]]) => {
        if ($(`kpi-${id}`)) $(`kpi-${id}`).textContent = (total ? Math.round(v / total * 100) : 0) + '% ' + msg;
    });
}

function aggregate(key, metrics) {
    return filteredData.reduce((acc, d) => {
        const k = d[key] || 'S/I';
        if (!acc[k]) acc[k] = Object.fromEntries(metrics.map(m => [m, 0]));
        metrics.forEach(m => acc[k][m] += Number(d[m]) || 0);
        return acc;
    }, {});
}

function updateCharts() {
    if (typeof Chart === 'undefined') return;
    const isAll = ($('filter-unidade')?.value === 'ALL');

    // 1. Alunos por Escola
    const dAlunos = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_total_alunos', '_alunos_aptos']);
    renderBar('chart-alunos-escola', Object.keys(dAlunos), [
        { label: 'Total de Alunos', data: Object.values(dAlunos).map(v => v._total_alunos), color: '#005599' },
        { label: 'Alunos Aptos/Agendados', data: Object.values(dAlunos).map(v => v._alunos_aptos), color: '#94a3b8' }
    ], 'x', true);

    // 2 & 3. Aptos/Agendados vs Inaptos/Não Agendados (Seção 2)
    const dApp = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_alunos_aptos', '_alunos_inaptos']);
    renderBar('chart-aplicacao', Object.keys(dApp), [
        { label: 'Aptos/Agendados', data: Object.values(dApp).map(v => v._alunos_aptos), color: '#005599' },
        { label: 'Inaptos/Não Agendados', data: Object.values(dApp).map(v => v._alunos_inaptos), color: '#94a3b8' }
    ], 'x', true);

    const sumMetric = m => filteredData.reduce((a, b) => a + (Number(b[m]) || 0), 0);
    renderPie('chart-aplicacao-pizza', ['Aptos/Agendados', 'Inaptos/Não Agendados'], [sumMetric('_alunos_aptos'), sumMetric('_alunos_inaptos')], ['#005599', '#94a3b8']);

    // 4 & 5. Presentes vs Ausentes (Seção 3)
    const dTab = aggregate(keysMap.curso, ['_presentes', '_ausentes']);
    renderBar('chart-tabulacao', Object.keys(dTab).map(k => k.slice(0, 28)), [
        { label: 'Presentes', data: Object.values(dTab).map(v => v._presentes), color: '#003DA5' },
        { label: 'Ausentes', data: Object.values(dTab).map(v => v._ausentes), color: '#00aaff' }
    ], 'y', true);

    renderPie('chart-tabulacao-pizza', ['Presentes', 'Ausentes'], [sumMetric('_presentes'), sumMetric('_ausentes')], ['#003DA5', '#00aaff']);

    // 6. Participação de Aptos/Agendados (Seção 4)
    const dPct = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_total_alunos', '_alunos_aptos']);
    const resPct = Object.entries(dPct).map(([n, v]) => {
        const p = v._total_alunos ? (v._alunos_aptos / v._total_alunos * 100) : 0;
        const l = v._alunos_aptos;
        return { n, p: Number(p.toFixed(1)), l };
    }).sort((a, b) => a.p - b.p);

    renderBar('chart-percentual-pratica', resPct.map(r => r.n), [{
        label: '% em relação ao Total de Alunos',
        data: resPct.map(r => r.p),
        alunos: resPct.map(r => r.l),
        backgroundColor: resPct.map(r => r.p < 50 ? '#ef4444' : (r.p < 80 ? '#94a3b8' : '#003DA5'))
    }], 'y', false, true);

    // 7. Aptos/Agendados por Curso (Seção 4)
    const dConf = aggregate(keysMap.curso, ['_alunos_aptos', '_alunos_inaptos']);
    const sortedConf = Object.entries(dConf).sort((a, b) => (b[1]._alunos_aptos + b[1]._alunos_inaptos) - (a[1]._alunos_aptos + a[1]._alunos_inaptos));
    const confLabels = sortedConf.map(e => e[0]), isConfH = confLabels.length > 8;

    const vpConf = $('chart-confirmacao-viewport'), wrConf = $('chart-confirmacao-wrapper');
    if (vpConf && wrConf) {
        vpConf.classList.toggle('is-scrollable', isConfH);
        wrConf.style.height = isConfH ? (confLabels.length * 30 + 40) + 'px' : '350px';
        if ($('chart-confirmacao-hint')) $('chart-confirmacao-hint').textContent = isConfH ? `${confLabels.length} cursos — role.` : '';
    }

    renderBar(
        'chart-confirmacao',
        confLabels,
        [
            {
                label: 'Aptos/Agendados',
                data: sortedConf.map(e => e[1]._alunos_aptos),
                color: '#003DA5'
            },
            {
                label: 'Inaptos/Não Agendados',
                data: sortedConf.map(e => e[1]._alunos_inaptos),
                color: '#94a3b8'
            }
        ],
        'y',
        true
    );
    if (isConfH) setupScrollspy(vpConf, wrConf, $('chart-confirmacao-indicator'), $('chart-confirmacao-top'), confLabels.length, 30);
}

function renderBar(id, labels, datasets, axis = 'x', stacked = false, isPercent = false) {
    const ctx = $(id); if (!ctx || typeof Chart === 'undefined') return;
    if (charts[id]) charts[id].destroy();

    const options = {
        responsive: true, maintainAspectRatio: false, indexAxis: axis,
        scales: {
            x: { stacked, beginAtZero: true },
            y: { stacked, beginAtZero: true }
        },
        plugins: {
            legend: { position: 'bottom' },
            tooltip: {
                callbacks: {
                    label: c => {
                        if (isPercent) {
                            const alunos = c.dataset.alunos?.[c.dataIndex];
                            return `${c.dataset.label}: ${c.raw}% (${alunos} alunos)`;
                        }
                        return `${c.dataset.label}: ${c.raw.toLocaleString('pt-BR')}`;
                    }
                }
            }
        }
    };

    if (isPercent) {
        const scale = axis === 'x' ? 'y' : 'x';
        options.scales[scale].max = 100;
        options.plugins.annotation = {
            annotations: {
                line1: {
                    type: 'line',
                    [axis === 'x' ? 'yMin' : 'xMin']: 50,
                    [axis === 'x' ? 'yMax' : 'xMax']: 50,
                    borderColor: 'rgba(0, 0, 0, 0.5)',
                    borderWidth: 2,
                    borderDash: [6, 6],
                    label: { display: true, content: 'Meta 50%', position: 'end' }
                }
            }
        };
    }

    charts[id] = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets: datasets.map(d => ({ ...d, backgroundColor: d.backgroundColor || d.color })) },
        options: options
    });
}

function renderPie(id, labels, data, colors) {
    const ctx = $(id); if (!ctx || typeof Chart === 'undefined') return;
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart(ctx, {
        type: 'doughnut', data: { labels, datasets: [{ data, backgroundColor: colors }] },
        options: {
            responsive: true, maintainAspectRatio: false, cutout: '60%',
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: c => `${c.label}: ${c.raw.toLocaleString('pt-BR')}`
                    }
                }
            }
        }
    });
}

function updateTable() {
    const body = $('table-body'); if (!body) return;
    body.innerHTML = filteredData.map(d =>
        `<tr>
            <td>${d[keysMap.unidade] || '-'}</td>
            <td>${(d[keysMap.curso] || '-').slice(0, 25)}</td>
            <td>${d[keysMap.categoria] || d['Categoria'] || '-'}</td>
            <td>${d._total_alunos}</td>
            <td>${d._alunos_aptos}</td>
            <td>${d._alunos_inaptos}</td>
            <td>${d._presentes}</td>
            <td>${d._ausentes}</td>
            </tr>`).join('');
    const vp = $('table-viewport'), ind = $('table-indicator'), hint = $('table-hint');
    if (vp && filteredData.length) {
        if (hint) hint.textContent = `${filteredData.length} registros — role.`;
        setupScrollspy(vp, body, ind, $('table-top'), filteredData.length, 42);
    } else if (ind) ind.style.display = 'none';
}

function setupScrollspy(vp, inner, ind, btn, total, rowH) {
    if (!vp || !ind) return;
    ind.style.display = 'block';
    const up = () => {
        const t = vp.scrollTop, v = vp.clientHeight;
        ind.textContent = `${Math.floor(t / rowH) + 1}–${Math.min(total, Math.ceil((t + v) / rowH))} de ${total}`;
        if (btn) btn.style.display = t > 80 ? 'flex' : 'none';
    };
    if (vp._h) vp.removeEventListener('scroll', vp._h);
    vp._h = up; vp.addEventListener('scroll', up, { passive: true });
    if (btn) btn.onclick = () => vp.scrollTo({ top: 0, behavior: 'smooth' });
    up();
}
