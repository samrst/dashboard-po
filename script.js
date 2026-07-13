let rawData = [], filteredData = [], charts = {};
const $ = id => document.getElementById(id);
const filters = ['unidade', 'curso', 'modalidade'];
const keysMap = { unidade: 'Unidade operacional (Escola)', curso: 'Curso', modalidade: 'Modalidade' };

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

            // Abordagem robusta: ler como array de arrays (header: 1), pular linha de título,
            // e construir os objetos manualmente mapeando pelos cabeçalhos
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

            const headers = allRows[headerRowIdx];
            const colIndex = {};
            headers.forEach((h, i) => {
                if (h != null) colIndex[String(h).trim()] = i;
            });

            // Mapeamento dos índices das colunas que precisamos
            const idx = {
                unidade: colIndex['Unidade operacional (Escola)'],
                curso: colIndex['Curso'],
                modalidade: colIndex['Modalidade'],
                alunos: colIndex['Alunos'] ?? colIndex['Alunos '],
                aptos: colIndex['Alunos Aptos'],
                inaptos: colIndex['Alunos Inaptos'],
                agendados: colIndex['Alunos Agendados'],
                naoAgendados: colIndex['Alunos Não Agendados'],
                presentes: colIndex['Presentes'],
                ausentes: colIndex['Ausentes']
            };

            // Validar que temos pelo menos a coluna de unidade
            if (idx.unidade === undefined) {
                throw new Error('Coluna "Unidade operacional (Escola)" não encontrada na planilha.');
            }

            // Processar as linhas de dados (pular título e cabeçalhos)
            rawData = [];
            for (let i = headerRowIdx + 1; i < allRows.length; i++) {
                const row = allRows[i];
                if (!row) continue;
                const unidade = row[idx.unidade];
                // Pular linhas vazias ou de total
                if (!unidade || String(unidade).trim() === '' || String(unidade).toLowerCase() === 'total') continue;

                rawData.push({
                    'Unidade operacional (Escola)': unidade,
                    'Curso': row[idx.curso] || '-',
                    'Modalidade': row[idx.modalidade] || '-',
                    '_total_alunos': parseInt(row[idx.alunos]) || 0,
                    '_alunos_aptos': parseInt(row[idx.aptos]) || 0,
                    '_alunos_inaptos': parseInt(row[idx.inaptos]) || 0,
                    '_agendados': parseInt(row[idx.agendados]) || 0,
                    '_nao_agendados': parseInt(row[idx.naoAgendados]) || 0,
                    '_presentes': parseInt(row[idx.presentes]) || 0,
                    '_ausentes': parseInt(row[idx.ausentes]) || 0
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
        aplicadas: '_agendados',
        feitas: '_nao_agendados',
        pendentes: '_presentes',
        confirmadas: '_ausentes'
    };
    const vals = Object.fromEntries(Object.entries(kpis).map(([k, v]) => [k, sum(v)]));

    Object.entries(vals).forEach(([k, v]) => { if ($(`kpi-${k}`)) $(`kpi-${k}`).textContent = v.toLocaleString('pt-BR'); });

    const pcts = {
        'homologados-pct': [vals.homologados, vals.objetivo, 'dos Alunos'],
        'total-pct': [vals.total, vals.objetivo, 'dos Alunos'],
        'aplicadas-pct': [vals.aplicadas, vals.homologados, 'dos Aptos'],
        'feitas-pct': [vals.feitas, vals.homologados, 'dos Aptos'],
        'pendentes-pct': [vals.pendentes, vals.aplicadas, 'dos Agendados'],
        'confirmadas-pct': [vals.confirmadas, vals.aplicadas, 'dos Agendados']
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
        { label: 'Alunos Aptos', data: Object.values(dAlunos).map(v => v._alunos_aptos), color: '#94a3b8' }
    ], 'x', true);

    // 2. Aptos vs Agendados por Curso
    const dHom = aggregate(keysMap.curso, ['_alunos_aptos', '_agendados']);
    const sortedHom = Object.entries(dHom).sort((a, b) => b[1]._alunos_aptos - a[1]._alunos_aptos);
    const labels = sortedHom.map(e => e[0]), isH = labels.length > 8;

    const vp = $('chart-homologacao-viewport'), wr = $('chart-homologacao-wrapper');
    if (vp && wr) {
        vp.classList.toggle('is-scrollable', isH);
        wr.style.height = isH ? (labels.length * 30 + 40) + 'px' : '350px';
        if ($('chart-homologacao-hint')) $('chart-homologacao-hint').textContent = isH ? `${labels.length} cursos — role.` : '';
    }
    renderBar('chart-homologacao', labels, [
        { label: 'Alunos Aptos', data: sortedHom.map(e => e[1]._alunos_aptos), color: '#003366' },
        { label: 'Agendados', data: sortedHom.map(e => e[1]._agendados), color: '#00aaff' }
    ], isH ? 'y' : 'x');
    if (isH) setupScrollspy(vp, wr, $('chart-homologacao-indicator'), $('chart-homologacao-top'), labels.length, 30);

    // 3 & 4. Agendados vs Não Agendados
    const dApp = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_agendados', '_nao_agendados']);
    renderBar('chart-aplicacao', Object.keys(dApp), [
        { label: 'Agendados', data: Object.values(dApp).map(v => v._agendados), color: '#005599' },
        { label: 'Não Agendados', data: Object.values(dApp).map(v => v._nao_agendados), color: '#94a3b8' }
    ], 'x', true);

    // 5 & 6. Presentes vs Ausentes
    const dTab = aggregate(keysMap.curso, ['_presentes', '_ausentes']);
    renderBar('chart-tabulacao', Object.keys(dTab).map(k => k.slice(0, 28)), [
        { label: 'Presentes', data: Object.values(dTab).map(v => v._presentes), color: '#003DA5' },
        { label: 'Ausentes', data: Object.values(dTab).map(v => v._ausentes), color: '#00aaff' }
    ], 'y', true);

    const sumMetric = m => filteredData.reduce((a, b) => a + (Number(b[m]) || 0), 0);
    renderPie('chart-aplicacao-pizza', ['Agendados', 'Não Agendados'], [sumMetric('_agendados'), sumMetric('_nao_agendados')], ['#005599', '#94a3b8']);
    renderPie('chart-tabulacao-pizza', ['Presentes', 'Ausentes'], [sumMetric('_presentes'), sumMetric('_ausentes')], ['#003DA5', '#00aaff']);

    // 7. Percentual de Agendamento
    const dPct = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_alunos_aptos', '_agendados']);
    const resPct = Object.entries(dPct).map(([n, v]) => {
        const p = v._alunos_aptos ? (v._agendados / v._alunos_aptos * 100) : 0;
        const l = v._agendados;
        return { n, p: Number(p.toFixed(1)), l };
    }).sort((a, b) => a.p - b.p);

    renderBar('chart-percentual-pratica', resPct.map(r => r.n), [{
    label: '% em relação aos Alunos Aptos',
    data: resPct.map(r => r.p),
    alunos: resPct.map(r => r.l),
    backgroundColor: resPct.map(r => r.p < 50 ? '#ef4444' : (r.p < 80 ? '#94a3b8' : '#003DA5'))
}], 'y', false, true);

    // 8. Agendamento por Curso
    const dConf = aggregate(
        isAll ? keysMap.unidade : keysMap.curso,
        ['_agendados', '_nao_agendados']
    );

    renderBar(
        'chart-confirmacao',
        Object.keys(dConf),
        [
            {
                label: 'Agendados',
                data: Object.values(dConf).map(v => v._agendados),
                color: '#003DA5'
            },
            {
                label: 'Não Agendados',
                data: Object.values(dConf).map(v => v._nao_agendados),
                color: '#94a3b8'
            }
        ],
        'y',
        true
    );
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
            <td>${d._total_alunos}</td>
            <td>${d._alunos_aptos}</td>
            <td>${d._alunos_inaptos}</td>
            <td>${d._agendados}</td>
            <td>${d._nao_agendados}</td>
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
