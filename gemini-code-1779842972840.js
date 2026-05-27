// Dados iniciais embutidos (Sua planilha anexada) para funcionar out-of-the-box
const defaultCSV = `Id,Data,Setor,Produto,Quant,Valor,Total
1,2026-05-01,Varejo,Notebook Lenovo IdeaPad,8,3200,25600
2,2026-05-02,Varejo,Smartphone Samsung Galaxy,15,1850,27750
3,2026-05-03,Supermercado,Arroz 5kg,120,28,3360
4,2026-05-04,Farmácia,Protetor Solar FPS 50,40,52,2080
5,2026-05-05,Moda,Camiseta Masculina,70,39.9,2793
6,2026-05-06,Informática,Mouse Gamer RGB,25,145,3625
7,2026-05-07,Papelaria,Caderno Universitário,90,18.5,1665
8,2026-05-08,Eletrodomésticos,Air Fryer 5L,12,420,5040
9,2026-05-09,Cosméticos,Perfume Importado,10,350,3500
10,2026-05-10,Bebidas,Refrigerante 2L,150,9.5,1425`;

// Estado global do App
let rawData = [];
let filteredData = [];
let chartInstances = {};
let columns = [];
let numericCols = [];
let categoryCols = [];

// Estado da Tabela
let currentPage = 1;
const rowsPerPage = 5;
let sortCol = '';
let sortAsc = true;

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    // Parseia os dados embutidos iniciais
    Papa.parse(defaultCSV, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
        complete: function(results) {
            processData(results.data);
        }
    });

    // Configura os Listeners dos Filtros
    document.querySelectorAll('.filter-card input, .filter-card select').forEach(el => {
        el.addEventListener('input', applyFilters);
    });

    // Listener de Upload de Excel/CSV local
    document.getElementById('fileUpload').addEventListener('change', handleFileUpload);
});

// Processamento Inicial e "Inteligência"
function processData(data) {
    if (!data || data.length === 0) return;
    rawData = data;
    columns = Object.keys(data[0]);

    // Detectar tipos de colunas
    numericCols = columns.filter(col => typeof data[0][col] === 'number');
    categoryCols = columns.filter(col => typeof data[0][col] === 'string' && !col.toLowerCase().includes('data'));

    populateDropdowns();
    applyFilters();
}

// Upload de novo arquivo (Excel ou CSV)
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, {raw: true});
        processData(json);
    };
    reader.readAsArrayBuffer(file);
}

// Popular Filtros
function populateDropdowns() {
    const setores = [...new Set(rawData.map(item => item.Setor || item.setor).filter(Boolean))];
    const produtos = [...new Set(rawData.map(item => item.Produto || item.produto).filter(Boolean))];

    const selSetor = document.getElementById('filterSetor');
    const selProduto = document.getElementById('filterProduto');
    
    selSetor.innerHTML = '<option value="">Todos</option>' + setores.map(s => `<option value="${s}">${s}</option>`).join('');
    selProduto.innerHTML = '<option value="">Todos</option>' + produtos.map(p => `<option value="${p}">${p}</option>`).join('');
}

// Filtros Principais
function applyFilters() {
    const start = document.getElementById('filterDateStart').value;
    const end = document.getElementById('filterDateEnd').value;
    const setor = document.getElementById('filterSetor').value;
    const produto = document.getElementById('filterProduto').value;
    const search = document.getElementById('filterSearch').value.toLowerCase();

    filteredData = rawData.filter(row => {
        let match = true;
        // Filtro Data
        const rowDate = row.Data || row.data;
        if (start && rowDate && rowDate < start) match = false;
        if (end && rowDate && rowDate > end) match = false;
        // Filtro Dropdown
        if (setor && (row.Setor || row.setor) !== setor) match = false;
        if (produto && (row.Produto || row.produto) !== produto) match = false;
        // Filtro Texto
        if (search) {
            const rowValues = Object.values(row).join(' ').toLowerCase();
            if (!rowValues.includes(search)) match = false;
        }
        return match;
    });

    currentPage = 1; // Reset paginação
    updateKPIs();
    renderCharts();
    renderTable();
}

// Atualizar KPIs
function updateKPIs() {
    const sumTotal = filteredData.reduce((acc, row) => acc + (row.Total || row.total || 0), 0);
    const sumQtd = filteredData.reduce((acc, row) => acc + (row.Quant || row.quant || row.Quantidade || 0), 0);
    const minVal = Math.min(...filteredData.map(r => r.Total || r.total || 0).filter(v => v > 0));
    const maxVal = Math.max(...filteredData.map(r => r.Total || r.total || 0));

    const formatCurrency = val => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);

    document.getElementById('kpiQtd').innerText = sumQtd.toLocaleString('pt-BR');
    document.getElementById('kpiTotal').innerText = formatCurrency(sumTotal);
    document.getElementById('kpiMedia').innerText = filteredData.length ? formatCurrency(sumTotal / sumQtd) : 'R$ 0,00';
    document.getElementById('kpiMax').innerText = filteredData.length ? formatCurrency(maxVal) : 'R$ 0,00';
    document.getElementById('kpiMin').innerText = (filteredData.length && minVal !== Infinity) ? formatCurrency(minVal) : 'R$ 0,00';
    document.getElementById('kpiCount').innerText = filteredData.length;
}

// Agrupamento para Gráficos
function groupData(key, sumKey) {
    return filteredData.reduce((acc, row) => {
        let k = row[key];
        if (!k) return acc;
        acc[k] = (acc[k] || 0) + (row[sumKey] || 1); // soma ou conta
        return acc;
    }, {});
}

// Renderizar Gráficos
function renderCharts() {
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    const textColor = isDark ? '#e0e0e0' : '#333';
    
    Chart.defaults.color = textColor;
    Chart.defaults.borderColor = isDark ? '#444' : '#e9ecef';

    const colors = ['#4361ee', '#3a0ca3', '#f72585', '#4cc9f0', '#ffb703', '#fb8500'];

    // 1. Linha/Área (Evolução Diária)
    const dataEvo = groupData(columns.find(c=>c.toLowerCase().includes('data')), 'Total');
    createChart('chartLineArea', 'line', {
        labels: Object.keys(dataEvo),
        datasets: [{
            label: 'Receita Diária',
            data: Object.values(dataEvo),
            borderColor: '#4361ee',
            backgroundColor: 'rgba(67, 97, 238, 0.2)',
            fill: true,
            tension: 0.4
        }]
    });

    // 2. Barras (Top Produtos)
    const prodCol = columns.find(c=>c.toLowerCase() === 'produto') || categoryCols[0];
    const dataProd = groupData(prodCol, 'Quant');
    createChart('chartBar', 'bar', {
        labels: Object.keys(dataProd),
        datasets: [{
            label: 'Quantidade Vendida',
            data: Object.values(dataProd),
            backgroundColor: colors
        }]
    });

    // 3. Pizza (Quantidade por Setor)
    const setorCol = columns.find(c=>c.toLowerCase() === 'setor') || categoryCols[1];
    const dataSetorQtd = groupData(setorCol, 'Quant');
    createChart('chartPie', 'pie', {
        labels: Object.keys(dataSetorQtd),
        datasets: [{
            data: Object.values(dataSetorQtd),
            backgroundColor: colors
        }]
    });

    // 4. Doughnut (Receita por Setor)
    const dataSetorTotal = groupData(setorCol, 'Total');
    createChart('chartDoughnut', 'doughnut', {
        labels: Object.keys(dataSetorTotal),
        datasets: [{
            data: Object.values(dataSetorTotal),
            backgroundColor: colors
        }]
    });
}

// Wrapper do Chart.js
function createChart(id, type, data) {
    const ctx = document.getElementById(id).getContext('2d');
    if (chartInstances[id]) chartInstances[id].destroy();
    
    chartInstances[id] = new Chart(ctx, {
        type: type,
        data: data,
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: (type==='pie'||type==='doughnut') ? 'right' : 'top' }
            }
        }
    });
}

// Lógica de Ordenação
function sortData(col) {
    if (sortCol === col) {
        sortAsc = !sortAsc;
    } else {
        sortCol = col;
        sortAsc = true;
    }
    
    filteredData.sort((a, b) => {
        let valA = a[col], valB = b[col];
        if (typeof valA === 'string') valA = valA.toLowerCase();
        if (typeof valB === 'string') valB = valB.toLowerCase();
        
        if (valA < valB) return sortAsc ? -1 : 1;
        if (valA > valB) return sortAsc ? 1 : -1;
        return 0;
    });
    
    renderTable();
}

// Renderizar Tabela com Paginação
function renderTable() {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    
    // Cabeçalho
    thead.innerHTML = '<tr>' + columns.map(col => 
        `<th onclick="sortData('${col}')">${col} ${sortCol === col ? (sortAsc ? '↑' : '↓') : ''}</th>`
    ).join('') + '</tr>';

    // Corpo (Paginação)
    const startIdx = (currentPage - 1) * rowsPerPage;
    const paginatedItems = filteredData.slice(startIdx, startIdx + rowsPerPage);
    
    tbody.innerHTML = paginatedItems.map(row => 
        '<tr>' + columns.map(col => {
            let val = row[col];
            if(typeof val === 'number' && val > 1000 && !col.toLowerCase().includes('id')) {
                 val = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
            }
            return `<td>${val}</td>`;
        }).join('') + '</tr>'
    ).join('');

    // Info da Paginação
    const totalPages = Math.ceil(filteredData.length / rowsPerPage) || 1;
    document.getElementById('pageInfo').innerText = `Página ${currentPage} de ${totalPages} (${filteredData.length} registros)`;
}

// Controles de Paginação
function prevPage() {
    if (currentPage > 1) { currentPage--; renderTable(); }
}
function nextPage() {
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    if (currentPage < totalPages) { currentPage++; renderTable(); }
}

// Modo Claro/Escuro
function toggleTheme() {
    const body = document.body;
    const btn = document.getElementById('themeToggle');
    if (body.getAttribute('data-theme') === 'light') {
        body.setAttribute('data-theme', 'dark');
        btn.innerHTML = '<i class="fa-solid fa-sun"></i>';
    } else {
        body.setAttribute('data-theme', 'light');
        btn.innerHTML = '<i class="fa-solid fa-moon"></i>';
    }
    renderCharts(); // Re-renderiza gráficos para ajustar cores do texto/bordas
}

// Exportação PDF
function exportToPDF() {
    const element = document.getElementById('dashboard-content');
    const opt = {
        margin:       10,
        filename:     'Relatorio_Dashboard.pdf',
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'landscape' }
    };
    html2pdf().set(opt).from(element).save();
}