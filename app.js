/**
 * GlassVibe Intelligent Client-Side Visualization Processor
 * Core Architecture: Clean Functional Modular Architecture
 */

// State Scope Object Wrapper
const AppState = {
    workbook: null,
    sheetNames: [],
    currentSheetName: '',
    rawSheetData: [],
    structuredData: [],
    columns: {
        categorical: [],
        numerical: []
    },
    instances: {
        barChart: null,
        lineChart: null,
        pieChart: null
    }
};

// Global Configuration Overrides for Glass Theme Styling Integration
const getChartTheming = () => {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    return {
        gridColor: isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(15, 23, 42, 0.08)',
        textColor: isDark ? '#94a3b8' : '#475569',
        accentColors: [
            'rgba(99, 102, 241, 0.85)',  // Indigo
            'rgba(236, 72, 153, 0.85)',  // Pink
            'rgba(20, 184, 166, 0.85)',  // Teal
            'rgba(245, 158, 11, 0.85)',  // Amber
            'rgba(139, 92, 246, 0.85)',  // Purple
            'rgba(59, 130, 246, 0.85)'   // Blue
        ],
        borderColors: isDark ? 'rgba(255,255,255,0.2)' : 'rgba(15,23,42,0.1)'
    };
};

/* ==========================================================================
   Document Initialization Entry Point
   ========================================================================== */
document.addEventListener('DOMContentLoaded', () => {
    initializeTheme();
    setupEventHandlers();
});

/**
 * Initializes and syncs application colors to match standard system states
 */
function initializeTheme() {
    const savedTheme = localStorage.getItem('gv-theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
}

/**
 * Maps event interactions back to core process loops
 */
function setupEventHandlers() {
    const themeBtn = document.getElementById('themeToggle');
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const sheetSelect = document.getElementById('sheetSelect');

    // Light-Dark State Toggle Router Action
    themeBtn.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme');
        const nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', nextTheme);
        localStorage.setItem('gv-theme', nextTheme);
        
        // Dynamic re-render to update native charts canvas borders/grid colors
        if (AppState.structuredData.length > 0) {
            rebuildAllCharts();
        }
    });

    // File Input Selectors Map Execution
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => handleFileSource(e.target.files[0]));

    // Drag Over Graphic Accent Controllers
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('drag-over'));
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        if (e.dataTransfer.files.length) {
            handleFileSource(e.dataTransfer.files[0]);
        }
    });

    // Active Sheet Navigation Execution Selection Control Map
    sheetSelect.addEventListener('change', (e) => loadWorkbookSheet(e.target.value));

    // Axis Selector Interface Element Controls Handlers Mapping
    setupAxisSelectors();
}

/* ==========================================================================
   Core Workbook & Sheet Processing Actions
   ========================================================================== */
function handleFileSource(file) {
    if (!file) return;

    const extension = file.name.split('.').pop().toLowerCase();
    if (extension !== 'xlsx' && extension !== 'xls') {
        alert('Validation Error: Please select an acceptable Microsoft Excel file schema type (.xlsx, .xls).');
        return;
    }

    // Render properties back inside visual state nodes
    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileSize').textContent = `${(file.size / 1024).toFixed(1)} KB`;
    document.getElementById('fileInfo').classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            AppState.workbook = workbook;
            AppState.sheetNames = workbook.SheetNames;
            
            populateSheetDropdown();
            loadWorkbookSheet(AppState.sheetNames[0]);
            
            // Switch layout shell frames visibility
            document.getElementById('emptyState').classList.add('hidden');
            document.getElementById('dashboardContent').classList.remove('hidden');
        } catch (error) {
            console.error('SheetJS Parsing Exception', error);
            alert('Parser Fault: Encountered an unrecoverable operational issue analyzing the structure internal to this Excel spreadsheet container.');
        }
    };
    reader.readAsArrayBuffer(file);
}

function populateSheetDropdown() {
    const container = document.getElementById('sheetSelectorContainer');
    const selector = document.getElementById('sheetSelect');
    
    selector.innerHTML = '';
    AppState.sheetNames.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        selector.appendChild(opt);
    });

    if (AppState.sheetNames.length > 1) {
        container.classList.remove('hidden');
    } else {
        container.classList.add('hidden');
    }
}

function loadWorkbookSheet(sheetName) {
    AppState.currentSheetName = sheetName;
    const worksheet = AppState.workbook.Sheets[sheetName];
    
    // Parse to array rows structures
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
    
    if (!rawData || rawData.length === 0) {
        alert(`Information: Target structural segment "${sheetName}" does not currently register accessible data matrices layout parameters.`);
        return;
    }

    AppState.rawSheetData = rawData;
    
    // Clean target null cells objects patterns
    AppState.structuredData = cleanDataStructure(rawData);
    
    // Identify internal column type properties
    analyzeDataSchema();
    
    // Render analytical outputs inside dashboard layouts
    calculateKPIs();
    populateAxisSelectors();
    rebuildAllCharts();
}

/**
 * Massages raw structures stripping completely empty layout row profiles
 */
function cleanDataStructure(data) {
    return data.filter(row => Object.values(row).some(val => val !== null && val !== ''));
}

/* ==========================================================================
   Structural Schema Engine Intelligence Analytical Workers
   ========================================================================== */
function analyzeDataSchema() {
    if (!AppState.structuredData.length) return;
    
    const sampleRow = AppState.structuredData[0];
    const keys = Object.keys(sampleRow);
    
    AppState.columns.categorical = [];
    AppState.columns.numerical = [];

    keys.forEach(key => {
        let numericCount = 0;
        let totalValids = 0;

        AppState.structuredData.forEach(row => {
            const value = row[key];
            if (value !== null && value !== undefined && value !== '') {
                totalValids++;
                if (!isNaN(Number(value))) {
                    numericCount++;
                }
            }
        });

        // Schema classification mapping rules logic heuristics execution
        if (totalValids > 0 && (numericCount / totalValids) > 0.7) {
            AppState.columns.numerical.push(key);
        } else {
            AppState.columns.categorical.push(key);
        }
    });

    // Fallback safety verification rules checks overrides
    if (AppState.columns.categorical.length === 0 && keys.length) {
        AppState.columns.categorical.push(keys[0]);
    }
    if (AppState.columns.numerical.length === 0 && keys.length) {
        AppState.columns.numerical.push(keys[keys.length - 1]);
    }
}

/**
 * Computes essential KPIs for display
 */
function calculateKPIs() {
    const data = AppState.structuredData;
    document.getElementById('kpiRows').textContent = data.length.toLocaleString();
    
    const totalCols = data.length ? Object.keys(data[0]).length : 0;
    document.getElementById('kpiCols').textContent = totalCols.toLocaleString();

    // Sum calculation logic execution loops
    if (AppState.columns.numerical.length > 0) {
        const targetMetric = AppState.columns.numerical[0];
        const aggregateSum = data.reduce((acc, row) => {
            const val = Number(row[targetMetric]);
            return acc + (!isNaN(val) ? val : 0);
        }, 0);

        document.getElementById('kpiMetricLabel1').textContent = `Total (${targetMetric})`;
        document.getElementById('kpiMetricValue1').textContent = 
            aggregateSum % 1 === 0 ? aggregateSum.toLocaleString() : aggregateSum.toFixed(2).toLocaleString();
    } else {
        document.getElementById('kpiMetricLabel1').textContent = "Aggregate Matrix";
        document.getElementById('kpiMetricValue1').textContent = "N/A";
    }
}

/* ==========================================================================
   Visual Dashboard Interface Selection Generation Controls
   ========================================================================== */
function populateAxisSelectors() {
    const selectors = document.querySelectorAll('.glass-select');
    selectors.forEach(select => select.innerHTML = '');

    const fillOptions = (selectId, colType) => {
        const element = document.getElementById(selectId);
        AppState.columns[colType].forEach(col => {
            const opt = document.createElement('option');
            opt.value = col;
            opt.textContent = col;
            element.appendChild(opt);
        });
    };

    // Populate Category Arrays
    ['barXSelect', 'lineXSelect', 'pieXSelect'].forEach(id => fillOptions(id, 'categorical'));
    // Populate Numerical Values Axis Arrays
    ['barYSelect', 'lineYSelect', 'pieYSelect'].forEach(id => fillOptions(id, 'numerical'));
}

function setupAxisSelectors() {
    document.getElementById('barXSelect').addEventListener('change', () => updateChartInstance('bar'));
    document.getElementById('barYSelect').addEventListener('change', () => updateChartInstance('bar'));
    
    document.getElementById('lineXSelect').addEventListener('change', () => updateChartInstance('line'));
    document.getElementById('lineYSelect').addEventListener('change', () => updateChartInstance('line'));
    
    document.getElementById('pieXSelect').addEventListener('change', () => updateChartInstance('pie'));
    document.getElementById('pieYSelect').addEventListener('change', () => updateChartInstance('pie'));
}

/* ==========================================================================
   Data Rendering Core Canvas Engine Methods (ChartJS Interface Wrapper)
   ========================================================================== */
function rebuildAllCharts() {
    updateChartInstance('bar');
    updateChartInstance('line');
    updateChartInstance('pie');
}

/**
 * Dynamically aggregates, computes, and renders a selected chart component type
 */
function updateChartInstance(type) {
    const data = AppState.structuredData;
    if (!data.length) return;

    const theme = getChartTheming();
    let xField, yField;

    if (type === 'bar') {
        xField = document.getElementById('barXSelect').value;
        yField = document.getElementById('barYSelect').value;
    } else if (type === 'line') {
        xField = document.getElementById('lineXSelect').value;
        yField = document.getElementById('lineYSelect').value;
    } else if (type === 'pie') {
        xField = document.getElementById('pieXSelect').value;
        yField = document.getElementById('pieYSelect').value;
    }

    if (!xField || !yField) return;

    // Aggregates numerical values by unique categorical labels
    const aggregationMap = new Map();
    data.forEach(row => {
        const label = String(row[xField] === null || row[xField] === undefined ? '(Blank)' : row[xField]);
        const val = Number(row[yField]);
        const validNum = !isNaN(val) ? val : 0;

        if (aggregationMap.has(label)) {
            aggregationMap.set(label, aggregationMap.get(label) + validNum);
        } else {
            aggregationMap.set(label, validNum);
        }
    });

    // Extract ordered structural matrices mappings arrays configurations, capped for performance
    const maxDataPoints = type === 'pie' ? 8 : 15; 
    let labels = Array.from(aggregationMap.keys());
    let values = Array.from(aggregationMap.values());

    if (labels.length > maxDataPoints) {
        labels = labels.slice(0, maxDataPoints);
        values = values.slice(0, maxDataPoints);
    }

    // Chart.js Configuration Definitions Mapping Block
    let chartConfig = {
        data: {
            labels: labels,
            datasets: [{
                label: yField,
                data: values,
                backgroundColor: type === 'pie' ? theme.accentColors : theme.accentColors[0],
                borderColor: type === 'pie' ? theme.borderColors : theme.accentColors[0],
                borderWidth: 1,
                borderRadius: type !== 'pie' ? 6 : 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: type === 'pie',
                    position: 'right',
                    labels: { color: theme.textColor, font: { family: 'Segoe UI' } }
                },
                tooltip: {
                    padding: 12,
                    borderRadius: 8,
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    titleFont: { size: 13, weight: 'bold' },
                    bodyFont: { size: 12 }
                }
            },
            scales: type !== 'pie' ? {
                x: {
                    grid: { color: theme.gridColor },
                    ticks: { color: theme.textColor, font: { size: 11 } }
                },
                y: {
                    grid: { color: theme.gridColor },
                    ticks: { color: theme.textColor, font: { size: 11 } }
                }
            } : {}
        }
    };

    // Inject contextual changes for specific styles
    if (type === 'bar') chartConfig.type = 'bar';
    if (type === 'line') {
        chartConfig.type = 'line';
        chartConfig.data.datasets[0].fill = true;
        chartConfig.data.datasets[0].backgroundColor = 'rgba(99, 102, 241, 0.15)';
        chartConfig.data.datasets[0].borderColor = theme.accentColors[0];
        chartConfig.data.datasets[0].tension = 0.3;
        chartConfig.data.datasets[0].pointRadius = 4;
    }
    if (type === 'pie') chartConfig.type = 'doughnut';

    // Garbage Collection Cycle - Terminate previous instances to avoid memory leaks
    if (AppState.instances[`${type}Chart`]) {
        AppState.instances[`${type}Chart`].destroy();
    }

    // Canvas instantiation execution map block paths
    const ctx = document.getElementById(`${type}Chart`).getContext('2d');
    AppState.instances[`${type}Chart`] = new Chart(ctx, chartConfig);
}