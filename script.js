// Global variables
let reservationData = [];
let instancesData = {};
let quotesData = {};
let originalReservationData = []; // Store unfiltered data
let dateFilter = {
    active: false,
    startDate: null,
    endDate: null
};
let uploadedFileName = ''; // Store the Excel filename

// DOM Elements
const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const fileName = document.getElementById('fileName');
const errorMessage = document.getElementById('errorMessage');
const loadingSpinner = document.getElementById('loadingSpinner');
const resultsSection = document.getElementById('resultsSection');

// Event Listeners
fileInput.addEventListener('change', handleFileSelect);
uploadArea.addEventListener('click', () => fileInput.click());

// Drag and drop functionality
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        fileInput.files = files;
        handleFileSelect({ target: { files } });
    }
});

// File handling
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    // Validate file type
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                       'application/vnd.ms-excel'];
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/)) {
        showError('Por favor, selecciona un archivo Excel válido (.xlsx o .xls)');
        return;
    }

    fileName.textContent = `📁 Archivo seleccionado: ${file.name}`;
    uploadedFileName = file.name; // Store filename
    hideError();
    loadingSpinner.classList.remove('hidden');
    resultsSection.classList.add('hidden');
    document.getElementById('filterSection').classList.add('hidden'); // Hide filter section on new upload

    // Read the file
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            processExcelFile(e.target.result);
        } catch (error) {
            showError('Error al procesar el archivo: ' + error.message);
            loadingSpinner.classList.add('hidden');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Process Excel file
function processExcelFile(data) {
    try {
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 'A' });

        // Skip header row
        reservationData = jsonData.slice(1);

        if (reservationData.length === 0) {
            showError('El archivo no contiene datos válidos');
            loadingSpinner.classList.add('hidden');
            return;
        }

        // Store original data
        originalReservationData = [...reservationData];
        
        // Reset filter state
        dateFilter.active = false;
        dateFilter.startDate = null;
        dateFilter.endDate = null;
        document.getElementById('startDate').value = '';
        document.getElementById('endDate').value = '';
        document.getElementById('filterStatus').classList.add('hidden');

        // Process data
        processReservations();
        displayResults();
        
        loadingSpinner.classList.add('hidden');
        resultsSection.classList.remove('hidden');
        document.getElementById('filterSection').classList.remove('hidden'); // Show filter section
    } catch (error) {
        showError('Error al leer el archivo Excel: ' + error.message);
        loadingSpinner.classList.add('hidden');
    }
}

// Process reservation data
function processReservations() {
    instancesData = {};

    reservationData.forEach(row => {
        const instancia = row.B || 'Sin Instancia';
        const estadoReserva = (row.U || '').trim();

        if (!instancesData[instancia]) {
            instancesData[instancia] = {
                name: instancia,
                confirmada: 0,
                pendiente: 0,
                fallida: 0,
                procesando: 0,
                total: 0,
                quotes: 0
            };
        }

        // Count by status (case-insensitive comparison)
        const estadoLower = estadoReserva.toLowerCase();
        if (estadoLower === 'confirmada') {
            instancesData[instancia].confirmada++;
        } else if (estadoLower === 'pendiente') {
            instancesData[instancia].pendiente++;
        } else if (estadoLower === 'fallida') {
            instancesData[instancia].fallida++;
        } else if (estadoLower === 'procesando') {
            instancesData[instancia].procesando++;
        }

        instancesData[instancia].total++;
    });

    // Initialize quotes data
    Object.keys(instancesData).forEach(instance => {
        if (!quotesData[instance]) {
            quotesData[instance] = 0;
        }
    });
}

// Display results
function displayResults() {
    displayOverallSummary();
    displayTop5Instances();
    displayGlobalConversion();
    displayInstanceDetails();
}

// Display overall summary
function displayOverallSummary() {
    let totalReservations = 0;
    let totalConfirmed = 0;
    let totalPending = 0;
    let totalFailed = 0;
    let totalProcessing = 0;

    Object.values(instancesData).forEach(instance => {
        totalReservations += instance.total;
        totalConfirmed += instance.confirmada;
        totalPending += instance.pendiente;
        totalFailed += instance.fallida;
        totalProcessing += instance.procesando;
    });

    document.getElementById('totalReservations').textContent = totalReservations;
    document.getElementById('totalConfirmed').textContent = totalConfirmed;
    document.getElementById('totalPending').textContent = totalPending;
    document.getElementById('totalFailed').textContent = totalFailed;
    document.getElementById('totalProcessing').textContent = totalProcessing;
    document.getElementById('totalInstances').textContent = Object.keys(instancesData).length;
}

// Display top 5 instances
function displayTop5Instances() {
    const sortedInstances = Object.values(instancesData)
        .sort((a, b) => b.total - a.total)
        .slice(0, 5);

    // Create chart
    const ctx = document.getElementById('top5Chart').getContext('2d');
    
    // Destroy existing chart if it exists
    if (window.top5Chart instanceof Chart) {
        window.top5Chart.destroy();
    }

    window.top5Chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sortedInstances.map(i => i.name),
            datasets: [{
                label: 'Total de Reservas',
                data: sortedInstances.map(i => i.total),
                backgroundColor: [
                    'rgba(74, 144, 226, 0.8)',
                    'rgba(52, 152, 219, 0.8)',
                    'rgba(41, 128, 185, 0.8)',
                    'rgba(39, 174, 96, 0.8)',
                    'rgba(46, 204, 113, 0.8)'
                ],
                borderColor: [
                    'rgba(74, 144, 226, 1)',
                    'rgba(52, 152, 219, 1)',
                    'rgba(41, 128, 185, 1)',
                    'rgba(39, 174, 96, 1)',
                    'rgba(46, 204, 113, 1)'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        stepSize: 1
                    }
                }
            }
        }
    });

    // Create table
    const top5Table = document.getElementById('top5Table');
    top5Table.innerHTML = sortedInstances.map((instance, index) => `
        <div class="top5-row">
            <div class="top5-rank">#${index + 1}</div>
            <div class="top5-name">${instance.name}</div>
            <div class="top5-count">${instance.total} reservas</div>
        </div>
    `).join('');
}

// Display global conversion
function displayGlobalConversion() {
    const totalQuotesInput = document.getElementById('totalQuotes');
    
    totalQuotesInput.addEventListener('input', updateGlobalConversion);
    updateGlobalConversion();
}

function updateGlobalConversion() {
    const totalQuotes = parseInt(document.getElementById('totalQuotes').value) || 0;
    const totalConfirmed = Object.values(instancesData)
        .reduce((sum, instance) => sum + instance.confirmada, 0);

    const conversionRate = totalConfirmed > 0 
        ? ((totalQuotes / totalConfirmed) * 100).toFixed(2)
        : '0.00';

    document.getElementById('globalConversionRate').textContent = `${conversionRate}%`;
}

// Display instance details
function displayInstanceDetails() {
    const container = document.getElementById('instancesContainer');
    
    // Sort instances by total reservations
    const sortedInstances = Object.entries(instancesData)
        .sort(([, a], [, b]) => b.total - a.total);

    container.innerHTML = sortedInstances.map(([instanceName, data]) => `
        <div class="instance-card">
            <div class="instance-header">
                <div class="instance-name">🏢 ${data.name}</div>
                <div class="instance-total">Total: ${data.total}</div>
            </div>

            <div class="instance-stats">
                <div class="stat-item">
                    <span class="stat-label">✅ Confirmada:</span>
                    <span class="stat-value confirmed">${data.confirmada}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">⏳ Pendiente:</span>
                    <span class="stat-value pending">${data.pendiente}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">❌ Fallida:</span>
                    <span class="stat-value failed">${data.fallida}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🔄 Procesando:</span>
                    <span class="stat-value processing">${data.procesando}</span>
                </div>
            </div>

            <div class="instance-conversion">
                <div class="quote-input-section">
                    <label>💼 Cantidad de Cotizaciones para esta instancia:</label>
                    <input 
                        type="number" 
                        class="quote-input" 
                        data-instance="${instanceName}"
                        value="${quotesData[instanceName] || 0}"
                        min="0"
                        onchange="updateInstanceConversion('${instanceName}')"
                    >
                </div>

                <div class="conversion-results" id="conversion-${instanceName}">
                    ${getConversionHTML(instanceName, quotesData[instanceName] || 0)}
                </div>
            </div>
        </div>
    `).join('');
}

function getConversionHTML(instanceName, quotes) {
    const data = instancesData[instanceName];
    
    const rate1 = data.total > 0 
        ? ((quotes / data.total) * 100).toFixed(2)
        : '0.00';
    
    const rate2 = data.confirmada > 0 
        ? ((quotes / data.confirmada) * 100).toFixed(2)
        : '0.00';

    return `
        <div class="conversion-result-item">
            <span class="conversion-result-label">Cotizaciones / Total Reservas:</span>
            <span class="conversion-result-value">${rate1}%</span>
        </div>
        <div class="conversion-result-item">
            <span class="conversion-result-label">Cotizaciones / Confirmadas:</span>
            <span class="conversion-result-value">${rate2}%</span>
        </div>
    `;
}

function updateInstanceConversion(instanceName) {
    const input = document.querySelector(`input[data-instance="${instanceName}"]`);
    const quotes = parseInt(input.value) || 0;
    quotesData[instanceName] = quotes;

    const conversionContainer = document.getElementById(`conversion-${instanceName}`);
    conversionContainer.innerHTML = getConversionHTML(instanceName, quotes);

    // Update global conversion as well
    updateGlobalConversion();
}

// Error handling
function showError(message) {
    errorMessage.textContent = '❌ ' + message;
    errorMessage.classList.remove('hidden');
}

function hideError() {
    errorMessage.classList.add('hidden');
}

// Date filter functions
function applyDateFilter() {
    const startDateInput = document.getElementById('startDate').value;
    const endDateInput = document.getElementById('endDate').value;
    
    if (!startDateInput || !endDateInput) {
        showError('Por favor selecciona ambas fechas (Desde y Hasta)');
        return;
    }
    
    const startDate = new Date(startDateInput);
    const endDate = new Date(endDateInput);
    
    if (startDate > endDate) {
        showError('La fecha "Desde" debe ser anterior a la fecha "Hasta"');
        return;
    }
    
    // Activate filter
    dateFilter.active = true;
    dateFilter.startDate = startDate;
    dateFilter.endDate = endDate;
    
    // Filter the data
    reservationData = originalReservationData.filter(row => {
        const fechaCreacion = row.Z; // Column Z: Fecha Creación
        if (!fechaCreacion) return false;
        
        // Parse date format "2026-02-16 12:14:21"
        const datePart = fechaCreacion.split(' ')[0]; // Get "2026-02-16"
        const rowDate = new Date(datePart);
        
        return rowDate >= startDate && rowDate <= endDate;
    });
    
    if (reservationData.length === 0) {
        showError('No se encontraron reservas en el rango de fechas seleccionado');
        reservationData = [...originalReservationData];
        return;
    }
    
    // Show filter status
    const filterStatus = document.getElementById('filterStatus');
    filterStatus.classList.remove('hidden');
    filterStatus.innerHTML = `
        <span class="filter-active-icon">✅</span>
        Mostrando <strong>${reservationData.length}</strong> reservas del 
        <strong>${formatDate(startDate)}</strong> al <strong>${formatDate(endDate)}</strong>
    `;
    
    // Reprocess and display filtered data
    processReservations();
    displayResults();
    
    hideError();
}

function clearDateFilter() {
    // Reset filter state
    dateFilter.active = false;
    dateFilter.startDate = null;
    dateFilter.endDate = null;
    
    // Clear date inputs
    document.getElementById('startDate').value = '';
    document.getElementById('endDate').value = '';
    
    // Hide filter status
    document.getElementById('filterStatus').classList.add('hidden');
    
    // Restore original data
    reservationData = [...originalReservationData];
    
    // Reprocess and display all data
    processReservations();
    displayResults();
    
    hideError();
}

function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}