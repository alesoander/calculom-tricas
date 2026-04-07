// Global variables
let reservationData = [];
let instancesData = {};
let quotesData = {};
let originalReservationData = []; // Store unfiltered data
let uploadedFileName = ''; // Store Excel filename
let dateFilter = {
    active: false,
    startDate: null,
    endDate: null
};

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
    hideError();
    loadingSpinner.classList.remove('hidden');
    resultsSection.classList.add('hidden');

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

        // Store original data copy
        originalReservationData = [...reservationData];
        
        // Store uploaded file name
        const fileInputElement = document.getElementById('fileInput');
        if (fileInputElement.files[0]) {
            uploadedFileName = fileInputElement.files[0].name;
        }

        // Show filter section
        document.getElementById('filterSection').classList.remove('hidden');

        // Process data
        processReservations();
        displayResults();
        
        loadingSpinner.classList.add('hidden');
        resultsSection.classList.remove('hidden');
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
    
    const rate1 = quotes > 0 
        ? ((data.total / quotes) * 100).toFixed(2)
        : '0.00';
    
    const rate2 = quotes > 0 
        ? ((data.confirmada / quotes) * 100).toFixed(2)
        : '0.00';

    return `
        <div class="conversion-result-item">
            <span class="conversion-result-label">Total Reservas / Cotizaciones:</span>
            <span class="conversion-result-value">${rate1}%</span>
        </div>
        <div class="conversion-result-item">
            <span class="conversion-result-label">Confirmadas / Cotizaciones:</span>
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

// ============================================
// Date Filter Functionality
// ============================================

// Apply date filter
function applyDateFilter() {
    const startDateInput = document.getElementById('startDate').value;
    const endDateInput = document.getElementById('endDate').value;
    
    if (!startDateInput || !endDateInput) {
        showError('Por favor selecciona ambas fechas (Desde y Hasta)');
        return;
    }
    
    // Parse dates in local timezone
    const startDate = new Date(startDateInput + 'T00:00:00');
    const endDate = new Date(endDateInput + 'T23:59:59');
    
    if (startDate > endDate) {
        showError('La fecha "Desde" debe ser anterior a la fecha "Hasta"');
        return;
    }
    
    // Activate filter
    dateFilter.active = true;
    dateFilter.startDate = startDate;
    dateFilter.endDate = endDate;
    
    // Filter the data - store in temp variable first for validation
    const filteredData = originalReservationData.filter(row => {
        const fechaCreacion = row.Z; // Column Z
        if (!fechaCreacion) return false;
        
        let rowDate;
        // Handle both string dates and Excel serial numbers
        if (typeof fechaCreacion === 'number') {
            // Excel serial date: days since 1900-01-01 (25569 is the difference between Excel epoch and Unix epoch in days)
            rowDate = new Date((fechaCreacion - 25569) * 86400 * 1000);
        } else if (typeof fechaCreacion === 'string') {
            // String format: "2026-02-16 12:14:21" - parse in local timezone
            const datePart = fechaCreacion.split(' ')[0];
            rowDate = new Date(datePart + 'T00:00:00');
        } else {
            return false;
        }
        
        return rowDate >= startDate && rowDate <= endDate;
    });
    
    // Validate filtered results before applying
    if (filteredData.length === 0) {
        showError('No se encontraron reservas en el rango de fechas seleccionado');
        dateFilter.active = false;
        return;
    }
    
    // Apply the filtered data
    reservationData = filteredData;
    
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

// Clear date filter
function clearDateFilter() {
    // Reset filter state
    dateFilter.active = false;
    dateFilter.startDate = null;
    dateFilter.endDate = null;
    
    // Clear inputs
    document.getElementById('startDate').value = '';
    document.getElementById('endDate').value = '';
    
    // Hide status
    document.getElementById('filterStatus').classList.add('hidden');
    
    // Restore original data
    reservationData = [...originalReservationData];
    
    // Reprocess and display
    processReservations();
    displayResults();
    
    hideError();
}

// Format date helper
function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// ============================================
// PDF Export Functionality
// ============================================

// PDF Export main function
async function exportToPDF() {
    try {
        showExportLoading();
        
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        
        let yPosition = 20;
        
        // Add header
        yPosition = addPDFHeader(doc, yPosition);
        
        // Add overall summary
        yPosition = addPDFSummary(doc, yPosition);
        
        // Check if we need a new page
        if (yPosition > 240) {
            doc.addPage();
            yPosition = 20;
        }
        
        // Capture and add chart
        yPosition = await addPDFChart(doc, yPosition);
        
        // Add Top 5 table
        yPosition = addPDFTop5Table(doc, yPosition);
        
        // Add global conversion
        yPosition = addPDFGlobalConversion(doc, yPosition);
        
        // Add instance details
        yPosition = await addPDFInstanceDetails(doc, yPosition);
        
        // Add footer to all pages
        addPDFFooter(doc);
        
        // Generate filename and download
        const filename = `reporte-reservas-${getFormattedDateTime()}.pdf`;
        doc.save(filename);
        
        hideExportLoading();
        showExportSuccess();
        
    } catch (error) {
        console.error('Error generating PDF:', error);
        hideExportLoading();
        showExportError();
    }
}

// Helper function to add PDF header
function addPDFHeader(doc, y) {
    // Title
    doc.setFontSize(20);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(74, 144, 226);
    doc.text('Análisis de Reservas - Reporte', 105, y, { align: 'center' });
    
    y += 10;
    
    // Generation date
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100, 100, 100);
    const now = new Date();
    const dateStr = `Generado el: ${now.toLocaleDateString('es-ES')} a las ${now.toLocaleTimeString('es-ES')}`;
    doc.text(dateStr, 105, y, { align: 'center' });
    
    y += 6;
    
    // File name
    if (uploadedFileName) {
        doc.text(`Archivo: ${uploadedFileName}`, 105, y, { align: 'center' });
        y += 6;
    }
    
    // Filter info if active
    if (dateFilter.active) {
        doc.setTextColor(39, 174, 96);
        doc.text(`Filtrado: ${formatDate(dateFilter.startDate)} - ${formatDate(dateFilter.endDate)}`, 105, y, { align: 'center' });
        y += 6;
    }
    
    y += 5;
    
    // Line separator
    doc.setDrawColor(74, 144, 226);
    doc.setLineWidth(0.5);
    doc.line(15, y, 195, y);
    
    y += 10;
    
    return y;
}

// Helper function to add overall summary
function addPDFSummary(doc, y) {
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0, 0, 0);
    doc.text('📈 Resumen General', 15, y);
    
    y += 8;
    
    const totalReservations = parseInt(document.getElementById('totalReservations').textContent);
    const totalConfirmed = parseInt(document.getElementById('totalConfirmed').textContent);
    const totalPending = parseInt(document.getElementById('totalPending').textContent);
    const totalFailed = parseInt(document.getElementById('totalFailed').textContent);
    const totalProcessing = parseInt(document.getElementById('totalProcessing').textContent);
    const totalInstances = parseInt(document.getElementById('totalInstances').textContent);
    
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    
    const summaryData = [
        { label: 'Total Reservas:', value: totalReservations, color: [74, 144, 226] },
        { label: 'Reservas Confirmadas:', value: totalConfirmed, color: [39, 174, 96] },
        { label: 'Reservas Pendientes:', value: totalPending, color: [243, 156, 18] },
        { label: 'Reservas Fallidas:', value: totalFailed, color: [231, 76, 60] },
        { label: 'Reservas Procesando:', value: totalProcessing, color: [155, 89, 182] },
        { label: 'Total Instancias:', value: totalInstances, color: [74, 144, 226] }
    ];
    
    summaryData.forEach((item, index) => {
        const xLeft = 15 + (index % 2) * 90;
        const yRow = y + Math.floor(index / 2) * 12;
        
        // Box background
        doc.setFillColor(245, 247, 250);
        doc.rect(xLeft, yRow - 5, 85, 10, 'F');
        
        // Label
        doc.setTextColor(127, 140, 141);
        doc.text(item.label, xLeft + 2, yRow);
        
        // Value
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(...item.color);
        doc.text(item.value.toString(), xLeft + 80, yRow, { align: 'right' });
        doc.setFont('helvetica', 'normal');
    });
    
    y += 45;
    
    return y;
}

// Helper function to capture and add chart
async function addPDFChart(doc, y) {
    try {
        // Check if we need a new page
        if (y > 180) {
            doc.addPage();
            y = 20;
        }
        
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 0, 0);
        doc.text('🏆 Top 5 Instancias con Más Ventas', 15, y);
        
        y += 8;
        
        const canvas = document.getElementById('top5Chart');
        if (canvas) {
            // Capture the chart using html2canvas
            const chartImage = await html2canvas(canvas, {
                backgroundColor: '#ffffff',
                scale: 2
            });
            
            const imgData = chartImage.toDataURL('image/png');
            
            // Add image to PDF
            const imgWidth = 180;
            const imgHeight = (chartImage.height * imgWidth) / chartImage.width;
            
            doc.addImage(imgData, 'PNG', 15, y, imgWidth, imgHeight);
            
            y += imgHeight + 5;
        }
        
        return y;
    } catch (error) {
        console.error('Error adding chart to PDF:', error);
        y += 10;
        doc.setFontSize(10);
        doc.setTextColor(231, 76, 60);
        doc.text('(Error al cargar gráfico)', 15, y);
        return y + 10;
    }
}

// Helper function to add Top 5 table
function addPDFTop5Table(doc, y) {
    const sortedInstances = Object.values(instancesData)
        .sort((a, b) => b.total - a.total)
        .slice(0, 5);
    
    if (sortedInstances.length === 0) return y;
    
    // Check if we need a new page
    if (y > 220) {
        doc.addPage();
        y = 20;
    }
    
    y += 5;
    
    // Table header
    doc.setFillColor(74, 144, 226);
    doc.rect(15, y - 5, 180, 8, 'F');
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(255, 255, 255);
    doc.text('Rank', 20, y);
    doc.text('Instancia', 40, y);
    doc.text('Total Reservas', 170, y, { align: 'right' });
    
    y += 8;
    
    // Table rows
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(0, 0, 0);
    
    sortedInstances.forEach((instance, index) => {
        if (y > 270) {
            doc.addPage();
            y = 20;
        }
        
        // Alternating row colors
        if (index % 2 === 0) {
            doc.setFillColor(245, 247, 250);
            doc.rect(15, y - 5, 180, 8, 'F');
        }
        
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(74, 144, 226);
        doc.text(`#${index + 1}`, 20, y);
        
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(0, 0, 0);
        const nameText = truncateText(instance.name, 40);
        doc.text(nameText, 40, y);
        
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(39, 174, 96);
        doc.text(instance.total.toString(), 170, y, { align: 'right' });
        
        y += 8;
    });
    
    y += 5;
    
    return y;
}

// Helper function to add global conversion rates
function addPDFGlobalConversion(doc, y) {
    // Check if we need a new page
    if (y > 240) {
        doc.addPage();
        y = 20;
    }
    
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0, 0, 0);
    doc.text('📊 Tasas de Conversión Globales', 15, y);
    
    y += 8;
    
    const totalQuotes = parseInt(document.getElementById('totalQuotes').value) || 0;
    const globalConversionRate = document.getElementById('globalConversionRate').textContent;
    
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    
    // Total de Cotizaciones
    doc.setFillColor(245, 247, 250);
    doc.rect(15, y - 5, 180, 10, 'F');
    doc.setTextColor(127, 140, 141);
    doc.text('Total de Cotizaciones:', 17, y);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(74, 144, 226);
    doc.text(totalQuotes.toString(), 190, y, { align: 'right' });
    
    y += 12;
    
    // Conversion Rate
    doc.setFont('helvetica', 'normal');
    doc.setFillColor(245, 247, 250);
    doc.rect(15, y - 5, 180, 10, 'F');
    doc.setTextColor(127, 140, 141);
    doc.text('Total Cotizaciones / Total Confirmadas:', 17, y);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(74, 144, 226);
    doc.text(globalConversionRate, 190, y, { align: 'right' });
    
    y += 15;
    
    return y;
}

// Helper function to add instance details
async function addPDFInstanceDetails(doc, y) {
    // Check if we need a new page
    if (y > 240) {
        doc.addPage();
        y = 20;
    }
    
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0, 0, 0);
    doc.text('🏢 Detalles por Instancia', 15, y);
    
    y += 10;
    
    // Sort instances by total reservations
    const sortedInstances = Object.entries(instancesData)
        .sort(([, a], [, b]) => b.total - a.total);
    
    for (const [instanceName, data] of sortedInstances) {
        // Check if we need a new page
        if (y > 230) {
            doc.addPage();
            y = 20;
        }
        
        // Instance name and total
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(74, 144, 226);
        const nameText = truncateText(data.name, 50);
        doc.text(`🏢 ${nameText}`, 15, y);
        doc.text(`Total: ${data.total}`, 190, y, { align: 'right' });
        
        y += 8;
        
        // Status breakdown
        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        
        const statusData = [
            { label: 'Confirmada:', value: data.confirmada, color: [39, 174, 96], emoji: '✅' },
            { label: 'Pendiente:', value: data.pendiente, color: [243, 156, 18], emoji: '⏳' },
            { label: 'Fallida:', value: data.fallida, color: [231, 76, 60], emoji: '❌' },
            { label: 'Procesando:', value: data.procesando, color: [155, 89, 182], emoji: '🔄' }
        ];
        
        statusData.forEach((status, index) => {
            const xPos = 20 + (index % 2) * 90;
            const yPos = y + Math.floor(index / 2) * 8;
            
            doc.setTextColor(127, 140, 141);
            doc.text(`${status.emoji} ${status.label}`, xPos, yPos);
            doc.setFont('helvetica', 'bold');
            doc.setTextColor(...status.color);
            doc.text(status.value.toString(), xPos + 80, yPos, { align: 'right' });
            doc.setFont('helvetica', 'normal');
        });
        
        y += 20;
        
        // Quote information
        const quotes = quotesData[instanceName] || 0;
        const rate1 = quotes > 0 ? ((data.total / quotes) * 100).toFixed(2) : '0.00';
        const rate2 = quotes > 0 ? ((data.confirmada / quotes) * 100).toFixed(2) : '0.00';
        
        doc.setFillColor(245, 247, 250);
        doc.rect(15, y - 5, 180, 20, 'F');
        
        doc.setTextColor(127, 140, 141);
        doc.text('💼 Cantidad de Cotizaciones:', 17, y);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(74, 144, 226);
        doc.text(quotes.toString(), 190, y, { align: 'right' });
        
        y += 6;
        
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(127, 140, 141);
        doc.text('Total Reservas / Cotizaciones:', 17, y);
        doc.setFont('helvetica', 'bold');
        doc.text(`${rate1}%`, 190, y, { align: 'right' });
        
        y += 6;
        
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(127, 140, 141);
        doc.text('Confirmadas / Cotizaciones:', 17, y);
        doc.setFont('helvetica', 'bold');
        doc.text(`${rate2}%`, 190, y, { align: 'right' });
        
        y += 12;
        
        // Separator line
        doc.setDrawColor(225, 232, 237);
        doc.setLineWidth(0.3);
        doc.line(15, y, 195, y);
        
        y += 8;
    }
    
    return y;
}

// Helper function to add footer to all pages
function addPDFFooter(doc) {
    const pageCount = doc.internal.getNumberOfPages();
    const pageHeight = doc.internal.pageSize.height;
    
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        
        doc.setFontSize(8);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(150, 150, 150);
        
        // Page number
        doc.text(`Página ${i} de ${pageCount}`, 105, pageHeight - 10, { align: 'center' });
        
        // Generation info
        const now = new Date();
        const timestamp = `Generado: ${now.toLocaleDateString('es-ES')} ${now.toLocaleTimeString('es-ES')}`;
        doc.text(timestamp, 15, pageHeight - 10);
        
        // Branding
        doc.text('Generado por Calculom-tricas', 195, pageHeight - 10, { align: 'right' });
    }
}

// Helper function to get formatted date time for filename
function getFormattedDateTime() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    
    return `${year}-${month}-${day}-${hours}${minutes}${seconds}`;
}

// Helper function to show export loading
function showExportLoading() {
    const btn = document.getElementById('exportPdfBtn');
    const exportStatus = document.getElementById('exportStatus');
    
    btn.classList.add('loading');
    btn.disabled = true;
    
    const icon = btn.querySelector('.btn-icon');
    const text = btn.querySelector('.btn-text');
    
    icon.textContent = '⏳';
    text.textContent = 'Generando PDF...';
    
    exportStatus.textContent = '';
    exportStatus.classList.add('hidden');
}

// Helper function to hide export loading
function hideExportLoading() {
    const btn = document.getElementById('exportPdfBtn');
    btn.classList.remove('loading');
    btn.disabled = false;
    
    const icon = btn.querySelector('.btn-icon');
    const text = btn.querySelector('.btn-text');
    
    icon.textContent = '📄';
    text.textContent = 'Exportar PDF';
}

// Helper function to show success message
function showExportSuccess() {
    const exportStatus = document.getElementById('exportStatus');
    exportStatus.textContent = '✅ PDF generado exitosamente';
    exportStatus.classList.remove('hidden');
    
    setTimeout(() => {
        exportStatus.classList.add('hidden');
    }, 3000);
}

// Helper function to show error message
function showExportError() {
    const exportStatus = document.getElementById('exportStatus');
    exportStatus.textContent = '❌ Error al generar PDF';
    exportStatus.classList.remove('hidden');
    
    setTimeout(() => {
        exportStatus.classList.add('hidden');
    }, 3000);
}

// Helper function to truncate text
function truncateText(text, maxLength) {
    if (!text) return '';
    return text.length > maxLength ? text.substring(0, maxLength - 3) + '...' : text;
}