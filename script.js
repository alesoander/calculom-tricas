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

// Helper: parse a cell value as a number, returning 0 if not numeric
function parseNumeric(val) {
    const n = parseFloat(val);
    return isNaN(n) ? 0 : n;
}

// Helper: normalise a Yes/No-style cell to a display string
function parseBool(val) {
    if (val === undefined || val === null || val === '') return '—';
    return String(val).trim();
}

// Process data with the new A→AK column schema
function processReservations() {
    instancesData = {};

    reservationData.forEach(row => {
        // Column B = Conexión is the primary key for each connection/instance
        const conexion = (row.B || 'Sin Conexión').toString().trim();

        instancesData[conexion] = {
            // Identity
            name: conexion,
            cliente: (row.A || '').toString().trim(),
            estado: (row.C || '').toString().trim(),
            api: (row.D || '').toString().trim(),
            equipo: (row.F || '').toString().trim(),
            plataforma: (row.G || '').toString().trim(),

            // Monthly metrics
            conversacionesMensual: parseNumeric(row.H),
            cantidadSeguimiento: parseNumeric(row.P),
            clicksReservaMensual: parseNumeric(row.R),

            // Feature flags
            automatizacion: parseBool(row.I),
            redesSociales: parseBool(row.J),
            preDisp: parseBool(row.K),
            disponible247: parseBool(row.L),
            convAutoManual: parseBool(row.M),
            seguimiento1: parseBool(row.N),
            seguimiento2: parseBool(row.O),
            campana: parseBool(row.Q),
            motor: parseBool(row.S),
            primerPausaModificado: parseBool(row.T),
            segundaPausaActivado: parseBool(row.U),
            entrenamientosFaltantes: parseBool(row.V),
            esMulticomplejo: parseBool(row.W),
            correoAlertas: parseBool(row.X),
            progresoCompletado: parseBool(row.Y),

            // Reviews
            reviewsTotal: parseNumeric(row.Z),
            reviewsCorregido: parseNumeric(row.AA),
            reviewsRecibido: parseNumeric(row.AB),
            reviewsPendiente: parseNumeric(row.AC),
            reviewsIrresoluble: parseNumeric(row.AD),
            reviewsErrorCount: parseNumeric(row.AE),
            reviewsErrorRate: parseNumeric(row.AF),

            // Support & surveys
            nivelSoporte: (row.AG || '').toString().trim(),
            pruebasCliente: parseBool(row.AH),
            calificacionEncuesta: parseNumeric(row.AI),
            respuestasEncuesta: parseNumeric(row.AJ),
            ultimaDesconexion: (row.AK || '').toString().trim(),

            total: 1
        };
    });

    // Keep quotes data consistent (not used in new schema but preserved for PDF export)
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
    const clients = Object.values(instancesData);

    const totalClients = clients.length;
    const totalConversaciones = clients.reduce((s, c) => s + c.conversacionesMensual, 0);
    const totalClicksReserva = clients.reduce((s, c) => s + c.clicksReservaMensual, 0);
    const totalReviews = clients.reduce((s, c) => s + c.reviewsTotal, 0);

    const calificaciones = clients.filter(c => c.calificacionEncuesta > 0);
    const avgCalificacion = calificaciones.length > 0
        ? (calificaciones.reduce((s, c) => s + c.calificacionEncuesta, 0) / calificaciones.length).toFixed(2)
        : '—';

    document.getElementById('totalReservations').textContent = totalClients;
    document.getElementById('totalConfirmed').textContent = totalConversaciones.toLocaleString('es-ES');
    document.getElementById('totalPending').textContent = totalClicksReserva.toLocaleString('es-ES');
    document.getElementById('totalFailed').textContent = totalReviews.toLocaleString('es-ES');
    document.getElementById('totalProcessing').textContent = avgCalificacion;
    document.getElementById('totalInstances').textContent = Object.keys(instancesData).length;
}

// Display top 5 instances
function displayTop5Instances() {
    const sortedInstances = Object.values(instancesData)
        .sort((a, b) => b.conversacionesMensual - a.conversacionesMensual)
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
            labels: sortedInstances.map(i => i.cliente || i.name),
            datasets: [{
                label: 'Conversaciones Mensuales',
                data: sortedInstances.map(i => i.conversacionesMensual),
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
            <div class="top5-name">${instance.cliente || instance.name}</div>
            <div class="top5-count">${instance.conversacionesMensual.toLocaleString('es-ES')} conv.</div>
        </div>
    `).join('');
}

// Display global metrics summary (replaces old conversion section)
function displayGlobalConversion() {
    updateGlobalConversion();
}

function updateGlobalConversion() {
    const clients = Object.values(instancesData);
    const totalQuotes = parseInt(document.getElementById('totalQuotes').value) || 0;

    // Avg reviews error rate
    const withRate = clients.filter(c => c.reviewsErrorRate > 0);
    const avgErrorRate = withRate.length > 0
        ? (withRate.reduce((s, c) => s + c.reviewsErrorRate, 0) / withRate.length).toFixed(2)
        : 0;

    // Total confirmed = totalConversaciones for global conversion calc
    const totalConversaciones = clients.reduce((s, c) => s + c.conversacionesMensual, 0);
    const conversionRate = totalConversaciones > 0 && totalQuotes > 0
        ? ((totalConversaciones / totalQuotes) * 100).toFixed(2)
        : '0.00';

    document.getElementById('globalConversionRate').textContent = `${conversionRate}%`;
    
    // Show avg error rate if element exists
    const avgErrorEl = document.getElementById('avgErrorRate');
    if (avgErrorEl) avgErrorEl.textContent = `${avgErrorRate}%`;
}

// Display instance details
function displayInstanceDetails() {
    const container = document.getElementById('instancesContainer');

    // Sort by conversaciones mensuales desc
    const sortedInstances = Object.entries(instancesData)
        .sort(([, a], [, b]) => b.conversacionesMensual - a.conversacionesMensual);

    container.innerHTML = sortedInstances.map(([, data]) => `
        <div class="instance-card">
            <div class="instance-header">
                <div class="instance-name">🏢 ${data.cliente || data.name}</div>
                <div class="instance-total">Conexión: ${data.name}</div>
            </div>

            <!-- Identity -->
            <div class="instance-stats">
                <div class="stat-item">
                    <span class="stat-label">📡 Estado:</span>
                    <span class="stat-value">${data.estado || '—'}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🔌 API:</span>
                    <span class="stat-value">${data.api || '—'}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">👥 Equipo:</span>
                    <span class="stat-value">${data.equipo || '—'}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📱 Plataforma:</span>
                    <span class="stat-value">${data.plataforma || '—'}</span>
                </div>
            </div>

            <!-- Monthly metrics -->
            <div class="instance-section-title">📊 Métricas Mensuales</div>
            <div class="instance-stats">
                <div class="stat-item">
                    <span class="stat-label">💬 Conversaciones:</span>
                    <span class="stat-value confirmed">${data.conversacionesMensual.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🔁 Seguimientos:</span>
                    <span class="stat-value">${data.cantidadSeguimiento.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🖱️ Clicks Reserva:</span>
                    <span class="stat-value">${data.clicksReservaMensual.toLocaleString('es-ES')}</span>
                </div>
            </div>

            <!-- Feature flags -->
            <div class="instance-section-title">⚙️ Configuración</div>
            <div class="instance-stats flags-grid">
                <div class="stat-item">
                    <span class="stat-label">🤖 Automatización:</span>
                    <span class="stat-value">${data.automatizacion}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📲 Redes Sociales:</span>
                    <span class="stat-value">${data.redesSociales}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📋 Pre/Dispo:</span>
                    <span class="stat-value">${data.preDisp}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🕐 24/7:</span>
                    <span class="stat-value">${data.disponible247}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🔀 Conv. Auto/Manual:</span>
                    <span class="stat-value">${data.convAutoManual}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📬 Seguimiento 1:</span>
                    <span class="stat-value">${data.seguimiento1}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📬 Seguimiento 2:</span>
                    <span class="stat-value">${data.seguimiento2}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📢 Campaña:</span>
                    <span class="stat-value">${data.campana}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">⚙️ Motor:</span>
                    <span class="stat-value">${data.motor}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">⏸️ Pausa 1 Mod.:</span>
                    <span class="stat-value">${data.primerPausaModificado}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">⏸️ Pausa 2 Act.:</span>
                    <span class="stat-value">${data.segundaPausaActivado}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🎓 Entrenam. Faltantes:</span>
                    <span class="stat-value">${data.entrenamientosFaltantes}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🏬 Multicomplejo:</span>
                    <span class="stat-value">${data.esMulticomplejo}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📧 Correo Alertas:</span>
                    <span class="stat-value">${data.correoAlertas}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">✅ Progreso Completado:</span>
                    <span class="stat-value">${data.progresoCompletado}</span>
                </div>
            </div>

            <!-- Reviews -->
            <div class="instance-section-title">⭐ Reviews</div>
            <div class="instance-stats">
                <div class="stat-item">
                    <span class="stat-label">📊 Total:</span>
                    <span class="stat-value">${data.reviewsTotal.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">✔️ Corregido:</span>
                    <span class="stat-value confirmed">${data.reviewsCorregido.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📥 Recibido:</span>
                    <span class="stat-value">${data.reviewsRecibido.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">⏳ Pendiente:</span>
                    <span class="stat-value pending">${data.reviewsPendiente.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🔒 Irresoluble:</span>
                    <span class="stat-value failed">${data.reviewsIrresoluble.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">❌ Error Count:</span>
                    <span class="stat-value failed">${data.reviewsErrorCount.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📉 Error Rate:</span>
                    <span class="stat-value ${data.reviewsErrorRate > 0 ? 'failed' : ''}">${data.reviewsErrorRate}%</span>
                </div>
            </div>

            <!-- Support & survey -->
            <div class="instance-section-title">🎯 Soporte & Encuesta</div>
            <div class="instance-stats">
                <div class="stat-item">
                    <span class="stat-label">🛡️ Nivel de Soporte:</span>
                    <span class="stat-value">${data.nivelSoporte || '—'}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">🧪 Pruebas del Cliente:</span>
                    <span class="stat-value">${data.pruebasCliente}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">⭐ Calificación Encuesta:</span>
                    <span class="stat-value confirmed">${data.calificacionEncuesta > 0 ? data.calificacionEncuesta : '—'}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📝 Respuestas Encuesta:</span>
                    <span class="stat-value">${data.respuestasEncuesta.toLocaleString('es-ES')}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📵 Última Desconexión WA:</span>
                    <span class="stat-value">${data.ultimaDesconexion || '—'}</span>
                </div>
            </div>
        </div>
    `).join('');
}

function getConversionHTML(instanceName, quotes) {
    const data = instancesData[instanceName];
    if (!data) return '';

    const rate1 = quotes > 0
        ? ((data.conversacionesMensual / quotes) * 100).toFixed(2)
        : '0.00';

    return `
        <div class="conversion-result-item">
            <span class="conversion-result-label">Conversaciones Mensuales / Cotizaciones:</span>
            <span class="conversion-result-value">${rate1}%</span>
        </div>
    `;
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
        const fechaCreacion = row.E; // Column E = Fecha Creación
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
        showError('No se encontraron registros en el rango de fechas seleccionado');
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
        Mostrando <strong>${reservationData.length}</strong> registros del 
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
        const filename = `reporte-metricas-${getFormattedDateTime()}.pdf`;
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
    doc.text('Análisis de Métricas de Clientes - Reporte', 105, y, { align: 'center' });
    
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
    doc.text('Resumen General', 15, y);

    y += 8;

    const clients = Object.values(instancesData);
    const totalClients = clients.length;
    const totalConversaciones = clients.reduce((s, c) => s + c.conversacionesMensual, 0);
    const totalClicksReserva = clients.reduce((s, c) => s + c.clicksReservaMensual, 0);
    const totalReviews = clients.reduce((s, c) => s + c.reviewsTotal, 0);
    const calificaciones = clients.filter(c => c.calificacionEncuesta > 0);
    const avgCalificacion = calificaciones.length > 0
        ? (calificaciones.reduce((s, c) => s + c.calificacionEncuesta, 0) / calificaciones.length).toFixed(2)
        : '—';

    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');

    const summaryData = [
        { label: 'Total Clientes:', value: totalClients, color: [74, 144, 226] },
        { label: 'Conversaciones (Mensual):', value: totalConversaciones.toLocaleString('es-ES'), color: [39, 174, 96] },
        { label: 'Clicks Reserva (Mensual):', value: totalClicksReserva.toLocaleString('es-ES'), color: [243, 156, 18] },
        { label: 'Reviews Total:', value: totalReviews.toLocaleString('es-ES'), color: [74, 144, 226] },
        { label: 'Prom. Calificacion Encuesta:', value: avgCalificacion, color: [155, 89, 182] },
        { label: 'Total Conexiones:', value: Object.keys(instancesData).length, color: [74, 144, 226] }
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
        doc.text('Top 5 Clientes por Conversaciones Mensuales', 15, y);
        
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
        .sort((a, b) => b.conversacionesMensual - a.conversacionesMensual)
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
    doc.text('Cliente', 40, y);
    doc.text('Conv. Mensual', 170, y, { align: 'right' });

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
        const nameText = truncateText(instance.cliente || instance.name, 40);
        doc.text(nameText, 40, y);

        doc.setFont('helvetica', 'bold');
        doc.setTextColor(39, 174, 96);
        doc.text(instance.conversacionesMensual.toLocaleString('es-ES'), 170, y, { align: 'right' });

        y += 8;
    });

    y += 5;

    return y;
}

// Helper function to add global metrics
function addPDFGlobalConversion(doc, y) {
    // Check if we need a new page
    if (y > 240) {
        doc.addPage();
        y = 20;
    }

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0, 0, 0);
    doc.text('Metricas Globales de Conversion', 15, y);

    y += 8;

    const totalQuotes = parseInt(document.getElementById('totalQuotes').value) || 0;
    const globalConversionRate = document.getElementById('globalConversionRate').textContent;
    const avgErrorRateEl = document.getElementById('avgErrorRate');
    const avgErrorRate = avgErrorRateEl ? avgErrorRateEl.textContent : '0.00%';

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
    doc.text('Conversaciones Mensuales / Cotizaciones:', 17, y);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(74, 144, 226);
    doc.text(globalConversionRate, 190, y, { align: 'right' });

    y += 12;

    // Avg error rate
    doc.setFont('helvetica', 'normal');
    doc.setFillColor(245, 247, 250);
    doc.rect(15, y - 5, 180, 10, 'F');
    doc.setTextColor(127, 140, 141);
    doc.text('Promedio Reviews Error Rate:', 17, y);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(231, 76, 60);
    doc.text(avgErrorRate, 190, y, { align: 'right' });

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
    doc.text('Detalles por Conexion', 15, y);

    y += 10;

    // Sort by conversaciones mensuales desc
    const sortedInstances = Object.entries(instancesData)
        .sort(([, a], [, b]) => b.conversacionesMensual - a.conversacionesMensual);

    for (const [, data] of sortedInstances) {
        // Check if we need a new page
        if (y > 220) {
            doc.addPage();
            y = 20;
        }

        // Client name and connection
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(74, 144, 226);
        const clientText = truncateText(data.cliente || data.name, 45);
        doc.text(clientText, 15, y);
        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(100, 100, 100);
        doc.text(`Conexion: ${truncateText(data.name, 30)}`, 15, y + 5);

        y += 12;

        // Identity row
        doc.setFontSize(9);
        const identityData = [
            { label: 'Estado:', value: data.estado || '—' },
            { label: 'API:', value: data.api || '—' },
            { label: 'Equipo:', value: data.equipo || '—' },
            { label: 'Nivel Soporte:', value: data.nivelSoporte || '—' }
        ];
        identityData.forEach((item, i) => {
            const xPos = 15 + (i % 2) * 90;
            const yPos = y + Math.floor(i / 2) * 7;
            doc.setFont('helvetica', 'normal');
            doc.setTextColor(127, 140, 141);
            doc.text(`${item.label}`, xPos, yPos);
            doc.setFont('helvetica', 'bold');
            doc.setTextColor(0, 0, 0);
            doc.text(truncateText(item.value, 20), xPos + 30, yPos);
        });

        y += 16;

        // Monthly metrics
        if (y > 270) { doc.addPage(); y = 20; }
        doc.setFillColor(245, 247, 250);
        doc.rect(15, y - 4, 180, 8, 'F');
        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(127, 140, 141);
        doc.text('Conv. Mensual:', 17, y);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(39, 174, 96);
        doc.text(data.conversacionesMensual.toLocaleString('es-ES'), 70, y);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(127, 140, 141);
        doc.text('Clicks Reserva:', 100, y);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(243, 156, 18);
        doc.text(data.clicksReservaMensual.toLocaleString('es-ES'), 145, y);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(127, 140, 141);
        doc.text('Seguimientos:', 160, y);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(74, 144, 226);
        doc.text(data.cantidadSeguimiento.toLocaleString('es-ES'), 192, y, { align: 'right' });

        y += 10;

        // Reviews row
        if (y > 270) { doc.addPage(); y = 20; }
        doc.setFillColor(250, 245, 245);
        doc.rect(15, y - 4, 180, 8, 'F');
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(127, 140, 141);
        doc.text(`Reviews: Total=${data.reviewsTotal}  Corregido=${data.reviewsCorregido}  Recibido=${data.reviewsRecibido}  Pendiente=${data.reviewsPendiente}  Error=${data.reviewsErrorRate}%`, 17, y);

        y += 10;

        // Survey row
        if (y > 270) { doc.addPage(); y = 20; }
        doc.setFillColor(245, 250, 245);
        doc.rect(15, y - 4, 180, 8, 'F');
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(127, 140, 141);
        const calif = data.calificacionEncuesta > 0 ? data.calificacionEncuesta : '—';
        doc.text(`Calif. Encuesta: ${calif}  |  Respuestas: ${data.respuestasEncuesta}  |  Ult. Desconexion WA: ${truncateText(data.ultimaDesconexion || '—', 25)}`, 17, y);

        y += 10;

        // Separator line
        doc.setDrawColor(225, 232, 237);
        doc.setLineWidth(0.3);
        doc.line(15, y, 195, y);

        y += 6;
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