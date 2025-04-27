import * as XLSX from 'xlsx';

const fileInput = document.getElementById('excelFile');
const fileNameDisplay = document.getElementById('fileName');
const outputDiv = document.getElementById('output');

// SVG for WhatsApp Icon (inline SVG better than storing as string if possible, but this works)
const whatsappIconSvg = `<svg viewBox="0 0 24 24" width="1em" height="1em" focusable="false" aria-hidden="true" fill="currentColor"><path d="M16.75 13.96c-.25-.12-1.48-.72-1.71-.81-.23-.08-.39-.12-.56.12-.16.24-.65.81-.8 1-.15.18-.29.2-.54.08-.25-.12-1.02-.37-1.95-1.2-.73-.66-1.22-1.48-1.36-1.72-.14-.24-.01-.37.11-.49.11-.11.25-.29.37-.43.12-.14.16-.24.24-.4.08-.16.04-.29-.02-.41-.06-.12-.56-1.34-.76-1.84-.2-.48-.4-.42-.55-.42-.15,0-.31,0-.47,0-.16,0-.42.06-.64.31-.22.25-.86.84-.86,2.05,0,1.21.88,2.37,1,2.53.12.16,1.75,2.67,4.24,3.73.59.25,1.05.4,1.41.51.59.19,1.13.16,1.56.1.48-.07,1.48-.6,1.69-1.18.21-.58.21-1.07.15-1.18-.06-.11-.22-.17-.47-.29zm-5.23 6.11c-3.18 0-6.14-1.03-8.68-2.93l-1.28.39 1.31-1.25c-2.11-2.63-3.24-5.84-3.24-9.21C.01 5.77 5.77.01 12.43.01 19.1.01 24 5.77 24 12.43c0 6.67-5.77 12.43-12.43 12.43zm0-22.5c-5.54 0-10.09 4.55-10.09 10.09 0 3.17 1.46 6.04 3.88 7.99L3.13 23l2.1-.61c1.89 1.32 4.15 2.03 6.55 2.03 5.54 0 10.09-4.55 10.09-10.09S17.17 2.33 11.52 2.33z"></path></svg>`;

fileInput.addEventListener('change', handleFile);

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) {
        fileNameDisplay.textContent = 'Ningún archivo seleccionado';
        outputDiv.innerHTML = ''; // Clear previous results
        return;
    }

    fileNameDisplay.textContent = file.name;
    outputDiv.innerHTML = '<p>Procesando archivo...</p>'; // Show loading message

    const reader = new FileReader();

    reader.onload = function(e) {
        const data = e.target.result; // Read as ArrayBuffer for binary or string for text
        const fileName = file.name.toLowerCase();
        let workbook;

        try {
            if (fileName.endsWith('.csv')) {
                 // For CSV, read as text and parse
                workbook = XLSX.read(data, { type: 'array', raw: false });
            } else {
                // For XLSX/XLS, read as ArrayBuffer
                const arrayBuffer = new Uint8Array(data);
                workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, dateNF:'dd/mm/yyyy;@', raw: false });
            }

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
             // Use cellDates and dateNF for Excel, raw: false for formatted strings from CSV
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { cellDates: true, dateNF:'dd/mm/yyyy;@', raw: false });

            displayData(jsonData);
        } catch (error) {
            console.error("Error processing file:", error);
            outputDiv.innerHTML = `<p class="error">Error al leer el archivo ${fileName.split('.').pop().toUpperCase()}. Asegúrate de que el formato es correcto y las columnas esperadas existen (Nombre Equipo, Celular, Hora Partido, Dia Partido, Nombre Cancha).</p>`;
        }
    };

     reader.onerror = function(error) {
        console.error("File reading error:", error);
        outputDiv.innerHTML = '<p class="error">No se pudo leer el archivo.</p>';
    };

    // Read based on file type for efficiency/compatibility
    if (file.name.toLowerCase().endsWith('.csv')) {
         reader.readAsText(file); // Read CSV as text
    } else {
         reader.readAsArrayBuffer(file); // Read Excel as array buffer
    }
}

// Function to format date consistently (DD/MM/YYYY)
function formatExcelDate(date) {
    if (date instanceof Date && !isNaN(date)) {
        return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(date);
    }
    // Handle potential string dates from CSV or non-date Excel values
    if (typeof date === 'string') {
        // Attempt to parse common string formats like YYYY-MM-DD or MM/DD/YYYY
        const parsedDate = new Date(date);
         if (!isNaN(parsedDate) && parsedDate.getFullYear() > 1900) { // Basic validation
             return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(parsedDate);
         }
         // If it looks like DD/MM/YYYY already, return it
         if (/^\d{2}\/\d{2}\/\d{4}$/.test(date)) {
             return date;
         }
    }
    // Handle Excel number dates if sheet_to_json raw:false fails somehow
    if (typeof date === 'number') {
         const unixEpoch = new Date(1899, 11, 30); // Excel epoch
         const dateValue = new Date(unixEpoch.getTime() + date * 24 * 60 * 60 * 1000);
         if (!isNaN(dateValue)) {
             return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(dateValue);
         }
    }

    return String(date || 'N/A'); // Return as string if unparseable
}

// Function to format time consistently (HH:MM)
function formatExcelTime(timeValue) {
    if (timeValue instanceof Date && !isNaN(timeValue)) {
        return new Intl.DateTimeFormat('es-ES', { hour: '2-digit', minute: '2-digit', hour12: false }).format(timeValue);
    }
     // Handle Excel fractional time (0 to <1) or number interpreted as milliseconds
    if (typeof timeValue === 'number') {
         // Check if it's a fraction of a day
         if (timeValue >= 0 && timeValue < 1) {
            const totalSeconds = Math.round(timeValue * 86400);
            const hours = Math.floor(totalSeconds / 3600);
            const minutes = Math.floor((totalSeconds % 3600) / 60);
            return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
         } else if (timeValue > 1000 && timeValue < 86400000 * 2) { // Heuristic: large number, might be milliseconds
             const dateFromMillis = new Date(timeValue);
             if (!isNaN(dateFromMillis)) {
                 return new Intl.DateTimeFormat('es-ES', { hour: '2-digit', minute: '2-digit', hour12: false }).format(dateFromMillis);
             }
         }
    }
    // Handle potential string times from CSV (HH:MM, HH:MM:SS etc.)
    if (typeof timeValue === 'string') {
        const match = timeValue.match(/^(\d{1,2}):(\d{1,2})/);
        if (match) {
            return `${String(match[1]).padStart(2, '0')}:${String(match[2]).padStart(2, '0')}`;
        }
         // Attempt parsing ISO-like time strings
         try {
             const [hours, minutes] = timeValue.split(':').map(Number);
             if (!isNaN(hours) && !isNaN(minutes) && hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60) {
                  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
             }
         } catch (e) { /* Ignore parsing errors */ }
        // Return as string if not easily parsed time
        return timeValue;
    }
    return 'N/A';
}

function displayData(data) {
    outputDiv.innerHTML = '';

    if (!data || data.length === 0) {
        outputDiv.innerHTML = '<p>El archivo está vacío o no contiene datos reconocibles.</p>';
        return;
    }

    // Create controls div
    const controlsDiv = document.createElement('div');
    controlsDiv.className = 'data-controls';
    
    const resetBtn = document.createElement('button');
    resetBtn.className = 'reset-btn';
    resetBtn.textContent = 'Resetear Datos';
    resetBtn.onclick = () => {
        if (confirm('¿Estás seguro de que deseas borrar todos los datos? Esta acción no se puede deshacer.')) {
            localStorage.removeItem('tableData');
            outputDiv.innerHTML = '';
        }
    };
    controlsDiv.appendChild(resetBtn);
    outputDiv.appendChild(controlsDiv);

    let validRows = [];
    
    data.forEach((row, index) => {
        const normalizedRow = {};
        for (const key in row) {
            const normalizedKey = key.trim().toLowerCase().replace(/\s+/g, '');
            normalizedRow[normalizedKey] = row[key];
        }

        const teamName = normalizedRow['nombreequipo'] || 'N/A';
        let phone = String(normalizedRow['celular'] || '').replace(/\D/g, '');
        const rawMatchTime = normalizedRow['horapartido'];
        const rawMatchDay = normalizedRow['diapartido'];
        const fieldName = normalizedRow['nombrecancha'] || 'N/A';

        const formattedMatchDay = formatExcelDate(rawMatchDay);
        const formattedMatchTime = formatExcelTime(rawMatchTime);

        if (!phone || teamName === 'N/A' || formattedMatchDay === 'N/A' || formattedMatchTime === 'N/A' || fieldName === 'N/A') {
            return;
        }

        validRows.push({
            id: `row-${index}`,
            teamName,
            phone,
            matchDay: formattedMatchDay,
            matchTime: formattedMatchTime,
            fieldName,
            confirmed: false,
            messagesSent: { confirmation: false, reminder: false }
        });
    });

    // Sort rows by time
    validRows.sort((a, b) => {
        const dateComparison = a.matchDay.localeCompare(b.matchDay);
        if (dateComparison !== 0) return dateComparison;
        return a.matchTime.localeCompare(b.matchTime);
    });

    // Find scheduled matches
    const scheduledMatches = findScheduledMatches(validRows);

    // Save to localStorage
    localStorage.setItem('tableData', JSON.stringify(validRows));
    localStorage.setItem('scheduledMatches', JSON.stringify(scheduledMatches));

    // Create and display table
    createTable(validRows);
    displayScheduledMatches(scheduledMatches);
}

function findScheduledMatches(rows) {
    const matches = [];
    const matchMap = new Map();

    // Group teams by match details (time, day, field)
    rows.forEach(row => {
        const matchKey = `${row.matchDay}-${row.matchTime}-${row.fieldName}`;
        if (!matchMap.has(matchKey)) {
            matchMap.set(matchKey, []);
        }
        matchMap.get(matchKey).push(row);
    });

    // Find matches with exactly two teams
    for (const [key, teams] of matchMap.entries()) {
        if (teams.length === 2) {
            const [team1, team2] = teams;
            matches.push({
                matchDay: team1.matchDay,
                matchTime: team1.matchTime,
                fieldName: team1.fieldName,
                team1: {
                    name: team1.teamName,
                    confirmed: team1.confirmed
                },
                team2: {
                    name: team2.teamName,
                    confirmed: team2.confirmed
                }
            });
        }
    }

    return matches;
}

function displayScheduledMatches(matches) {
    const matchesSection = document.createElement('div');
    matchesSection.className = 'scheduled-matches';
    matchesSection.innerHTML = '<h2>Partidos Programados</h2>';

    if (matches.length === 0) {
        matchesSection.innerHTML += '<p>No hay partidos programados actualmente.</p>';
    } else {
        const matchesList = document.createElement('div');
        matchesList.className = 'matches-list';

        matches.forEach(match => {
            const matchCard = document.createElement('div');
            matchCard.className = `match-card ${match.team1.confirmed && match.team2.confirmed ? 'ready' : ''}`;
            
            matchCard.innerHTML = `
                <div class="match-header">
                    <span class="match-date">${match.matchDay}</span>
                    <span class="match-time">${match.matchTime}</span>
                    <span class="match-field">${match.fieldName}</span>
                </div>
                <div class="match-teams">
                    <div class="team ${match.team1.confirmed ? 'confirmed' : ''}">
                        <span class="team-name">${match.team1.name}</span>
                        <span class="confirmation-status">${match.team1.confirmed ? '✓ Confirmado' : 'Pendiente'}</span>
                    </div>
                    <div class="match-vs">VS</div>
                    <div class="team ${match.team2.confirmed ? 'confirmed' : ''}">
                        <span class="team-name">${match.team2.name}</span>
                        <span class="confirmation-status">${match.team2.confirmed ? '✓ Confirmado' : 'Pendiente'}</span>
                    </div>
                </div>
                ${match.team1.confirmed && match.team2.confirmed ? 
                    '<div class="match-status ready">¡Partido confirmado!</div>' : 
                    '<div class="match-status pending">Esperando confirmaciones</div>'}
            `;
            
            matchesList.appendChild(matchCard);
        });

        matchesSection.appendChild(matchesList);
    }

    outputDiv.appendChild(matchesSection);
}

function createTable(rows) {
    if (rows.length === 0) {
        outputDiv.innerHTML += '<p>No se encontraron filas con datos válidos.</p>';
        return;
    }

    const table = document.createElement('table');
    table.className = 'data-table';

    // Create header
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th>Equipo</th>
            <th>Celular</th>
            <th>Día</th>
            <th>Hora</th>
            <th>Cancha</th>
            <th>Acciones</th>
        </tr>
    `;
    table.appendChild(thead);

    // Create body
    const tbody = document.createElement('tbody');
    rows.forEach(row => {
        const tr = document.createElement('tr');
        // Add confirmed class if row is confirmed
        if (row.confirmed) {
            tr.className = 'confirmed';
        }
        tr.innerHTML = `
            <td>${row.teamName}</td>
            <td>${row.phone}</td>
            <td>${row.matchDay}</td>
            <td>${row.matchTime}</td>
            <td>${row.fieldName}</td>
            <td>
                <div class="table-actions">
                    <a href="https://wa.me/${row.phone}?text=${encodeURIComponent(`Hola ${row.teamName}, por favor confirma tu asistencia al partido del día ${row.matchDay} a las ${row.matchTime} en ${row.fieldName}. Responde SI o NO.`)}"
                       target="_blank"
                       class="whatsapp-link confirm ${row.messagesSent.confirmation ? 'sent' : ''}"
                       onclick="markMessageSent(event, '${row.id}', 'confirmation')">
                       ${whatsappIconSvg} Solicitar Confirmación
                    </a>
                    <a href="https://wa.me/${row.phone}?text=${encodeURIComponent(`Estamos listos para iniciar los preparativos previos al partido de *${row.teamName}*. Les pedimos acercarse a la mesa de *${row.fieldName}* por favor. Requisitos: Presentarse como minimo 15min antes del comienzo del partido, puntualidad, uniforme completo. ¡Te esperamos!`)}"
                       target="_blank"
                       class="whatsapp-link reminder ${row.messagesSent.reminder ? 'sent' : ''}"
                       onclick="markMessageSent(event, '${row.id}', 'reminder')">
                       ${whatsappIconSvg} Enviar Recordatorio
                    </a>
                    <button class="confirm-attendance-btn ${row.confirmed ? 'confirmed' : ''}"
                            onclick="toggleConfirmation('${row.id}')">
                            ${row.confirmed ? 'Confirmado' : 'Marcar Confirmado'}
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    outputDiv.appendChild(table);
}

window.markMessageSent = function(event, rowId, messageType) {
    const savedData = JSON.parse(localStorage.getItem('tableData'));
    const row = savedData.find(r => r.id === rowId);
    if (row) {
        row.messagesSent[messageType] = true;
        localStorage.setItem('tableData', JSON.stringify(savedData));
    }
    setTimeout(function() {
        location.reload();
    }, 3000);
};

window.toggleConfirmation = function(rowId) {
    const savedData = JSON.parse(localStorage.getItem('tableData'));
    const row = savedData.find(r => r.id === rowId);
    if (row) {
        row.confirmed = !row.confirmed;
        localStorage.setItem('tableData', JSON.stringify(savedData));
        
        // Update scheduled matches
        const scheduledMatches = findScheduledMatches(savedData);
        localStorage.setItem('scheduledMatches', JSON.stringify(scheduledMatches));
        
        // Refresh the display
        createTable(savedData);
        displayScheduledMatches(scheduledMatches);
    }
    location.reload();
};

// Add this to initialize the page with saved data
document.addEventListener('DOMContentLoaded', loadSavedData);

function loadSavedData() {
    const savedData = localStorage.getItem('tableData');
    const savedMatches = localStorage.getItem('scheduledMatches');
    
    if (savedData) {
        const parsedData = JSON.parse(savedData);
        createTable(parsedData);
        
        // If we have saved matches, display them
        if (savedMatches) {
            displayScheduledMatches(JSON.parse(savedMatches));
        }
    }
}