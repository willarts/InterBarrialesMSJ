import * as XLSX from 'xlsx';

const fileInput = document.getElementById('excelFile');
const fileNameDisplay = document.getElementById('fileName');
const sheetUrlInput = document.getElementById('sheetUrl');
const loadSheetButton = document.getElementById('loadSheetButton');
const outputDiv = document.getElementById('output');

// SVG for WhatsApp Icon remains the same
const whatsappIconSvg = `<svg viewBox="0 0 24 24" width="1em" height="1em" focusable="false" aria-hidden="true" fill="currentColor"><path d="M16.75 13.96c-.25-.12-1.48-.72-1.71-.81-.23-.08-.39-.12-.56.12-.16.24-.65.81-.8 1-.15.18-.29.2-.54.08-.25-.12-1.02-.37-1.95-1.2-.73-.66-1.22-1.48-1.36-1.72-.14-.24-.01-.37.11-.49.11-.11.25-.29.37-.43.12-.14.16-.24.24-.4.08-.16.04-.29-.02-.41-.06-.12-.56-1.34-.76-1.84-.2-.48-.4-.42-.55-.42-.15,0-.31,0-.47,0-.16,0-.42.06-.64.31-.22.25-.86.84-.86,2.05,0,1.21.88,2.37,1,2.53.12.16,1.75,2.67,4.24,3.73.59.25,1.05.4,1.41.51.59.19,1.13.16,1.56.1.48-.07,1.48-.6,1.69-1.18.21-.58.21-1.07.15-1.18-.06-.11-.22-.17-.47-.29zm-5.23 6.11c-3.18 0-6.14-1.03-8.68-2.93l-1.28.39 1.31-1.25c-2.11-2.63-3.24-5.84-3.24-9.21C.01 5.77 5.77.01 12.43.01 19.1.01 24 5.77 24 12.43c0 6.67-5.77 12.43-12.43 12.43zm0-22.5c-5.54 0-10.09 4.55-10.09 10.09 0 3.17 1.46 6.04 3.88 7.99L3.13 23l2.1-.61c1.89 1.32 4.15 2.03 6.55 2.03 5.54 0 10.09-4.55 10.09-10.09S17.17 2.33 11.52 2.33z"></path></svg>`;

fileInput.addEventListener('change', handleFile);
loadSheetButton.addEventListener('click', handleSheetUrl);

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) {
        fileNameDisplay.textContent = 'Ningún archivo seleccionado';
        // outputDiv.innerHTML = ''; // Don't clear output if just deselecting
        return;
    }

    // Clear URL input when a file is selected
    sheetUrlInput.value = '';
    fileNameDisplay.textContent = file.name;
    outputDiv.innerHTML = '<p>Procesando archivo...</p>'; // Show loading message

    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            // Use XLSX library for Excel files
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF:'dd/mm/yyyy;@', raw: false });
            processWorkbook(workbook);
        } catch (error) {
            console.error("Error processing Excel file:", error);
            outputDiv.innerHTML = '<p class="error">Error al leer el archivo Excel. Asegúrate de que el formato es correcto y las columnas esperadas existen (Nombre Equipo, Celular, Hora Partido, Dia Partido, Nombre Cancha).</p>';
        }
    };

    reader.onerror = function(error) {
        console.error("File reading error:", error);
        outputDiv.innerHTML = '<p class="error">No se pudo leer el archivo.</p>';
    };

    reader.readAsArrayBuffer(file);
}

async function handleSheetUrl() {
    const url = sheetUrlInput.value.trim();
    if (!url) {
        outputDiv.innerHTML = '<p class="error">Por favor, ingresa una URL de Google Sheets.</p>';
        return;
    }

    // Clear file input when URL is used
    fileInput.value = ''; // Reset file input
    fileNameDisplay.textContent = 'Ningún archivo seleccionado';
    outputDiv.innerHTML = '<p>Cargando datos desde Google Sheet...</p>';

    // Regular expression to extract the Sheet ID from various Google Sheet URL formats
    const sheetIdRegex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
    const match = url.match(sheetIdRegex);

    if (!match || !match[1]) {
        outputDiv.innerHTML = '<p class="error">URL de Google Sheet no válida. Asegúrate de que contenga "/spreadsheets/d/ID_DE_LA_HOJA".</p>';
        return;
    }

    const sheetId = match[1];
    // Construct the CSV export URL (exports the first visible sheet by default)
    // Note: The sheet must be public or "anyone with the link can view" for this to work without authentication.
    const csvExportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`;

    try {
        const response = await fetch(csvExportUrl);
        if (!response.ok) {
            throw new Error(`Error al descargar la hoja de cálculo (Código: ${response.status}). Asegúrate de que la hoja sea pública o accesible con el enlace.`);
        }
        const csvText = await response.text();

        // Use XLSX library to parse the fetched CSV text
        // Important: Set cellDates: false for CSV initially, date/time parsing handled later
        const workbook = XLSX.read(csvText, { type: 'string', raw: false }); // Let displayData handle date/time conversion
        processWorkbook(workbook);

    } catch (error) {
        console.error("Error loading/processing Google Sheet:", error);
        outputDiv.innerHTML = `<p class="error">Error al cargar desde Google Sheet: ${error.message}</p>`;
    }
}

// Central function to process workbook data (from file or URL)
function processWorkbook(workbook) {
     try {
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        // Explicitly tell sheet_to_json about date format expectation for XLSX source
        // For CSV source, date/time parsing relies more on formatExcelDate/Time functions
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            // raw: false helps interpret numbers/dates, but we double-check formats below
             raw: false,
             // cellDates: true // Keep this potentially, or rely solely on custom formatters
        });

        displayData(jsonData);
    } catch (error) {
        console.error("Error converting sheet to JSON:", error);
        outputDiv.innerHTML = '<p class="error">Error al procesar los datos de la hoja. Verifica el formato.</p>';
    }
}


// Function to format date consistently (DD/MM/YYYY)
function formatExcelDate(dateInput) {
    if (dateInput instanceof Date && !isNaN(dateInput)) {
        // Format Date object directly
        return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(dateInput);
    }

    // Handle numbers (Excel serial dates) - Approximation
    if (typeof dateInput === 'number' && dateInput > 1) {
         try {
             // Excel serial date starts from 1 (for 1/1/1900), JS Date starts from 0ms (1/1/1970 UTC)
             // Convert Excel serial date number to milliseconds since Unix epoch
             // Subtract 25569: (days between 1/1/1900 and 1/1/1970) + adjustment for Excel's leap year bug
             const utc_days = Math.floor(dateInput - 25569);
             const utc_value = utc_days * 86400; // Seconds
             const dateInfo = new Date(utc_value * 1000); // Convert seconds to milliseconds

             // Adjust for timezone offset to get local date parts
             const localDate = new Date(dateInfo.getTime() + (dateInfo.getTimezoneOffset() * 60000));

             if (!isNaN(localDate)) {
                 return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(localDate);
             }
         } catch (e) { /* Fall through if conversion fails */ }
    }

    // Handle strings (attempt parsing common formats)
    if (typeof dateInput === 'string') {
        // Try DD/MM/YYYY or YYYY-MM-DD first
        let parsedDate = null;
        const partsDMY = dateInput.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
        const partsYMD = dateInput.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);

        if (partsDMY) {
            // Assume DD/MM/YYYY
            const year = partsDMY[3].length === 2 ? parseInt('20' + partsDMY[3]) : parseInt(partsDMY[3]);
            parsedDate = new Date(year, parseInt(partsDMY[2]) - 1, parseInt(partsDMY[1]));
        } else if (partsYMD) {
            // Assume YYYY-MM-DD
            parsedDate = new Date(parseInt(partsYMD[1]), parseInt(partsYMD[2]) - 1, parseInt(partsYMD[3]));
        } else {
             // Try generic parsing (less reliable for specific formats)
             try {
                 parsedDate = new Date(dateInput);
             } catch (e) { /* Ignore */ }
        }

        if (parsedDate instanceof Date && !isNaN(parsedDate)) {
            return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(parsedDate);
        }
        // If it looks like DD/MM/YYYY but failed Date parsing (e.g. 31/04/2023), return original string
        if (partsDMY || /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(dateInput)) {
            return dateInput;
        }
    }

    // Fallback if input is not recognized
    return String(dateInput || 'N/A');
}

// Function to format time consistently (HH:MM)
function formatExcelTime(timeInput) {
    // Handle Date objects (if time was parsed as part of a date)
    if (timeInput instanceof Date && !isNaN(timeInput)) {
        return new Intl.DateTimeFormat('es-ES', { hour: '2-digit', minute: '2-digit', hour12: false }).format(timeInput);
    }

    // Handle numbers (Excel time fractions: 0.0 to 1.0)
    if (typeof timeInput === 'number' && timeInput >= 0 && timeInput < 1) {
        // Convert Excel time fraction to HH:MM
        const totalSeconds = Math.round(timeInput * 86400);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }

    // Handle strings (HH:MM, H:MM, HH:MM:SS, etc.)
    if (typeof timeInput === 'string') {
        // Match HH:MM or H:MM at the start of the string
        const match = timeInput.match(/^(\d{1,2}):(\d{1,2})/);
        if (match) {
            return `${String(match[1]).padStart(2, '0')}:${String(match[2]).padStart(2, '0')}`;
        }
         // Handle potential time strings from Date.toString() or ISO strings
         try {
             const parsedDate = new Date(`1970-01-01T${timeInput}Z`); // Try parsing as time part
             if (!isNaN(parsedDate)) {
                  return new Intl.DateTimeFormat('es-ES', { hour: '2-digit', minute: '2-digit', hour12: false }).format(parsedDate);
             }
         } catch(e) { /* ignore */ }

        // Return the original string if it looks like a time but wasn't parsed
         if (/\d{1,2}:\d{2}/.test(timeInput)) {
            return timeInput;
         }
    }

    // Fallback
    return String(timeInput || 'N/A');
}


function displayData(data) {
    outputDiv.innerHTML = '';

    if (!data || data.length === 0) {
        outputDiv.innerHTML = '<p>No se encontraron datos procesables en la fuente seleccionada.</p>';
        return;
    }

    let validRowsFound = false;

    data.forEach((row, index) => {
        const normalizedRow = {};
        // Normalize keys: lowercase, remove spaces
        for (const key in row) {
            if (Object.hasOwnProperty.call(row, key)) {
                const normalizedKey = String(key).trim().toLowerCase().replace(/\s+/g, '');
                normalizedRow[normalizedKey] = row[key];
            }
        }

        // Log normalized keys for debugging
        // console.log("Normalized keys:", Object.keys(normalizedRow));

        const teamName = normalizedRow['nombreequipo'] || 'N/A';
        let phone = normalizedRow['celular'] || '';
        const rawMatchTime = normalizedRow['horapartido']; // Keep raw value for flexible formatting
        const rawMatchDay = normalizedRow['diapartido'];   // Keep raw value for flexible formatting
        const fieldName = normalizedRow['nombrecancha'] || 'N/A';

        phone = String(phone).replace(/\D/g, ''); // Clean phone number

        // Apply formatting functions
        const formattedMatchDay = formatExcelDate(rawMatchDay);
        const formattedMatchTime = formatExcelTime(rawMatchTime);

        // Check for essential data *after* formatting attempts
        if (!phone || phone.length < 8 || teamName === 'N/A' || formattedMatchDay === 'N/A' || formattedMatchTime === 'N/A' || fieldName === 'N/A') {
             console.warn("Skipping row due to missing/invalid essential data:",
                { original: row, normalized: normalizedRow, phone, teamName, formattedMatchDay, formattedMatchTime, fieldName });
            return; // Skip row if essential data is missing or unparseable
        }

        validRowsFound = true;

        const card = document.createElement('div');
        card.className = 'player-card';
        card.dataset.id = `player-${index}`; // Add a unique identifier

        card.innerHTML = `
            <h3>${teamName}</h3>
            <p><strong>Celular:</strong> ${phone}</p>
            <p><strong>Día:</strong> ${formattedMatchDay}</p>
            <p><strong>Hora:</strong> ${formattedMatchTime}</p>
            <p><strong>Cancha:</strong> ${fieldName}</p>
            <div class="card-actions">
                <!-- Links and button will be added here -->
            </div>
        `;

        const actionsContainer = card.querySelector('.card-actions');

        // Generate WhatsApp message texts
        const confirmationText = encodeURIComponent(`Hola ${teamName}, por favor confirma tu asistencia al partido del día ${formattedMatchDay} a las ${formattedMatchTime} en ${fieldName}. Responde SI o NO.`);
        const confirmationLink = `https://wa.me/${phone}?text=${confirmationText}`;

        const reminderText = encodeURIComponent(`Recordatorio: Partido hoy ${formattedMatchDay} a las ${formattedMatchTime} en ${fieldName}. Requisitos: puntualidad, uniforme completo. ¡Te esperamos!`);
        const reminderLink = `https://wa.me/${phone}?text=${reminderText}`;

        // Create "Confirmación" link/button
        const confirmBtn = document.createElement('a');
        confirmBtn.href = confirmationLink;
        confirmBtn.target = "_blank";
        confirmBtn.className = "whatsapp-link confirm";
        confirmBtn.innerHTML = `${whatsappIconSvg} Enviar Solicitud Confirmación`;
        confirmBtn.addEventListener('click', function(e) {
            // Visual feedback, doesn't prevent link opening
            this.classList.add('sent-feedback'); // Use a different class to avoid disabling
             this.innerHTML = `${whatsappIconSvg} Abriendo WhatsApp...`;
             setTimeout(() => {
                 this.classList.remove('sent-feedback');
                 this.innerHTML = `${whatsappIconSvg} Enviar Solicitud Confirmación`;
             }, 2500); // Revert after a delay
        });
        actionsContainer.appendChild(confirmBtn);

        // Create "Recordatorio" link/button
        const reminderBtn = document.createElement('a');
        reminderBtn.href = reminderLink;
        reminderBtn.target = "_blank";
        reminderBtn.className = "whatsapp-link reminder";
        reminderBtn.innerHTML = `${whatsappIconSvg} Enviar Recordatorio`;
         reminderBtn.addEventListener('click', function(e) {
            this.classList.add('sent-feedback');
            this.innerHTML = `${whatsappIconSvg} Abriendo WhatsApp...`;
             setTimeout(() => {
                  this.classList.remove('sent-feedback');
                 this.innerHTML = `${whatsappIconSvg} Enviar Recordatorio`;
             }, 2500);
        });
        actionsContainer.appendChild(reminderBtn);

        // Create Manual Confirmation Button
        const manualConfirmBtn = document.createElement('button');
        manualConfirmBtn.className = 'confirm-attendance-btn';
        manualConfirmBtn.textContent = 'Marcar Confirmado'; // Initial text
        manualConfirmBtn.type = 'button'; // Ensure it's not a submit button

        manualConfirmBtn.addEventListener('click', function() {
            this.classList.toggle('confirmed');
            if (this.classList.contains('confirmed')) {
                this.textContent = 'Confirmado'; // Update text when confirmed
            } else {
                this.textContent = 'Marcar Confirmado'; // Revert text when unconfirmed
            }
            // No backend, so this state is visual only for the current session
        });

        actionsContainer.appendChild(manualConfirmBtn);


        outputDiv.appendChild(card);
    });

    if (!validRowsFound) {
        // Display error if data was present but no valid rows were found after processing
        outputDiv.innerHTML = '<p class="error">No se encontraron filas con datos válidos completos (Equipo, Celular, Día, Hora, Cancha) después de procesar. Por favor, revisa el contenido y formato de las columnas/celdas en tu archivo o Google Sheet.</p>';
    }
}