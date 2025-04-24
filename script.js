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
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { cellDates: true, dateNF:'dd/mm/yyyy;@', raw: false });

            displayData(jsonData);
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

// Function to format date consistently (DD/MM/YYYY)
function formatExcelDate(date) {
    if (date instanceof Date && !isNaN(date)) {
        return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(date);
    }
    if (typeof date === 'number' || typeof date === 'string') {
        try {
            const parsedDate = new Date(date);
            if (!isNaN(parsedDate)) {
                 return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(parsedDate);
            }
        } catch (e) { /* Ignore parsing errors */ }
    }
    return String(date || 'N/A'); // Return as string if unparseable
}

// Function to format time consistently (HH:MM)
function formatExcelTime(timeValue) {
    if (timeValue instanceof Date && !isNaN(timeValue)) {
        return new Intl.DateTimeFormat('es-ES', { hour: '2-digit', minute: '2-digit', hour12: false }).format(timeValue);
    }
    if (typeof timeValue === 'number' && timeValue >= 0 && timeValue < 1) {
        const totalSeconds = Math.round(timeValue * 86400);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }
    if (typeof timeValue === 'string') {
        const match = timeValue.match(/^(\d{1,2}):(\d{1,2})/);
        if (match) {
            return `${String(match[1]).padStart(2, '0')}:${String(match[2]).padStart(2, '0')}`;
        }
        // Handle potential ISO time strings or just return string
        if (/\d{2}:\d{2}(:\d{2})?/.test(timeValue)) {
             return timeValue.substring(0, 5); // Assume HH:MM start
        }
        return timeValue; // Return as string if not easily parsed time
    }
    return 'N/A';
}


function displayData(data) {
    outputDiv.innerHTML = '';

    if (!data || data.length === 0) {
        outputDiv.innerHTML = '<p>El archivo está vacío o no contiene datos reconocibles.</p>';
        return;
    }

    let validRowsFound = false;

    data.forEach((row, index) => { // Add index for unique IDs if needed
        const normalizedRow = {};
        for (const key in row) {
            if (Object.hasOwnProperty.call(row, key)) {
                const normalizedKey = key.trim().toLowerCase().replace(/\s+/g, '');
                normalizedRow[normalizedKey] = row[key];
            }
        }

        const teamName = normalizedRow['nombreequipo'] || 'N/A';
        let phone = normalizedRow['celular'] || '';
        const rawMatchTime = normalizedRow['horapartido'];
        const rawMatchDay = normalizedRow['diapartido'];
        const fieldName = normalizedRow['nombrecancha'] || 'N/A';

        phone = String(phone).replace(/\D/g, '');

        const formattedMatchDay = formatExcelDate(rawMatchDay);
        const formattedMatchTime = formatExcelTime(rawMatchTime);

        if (!phone || teamName === 'N/A' || formattedMatchDay === 'N/A' || formattedMatchTime === 'N/A' || fieldName === 'N/A') {
            console.warn("Skipping row due to missing/invalid data:", row, { phone, teamName, formattedMatchDay, formattedMatchTime, fieldName });
            return;
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
            // Mark as sent visually, but allow the default link action
            this.classList.add('sent');
            // Optionally, update text or icon further if needed
             this.innerHTML = `${whatsappIconSvg} Enviando...`; // Temp text
             // Restore text after a short delay to allow navigation
             setTimeout(() => {
                 if (this.classList.contains('sent')) { // Check if still marked as sent
                     this.innerHTML = `${whatsappIconSvg} Enviar Solicitud Confirmación`;
                 }
             }, 1500);
        }, { once: false }); // Allow clicking again if needed, though 'sent' style persists visually
        actionsContainer.appendChild(confirmBtn);

        // Create "Recordatorio" link/button
        const reminderBtn = document.createElement('a');
        reminderBtn.href = reminderLink;
        reminderBtn.target = "_blank";
        reminderBtn.className = "whatsapp-link reminder";
        reminderBtn.innerHTML = `${whatsappIconSvg} Enviar Recordatorio`;
         reminderBtn.addEventListener('click', function(e) {
            this.classList.add('sent');
            this.innerHTML = `${whatsappIconSvg} Enviando...`;
             setTimeout(() => {
                 if (this.classList.contains('sent')) {
                    this.innerHTML = `${whatsappIconSvg} Enviar Recordatorio`;
                 }
             }, 1500);
        }, { once: false });
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

    if (!validRowsFound && data.length > 0) { // Ensure check happens only if data existed
         outputDiv.innerHTML = '<p class="error">No se encontraron filas con datos válidos completos (Equipo, Celular, Día, Hora, Cancha) en el archivo. Por favor, revisa el contenido y formato de las columnas.</p>';
    } else if (!validRowsFound && data.length === 0) {
         // Message for empty file remains the same
         outputDiv.innerHTML = '<p>El archivo está vacío o no contiene datos reconocibles.</p>';
    }
}