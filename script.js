import * as XLSX from 'xlsx';

// --- Google API Configuration ---
const API_KEY = 'YOUR_GOOGLE_API_KEY'; // Replace!
const CLIENT_ID = 'YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com'; // Replace!

const SCOPES = 'https://www.googleapis.com/auth/drive.readonly'; // Read-only access is sufficient
const DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"];
const APP_ID = CLIENT_ID.split('-')[0]; // Often needed for Picker, derived from Client ID

// --- DOM Elements ---
const fileInput = document.getElementById('excelFile');
const fileNameDisplay = document.getElementById('fileName');
const outputDiv = document.getElementById('output');
const driveButton = document.getElementById('driveButton');
const authStatusSpan = document.getElementById('authStatus');

// --- Global Variables ---
let tokenClient = null; // For Google Identity Services (GIS)
let gapiInited = false;
let gisInited = false;
let pickerApiLoaded = false;
let oauthToken = null; // Store the OAuth token

// --- Existing Code ---
const whatsappIconSvg = `<svg viewBox="0 0 24 24" width="1em" height="1em" focusable="false" aria-hidden="true" fill="currentColor"><path d="M16.75 13.96c-.25-.12-1.48-.72-1.71-.81-.23-.08-.39-.12-.56.12-.16.24-.65.81-.8 1-.15.18-.29.2-.54.08-.25-.12-1.02-.37-1.95-1.2-.73-.66-1.22-1.48-1.36-1.72-.14-.24-.01-.37.11-.49.11-.11.25-.29.37-.43.12-.14.16-.24.24-.4.08-.16.04-.29-.02-.41-.06-.12-.56-1.34-.76-1.84-.2-.48-.4-.42-.55-.42-.15,0-.31,0-.47,0-.16,0-.42.06-.64.31-.22.25-.86.84-.86,2.05,0,1.21.88,2.37,1,2.53.12.16,1.75,2.67,4.24,3.73.59.25,1.05.4,1.41.51.59.19,1.13.16,1.56.1.48-.07,1.48-.6,1.69-1.18.21-.58.21-1.07.15-1.18-.06-.11-.22-.17-.47-.29zm-5.23 6.11c-3.18 0-6.14-1.03-8.68-2.93l-1.28.39 1.31-1.25c-2.11-2.63-3.24-5.84-3.24-9.21C.01 5.77 5.77.01 12.43.01 19.1.01 24 5.77 24 12.43c0 6.67-5.77 12.43-12.43 12.43zm0-22.5c-5.54 0-10.09 4.55-10.09 10.09 0 3.17 1.46 6.04 3.88 7.99L3.13 23l2.1-.61c1.89 1.32 4.15 2.03 6.55 2.03 5.54 0 10.09-4.55 10.09-10.09S17.17 2.33 11.52 2.33z"></path></svg>`;

fileInput.addEventListener('change', handleFileSelectLocal);
driveButton.addEventListener('click', handleAuthClick); // Will trigger auth first, then picker


// --- Google API Initialization ---

// Called after Google API script loads
window.gapiLoaded = () => {
    // Check if placeholders are replaced
    if (API_KEY === 'YOUR_GOOGLE_API_KEY' || CLIENT_ID === 'YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com') {
        console.warn("Google API Key or Client ID not set. Google Drive functionality disabled.");
        authStatusSpan.textContent = "Google API no configurada.";
        driveButton.disabled = true;
        return;
    }
    gapi.load('client:picker', initializeGapiClient);
    gisInited = true;
    // Initialize the GIS client
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: /* @param {Object} resp */ (resp) => {
            if (resp.error) {
                console.error('Error getting access token:', resp.error);
                authStatusSpan.textContent = 'Error de autenticación.';
                oauthToken = null;
            } else {
                console.log('Access token received.');
                oauthToken = resp.access_token;
                // Load picker API after successful auth if not already loaded
                gapi.load('picker', onPickerApiLoad);
            }
            updateSignInStatus();
        }, // Function to handle the response token
    });
    // Check initial auth state silently
    // tokenClient.requestAccessToken({prompt: 'none'}); // Try silent login first
};

async function initializeGapiClient() {
    try {
        await gapi.client.init({
            apiKey: API_KEY,
            discoveryDocs: DISCOVERY_DOCS,
        });
        gapiInited = true;
        console.log("GAPI client initialized.");
        // Initial check (might already have a token via GIS silent check if uncommented above)
        updateSignInStatus();
    } catch (error) {
        console.error("Error initializing GAPI client:", error);
        authStatusSpan.textContent = 'Error iniciando Google API.';
        driveButton.disabled = true;
    }
}

function updateSignInStatus() {
    if (oauthToken) {
        authStatusSpan.textContent = 'Autenticado con Google.';
        driveButton.textContent = 'Seleccionar de Google Drive';
        driveButton.disabled = false;
        driveButton.onclick = loadPicker; // Change button action to load picker directly
    } else {
        authStatusSpan.textContent = 'Necesita autenticación.';
        driveButton.textContent = 'Autenticar con Google';
        driveButton.disabled = false; // Enable for authentication
        driveButton.onclick = handleAuthClick; // Set action to authenticate
    }
}

// --- Authentication Flow ---

function handleAuthClick() {
    if (!gisInited || !gapiInited) {
        console.error("Google libraries not fully initialized.");
        authStatusSpan.textContent = 'Error: Bibliotecas no cargadas.';
        return;
    }
    if (oauthToken) {
        // If already authenticated, proceed to picker
        loadPicker();
    } else {
        // Prompt the user to select an account and grant access if not already signed in
        tokenClient.requestAccessToken({ prompt: 'consent' });
    }
}

// --- Google Picker Logic ---

function onPickerApiLoad() {
    pickerApiLoaded = true;
    console.log("Picker API loaded.");
    // If user clicked the button before API loaded, create picker now
    if (oauthToken) { // Ensure we are still authenticated
        createPicker();
    }
}

function loadPicker() {
    if (!oauthToken) {
        console.warn("Cannot load picker: Not authenticated.");
        handleAuthClick(); // Re-trigger auth if token is missing
        return;
    }
    if (pickerApiLoaded) {
        createPicker();
    } else {
        console.log("Picker API not loaded yet, loading...");
        gapi.load('picker', onPickerApiLoad);
        // The callback 'onPickerApiLoad' will call createPicker
    }
}

function createPicker() {
    if (!pickerApiLoaded || !oauthToken) {
        console.error("Picker API not loaded or not authenticated.");
        authStatusSpan.textContent = 'Error al abrir selector.';
        updateSignInStatus(); // Refresh UI state
        return;
    }
    console.log("Creating Picker...");
    const view = new google.picker.View(google.picker.ViewId.SPREADSHEETS);
    // view.setMimeTypes("application/vnd.google-apps.spreadsheet"); // Only Google Sheets

    const picker = new google.picker.PickerBuilder()
        .setAppId(APP_ID) // Use derived App ID
        .setOAuthToken(oauthToken)
        .addView(view)
        .setDeveloperKey(API_KEY)
        .setCallback(pickerCallback)
        .build();
    picker.setVisible(true);
}

// --- Picker Callback & File Processing ---

async function pickerCallback(data) {
    if (data.action === google.picker.Action.PICKED) {
        const file = data.docs[0];
        const fileId = file.id;
        const fileName = file.name;
        console.log(`Picked file: ${fileName} (ID: ${fileId})`);
        fileNameDisplay.textContent = `Drive: ${fileName}`;
        outputDiv.innerHTML = '<p>Descargando y procesando archivo de Google Drive...</p>';

        try {
            // Use Drive API v3 to export the Google Sheet as an .xlsx file
            const exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;

            const response = await fetch(exportUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${oauthToken}`
                }
            });

            if (!response.ok) {
                throw new Error(`Error exporting file: ${response.statusText} (Status: ${response.status})`);
            }

            const fileData = await response.arrayBuffer();
            console.log("File downloaded successfully from Drive.");
            processExcelData(fileData);

        } catch (error) {
            console.error("Error fetching/processing Google Drive file:", error);
            outputDiv.innerHTML = `<p class="error">Error al descargar o procesar el archivo de Google Drive: ${error.message}. Asegúrate de que la API de Google Drive esté habilitada.</p>`;
            fileNameDisplay.textContent = 'Error al cargar archivo de Drive';
        }
    } else if (data.action === google.picker.Action.CANCEL) {
        console.log("Google Picker cancelled by user.");
        // Optionally reset file name display if needed
        // fileNameDisplay.textContent = 'Ningún archivo seleccionado';
    }
}

// Handle local file selection
function handleFileSelectLocal(event) {
    const file = event.target.files[0];
    if (!file) {
        fileNameDisplay.textContent = 'Ningún archivo seleccionado';
        outputDiv.innerHTML = ''; // Clear previous results
        return;
    }

    fileNameDisplay.textContent = `Local: ${file.name}`;
    outputDiv.innerHTML = '<p>Procesando archivo local...</p>'; // Show loading message

    const reader = new FileReader();

    reader.onload = function (e) {
        const arrayBuffer = e.target.result;
        processExcelData(arrayBuffer);
    };

    reader.onerror = function (error) {
        console.error("File reading error:", error);
        outputDiv.innerHTML = '<p class="error">No se pudo leer el archivo local.</p>';
        fileNameDisplay.textContent = 'Error al leer archivo local';
    };

    reader.readAsArrayBuffer(file);

    // Reset file input value so the 'change' event fires even if the same file is selected again
    event.target.value = null;
}

// Refactored core processing logic
function processExcelData(arrayBuffer) {
    try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        // Ensure date/time parsing options are correct
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            cellDates: true, // Attempt to parse dates
            dateNF: 'dd/mm/yyyy;@', // Format preference for Excel dates if cellDates works
            raw: false // Use formatted strings from Excel where possible
        });

        console.log("Raw JSON Data:", jsonData); // Log raw data for debugging dates/times
        displayData(jsonData);
    } catch (error) {
        console.error("Error processing Excel data:", error);
        outputDiv.innerHTML = '<p class="error">Error al procesar los datos del archivo. Asegúrate de que el formato es correcto y las columnas esperadas existen (Nombre Equipo, Celular, Hora Partido, Dia Partido, Nombre Cancha).</p>';
        // Keep the filename displayed even on error
    }
}

// --- Date/Time Formatting (Minor adjustments possible) ---
function formatExcelDate(date) {
    // Prioritize Date objects first
    if (date instanceof Date && !isNaN(date)) {
        // Check if time part is significant (more than just midnight UTC)
        // Excel dates might import as date object at midnight UTC
        // If it's exactly midnight UTC, it's likely just a date.
        // If it has time components, it might be a datetime cell.
        return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(date);
    }
    // Handle strings that look like dates (e.g., from raw:false or manual entry)
    if (typeof date === 'string') {
        // Try parsing common formats if needed, but rely on cellDates first
        // Basic check for DD/MM/YYYY or YYYY-MM-DD etc.
        if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(date) || /^\d{4}-\d{1,2}-\d{1,2}/.test(date)) {
            try {
                // Attempt to parse and reformat to ensure DD/MM/YYYY
                const parsedDate = new Date(date.includes('/') ? date.split('/').reverse().join('-') : date); // Handle DD/MM/YYYY -> YYYY-MM-DD for Date constructor
                if (!isNaN(parsedDate)) {
                    return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(parsedDate);
                }
            } catch (e) { /* Ignore parsing error */ }
            // If direct parsing fails or it's already DD/MM/YYYY, return as is (assuming it's correct)
            return date;
        }
    }
    // Handle Excel serial dates (numbers) if cellDates fails
    if (typeof date === 'number' && date > 1) {
        try {
            const excelEpoch = new Date(1899, 11, 30); // Excel epoch starts Dec 30, 1899
            const jsDate = new Date(excelEpoch.getTime() + date * 86400000);
            if (!isNaN(jsDate)) {
                return new Intl.DateTimeFormat('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(jsDate);
            }
        } catch (e) { /* Ignore conversion error */ }
    }
    console.warn("Could not format date:", date);
    return 'N/A'; // Return N/A if unparseable
}

function formatExcelTime(timeValue) {
    // Prioritize Date objects (might be datetime cells)
    if (timeValue instanceof Date && !isNaN(timeValue)) {
        // Check if it has a non-zero time component
        if (timeValue.getHours() !== 0 || timeValue.getMinutes() !== 0 || timeValue.getSeconds() !== 0) {
            return new Intl.DateTimeFormat('es-ES', { hour: '2-digit', minute: '2-digit', hour12: false }).format(timeValue);
        }
        // If it's midnight, maybe it wasn't a time cell - treat as N/A for time? Or maybe depends on context?
        // Let's fall through for now.
    }
    // Handle Excel time serial numbers (fraction of a day)
    if (typeof timeValue === 'number' && timeValue >= 0 && timeValue < 1) {
        const totalSeconds = Math.round(timeValue * 86400);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }
    // Handle time strings (HH:MM or HH:MM:SS)
    if (typeof timeValue === 'string') {
        const match = timeValue.match(/^(\d{1,2}):(\d{2})(:(\d{2}))?/);
        if (match) {
            const hours = String(match[1]).padStart(2, '0');
            const minutes = String(match[2]).padStart(2, '0');
            return `${hours}:${minutes}`;
        }
    }
    console.warn("Could not format time:", timeValue);
    return 'N/A';
}

// --- Display Logic (Mostly Unchanged) ---
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
        confirmBtn.addEventListener('click', function (e) {
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
        reminderBtn.addEventListener('click', function (e) {
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

        manualConfirmBtn.addEventListener('click', function () {
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