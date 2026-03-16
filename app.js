const STORAGE_KEY = 'terumo_pipetas_data';

// Datos iniciales de prueba basados en el certificado analizado (FJ25431)
const INITIAL_DATA = [
    { 
        "id": 1,
        "serie": "FJ25431", 
        "marca": "THERMO", 
        "modelo": "MONOCANAL", 
        "volumen": "100-1000 µl", 
        "sector": "LABORATORIO", 
        "codigo_interno": "89",
        "calibracion": "20/06/2024", 
        "vencimiento": "20/06/2025", 
        "certificado": "FJ25431.pdf" 
    }
];

const api = {
    getData: () => {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (!stored) {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(INITIAL_DATA));
            return INITIAL_DATA;
        }
        return JSON.parse(stored);
    },
    saveData: (data) => {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    },
    addItem: (item) => {
        const data = api.getData();
        const newItem = { ...item, id: Date.now() };
        data.push(newItem);
        api.saveData(data);
        return newItem;
    },
    updateItem: (id, updatedItem) => {
        const data = api.getData();
        const index = data.findIndex(item => item.id === id);
        if (index !== -1) {
            data[index] = { ...updatedItem, id };
            api.saveData(data);
        }
    },
    deleteItem: (id) => {
        const data = api.getData();
        const filtered = data.filter(item => item.id !== id);
        api.saveData(filtered);
    }
};

const listElement = document.getElementById('pipette-list');
const addBtn = document.getElementById('add-btn');
const cancelBtn = document.getElementById('cancel-btn');
const modalOverlay = document.getElementById('modal-overlay');
const form = document.getElementById('pipette-form');
const modalTitle = document.querySelector('#modal-overlay h2');
const totalCountElement = document.getElementById('total-count');

let editingId = null;
let selectedPipettes = new Set();

const searchText = document.getElementById('search-text');
const filterSector = document.getElementById('filter-sector');

// Bulk Actions Variables
const selectAllCheckbox = document.getElementById('select-all-checkbox');
const bulkActionsBar = document.getElementById('bulk-actions-bar');
const selectedCount = document.getElementById('selected-count');
const bulkSectorSelect = document.getElementById('bulk-sector-select');
const bulkApplyBtn = document.getElementById('bulk-apply-btn');

// File Upload Simulation
const fileInput = document.getElementById('file-input');
const attachBtn = document.getElementById('attach-btn');
const certificadoInput = document.getElementById('certificado');
const removeCertBtn = document.getElementById('remove-cert-btn');

// Bulk Import Action
const importModalBtn = document.getElementById('import-modal-btn');
const importFileInput = document.getElementById('import-file-input');
const importModalOverlay = document.getElementById('import-modal-overlay');
const importCancelBtn = document.getElementById('import-cancel-btn');
const sectorImportBtns = document.querySelectorAll('.sector-import-btn');

// Excel Import Elements
const importExcelBtn = document.getElementById('import-excel-btn');
const importExcelInput = document.getElementById('import-excel-input');

let targetImportSector = null;
let importMode = 'pdf'; // 'pdf' or 'excel'

// Configuración del worker de PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

function updateRemoveBtnVisibility() {
    if (removeCertBtn) {
        removeCertBtn.style.display = certificadoInput.value ? 'flex' : 'none';
    }
}

if (attachBtn) {
    attachBtn.onclick = () => fileInput.click();
    fileInput.onchange = async (e) => {
        const file = e.target.files[0];
        if (file) {
            certificadoInput.value = file.name;
            updateRemoveBtnVisibility();
            
            // Si el nombre del archivo contiene '.pdf' procesarlo
            if (file.name.toLowerCase().endsWith('.pdf')) {
                try {
                    // Cambiamos el texto del botón a "Leyendo..." para feedback
                    const originalText = attachBtn.innerText;
                    attachBtn.innerText = "LEYENDO PDF...";
                    attachBtn.style.opacity = "0.7";

                    const arrayBuffer = await file.arrayBuffer();
                    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
                    let fullText = '';
                    
                    for(let i = 1; i <= pdf.numPages; i++) {
                        const page = await pdf.getPage(i);
                        const textContent = await page.getTextContent();
                        const pageText = textContent.items.map(item => item.str).join(' ');
                        fullText += pageText + ' ';
                    }

                    const dataObj = extractDataFromText(fullText, file.name);
                    
                    if(dataObj.serie && !document.getElementById('serie').value) document.getElementById('serie').value = dataObj.serie;
                    if(dataObj.marca && !document.getElementById('marca').value) document.getElementById('marca').value = dataObj.marca;
                    if(dataObj.modelo && !document.getElementById('modelo').value) document.getElementById('modelo').value = dataObj.modelo;
                    if(dataObj.volumen && !document.getElementById('volumen').value) document.getElementById('volumen').value = dataObj.volumen;
                    if(dataObj.sector && !document.getElementById('sector').value) document.getElementById('sector').value = dataObj.sector;
                    if(dataObj.calibracion && !document.getElementById('calibracion').value) document.getElementById('calibracion').value = dataObj.calibracion;
                    if(dataObj.vencimiento && !document.getElementById('vencimiento').value) document.getElementById('vencimiento').value = dataObj.vencimiento;
                    
                    attachBtn.innerText = originalText;
                    attachBtn.style.opacity = "1";
                } catch (err) {
                    console.error("Error al procesar o leer el PDF:", err);
                    attachBtn.innerText = "ADJUNTAR PDF";
                    attachBtn.style.opacity = "1";
                }
            }
        }
    };
}

if (importModalBtn && importFileInput) {
    importModalBtn.onclick = () => {
        importMode = 'pdf';
        if (importModalOverlay) importModalOverlay.style.display = 'flex';
    };

    if (importCancelBtn) {
        importCancelBtn.onclick = () => {
            importModalOverlay.style.display = 'none';
            targetImportSector = null;
        };
    }

    sectorImportBtns.forEach(btn => {
        btn.onclick = () => {
            targetImportSector = btn.getAttribute('data-sector');
            importModalOverlay.style.display = 'none';
            
            if (importMode === 'pdf') {
                importFileInput.click();
            } else if (importMode === 'excel' && importExcelInput) {
                importExcelInput.click();
            }
        };
    });
    
    importFileInput.onchange = async (e) => {
        const files = e.target.files;
        if (!files || files.length === 0) return;
        
        const originalText = importModalBtn.innerText;
        importModalBtn.innerText = "IMPORTANDO...";
        importModalBtn.style.opacity = "0.7";
        
        let importedCount = 0;
        
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (!file.name.toLowerCase().endsWith('.pdf')) continue;
            
            try {
                const arrayBuffer = await file.arrayBuffer();
                const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
                let fullText = '';
                
                for(let j = 1; j <= pdf.numPages; j++) {
                    const page = await pdf.getPage(j);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + ' ';
                }
                
                const dataObj = extractDataFromText(fullText, file.name);
                
                // Solo si encontramos al menos el N° Serie o Marca lo guardamos
                if (dataObj.serie || dataObj.marca) {
                    const currentData = api.getData();
                    const existingIndex = dataObj.serie ? currentData.findIndex(item => item.serie === dataObj.serie) : -1;
                    
                    if (existingIndex === -1) {
                        api.addItem({
                            serie: dataObj.serie || "S/D",
                            marca: dataObj.marca || "S/D",
                            modelo: dataObj.modelo || "S/D",
                            volumen: dataObj.volumen || "S/D",
                            sector: targetImportSector || dataObj.sector || "S/D",
                            codigo_interno: "",
                            calibracion: dataObj.calibracion || "",
                            vencimiento: dataObj.vencimiento || "",
                            certificado: file.name
                        });
                        importedCount++;
                    } else {
                        const existingPipette = currentData[existingIndex];
                        if (!existingPipette.certificado || existingPipette.certificado.trim() === '') {
                            // Update existing record with the new certificate
                            existingPipette.certificado = file.name;
                            
                            // Also update dates if missing
                            if (!existingPipette.calibracion && dataObj.calibracion) {
                                existingPipette.calibracion = dataObj.calibracion;
                            }
                            if (!existingPipette.vencimiento && dataObj.vencimiento) {
                                existingPipette.vencimiento = dataObj.vencimiento;
                            }
                            
                            api.saveData(currentData);
                            importedCount++;
                            console.log(`Pipeta actualizada: N° Serie ${dataObj.serie} se le adjuntó el certificado.`);
                        } else {
                            console.log(`Pipeta ignorada: N° Serie ${dataObj.serie} ya está registrado y ya posee un certificado.`);
                        }
                    }
                }
                
            } catch (err) {
                console.error("Error al procesar el archivo:", file.name, err);
            }
        }
        
        importModalBtn.innerText = originalText;
        importModalBtn.style.opacity = "1";
        importFileInput.value = ''; // clean input
        targetImportSector = null; // reset
        
        if (importedCount > 0) {
            updateSectorFilter();
            renderList();
            alert(`Se han importado exitosamente ${importedCount} pipetas.`);
        } else {
            alert("No se pudo extraer información válida de los PDFs seleccionados o ya estaban registrados.");
        }
    };
}

// ================= EXCEL IMPORT LOGIC =================
if (importExcelBtn && importExcelInput) {
    importExcelBtn.onclick = () => {
        importMode = 'excel';
        if (importModalOverlay) importModalOverlay.style.display = 'flex';
    };
    
    importExcelInput.onchange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const originalText = importExcelBtn.innerText;
        importExcelBtn.innerText = "IMPORTANDO...";
        importExcelBtn.style.opacity = "0.7";
        
        let importedCount = 0;
        
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON, using the first row as headers
            const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            const currentData = api.getData();
            
            for (let row of rawData) {
                // Mapeo sugerido
                // Sector = "Unidad"
                // Serie = "Campo Equipo"
                // Rango = "Rango"
                // Modelo = "Canales o Tipo"
                // Calibracion = "Fecha de calibración"
                // Codigo Interno = "Ubicación tecnica"
                
                let rawSerie = String(row["Campo Equipo"] || "").trim();
                if (!rawSerie) continue; // Si no hay serie no agregamos nada
                
                // Limpiar prefijos de serie comunes
                rawSerie = rawSerie.replace(/^(PIPETA-|PIP-|MIC-|S\/N:\s*)/i, '').trim();
                
                // Chequear si ya existe
                const exists = currentData.some(item => item.serie === rawSerie || item.codigo_interno === String(row["Ubicación tecnica"] || "").trim());
                
                let parsedSectorFromExcel = String(row["Unidad"] || "").trim() || "S/D";
                let finalSector = targetImportSector || parsedSectorFromExcel;

                if (!exists) {
                    // Procesar Modelo/Canales
                    let modeloRaw = String(row["Canales o Tipo"] || "").trim().toUpperCase();
                    let finalModelo = modeloRaw;
                    if (modeloRaw === "1") finalModelo = "MONOCANAL";
                    else if (modeloRaw === "8" || modeloRaw === "12") finalModelo = "MULTICANAL";
                    
                    // Procesar Rango
                    let finalVolumen = String(row["Rango"] || "").trim();
                    if (finalVolumen && !finalVolumen.includes("µl") && !finalVolumen.includes("ul") && finalModelo !== "DISPENSER") {
                        finalVolumen += " µl";
                    }
                    
                    // Procesar Fecha de Calibración y Vencimiento
                    let calibDateStr = "";
                    let vencDateStr = "";
                    
                    if (row["Fecha de calibración"]) {
                        const dateVal = row["Fecha de calibración"];
                        // SheetJS sometimes parses dates directly to JS Date objects if `cellDates: true`
                        if (dateVal instanceof Date) {
                            const dStr = dateVal.getDate().toString().padStart(2, '0');
                            const mStr = (dateVal.getMonth() + 1).toString().padStart(2, '0');
                            calibDateStr = `${dStr}/${mStr}/${dateVal.getFullYear()}`;
                            
                            // Sumar 1 año para vencimiento
                            const vYear = dateVal.getFullYear() + 1;
                            vencDateStr = `${dStr}/${mStr}/${vYear}`;
                        } else if (typeof dateVal === 'string') {
                            // Si vino como string "25/07/2024"
                            const parts = dateVal.split(/[-\/.]/);
                            if (parts.length === 3) {
                                // Asumimos DD/MM/YYYY por el formato local de Excel
                                let day = parseInt(parts[0]);
                                let month = parseInt(parts[1]);
                                let year = parseInt(parts[2]);
                                
                                if (year < 100) year += 2000; // Por si viene "24" en vez de "2024"
                                
                                const dStr = day.toString().padStart(2, '0');
                                const mStr = month.toString().padStart(2, '0');
                                calibDateStr = `${dStr}/${mStr}/${year}`;
                                vencDateStr = `${dStr}/${mStr}/${year + 1}`;
                            } else {
                                calibDateStr = dateVal;
                            }
                        }
                    }
                    
                    api.addItem({
                        serie: rawSerie,
                        marca: "S/D", // Por defecto
                        modelo: finalModelo || "S/D",
                        volumen: finalVolumen || "S/D",
                        sector: finalSector,
                        codigo_interno: String(row["Ubicación tecnica"] || "").trim(),
                        calibracion: calibDateStr,
                        vencimiento: vencDateStr,
                        certificado: ""
                    });
                    
                    importedCount++;
                } else {
                    console.log(`Pipeta ignorada desde Excel: El equipo ${rawSerie} ya está registrado.`);
                }
            }
        } catch (err) {
            console.error("Error al procesar el archivo Excel:", err);
            alert("Error al procesar el archivo Excel. Asegúrate de que el formato sea correcto.");
        }
        
        importExcelBtn.innerText = originalText;
        importExcelBtn.style.opacity = "1";
        importExcelInput.value = ''; // clean input
        targetImportSector = null; // reset
        
        if (importedCount > 0) {
            updateSectorFilter();
            renderList();
            alert(`Se han importado exitosamente ${importedCount} pipetas desde el Excel.`);
        } else {
            alert("No se importó ninguna pipeta. Puede que todas ya existan o que el formato no coincida con los encabezados esperados.");
        }
    };
}
// ==============================================

function extractDataFromText(text, filename) {
    const cleanText = text.replace(/\s+/g, ' ');
    const result = { serie: "", marca: "", modelo: "", volumen: "", sector: "", calibracion: "", vencimiento: "", certificado: filename };
    
    // 1. Extraer Nro de Serie
    const serieMatch = cleanText.match(/([A-Z]{2}\d{4,5}|\b\d{6,9}\b)/i);
    if(serieMatch) result.serie = serieMatch[1];
    
    // 2. Marca
    const marcasMatches = ["THERMO", "EPPENDORF", "GILSON", "SOCOREX", "BRAND", "DRAGON", "HTL", "FINNPIPETTE"];
    for(let m of marcasMatches) {
        if(cleanText.toUpperCase().includes(m)) {
            result.marca = m;
            break;
        }
    }
    
    // 3. Modelo
    if(cleanText.toUpperCase().includes("MONOCANAL")) {
        result.modelo = "MONOCANAL";
    } else if(cleanText.toUpperCase().includes("MULTICANAL")) {
        result.modelo = "MULTICANAL";
    }
    
    // 4. Sector
    if(cleanText.toUpperCase().includes("LABORATORIO")) {
        result.sector = "LABORATORIO";
    }
    
    // 5. Volumen / Rango
    const volMatch = cleanText.match(/(\d+\s*[-a]\s*\d+\s*[µu]l|\d+\s*[-a]\s*\d+\s*ml|\b\d+\s*[µu]l\b|\b\d+\s*ml\b)/i);
    if(volMatch) {
        result.volumen = volMatch[1].replace('ul', 'µl').replace('uL', 'µl').trim();
    }
    
    // 6. Fechas
    // Patrón 1: DD/MM/YYYY, DD.MM.YYYY, DD-MM-YYYY (con o sin espacios)
    const datePat1 = /\b(\d{1,2})\s*[\/\-\.\s]+\s*(\d{1,2})\s*[\/\-\.\s]+\s*(202\d)\b/g;
    // Patrón 2: YYYY-MM-DD, YYYY/MM/DD
    const datePat2 = /\b(202\d)\s*[\/\-\.\s]+\s*(\d{1,2})\s*[\/\-\.\s]+\s*(\d{1,2})\b/g;
    // Patrón 3: DD MMMM YYYY (ej. 15 Feb 2024, 15 de Abril 2024, 15.02.2024 where 02 is letters)
    const datePat3 = /\b(\d{1,2})\s*(?:de\s*)?([a-zA-Z]{3,10})\s*(?:de\s*,?\s*)?(202\d)\b/gi;
    // Patrón 4: MMMM DD, YYYY (ej. February 15, 2024)
    const datePat4 = /\b([a-zA-Z]{3,10})\s+(\d{1,2})(?:st|nd|rd|th|,)?[ \t]+(202\d)\b/gi;

    let calibDay = null, calibMonth = null, calibYear = null;

    const match1 = [...cleanText.matchAll(datePat1)];
    const match2 = [...cleanText.matchAll(datePat2)];
    const match3 = [...cleanText.matchAll(datePat3)];
    const match4 = [...cleanText.matchAll(datePat4)];

    const monthMap = {
        'jan': 1, 'january': 1, 'ene': 1, 'enero': 1,
        'feb': 2, 'february': 2, 'febrero': 2,
        'mar': 3, 'march': 3, 'marzo': 3,
        'apr': 4, 'april': 4, 'abr': 4, 'abril': 4,
        'may': 5, 'mayo': 5,
        'jun': 6, 'june': 6, 'junio': 6,
        'jul': 7, 'july': 7, 'julio': 7,
        'aug': 8, 'august': 8, 'ago': 8, 'agosto': 8,
        'sep': 9, 'september': 9, 'septiembre': 9,
        'oct': 10, 'october': 10, 'octubre': 10,
        'nov': 11, 'november': 11, 'noviembre': 11,
        'dec': 12, 'december': 12, 'dic': 12, 'diciembre': 12
    };

    if (match1.length > 0) {
        let best = match1.find(d => parseInt(d[2]) <= 12) || match1[0];
        calibDay = parseInt(best[1]); calibMonth = parseInt(best[2]); calibYear = parseInt(best[3]);
    } else if (match2.length > 0) {
        let best = match2.find(d => parseInt(d[2]) <= 12) || match2[0];
        calibYear = parseInt(best[1]); calibMonth = parseInt(best[2]); calibDay = parseInt(best[3]);
    } else if (match3.length > 0) {
        calibDay = parseInt(match3[0][1]); 
        calibMonth = monthMap[match3[0][2].toLowerCase()] || 1; 
        calibYear = parseInt(match3[0][3]);
    } else if (match4.length > 0) {
        calibMonth = monthMap[match4[0][1].toLowerCase()] || 1; 
        calibDay = parseInt(match4[0][2]); 
        calibYear = parseInt(match4[0][3]);
    }

    if (calibYear && calibMonth && calibDay) {
        const dStr = calibDay.toString().padStart(2, '0');
        const mStr = calibMonth.toString().padStart(2, '0');
        result.calibracion = `${dStr}/${mStr}/${calibYear}`;
        
        // 7. Vencimiento
        // Por defecto asumimos 1 Año (12 meses) de vigencia para pipetas si no se especifica.
        let vYear = calibYear + 1;
        let vMonth = calibMonth;
        let vDay = calibDay;

        // Verificar si especifica algo distinto a 12 meses (ej 6 meses)
        const mesesMatch = cleanText.match(/(\d+)\s*meses/i);
        if (mesesMatch) {
            const meses = parseInt(mesesMatch[1]);
            const dateObj = new Date(calibYear, calibMonth - 1, calibDay);
            dateObj.setMonth(dateObj.getMonth() + meses);
            vYear = dateObj.getFullYear();
            vMonth = dateObj.getMonth() + 1;
            vDay = dateObj.getDate();
        }

        const vdStr = vDay.toString().padStart(2, '0');
        const vmStr = vMonth.toString().padStart(2, '0');
        result.vencimiento = `${vdStr}/${vmStr}/${vYear}`;
    }
    
    return result;
}

if (removeCertBtn) {
    removeCertBtn.onclick = () => {
        certificadoInput.value = '';
        fileInput.value = '';
        updateRemoveBtnVisibility();
    };
}

// Helper: Parse DD/MM/YYYY to Date object
function parseDateStr(dateStr) {
    if (!dateStr || dateStr === 'N/A' || dateStr === 'Pendiente') return null;
    const parts = dateStr.split('/');
    if (parts.length === 3) {
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return null;
}

// Load Unique Sectors for Filter
function updateSectorFilter() {
    const data = api.getData();
    const sectors = [...new Set(data.map(item => item.sector))].sort();

    filterSector.innerHTML = '<option value="">TODOS LOS SECTORES</option>';
    sectors.forEach(sec => {
        if (!sec) return;
        const opt = document.createElement('option');
        opt.value = sec;
        opt.textContent = sec;
        filterSector.appendChild(opt);
    });
}

function renderList() {
    const data = api.getData();
    const searchVal = searchText.value.toLowerCase();
    const sectorVal = filterSector.value;

    const now = new Date();
    // Consideramos "Próximo a vencer" a los 40 días antes (aprox mes y medio)
    const warningDate = new Date();
    warningDate.setDate(now.getDate() + 40);

    listElement.innerHTML = '';

    const filteredData = data.filter(item => {
        const matchesSearch = 
            item.serie.toLowerCase().includes(searchVal) || 
            item.marca.toLowerCase().includes(searchVal) ||
            item.modelo.toLowerCase().includes(searchVal) ||
            item.sector.toLowerCase().includes(searchVal);
        const matchesSector = !sectorVal || item.sector === sectorVal;
        return matchesSearch && matchesSector;
    });

    totalCountElement.textContent = filteredData.length;

    if (filteredData.length === 0) {
        listElement.innerHTML = '<tr><td colspan="10" style="text-align:center; padding: 2rem; color: var(--text-muted);">No se encontraron pipetas.</td></tr>';
        return;
    }

    filteredData.forEach((item, index) => {
        const expDate = parseDateStr(item.vencimiento);
        
        let rowClass = '';
        let badgeClass = 'badge-success'; // Vigente por defecto si tiene fecha en el futuro
        let estadoText = 'Vigente';

        if (expDate) {
            if (expDate < now) {
                rowClass = 'row-alert';
                badgeClass = 'badge-danger';
                estadoText = 'Vencida';
            } else if (expDate < warningDate) {
                rowClass = 'row-warning';
                badgeClass = 'badge-warning';
                estadoText = 'Por Vencer';
            }
        } else {
            badgeClass = '';
            estadoText = item.vencimiento || 'N/A';
        }

        const tr = document.createElement('tr');
        if (rowClass) tr.className = rowClass;

        // Intentar buscar en la carpeta superior "Pipetas 2025", o donde el usuario lo asigne. 
        // Por defecto asumimos que están en "../Pipetas 2025/"
        const certLink = item.certificado
            ? `<a href="../Pipetas 2025/${item.certificado}" target="_blank" class="btn-success-outline">📄 Ver PDF</a>`
            : '<span style="color: var(--text-muted); font-size: 0.8rem;">Sin adjunto</span>';

        tr.innerHTML = `
            <td style="text-align: center;">
                <input type="checkbox" class="row-checkbox" data-id="${item.id}" ${selectedPipettes.has(item.id) ? 'checked' : ''} style="cursor: pointer; width: auto;">
            </td>
            <td style="color: var(--text-muted); font-weight: 600;">${index + 1}</td>
            <td style="font-weight: bold; color: #38bdf8;">${item.serie}</td>
            <td>${item.marca} ${item.modelo}</td>
            <td>${item.volumen}</td>
            <td>${item.sector} ${item.codigo_interno ? `<br><small style="color: #94a3b8">Cod: ${item.codigo_interno}</small>` : ''}</td>
            <td>${item.calibracion || 'N/A'}</td>
            <td><span class="badge ${badgeClass}" title="${estadoText}">${item.vencimiento || 'N/A'}</span></td>
            <td>${certLink}</td>
            <td>
                <div style="display: flex; gap: 0.5rem;">
                    <button class="btn-primary edit-btn" data-id="${item.id}" style="padding: 0.4rem 0.8rem; font-size: 0.75rem;">EDITAR</button>
                    <button class="btn-danger delete-btn" data-id="${item.id}" style="padding: 0.4rem 0.8rem; font-size: 0.75rem;">ELIMINAR</button>
                </div>
            </td>
        `;
        listElement.appendChild(tr);
    });

    // Bind Edit buttons
    document.querySelectorAll('.edit-btn').forEach(btn => {
        btn.onclick = () => {
            const id = parseInt(btn.getAttribute('data-id'));
            const data = api.getData();
            const item = data.find(i => i.id === id);
            if (item) {
                editingId = id;
                modalTitle.textContent = 'Editar Pipeta';
                document.getElementById('serie').value = item.serie;
                document.getElementById('marca').value = item.marca;
                document.getElementById('modelo').value = item.modelo;
                document.getElementById('volumen').value = item.volumen;
                document.getElementById('sector').value = item.sector;
                document.getElementById('codigo_interno').value = item.codigo_interno || '';
                document.getElementById('calibracion').value = item.calibracion || '';
                document.getElementById('vencimiento').value = item.vencimiento || '';
                document.getElementById('certificado').value = item.certificado || '';
                updateRemoveBtnVisibility();
                modalOverlay.style.display = 'flex';
            }
        };
    });

    // Bind Delete buttons
    document.querySelectorAll('.delete-btn').forEach(btn => {
        btn.onclick = () => {
            const id = parseInt(btn.getAttribute('data-id'));
            if (confirm('¿Estás seguro de eliminar el registro de esta pipeta?')) {
                api.deleteItem(id);
                renderList();
            }
        };
    });
    // Bind Row Checkboxes
    document.querySelectorAll('.row-checkbox').forEach(chk => {
        chk.onchange = (e) => {
            const id = parseInt(e.target.getAttribute('data-id'));
            if (e.target.checked) {
                selectedPipettes.add(id);
            } else {
                selectedPipettes.delete(id);
            }
            updateBulkActionBar();
        };
    });

    if (selectAllCheckbox) {
        // Sync "select all" state based on currently visible/filtered rows
        selectAllCheckbox.checked = filteredData.length > 0 && 
            filteredData.every(item => selectedPipettes.has(item.id));
        
        selectAllCheckbox.onchange = (e) => {
            const isChecked = e.target.checked;
            filteredData.forEach(item => {
                const cb = document.querySelector(`.row-checkbox[data-id="${item.id}"]`);
                if (cb) cb.checked = isChecked;
                if (isChecked) {
                    selectedPipettes.add(item.id);
                } else {
                    selectedPipettes.delete(item.id);
                }
            });
            updateBulkActionBar();
        };
    }
}

function updateBulkActionBar() {
    if (selectedPipettes.size > 0) {
        bulkActionsBar.style.display = 'block';
        selectedCount.textContent = selectedPipettes.size;
    } else {
        bulkActionsBar.style.display = 'none';
    }
}

// Bulk Apply action
if (bulkApplyBtn) {
    bulkApplyBtn.onclick = () => {
        if (selectedPipettes.size === 0) return;
        const newSector = bulkSectorSelect.value;
        if (!newSector) {
            alert('Por favor, selecciona un sector al cual mover las pipetas.');
            return;
        }

        if (confirm(`¿Estás seguro de mover ${selectedPipettes.size} pipetas al sector ${newSector}?`)) {
            const data = api.getData();
            data.forEach((item, index) => {
                if (selectedPipettes.has(item.id)) {
                    data[index].sector = newSector;
                }
            });
            api.saveData(data);
            selectedPipettes.clear();
            bulkSectorSelect.value = '';
            if(selectAllCheckbox) selectAllCheckbox.checked = false;
            
            updateSectorFilter();
            renderList();
            updateBulkActionBar();
            alert('¡Sector actualizado masivamente!');
        }
    }
}

// Event Listeners for Filters
searchText.oninput = renderList;
filterSector.onchange = renderList;

addBtn.onclick = () => {
    editingId = null;
    modalTitle.textContent = 'Registrar Nueva Pipeta';
    form.reset();
    updateRemoveBtnVisibility();
    modalOverlay.style.display = 'flex';
}
cancelBtn.onclick = () => modalOverlay.style.display = 'none';

window.onclick = (e) => {
    if (e.target === modalOverlay) modalOverlay.style.display = 'none';
};

form.onsubmit = (e) => {
    e.preventDefault();
    const itemData = {
        serie: document.getElementById('serie').value.trim(),
        marca: document.getElementById('marca').value.trim(),
        modelo: document.getElementById('modelo').value.trim(),
        volumen: document.getElementById('volumen').value.trim(),
        sector: document.getElementById('sector').value.trim(),
        codigo_interno: document.getElementById('codigo_interno').value.trim(),
        calibracion: document.getElementById('calibracion').value.trim(),
        vencimiento: document.getElementById('vencimiento').value.trim(),
        certificado: document.getElementById('certificado').value.trim()
    };

    if (editingId) {
        api.updateItem(editingId, itemData);
    } else {
        const currentData = api.getData();
        const exists = currentData.some(item => item.serie === itemData.serie);
        if (exists) {
            alert('Ups... Ya existe una pipeta registrada con este N° de Serie.');
            return;
        }
        api.addItem(itemData);
    }

    modalOverlay.style.display = 'none';
    updateSectorFilter();
    renderList();
};

// PDF Export Function
function exportToPDF(customFilename) {
    const data = api.getData();
    const searchVal = searchText.value.toLowerCase();
    const sectorVal = filterSector.value;

    const filteredData = data.filter(item => {
        const matchesSearch = 
            item.serie.toLowerCase().includes(searchVal) || 
            item.marca.toLowerCase().includes(searchVal) ||
            item.modelo.toLowerCase().includes(searchVal);
        const matchesSector = !sectorVal || item.sector === sectorVal;
        return matchesSearch && matchesSector;
    });

    const printContainer = document.createElement('div');
    printContainer.style.padding = '20px';
    printContainer.style.fontFamily = "'Inter', sans-serif";
    printContainer.style.background = 'white';
    printContainer.style.color = 'black';

    const date = new Date().toLocaleDateString();

    printContainer.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0ea5e9; padding-bottom: 10px; margin-bottom: 20px;">
            <h1 style="margin: 0; font-size: 24px; color: #0ea5e9;">Reporte de Control de Pipetas</h1>
            <div style="text-align: right; font-size: 12px; color: #666;">
                <p>Fecha de Reporte: ${date}</p>
                <p>Total Pipetas: ${filteredData.length}</p>
            </div>
        </div>
        <table style="width: 100%; border-collapse: collapse; font-size: 10px;">
            <thead>
                <tr style="background: #f0f9ff; color: #0369a1;">
                    <th style="border: 1px solid #bae6fd; padding: 8px; text-align: left;">#</th>
                    <th style="border: 1px solid #bae6fd; padding: 8px; text-align: left;">N° SERIE</th>
                    <th style="border: 1px solid #bae6fd; padding: 8px; text-align: left;">MARCA / MODELO</th>
                    <th style="border: 1px solid #bae6fd; padding: 8px; text-align: left;">VOLUMEN</th>
                    <th style="border: 1px solid #bae6fd; padding: 8px; text-align: left;">SECTOR</th>
                    <th style="border: 1px solid #bae6fd; padding: 8px; text-align: left;">CALIBRACIÓN / VTO</th>
                </tr>
            </thead>
            <tbody>
                ${filteredData.map((item, index) => {
        return `
                        <tr>
                            <td style="border: 1px solid #e0f2fe; padding: 6px; width: 30px;">${index + 1}</td>
                            <td style="border: 1px solid #e0f2fe; padding: 6px; font-weight: 600;">${item.serie}</td>
                            <td style="border: 1px solid #e0f2fe; padding: 6px;">${item.marca} ${item.modelo}</td>
                            <td style="border: 1px solid #e0f2fe; padding: 6px;">${item.volumen}</td>
                            <td style="border: 1px solid #e0f2fe; padding: 6px;">${item.sector}</td>
                            <td style="border: 1px solid #e0f2fe; padding: 6px;">
                                Cal: ${item.calibracion || 'N/A'}<br>
                                Vto: <strong style="${item.vencimiento ? '' : 'color:red;'}">${item.vencimiento || 'Pendiente'}</strong>
                            </td>
                        </tr>
                    `;
    }).join('')}
            </tbody>
        </table>
        <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 10px; font-size: 9px; color: #999; text-align: center;">
            <p>Sistema de Gestión de Laboratorio - Inventario de Pipetas</p>
        </div>
    `;

    const opt = {
        margin: [10, 10, 10, 10],
        filename: customFilename || `Reporte_Pipetas_${new Date().toISOString().split('T')[0]}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 3, useCORS: true, letterRendering: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
    };

    html2pdf().set(opt).from(printContainer).save();
}

// Modal PDF Elements
const modalPdfOverlay = document.getElementById('modal-pdf-overlay');
const pdfFilenameInput = document.getElementById('pdf-filename');
const pdfConfirmBtn = document.getElementById('pdf-confirm-btn');
const pdfCancelBtn = document.getElementById('pdf-cancel-btn');

document.getElementById('pdf-btn').onclick = () => {
    pdfFilenameInput.value = `Reporte_Pipetas_${new Date().toISOString().split('T')[0]}`;
    modalPdfOverlay.style.display = 'flex';
};

pdfConfirmBtn.onclick = () => {
    let filename = pdfFilenameInput.value.trim();
    if (filename && !filename.toLowerCase().endsWith('.pdf')) {
        filename += '.pdf';
    }
    exportToPDF(filename);
    modalPdfOverlay.style.display = 'none';
};

pdfCancelBtn.onclick = () => {
    modalPdfOverlay.style.display = 'none';
};

window.addEventListener('click', (e) => {
    if (e.target === modalPdfOverlay) {
        modalPdfOverlay.style.display = 'none';
    }
});

// Initial Load
updateSectorFilter();
renderList();
