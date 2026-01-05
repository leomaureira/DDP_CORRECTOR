// Variables globales
let selectedFile = null;

// Elementos del DOM
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const processBtn = document.getElementById('processBtn');
const status = document.getElementById('status');

// Eventos de drag and drop
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

processBtn.addEventListener('click', processFile);

// Manejar archivo seleccionado
function handleFile(file) {
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
        showStatus('Por favor seleccioná un archivo Excel (.xlsx o .xls)', 'error');
        return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.classList.add('visible');
    processBtn.disabled = false;
    hideStatus();
}

// Mostrar estado
function showStatus(message, type) {
    status.textContent = message;
    status.className = 'status visible ' + type;
}

function hideStatus() {
    status.className = 'status';
}

// Procesar archivo

// Exportar DataFrame original para debugging
function exportDataFrameForDebug(df) {
    const workbook = XLSX.utils.book_new();
    
    // Crear hoja con los datos originales
    const worksheet = XLSX.utils.json_to_sheet(df);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Datos Originales');
    
    // Crear hoja con análisis de tipos de datos
    const dataTypesAnalysis = df.slice(0, 10).map((row, index) => {
        const analysis = { 'Fila': index + 1 };
        Object.keys(row).forEach(key => {
            const value = row[key];
            analysis[key] = value;
            analysis[`${key}_TIPO`] = typeof value;
            analysis[`${key}_VACIO`] = !value ? 'SÍ' : 'NO';
        });
        return analysis;
    });

    const wsTypes = XLSX.utils.json_to_sheet(dataTypesAnalysis);
    XLSX.utils.book_append_sheet(workbook, wsTypes, 'Análisis de Tipos');
    
    // Crear hoja con estadísticas
    const stats = [{
        'Total de Filas': df.length,
        'Total de Columnas': Object.keys(df[0] || {}).length,
        'Columnas': Object.keys(df[0] || {}).join(', ')
    }];
    
    const wsStats = XLSX.utils.json_to_sheet(stats);
    XLSX.utils.book_append_sheet(workbook, wsStats, 'Estadísticas');
    
    // Descargar con timestamp
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    XLSX.writeFile(workbook, `DEBUG_dataframe_${timestamp}.xlsx`);
}



function processFile() {
    if (!selectedFile) {
        showStatus('Primero debés importar un archivo.', 'error');
        return;
    }

    showStatus('Procesando archivo...', 'processing');
    processBtn.disabled = true;

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: ''
            });

            const timmedData = jsonData.map(row => row.slice(1))
            const headers = jsonData[12];
            const dataRows = jsonData.slice(13, -1);

            const df = dataRows.map(row => {
                const obj = {};
                for (let i = 1; i < headers.length; i++) {
                    const header = headers[i];
                    if (header) {
                        obj[header] = row[i] !== undefined ? row[i] : '';
                    }
                }
                return obj;
            });

            // 🔍 EXPORTAR DATAFRAME PARA DEBUG
            // Descomentá esta línea si querés ver el DataFrame original:
            exportDataFrameForDebug(df);

            // Aplicar filtros
            const results = applyFilters(df);

            // Crear y descargar Excel con resultados
            createAndDownloadExcel(results);

            showStatus('Archivo procesado correctamente', 'success');

        } catch (error) {
            console.error('Error:', error);
            showStatus('Error al procesar el archivo: ' + error.message, 'error');
        }

        processBtn.disabled = false;
    };

    reader.onerror = function() {
        showStatus('Error al leer el archivo', 'error');
        processBtn.disabled = false;
    };

    reader.readAsArrayBuffer(selectedFile);
}

// Aplicar filtros al DataFrame
function applyFilters(df) {
    // Cantidad nula
    const Qty_0 = df.filter(row => row['Cantidad'] === 0);

    // Precio Unitario en 0
    const PU_0 = df.filter(row => row['Precio Unitario'] === 0);

    // Sin Validez
    const VAL_0 = df.filter(row => row['Validez'] === 0);

    // Análisis de oferta nulos (Evaluación Técnica vacía)
    const AO_0 = df.filter(row => row['Evaluación Técnica'] === '');

    // Sin Lugar de entrega
    const LE_0 = df.filter(row => row['Lugar de entrega'] === '');

    // Incoterms nulos (Condición de entrega vacía)
    const INC_0 = df.filter(row => row['Condición de entrega'] === '');

    // Sin medio de transporte
    const TRA_0 = df.filter(row => row['Medio de transporte'] === '');

    // AT SITE con logística cargada
    const AT_SITE = df.filter(row =>
        row['Lugar de entrega'] === 'AT-SITE' &&
        row['Total AT SITE'] !== row['Subtotal materiales'] &&
        row['TIPO DE ITEM'] === ''
    );

    // Sin logística cargada
    const NO_AT_SITE = df.filter(row =>
        row['Lugar de entrega'] !== 'AT-SITE' &&
        row['Total AT SITE'] === row['Subtotal materiales'] &&
        row['TIPO DE ITEM'] === ''
    );

    // PU distinto según descripción
    const PU_MAY_2 = findDifferentPricesByField(df, 'Descripción Item');

    // PU distinto según TAG
    const PU_MAY_3 = findDifferentPricesByField(df, 'Código TAG');

    return {
        'Incoterms Nulos': INC_0,
        'Analisis de oferta Nulos': AO_0,
        'Sin Validez': VAL_0,
        'Precios Unitarios en 0': PU_0,
        'Sin Cantidad': Qty_0,
        'PU Distinto según descripción': PU_MAY_2,
        'PU Distinto según TAG': PU_MAY_3,
        'Sin medio de transporte': TRA_0,
        'Sin Lugar de Entrega': LE_0,
        'Sin logistica cargada': NO_AT_SITE,
        'AT SITE CON LOGISTICA CARGADA': AT_SITE
    };
}

// Encontrar items con precios diferentes agrupados por un campo
function findDifferentPricesByField(df, fieldName) {
    // Agrupar por campo y contar precios únicos
    const groups = {};

    df.forEach(row => {
        const key = row[fieldName];
        if (key === undefined || key === '') return;

        if (!groups[key]) {
            groups[key] = new Set();
        }
        groups[key].add(row['Precio Unitario']);
    });

    // Encontrar grupos con más de un precio único
    const keysWithMultiplePrices = Object.keys(groups).filter(key => groups[key].size > 1);

    // Filtrar el DataFrame original
    return df.filter(row => keysWithMultiplePrices.includes(String(row[fieldName])));
}

// Crear y descargar archivo Excel
function createAndDownloadExcel(results) {
    const workbook = XLSX.utils.book_new();

    // Agregar cada resultado como una hoja
    Object.entries(results).forEach(([sheetName, data]) => {
        // Truncar nombre de hoja si es muy largo (máximo 31 caracteres en Excel)
        const truncatedName = sheetName.substring(0, 31);

        if (data.length > 0) {
            const worksheet = XLSX.utils.json_to_sheet(data);
            XLSX.utils.book_append_sheet(workbook, worksheet, truncatedName);
        } else {
            // Crear hoja vacía con headers si no hay datos
            const worksheet = XLSX.utils.aoa_to_sheet([['Sin resultados']]);
            XLSX.utils.book_append_sheet(workbook, worksheet, truncatedName);
        }
    });

    // Generar nombre de archivo con timestamp
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    const outputFileName = `resultado_${timestamp}.xlsx`;

    // Descargar
    XLSX.writeFile(workbook, outputFileName);
}
