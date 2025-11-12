/**
 * TaskAssignmentModule.js
 *
 * Módulo para crear tablas de asignación de tareas de diseño.
 * Distribuye números correlativos (S1, S2, C1, C2...) entre diseñadores.
 */

/**
 * Genera una tabla de asignación en un nuevo slide
 *
 * @param {Object} assignmentData - Datos de la asignación
 * @param {Array} assignmentData.designers - Array de objetos {name: string}
 * @param {number} assignmentData.stills - Cantidad de Stills por diseñador
 * @param {number} assignmentData.conceptual - Cantidad de Conceptuales por diseñador
 * @returns {Object} - {success: boolean, log: string}
 */
function generateTaskAssignmentTable(assignmentData) {
  var debugLog = "";

  function log(msg) {
    debugLog += msg + "\n";
    Logger.log(msg);
  }

  try {
    log("=== GENERANDO TABLA DE ASIGNACIÓN ===");

    // Validaciones
    if (!assignmentData || !assignmentData.designers) {
      throw new Error("Datos incompletos. Se requieren diseñadores.");
    }

    if (assignmentData.designers.length === 0) {
      throw new Error("Debe haber al menos un diseñador.");
    }

    var designers = assignmentData.designers;
    var stills = parseInt(assignmentData.stills) || 0;
    var conceptual = parseInt(assignmentData.conceptual) || 0;

    if (stills === 0 && conceptual === 0) {
      throw new Error("Debe haber al menos un Still o Conceptual.");
    }

    log("Diseñadores: " + designers.length);
    log("Stills por diseñador: " + stills);
    log("Conceptuales por diseñador: " + conceptual);
    log("Piezas por diseñador: " + (stills + conceptual));
    log("Total de piezas: " + (designers.length * (stills + conceptual)));

    // Crear slide y tabla usando API avanzada
    var presentation = SlidesApp.getActivePresentation();
    var presentationId = presentation.getId();
    var slideId = Utilities.getUuid();

    log("Creando slide con ID: " + slideId);

    // Crear slide y tabla
    createSlideAndTable(presentationId, slideId, designers, stills, conceptual, log);

    log("✓ Tabla de asignación generada exitosamente");

    return {
      success: true,
      log: debugLog,
      slideId: slideId
    };

  } catch (error) {
    log("✗ ERROR: " + error.toString());
    return {
      success: false,
      log: debugLog,
      error: error.toString()
    };
  }
}

/**
 * Crea el slide y la tabla en una sola operación usando API avanzada
 */
function createSlideAndTable(presentationId, slideId, designers, stills, conceptual, logFunction) {
  var rows = designers.length; // Sin encabezado
  var cols = 1 + stills + conceptual; // 1 para nombre + categorías

  logFunction("Creando tabla de " + rows + " filas x " + cols + " columnas");

  var tableId = Utilities.getUuid();
  var requests = [];

  // 1. Crear el slide
  requests.push({
    createSlide: {
      objectId: slideId,
      slideLayoutReference: {
        predefinedLayout: 'BLANK'
      }
    }
  });

  // 2. Crear la tabla en el slide
  requests.push({
    createTable: {
      objectId: tableId,
      elementProperties: {
        pageObjectId: slideId,
        size: {
          width: { magnitude: 9 * 914400, unit: 'EMU' }, // 9 pulgadas
          height: { magnitude: (rows * 0.5) * 914400, unit: 'EMU' }
        },
        transform: {
          scaleX: 1,
          scaleY: 1,
          translateX: 0.5 * 914400,
          translateY: 0.5 * 914400,
          unit: 'EMU'
        }
      },
      rows: rows,
      columns: cols
    }
  });

  // 3. Llenar contenido de la tabla
  var stillCounter = 1;
  var conceptualCounter = 1;

  for (var i = 0; i < designers.length; i++) {
    var designer = designers[i];
    var colIndex = 0;

    // Primera columna: nombre del diseñador
    requests.push(createCellTextRequest(tableId, i, colIndex, designer.name));
    colIndex++;

    // Columnas de Stills: S1, S2, S3...
    for (var s = 0; s < stills; s++) {
      requests.push(createCellTextRequest(tableId, i, colIndex, "S" + stillCounter));
      stillCounter++;
      colIndex++;
    }

    // Columnas de Conceptuales: C1, C2, C3...
    for (var c = 0; c < conceptual; c++) {
      requests.push(createCellTextRequest(tableId, i, colIndex, "C" + conceptualCounter));
      conceptualCounter++;
      colIndex++;
    }
  }

  // Ejecutar TODO en una sola llamada
  try {
    Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
    logFunction("✓ Slide y tabla creados exitosamente");
  } catch (e) {
    logFunction("✗ Error creando slide y tabla: " + e.toString());
    throw e;
  }
}

/**
 * Crea un request para insertar texto en una celda de tabla
 */
function createCellTextRequest(tableId, rowIndex, colIndex, text) {
  var request = {
    insertText: {
      objectId: tableId,
      cellLocation: {
        rowIndex: rowIndex,
        columnIndex: colIndex
      },
      text: text,
      insertionIndex: 0
    }
  };

  return request;
}

/**
 * Función de prueba con datos de ejemplo
 */
function testTaskAssignment() {
  var testData = {
    designers: [
      { name: "Ana García" },
      { name: "Carlos López" },
      { name: "María Torres" },
      { name: "Juan Pérez" },
      { name: "Laura Martínez" }
    ],
    stills: 3,
    conceptual: 2
  };

  var result = generateTaskAssignmentTable(testData);
  Logger.log(result.log);

  return result;
}
