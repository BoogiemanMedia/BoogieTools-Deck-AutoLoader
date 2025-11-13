/**
 * TaskAssignmentModule.js
 *
 * Módulo para crear tablas de asignación de tareas de diseño.
 * Distribuye categorías (con posibles subcategorías) entre diseñadores.
 */

/**
 * Genera una tabla de asignación en un nuevo slide
 *
 * @param {Object} assignmentData - Datos de la asignación
 * @param {Array} assignmentData.designers - Array de objetos {name: string}
 * @param {Array} assignmentData.categories - Array de strings con categorías (ej: ["1", "2a", "2b", "3"])
 * @param {number} assignmentData.stillsCount - Cantidad de categorías que son Stills (primeras N)
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
    if (!assignmentData || !assignmentData.designers || !assignmentData.categories) {
      throw new Error("Datos incompletos. Se requieren diseñadores y categorías.");
    }

    if (assignmentData.designers.length === 0) {
      throw new Error("Debe haber al menos un diseñador.");
    }

    if (assignmentData.categories.length === 0) {
      throw new Error("Debe haber al menos una categoría.");
    }

    var designers = assignmentData.designers;
    var categories = assignmentData.categories;
    var stillsCount = parseInt(assignmentData.stillsCount) || 0;
    var conceptualsCount = categories.length - stillsCount;

    log("Diseñadores: " + designers.length);
    log("Categorías totales: " + categories.length);
    log("Categorías Stills: " + stillsCount + " (primeras)");
    log("Categorías Conceptuales: " + conceptualsCount + " (restantes)");
    log("Casilleros por diseñador: " + categories.length);
    log("Total de casilleros: " + (designers.length * categories.length));

    // Crear slide y tabla usando API avanzada
    var presentation = SlidesApp.getActivePresentation();
    var presentationId = presentation.getId();
    var slideId = Utilities.getUuid();

    log("Creando slide con ID: " + slideId);

    // Crear slide y tabla
    createSlideAndTable(presentationId, slideId, designers, categories, log);

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
function createSlideAndTable(presentationId, slideId, designers, categories, logFunction) {
  var rows = designers.length; // Sin encabezado
  var cols = 1 + categories.length; // 1 para nombre + categorías

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
  for (var i = 0; i < designers.length; i++) {
    var designer = designers[i];
    var colIndex = 0;

    // Primera columna: nombre del diseñador
    requests.push(createCellTextRequest(tableId, i, colIndex, designer.name));
    colIndex++;

    // Resto de columnas: categorías (se repiten para cada diseñador)
    for (var c = 0; c < categories.length; c++) {
      requests.push(createCellTextRequest(tableId, i, colIndex, categories[c]));
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
 * Valida el formato de una categoría (permite números y subcategorías)
 * Ejemplos válidos: "1", "1a", "1b", "2", "10c", etc.
 */
function validateCategoryFormat(category) {
  var pattern = /^\d+[a-z]?$/i;
  return pattern.test(category.trim());
}

/**
 * Ordena categorías en orden natural (1, 1a, 1b, 2, 2a, 3, etc.)
 */
function sortCategories(categories) {
  return categories.sort(function(a, b) {
    // Extraer número y letra
    var matchA = a.match(/^(\d+)([a-z]?)$/i);
    var matchB = b.match(/^(\d+)([a-z]?)$/i);

    if (!matchA || !matchB) return 0;

    var numA = parseInt(matchA[1]);
    var numB = parseInt(matchB[1]);
    var letterA = matchA[2].toLowerCase();
    var letterB = matchB[2].toLowerCase();

    // Comparar primero por número
    if (numA !== numB) {
      return numA - numB;
    }

    // Si el número es igual, comparar por letra
    if (letterA < letterB) return -1;
    if (letterA > letterB) return 1;
    return 0;
  });
}

/**
 * Función de prueba con datos de ejemplo
 */
function testTaskAssignment() {
  var testData = {
    designers: [
      { name: "Ana García" },
      { name: "Carlos López" },
      { name: "María Torres" }
    ],
    categories: ["1", "2a", "2b", "2c", "3", "4", "5"],
    stillsCount: 5 // Primeras 5 son Stills (1, 2a, 2b, 2c, 3), resto Conceptuales (4, 5)
  };

  var result = generateTaskAssignmentTable(testData);
  Logger.log(result.log);

  return result;
}
