/**
 * TaskAssignmentModule.js
 *
 * Módulo para crear tablas de asignación de tareas de diseño.
 * Distribuye categorías correlativamente entre diseñadores según su capacidad.
 */

/**
 * Genera una tabla de asignación en un nuevo slide
 *
 * @param {Object} assignmentData - Datos de la asignación
 * @param {Array} assignmentData.designers - Array de objetos {name: string}
 * @param {Array} assignmentData.categories - Array de strings con categorías (ej: ["1", "2a", "2b", "3"])
 * @param {number} assignmentData.piecesPerDesigner - Máximo de piezas por diseñador
 * @param {number} assignmentData.stillsCount - Cantidad de COLUMNAS que son Stills (primeras N columnas)
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
    var piecesPerDesigner = parseInt(assignmentData.piecesPerDesigner) || 0;
    var stillsCount = parseInt(assignmentData.stillsCount) || 0;
    var conceptualsCount = piecesPerDesigner - stillsCount;

    if (piecesPerDesigner === 0) {
      throw new Error("Debe definir cuántas piezas hace cada diseñador.");
    }

    log("Diseñadores: " + designers.length);
    log("Categorías totales: " + categories.length);
    log("Piezas por diseñador: " + piecesPerDesigner);
    log("Columnas Stills (primeras): " + stillsCount);
    log("Columnas Conceptuales (restantes): " + conceptualsCount);
    log("Capacidad total: " + (designers.length * piecesPerDesigner));

    // Crear matriz de asignación correlativa
    var assignmentMatrix = createAssignmentMatrix(designers.length, piecesPerDesigner, categories);

    // Usar el slide actual
    var presentation = SlidesApp.getActivePresentation();
    var presentationId = presentation.getId();
    var currentSlide = presentation.getSelection().getCurrentPage();

    if (!currentSlide) {
      throw new Error("No hay un slide seleccionado. Por favor selecciona un slide.");
    }

    var slideId = currentSlide.getObjectId();
    log("Usando slide actual con ID: " + slideId);

    // Crear tabla en el slide actual
    createTableInSlide(presentationId, slideId, designers, piecesPerDesigner, assignmentMatrix, log);

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
 * Crea la matriz de asignación distribuyendo categorías correlativamente
 * Las categorías se reciclan/repiten hasta llenar toda la tabla
 * @returns {Array} - Matriz [diseñadores][columnas] con las categorías asignadas
 */
function createAssignmentMatrix(designersCount, piecesPerDesigner, categories) {
  var matrix = [];
  var categoryIndex = 0;

  // Distribuir categorías correlativamente (reciclándolas si es necesario)
  for (var i = 0; i < designersCount; i++) {
    matrix[i] = [];
    for (var j = 0; j < piecesPerDesigner; j++) {
      // Usar módulo para reciclar las categorías cuando se acaben
      matrix[i][j] = categories[categoryIndex % categories.length];
      categoryIndex++;
    }
  }

  return matrix;
}

/**
 * Crea la tabla en el slide actual usando API avanzada
 */
function createTableInSlide(presentationId, slideId, designers, piecesPerDesigner, matrix, logFunction) {
  var rows = designers.length; // Sin encabezado
  var cols = 1 + piecesPerDesigner; // 1 para nombre + piezas por diseñador

  logFunction("Creando tabla de " + rows + " filas x " + cols + " columnas");

  var tableId = Utilities.getUuid();
  var requests = [];

  // 1. Crear la tabla en el slide actual
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

    // Resto de columnas: categorías asignadas correlativamente (siempre tienen valor)
    for (var j = 0; j < piecesPerDesigner; j++) {
      requests.push(createCellTextRequest(tableId, i, colIndex, matrix[i][j]));
      colIndex++;
    }
  }

  // 4. Cambiar color de texto a blanco en todas las celdas
  for (var i = 0; i < designers.length; i++) {
    for (var colIndex = 0; colIndex < cols; colIndex++) {
      requests.push(createWhiteTextStyleRequest(tableId, i, colIndex));
    }
  }

  // Ejecutar TODO en una sola llamada
  try {
    Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
    logFunction("✓ Tabla creada exitosamente en el slide actual");
  } catch (e) {
    logFunction("✗ Error creando tabla: " + e.toString());
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
 * Crea un request para cambiar el color del texto a blanco
 */
function createWhiteTextStyleRequest(tableId, rowIndex, colIndex) {
  var request = {
    updateTextStyle: {
      objectId: tableId,
      cellLocation: {
        rowIndex: rowIndex,
        columnIndex: colIndex
      },
      style: {
        foregroundColor: {
          opaqueColor: {
            rgbColor: {
              red: 1.0,
              green: 1.0,
              blue: 1.0
            }
          }
        }
      },
      fields: "foregroundColor"
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
    categories: ["1", "2", "3", "4", "5", "6"],
    piecesPerDesigner: 6, // Cada diseñador hace 6 piezas (6 columnas)
    stillsCount: 4 // Primeras 4 COLUMNAS son Stills, últimas 2 son Conceptuales
  };

  var result = generateTaskAssignmentTable(testData);
  Logger.log(result.log);

  return result;
}
