/**
 * TaskAssignmentModule.js
 *
 * Módulo para crear tablas de asignación de tareas de diseño.
 * Distribuye temas entre diseñadores según su carga de trabajo (Stills y Conceptuales).
 */

/**
 * Genera una tabla de asignación en un nuevo slide
 *
 * @param {Object} assignmentData - Datos de la asignación
 * @param {Array} assignmentData.designers - Array de objetos {name: string, stills: number, conceptual: number}
 * @param {Array} assignmentData.topics - Array de strings con los temas (ej: ["1a", "1b", "2", "3a", "3b"])
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
    if (!assignmentData || !assignmentData.designers || !assignmentData.topics) {
      throw new Error("Datos incompletos. Se requieren diseñadores y temas.");
    }

    if (assignmentData.designers.length === 0) {
      throw new Error("Debe haber al menos un diseñador.");
    }

    if (assignmentData.topics.length === 0) {
      throw new Error("Debe haber al menos un tema.");
    }

    var designers = assignmentData.designers;
    var topics = assignmentData.topics;

    log("Diseñadores: " + designers.length);
    log("Temas: " + topics.length);

    // Calcular totales por diseñador
    var designersWithTotal = designers.map(function(d) {
      return {
        name: d.name,
        stills: parseInt(d.stills) || 0,
        conceptual: parseInt(d.conceptual) || 0,
        total: (parseInt(d.stills) || 0) + (parseInt(d.conceptual) || 0)
      };
    });

    // Calcular total de piezas
    var totalPieces = designersWithTotal.reduce(function(sum, d) {
      return sum + d.total;
    }, 0);

    log("Total de piezas a asignar: " + totalPieces);

    // Validar que hay suficientes temas
    if (topics.length > totalPieces) {
      log("⚠️ Advertencia: Hay más temas (" + topics.length + ") que piezas totales (" + totalPieces + ")");
    }

    // Crear matriz de asignación
    var assignmentMatrix = createAssignmentMatrix(designersWithTotal, topics);

    // Crear slide y tabla usando API avanzada
    var presentation = SlidesApp.getActivePresentation();
    var presentationId = presentation.getId();
    var slideId = Utilities.getUuid();

    log("Creando slide con ID: " + slideId);

    // Crear slide usando API avanzada
    createSlideAndTable(presentationId, slideId, designersWithTotal, topics, assignmentMatrix, log);

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
 * Crea la matriz de asignación (quién hace qué tema)
 * Distribuye los temas correlativamente entre diseñadores
 */
function createAssignmentMatrix(designers, topics) {
  var matrix = [];
  var topicIndex = 0;

  // Inicializar matriz vacía
  for (var i = 0; i < designers.length; i++) {
    matrix[i] = [];
    for (var j = 0; j < topics.length; j++) {
      matrix[i][j] = false;
    }
  }

  // Asignar temas correlativamente
  for (var i = 0; i < designers.length; i++) {
    var designer = designers[i];
    var piecesToAssign = designer.total;

    for (var p = 0; p < piecesToAssign; p++) {
      if (topicIndex < topics.length) {
        matrix[i][topicIndex] = true;
        topicIndex++;
      } else {
        // Ya no hay más temas para asignar
        break;
      }
    }
  }

  return matrix;
}

/**
 * Crea el slide y la tabla en una sola operación usando API avanzada
 */
function createSlideAndTable(presentationId, slideId, designers, topics, matrix, logFunction) {
  var rows = designers.length + 1; // +1 para header
  var cols = topics.length + 1; // +1 para nombres

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
          height: { magnitude: (rows * 0.4) * 914400, unit: 'EMU' }
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

  // Header: primera celda vacía
  requests.push(createCellTextRequest(tableId, 0, 0, "", true));

  // Header: temas en primera fila
  for (var j = 0; j < topics.length; j++) {
    requests.push(createCellTextRequest(tableId, 0, j + 1, topics[j], true));
  }

  // Filas de diseñadores
  for (var i = 0; i < designers.length; i++) {
    var designer = designers[i];

    // Primera columna: nombre del diseñador + cantidades
    var nameText = designer.name + "\n(S:" + designer.stills + " C:" + designer.conceptual + ")";
    requests.push(createCellTextRequest(tableId, i + 1, 0, nameText, false));

    // Resto de columnas: marcar asignaciones
    for (var j = 0; j < topics.length; j++) {
      var cellText = matrix[i][j] ? "✓" : "";
      requests.push(createCellTextRequest(tableId, i + 1, j + 1, cellText, false));
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
function createCellTextRequest(tableId, rowIndex, colIndex, text, isHeader) {
  var request = {
    insertText: {
      objectId: tableId,
      cellLocation: {
        tableCellLocation: {
          tableObjectId: tableId,
          rowIndex: rowIndex,
          columnIndex: colIndex
        }
      },
      text: text,
      insertionIndex: 0
    }
  };

  return request;
}

/**
 * Valida el formato de un tema (permite números y letras)
 * Ejemplos válidos: "1", "1a", "1b", "2", "10c", etc.
 */
function validateTopicFormat(topic) {
  var pattern = /^\d+[a-z]?$/i;
  return pattern.test(topic.trim());
}

/**
 * Ordena temas en orden natural (1, 1a, 1b, 2, 2a, 3, etc.)
 */
function sortTopics(topics) {
  return topics.sort(function(a, b) {
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
      { name: "Ana García", stills: 3, conceptual: 2 },
      { name: "Carlos López", stills: 2, conceptual: 1 },
      { name: "María Torres", stills: 1, conceptual: 2 }
    ],
    topics: ["1a", "1b", "2", "3a", "3b", "4", "5", "6"]
  };

  var result = generateTaskAssignmentTable(testData);
  Logger.log(result.log);

  return result;
}
