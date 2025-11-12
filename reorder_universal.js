/**
 * ReorderUniversal.js
 * 
 * Funciona con cualquier tipo de índice: automático, manual, o sin mapeo previo
 */

function reorderFromLinkedIndexModule() {
  var debugLog = "";
  var startTime = new Date().getTime();
  
  function log(msg) {
    var timestamp = new Date().getTime() - startTime;
    var logMsg = "[" + timestamp + "ms] " + msg;
    debugLog += logMsg + "\n";
    Logger.log(logMsg);
  }
  
  function returnResult(success, message) {
    if (message) {
      log(message);
    }
    return { 
      success: success, 
      log: debugLog 
    };
  }
  
  log("=== REORDENAMIENTO UNIVERSAL ===");
  log("Fecha/Hora: " + new Date().toString());
  
  try {
    var presentation = SlidesApp.getActivePresentation();
    var slides = presentation.getSlides();
    log("✓ Presentación obtenida: " + slides.length + " slides");
    
    // 1. Intentar cargar mapeo guardado (si existe)
    log("\n1. VERIFICANDO MAPEO EXISTENTE");
    var savedMapping = loadMappingFromPresentation();
    var hasSavedMapping = savedMapping !== null;
    
    if (hasSavedMapping) {
      log("✓ Mapeo automático encontrado - Usando método con correlación");
      var result = reorderWithSavedMapping(presentation, savedMapping, log);
      return returnResult(result.success, result.success ? "Reordenamiento con mapeo completado" : "Error en reordenamiento con mapeo");
    } else {
      log("⚠️ No hay mapeo guardado - Usando método de detección automática");
      var result = reorderWithoutMapping(presentation, log);
      return returnResult(result.success, result.success ? "Reordenamiento sin mapeo completado" : "Error en reordenamiento sin mapeo");
    }
    
  } catch (error) {
    log("\n✗ ERROR GENERAL: " + error.toString());
    return returnResult(false, "Error: " + error.toString());
  }
}

/**
 * Método 1: Con mapeo guardado (índice generado automáticamente)
 */
function reorderWithSavedMapping(presentation, savedMapping, logFunction) {
  var slides = presentation.getSlides();
  
  logFunction("🔧 Usando mapeo guardado para correlación precisa...");
  
  // 1. Detectar slides de índice usando el mapeo
  logFunction("\n2. DETECTANDO SLIDES DE ÍNDICE DESDE MAPEO");
  var indexSlides = detectIndexSlidesFromMapping(slides, savedMapping);
  logFunction("✓ Slides de índice detectados: " + indexSlides.length + " slides");
  
  // 2. Extraer orden visual actual y correlacionar con grupos originales
  logFunction("\n3. EXTRAYENDO ORDEN VISUAL Y CORRELACIONANDO");
  var visualGroupOrder = extractVisualGroupOrderFromMapping(indexSlides, savedMapping, logFunction);
  logFunction("✓ Orden visual correlacionado: " + visualGroupOrder.length + " grupos");
  
  // 3. Detectar grupos UBA actuales
  logFunction("\n4. DETECTANDO GRUPOS UBA ACTUALES");
  var currentUBAGroups = detectCurrentUBAGroupsSimple(slides, indexSlides.length, logFunction);
  logFunction("✓ Grupos UBA actuales: " + currentUBAGroups.length + " grupos");
  
  // 4. Crear plan de reordenamiento
  logFunction("\n5. CREANDO PLAN DE REORDENAMIENTO");
  var reorderPlan = createMappingBasedReorderPlan(slides, visualGroupOrder, currentUBAGroups, indexSlides.length, logFunction);
  
  if (reorderPlan.operations.length === 0) {
    logFunction("✓ Los grupos ya están en el orden correcto");
    return { success: true, log: logFunction.debugLog };
  }
  
  // 5. Ejecutar reordenamiento
  logFunction("\n6. EJECUTANDO REORDENAMIENTO");
  var result = executeReorderPlan(presentation, reorderPlan, logFunction);
  
  // 6. Actualizar el mapeo guardado después del reordenamiento
  logFunction("\n7. ACTUALIZANDO MAPEO GUARDADO");
  updateSavedMappingAfterReorder(savedMapping, visualGroupOrder, logFunction);
  
  logFunction("\n=== REORDENAMIENTO CON MAPEO COMPLETADO ===");
  logFunction("📊 Operaciones ejecutadas: " + result.operationsExecuted);
  logFunction("📊 Slides movidos: " + result.slidesMoved);
  
  return { success: true, log: logFunction.debugLog };
}

/**
 * Detecta slides de índice usando el mapeo guardado
 */
function detectIndexSlidesFromMapping(slides, mapping) {
  var indexSlides = [];
  
  for (var i = 0; i < mapping.slides.length; i++) {
    var slideMapping = mapping.slides[i];
    
    // Encontrar slide por ID
    for (var j = 0; j < slides.length; j++) {
      if (slides[j].getObjectId() === slideMapping.slideId) {
        indexSlides.push({
          slide: slides[j],
          position: j,
          slideId: slideMapping.slideId,
          originalMapping: slideMapping
        });
        break;
      }
    }
  }
  
  return indexSlides;
}

/**
 * Extrae el orden visual actual y devuelve qué grupos deben ir en cada posición
 * MEJORADO: Ahora lee los títulos de las imágenes como método principal
 */
function extractVisualGroupOrderFromMapping(indexSlides, originalMapping, logFunction) {
  var visualGroupOrder = [];
  
  for (var i = 0; i < indexSlides.length; i++) {
    var indexSlide = indexSlides[i];
    var slideMapping = indexSlide.originalMapping;
    
    logFunction("  📄 Analizando slide de índice " + (i + 1) + "...");
    
    try {
      var elements = indexSlide.slide.getPageElements();
      var imageElements = [];
      
      // Extraer todas las imágenes con sus posiciones actuales
      for (var j = 0; j < elements.length; j++) {
        if (elements[j].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
          var element = elements[j];
          var image = element.asImage();
          
          var imageInfo = {
            left: element.getLeft(),
            top: element.getTop(),
            width: element.getWidth(),
            height: element.getHeight(),
            objectId: element.getObjectId(),
            driveId: null,
            title: null,
            groupFromTitle: null
          };
          
          // NUEVO: Leer el título de la imagen
          try {
            imageInfo.title = image.getTitle();
            if (imageInfo.title && imageInfo.title.startsWith("BOOGIE_GROUP_")) {
              var groupNum = parseInt(imageInfo.title.split("_")[2]);
              if (!isNaN(groupNum)) {
                imageInfo.groupFromTitle = groupNum - 1; // Convertir a índice base-0
                logFunction("    🏷️ Título encontrado: " + imageInfo.title + " → Grupo " + groupNum);
              }
            }
          } catch (titleError) {
            // Ignorar error al leer título
          }
          
          // Intentar obtener Drive ID como fallback
          try {
            var sourceUrl = image.getSourceUrl();
            if (sourceUrl) {
              var match = sourceUrl.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
              if (match) {
                imageInfo.driveId = match[1];
              }
            }
          } catch (e) {
            // Ignorar errores de sourceUrl
          }
          
          imageElements.push(imageInfo);
        }
      }
      
      // Ordenar por posición visual ACTUAL
      imageElements.sort(function(a, b) {
        var topDiff = a.top - b.top;
        var tolerance = 50;
        
        if (Math.abs(topDiff) < tolerance) {
          return a.left - b.left;
        }
        return topDiff;
      });
      
      logFunction("    → " + imageElements.length + " imágenes en orden visual actual");
      
      // Para cada imagen en orden visual, determinar qué grupo original representa
      for (var k = 0; k < imageElements.length; k++) {
        var imgElement = imageElements[k];
        var originalGroup = null;
        
        // PRIORIDAD 1: Usar el título si existe
        if (imgElement.groupFromTitle !== null) {
          originalGroup = imgElement.groupFromTitle;
          logFunction("      ✓ Posición visual " + (visualGroupOrder.length + 1) + 
                      " → Grupo " + (originalGroup + 1) + " (por título)");
        } else {
          // FALLBACK: Usar métodos anteriores si no hay título
          originalGroup = findOriginalGroupForImageEnhanced(slideMapping, imgElement, k);
          if (originalGroup !== null) {
            logFunction("      ✓ Posición visual " + (visualGroupOrder.length + 1) + 
                        " → Grupo " + (originalGroup + 1) + " (por correlación)");
          }
        }
        
        if (originalGroup !== null) {
          visualGroupOrder.push({
            visualPosition: visualGroupOrder.length + 1,
            originalGroupIndex: originalGroup,
            imageObjectId: imgElement.objectId,
            driveId: imgElement.driveId,
            positionInGrid: k,
            title: imgElement.title
          });
        }
      }
      
    } catch (e) {
      logFunction("    ✗ Error procesando slide: " + e.toString());
    }
  }
  
  return visualGroupOrder;
}

/**
 * Encuentra qué grupo original representa una imagen (MEJORADO)
 * Ahora maneja casos donde el Object ID cambió por copy-paste
 */
function findOriginalGroupForImageEnhanced(slideMapping, imgElement, positionIndex) {
  // Método 1: Buscar por Object ID (si no cambió)
  for (var i = 0; i < slideMapping.thumbnails.length; i++) {
    var thumb = slideMapping.thumbnails[i];
    if (thumb.imageObjectId === imgElement.objectId) {
      return thumb.groupIndex;
    }
  }
  
  // Método 2: Buscar por Drive ID (más confiable si existe)
  if (imgElement.driveId) {
    for (var i = 0; i < slideMapping.thumbnails.length; i++) {
      var thumb = slideMapping.thumbnails[i];
      if (thumb.driveId === imgElement.driveId) {
        Logger.log("    → Match por Drive ID para posición " + (positionIndex + 1));
        return thumb.groupIndex;
      }
    }
  }
  
  // Método 3: Buscar por posición en el grid (cuando falla todo lo demás)
  // Esto asume que el orden no cambió, solo los IDs
  var gridPositions = [
    {x: 58,  y: 15,  pos: 0}, {x: 274, y: 15,  pos: 1}, {x: 490, y: 15,  pos: 2},
    {x: 58,  y: 135, pos: 3}, {x: 274, y: 135, pos: 4}, {x: 490, y: 135, pos: 5},
    {x: 58,  y: 256, pos: 6}, {x: 274, y: 256, pos: 7}, {x: 490, y: 256, pos: 8}
  ];
  
  var TOLERANCE = 60;
  
  for (var i = 0; i < gridPositions.length; i++) {
    var gridPos = gridPositions[i];
    var distance = Math.sqrt(
      Math.pow(imgElement.left - gridPos.x, 2) + 
      Math.pow(imgElement.top - gridPos.y, 2)
    );
    
    if (distance < TOLERANCE) {
      // Buscar thumbnail en esa posición del grid
      for (var j = 0; j < slideMapping.thumbnails.length; j++) {
        if (slideMapping.thumbnails[j].position === gridPos.pos) {
          Logger.log("    → Match por posición grid para índice " + (positionIndex + 1));
          return slideMapping.thumbnails[j].groupIndex;
        }
      }
    }
  }
  
  // Método 4: Último recurso - usar el índice de posición directamente
  // Esto asume que las imágenes están en el mismo orden
  if (positionIndex < slideMapping.thumbnails.length) {
    var thumb = slideMapping.thumbnails[positionIndex];
    if (thumb) {
      Logger.log("    → Match por índice secuencial para posición " + (positionIndex + 1));
      return thumb.groupIndex;
    }
  }
  
  Logger.log("    ✗ No se pudo determinar grupo para imagen en posición " + (positionIndex + 1));
  return null;
}

/**
 * Detecta grupos UBA actuales de forma simple
 */
function detectCurrentUBAGroupsSimple(slides, startIndex, logFunction) {
  var ubaGroups = [];
  var MIN_UBA_WIDTH = 400;
  var MIN_UBA_HEIGHT = 200;
  
  logFunction("  🔍 Detectando grupos UBA desde posición " + (startIndex + 1) + "...");
  
  for (var i = startIndex; i < slides.length - 2; i++) {
    try {
      var slide = slides[i];
      
      if (slideHasLargeImage(slide, MIN_UBA_WIDTH, MIN_UBA_HEIGHT)) {
        var group = {
          slides: [i, i + 1, i + 2],
          currentIndex: ubaGroups.length
        };
        
        ubaGroups.push(group);
        logFunction("    ✓ Grupo actual " + (ubaGroups.length) + " en slides " + 
                    (i + 1) + "-" + (i + 3));
        
        i += 2;
      }
    } catch (e) {
      // Ignorar errores
    }
  }
  
  return ubaGroups;
}

/**
 * Crea plan de reordenamiento basado en mapeo (CORREGIDO)
 */
function createMappingBasedReorderPlan(slides, visualGroupOrder, currentUBAGroups, indexSlideCount, logFunction) {
  var plan = { operations: [] };
  
  logFunction("📋 Creando plan basado en mapeo (optimizado)...");
  logFunction("  📊 Orden visual deseado: " + visualGroupOrder.length + " grupos");
  logFunction("  📊 Grupos UBA actuales: " + currentUBAGroups.length + " grupos");
  
  // CORRECCIÓN: Crear un mapeo de reordenamiento completo primero
  var reorderMapping = [];
  
  for (var visualPos = 0; visualPos < visualGroupOrder.length; visualPos++) {
    var desiredGroup = visualGroupOrder[visualPos];
    var targetOriginalIndex = desiredGroup.originalGroupIndex;
    var newBasePosition = indexSlideCount + (visualPos * 3);
    
    if (targetOriginalIndex < currentUBAGroups.length) {
      var currentGroup = currentUBAGroups[targetOriginalIndex];
      
      // Convertir posiciones a base-1 para el log
      var currentSlidesDisplay = currentGroup.slides.map(function(s) { return s + 1; });
      var targetSlidesDisplay = [newBasePosition + 1, newBasePosition + 2, newBasePosition + 3];
      
      reorderMapping.push({
        visualPosition: visualPos + 1,
        originalGroupIndex: targetOriginalIndex,
        currentSlides: currentGroup.slides.slice(), // Copiar array
        newBasePosition: newBasePosition,
        targetSlides: [newBasePosition, newBasePosition + 1, newBasePosition + 2]
      });
      
      logFunction("  📍 Visual " + (visualPos + 1) + ": Grupo original " + (targetOriginalIndex + 1) + 
                  " de slides " + currentSlidesDisplay.join("-") + 
                  " → slides " + targetSlidesDisplay.join("-"));
    }
  }
  
  // CORRECCIÓN: Detectar movimientos que realmente necesitan hacerse
  var movementsNeeded = [];
  
  for (var i = 0; i < reorderMapping.length; i++) {
    var mapping = reorderMapping[i];
    var needsMovement = false;
    
    // Verificar si algún slide está en posición incorrecta
    for (var j = 0; j < 3; j++) {
      if (mapping.currentSlides[j] !== mapping.targetSlides[j]) {
        needsMovement = true;
        break;
      }
    }
    
    if (needsMovement) {
      movementsNeeded.push(mapping);
    }
  }
  
  logFunction("  📊 Grupos que necesitan movimiento: " + movementsNeeded.length + "/" + reorderMapping.length);
  
  // CORRECCIÓN: Optimizar orden de movimientos para evitar conflictos
  // Mover primero los grupos que van hacia atrás, luego los que van hacia adelante
  movementsNeeded.sort(function(a, b) {
    var aGoingBackward = a.newBasePosition < a.currentSlides[0];
    var bGoingBackward = b.newBasePosition < b.currentSlides[0];
    
    if (aGoingBackward && !bGoingBackward) return -1;
    if (!aGoingBackward && bGoingBackward) return 1;
    
    // Si ambos van en la misma dirección, ordenar por posición actual
    if (aGoingBackward) {
      return a.currentSlides[0] - b.currentSlides[0]; // Hacia atrás: orden ascendente
    } else {
      return b.currentSlides[0] - a.currentSlides[0]; // Hacia adelante: orden descendente
    }
  });
  
  // Crear operaciones de movimiento
  for (var i = 0; i < movementsNeeded.length; i++) {
    var mapping = movementsNeeded[i];
    
    logFunction("    → Movimiento " + (i + 1) + ": Grupo " + (mapping.originalGroupIndex + 1) + 
                " a posición visual " + mapping.visualPosition);
    
    // Crear operaciones para los 3 slides del grupo
    for (var slideIndex = 0; slideIndex < 3; slideIndex++) {
      var currentPos = mapping.currentSlides[slideIndex];
      var newPos = mapping.targetSlides[slideIndex];
      
      plan.operations.push({
        slideId: slides[currentPos].getObjectId(),
        currentPosition: currentPos,
        newPosition: newPos,
        visualOrder: mapping.visualPosition,
        originalGroupIndex: mapping.originalGroupIndex,
        slideType: slideIndex === 0 ? "UBA" : (slideIndex === 1 ? "Eclipse" : "Others"),
        movementGroup: i,
        priority: mapping.newBasePosition < mapping.currentSlides[0] ? "backward" : "forward"
      });
    }
  }
  
  logFunction("  📊 Operaciones requeridas: " + plan.operations.length);
  
  // Log de operaciones agrupadas
  for (var i = 0; i < movementsNeeded.length; i++) {
    var mapping = movementsNeeded[i];
    logFunction("    " + (i + 1) + ". Grupo " + (mapping.originalGroupIndex + 1) + 
                " (" + mapping.currentSlides.join("-") + " → " + 
                mapping.targetSlides.map(function(s) { return s + 1; }).join("-") + ")");
  }
  
  return plan;
}

/**
 * Ejecuta el plan de reordenamiento (OPTIMIZADO)
 */
function executeReorderPlan(presentation, plan, logFunction) {
  var result = {
    operationsExecuted: 0,
    slidesMoved: 0,
    errors: []
  };
  
  if (plan.operations.length === 0) {
    logFunction("  ✓ No hay operaciones que ejecutar");
    return result;
  }
  
  // CORRECCIÓN: Ejecutar movimientos por grupos completos para mantener integridad
  logFunction("  🚀 Ejecutando movimientos por grupos completos...");
  
  // Agrupar operaciones por movementGroup
  var groupedOperations = {};
  for (var i = 0; i < plan.operations.length; i++) {
    var op = plan.operations[i];
    var groupKey = op.movementGroup;
    
    if (!groupedOperations[groupKey]) {
      groupedOperations[groupKey] = [];
    }
    groupedOperations[groupKey].push(op);
  }
  
  // Obtener grupos ordenados
  var groupKeys = Object.keys(groupedOperations);
  groupKeys.sort(function(a, b) { return parseInt(a) - parseInt(b); });
  
  logFunction("  📊 Ejecutando " + groupKeys.length + " grupos de movimientos...");
  
  // Ejecutar cada grupo de movimientos
  for (var g = 0; g < groupKeys.length; g++) {
    var groupKey = groupKeys[g];
    var groupOps = groupedOperations[groupKey];
    
    // Ordenar operaciones dentro del grupo por tipo (UBA, Eclipse, Others)
    groupOps.sort(function(a, b) {
      var order = { "UBA": 0, "Eclipse": 1, "Others": 2 };
      return order[a.slideType] - order[b.slideType];
    });
    
    var firstOp = groupOps[0];
    logFunction("    🔄 Grupo " + (g + 1) + "/" + groupKeys.length + 
                ": Moviendo grupo original " + (firstOp.originalGroupIndex + 1) + 
                " a posición visual " + firstOp.visualOrder + 
                " (" + firstOp.priority + ")");
    
    // IMPORTANTE: Obtener las posiciones actuales ANTES de mover el grupo
    var currentSlides = presentation.getSlides();
    var groupSlidesInfo = [];
    
    // Recopilar información actual de todos los slides del grupo
    for (var i = 0; i < groupOps.length; i++) {
      var operation = groupOps[i];
      var currentPos = -1;
      
      // Encontrar posición actual del slide
      for (var j = 0; j < currentSlides.length; j++) {
        if (currentSlides[j].getObjectId() === operation.slideId) {
          currentPos = j;
          break;
        }
      }
      
      groupSlidesInfo.push({
        slideId: operation.slideId,
        currentPosition: currentPos,
        targetPosition: operation.newPosition,
        slideType: operation.slideType
      });
    }
    
    // Mover los slides del grupo en orden correcto
    // Primero ordenar por posición actual (de mayor a menor si van hacia adelante)
    var isMovingForward = groupSlidesInfo[0].targetPosition > groupSlidesInfo[0].currentPosition;
    
    if (isMovingForward) {
      // Si se mueven hacia adelante, mover primero el último
      groupSlidesInfo.sort(function(a, b) {
        return b.currentPosition - a.currentPosition;
      });
    } else {
      // Si se mueven hacia atrás, mover primero el primero
      groupSlidesInfo.sort(function(a, b) {
        return a.currentPosition - b.currentPosition;
      });
    }
    
    // Ejecutar movimientos
    for (var i = 0; i < groupSlidesInfo.length; i++) {
      var slideInfo = groupSlidesInfo[i];
      
      try {
        // Obtener el slide actual
        var slideToMove = null;
        var updatedSlides = presentation.getSlides();
        
        for (var j = 0; j < updatedSlides.length; j++) {
          if (updatedSlides[j].getObjectId() === slideInfo.slideId) {
            slideToMove = updatedSlides[j];
            break;
          }
        }
        
        if (slideToMove) {
          logFunction("      → " + slideInfo.slideType + ": slide pos " + 
                      (slideInfo.currentPosition + 1) + " → " + (slideInfo.targetPosition + 1));
          
          slideToMove.move(slideInfo.targetPosition);
          result.operationsExecuted++;
          result.slidesMoved++;
          
          // Pausa pequeña entre slides para estabilidad
          Utilities.sleep(100);
          
        } else {
          var error = "Slide no encontrado: " + slideInfo.slideId;
          logFunction("        ✗ " + error);
          result.errors.push(error);
        }
        
      } catch (e) {
        var error = "Error moviendo slide: " + e.toString();
        logFunction("        ✗ " + error);
        result.errors.push(error);
      }
    }
    
    // Pausa más larga entre grupos para que se estabilice
    if (g < groupKeys.length - 1) {
      Utilities.sleep(300);
    }
  }
  
  logFunction("  ✅ Completado: " + result.operationsExecuted + "/" + plan.operations.length + " operaciones");
  if (result.errors.length > 0) {
    logFunction("  ⚠️ Errores: " + result.errors.length);
  }
  
  return result;
}

/**
 * Actualiza el mapeo guardado después del reordenamiento
 */
function updateSavedMappingAfterReorder(savedMapping, visualGroupOrder, logFunction) {
  try {
    logFunction("🔄 Actualizando mapeo guardado...");
    
    // Actualizar las posiciones de los grupos en el mapeo
    for (var i = 0; i < savedMapping.slides.length; i++) {
      var slideMapping = savedMapping.slides[i];
      
      for (var j = 0; j < slideMapping.thumbnails.length; j++) {
        var thumb = slideMapping.thumbnails[j];
        
        // Encontrar el nuevo orden de este grupo
        for (var k = 0; k < visualGroupOrder.length; k++) {
          if (visualGroupOrder[k].originalGroupIndex === thumb.groupIndex) {
            // Actualizar con la nueva posición visual
            thumb.visualPosition = visualGroupOrder[k].visualPosition;
            break;
          }
        }
      }
    }
    
    // Guardar el mapeo actualizado
    savedMapping.lastReorder = new Date().toISOString();
    var properties = PropertiesService.getDocumentProperties();
    properties.setProperty('BOOGIE_THUMBNAIL_MAPPING', JSON.stringify(savedMapping));
    
    logFunction("✓ Mapeo actualizado exitosamente");
    
    // NUEVO: Actualizar los títulos de las imágenes para reflejar el nuevo orden
    logFunction("\n🏷️ Actualizando títulos de thumbnails...");
    updateThumbnailTitles(visualGroupOrder, logFunction);
    
  } catch (e) {
    logFunction("❌ Error actualizando mapeo: " + e.toString());
  }
}

/**
 * Actualiza los títulos de los thumbnails para reflejar el nuevo orden
 */
function updateThumbnailTitles(visualGroupOrder, logFunction) {
  try {
    var presentation = SlidesApp.getActivePresentation();
    var slides = presentation.getSlides();
    var updatedCount = 0;
    
    // Agrupar por slide para eficiencia
    var updatesBySlide = {};
    
    for (var i = 0; i < visualGroupOrder.length; i++) {
      var item = visualGroupOrder[i];
      var newTitle = "BOOGIE_GROUP_" + item.visualPosition;
      
      // Buscar la imagen por ObjectId
      for (var s = 0; s < slides.length; s++) {
        var elements = slides[s].getPageElements();
        
        for (var e = 0; e < elements.length; e++) {
          if (elements[e].getObjectId() === item.imageObjectId) {
            try {
              var image = elements[e].asImage();
              var oldTitle = image.getTitle();
              image.setTitle(newTitle);
              updatedCount++;
              
              if (oldTitle !== newTitle) {
                logFunction("  ✓ Actualizado: " + (oldTitle || "sin título") + " → " + newTitle);
              }
              
              break;
            } catch (e) {
              logFunction("  ⚠️ Error actualizando título: " + e.toString());
            }
          }
        }
      }
    }
    
    logFunction("  📊 Títulos actualizados: " + updatedCount + "/" + visualGroupOrder.length);
    
  } catch (e) {
    logFunction("  ❌ Error actualizando títulos: " + e.toString());
  }
}

/**
 * Método 2: Sin mapeo (índice manual o sin índice previo)
 */
function reorderWithoutMapping(presentation, logFunction) {
  var slides = presentation.getSlides();
  
  logFunction("🔧 Detectando estructura automáticamente...");
  
  // 1. Detectar slides de índice
  logFunction("\n2. DETECTANDO SLIDES DE ÍNDICE");
  var indexSlides = detectIndexSlidesAutomatically(slides);
  
  if (indexSlides.length === 0) {
    return { 
      success: false, 
      log: logFunction.debugLog + "\n✗ No se encontraron slides de índice válidos" 
    };
  }
  
  logFunction("✓ Slides de índice detectados: " + indexSlides.length);
  
  // 2. Extraer orden visual actual
  logFunction("\n3. EXTRAYENDO ORDEN VISUAL");
  var visualOrder = extractVisualOrderWithoutMapping(indexSlides, logFunction);
  logFunction("✓ Orden visual extraído: " + visualOrder.length + " imágenes");
  
  // 3. Detectar grupos UBA
  logFunction("\n4. DETECTANDO GRUPOS UBA");
  var ubaGroups = detectUBAGroupsAfterIndex(slides, indexSlides.length, logFunction);
  logFunction("✓ Grupos UBA detectados: " + ubaGroups.length);
  
  // 4. Método de correlación inteligente
  logFunction("\n5. CORRELACIONANDO IMÁGENES CON GRUPOS");
  var correlation = correlateImagesWithGroups(visualOrder, ubaGroups, slides, logFunction);
  
  if (correlation.method === "failed") {
    return { 
      success: false, 
      log: logFunction.debugLog + "\n✗ No se pudo correlacionar automáticamente" 
    };
  }
  
  // 5. Crear y ejecutar plan
  logFunction("\n6. CREANDO PLAN DE REORDENAMIENTO");
  var plan = createUniversalReorderPlan(slides, correlation.mapping, indexSlides.length, logFunction);
  
  if (plan.operations.length === 0) {
    logFunction("✓ Los grupos ya están en orden correcto");
    return { success: true, log: logFunction.debugLog };
  }
  
  logFunction("\n7. EJECUTANDO REORDENAMIENTO");
  var result = executeUniversalPlan(presentation, plan, logFunction);
  
  logFunction("\n=== REORDENAMIENTO COMPLETADO ===");
  logFunction("📊 Método usado: " + correlation.method);
  logFunction("📊 Operaciones ejecutadas: " + result.operationsExecuted);
  
  return { success: true, log: logFunction.debugLog };
}

/**
 * Detecta slides de índice automáticamente
 */
function detectIndexSlidesAutomatically(slides) {
  var indexSlides = [];
  var MIN_IMAGES = 2;
  var MAX_IMAGE_SIZE = 350;
  
  for (var i = 0; i < slides.length; i++) {
    try {
      var slide = slides[i];
      var elements = slide.getPageElements();
      var smallImages = 0;
      var largeImages = 0;
      var totalImages = 0;
      
      for (var j = 0; j < elements.length; j++) {
        if (elements[j].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
          totalImages++;
          var width = elements[j].getWidth();
          var height = elements[j].getHeight();
          
          if (width <= MAX_IMAGE_SIZE && height <= MAX_IMAGE_SIZE) {
            smallImages++;
          } else {
            largeImages++;
          }
        }
      }
      
      // Criterios para ser slide de índice:
      // 1. Múltiples imágenes pequeñas
      // 2. Pocas o ninguna imagen grande
      // 3. Densidad alta de imágenes
      var isIndexSlide = (
        smallImages >= MIN_IMAGES && 
        largeImages <= 1 && 
        totalImages >= MIN_IMAGES
      );
      
      if (isIndexSlide) {
        indexSlides.push({
          slide: slide,
          position: i,
          imageCount: smallImages
        });
      } else if (largeImages > 1) {
        // Si encontramos muchas imágenes grandes, probablemente terminaron los índices
        break;
      }
    } catch (e) {
      // Ignorar errores
    }
  }
  
  return indexSlides;
}

/**
 * Extrae orden visual sin mapeo previo
 */
function extractVisualOrderWithoutMapping(indexSlides, logFunction) {
  var allImages = [];
  
  for (var i = 0; i < indexSlides.length; i++) {
    var indexSlide = indexSlides[i];
    
    try {
      var elements = indexSlide.slide.getPageElements();
      var slideImages = [];
      
      for (var j = 0; j < elements.length; j++) {
        if (elements[j].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
          var element = elements[j];
          var image = element.asImage();
          
          var imageInfo = {
            left: element.getLeft(),
            top: element.getTop(),
            width: element.getWidth(),
            height: element.getHeight(),
            objectId: element.getObjectId(),
            slideIndex: i,
            element: element,
            driveId: null,
            title: null,
            groupFromTitle: null
          };
          
          // NUEVO: Intentar leer título
          try {
            imageInfo.title = image.getTitle();
            if (imageInfo.title && imageInfo.title.startsWith("BOOGIE_GROUP_")) {
              var groupNum = parseInt(imageInfo.title.split("_")[2]);
              if (!isNaN(groupNum)) {
                imageInfo.groupFromTitle = groupNum - 1; // Base-0
                logFunction("    🏷️ Título encontrado: " + imageInfo.title);
              }
            }
          } catch (e) {
            // Ignorar error de título
          }
          
          // Intentar obtener Drive ID
          try {
            var sourceUrl = image.getSourceUrl();
            if (sourceUrl) {
              var match = sourceUrl.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
              if (match) {
                imageInfo.driveId = match[1];
              }
            }
          } catch (e) {
            // Ignorar error
          }
          
          slideImages.push(imageInfo);
        }
      }
      
      // Ordenar por posición visual
      slideImages.sort(function(a, b) {
        var topDiff = a.top - b.top;
        var tolerance = 60;
        
        if (Math.abs(topDiff) < tolerance) {
          return a.left - b.left;
        }
        return topDiff;
      });
      
      logFunction("  📄 Slide " + (i + 1) + ": " + slideImages.length + " imágenes ordenadas");
      allImages = allImages.concat(slideImages);
      
    } catch (e) {
      logFunction("  ✗ Error en slide " + (i + 1) + ": " + e.toString());
    }
  }
  
  return allImages;
}

/**
 * Detecta grupos UBA después de los slides de índice
 */
function detectUBAGroupsAfterIndex(slides, startIndex, logFunction) {
  var ubaGroups = [];
  var MIN_WIDTH = 400;
  var MIN_HEIGHT = 200;
  
  logFunction("  🔍 Buscando grupos UBA desde slide " + (startIndex + 1) + "...");
  
  for (var i = startIndex; i < slides.length - 2; i++) {
    if (slideHasLargeImage(slides[i], MIN_WIDTH, MIN_HEIGHT)) {
      var ubaInfo = extractUBAInfoForCorrelation(slides[i]);
      
      ubaGroups.push({
        slides: [i, i + 1, i + 2],
        startPos: i,
        index: ubaGroups.length,
        ubaInfo: ubaInfo
      });
      
      logFunction("    ✓ Grupo " + (ubaGroups.length) + " en slides " + 
                  (i + 1) + "-" + (i + 3) + 
                  (ubaInfo.driveId ? " (Drive ID: " + ubaInfo.driveId.substring(0, 8) + "...)" : ""));
      i += 2;
    }
  }
  
  return ubaGroups;
}

/**
 * Extrae información del UBA para correlación
 */
function extractUBAInfoForCorrelation(slide) {
  var info = {
    driveId: null,
    width: 0,
    height: 0,
    aspectRatio: 0
  };
  
  try {
    var elements = slide.getPageElements();
    var largestArea = 0;
    var largestImage = null;
    
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
        var width = elements[i].getWidth();
        var height = elements[i].getHeight();
        var area = width * height;
        
        if (area > largestArea) {
          largestArea = area;
          largestImage = elements[i];
          info.width = width;
          info.height = height;
          info.aspectRatio = width / height;
        }
      }
    }
    
    if (largestImage) {
      try {
        var image = largestImage.asImage();
        var sourceUrl = image.getSourceUrl();
        if (sourceUrl) {
          var match = sourceUrl.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
          if (match) {
            info.driveId = match[1];
          }
        }
      } catch (e) {
        // Ignorar error
      }
    }
    
  } catch (e) {
    Logger.log("Error extrayendo info UBA: " + e.toString());
  }
  
  return info;
}

/**
 * Correlaciona imágenes del índice con grupos UBA usando múltiples métodos
 */
function correlateImagesWithGroups(visualOrder, ubaGroups, slides, logFunction) {
  logFunction("🧠 Intentando correlación automática...");
  
  // Método 0: Correlación por títulos (más confiable si existe)
  var titleMapping = tryCorrelationByTitles(visualOrder, ubaGroups, logFunction);
  if (titleMapping.success && titleMapping.mappedCount === visualOrder.length) {
    return {
      method: "títulos",
      mapping: titleMapping.mapping,
      confidence: "perfecta"
    };
  }
  
  // Método 1: Correlación por Drive ID (más confiable)
  var driveIdMapping = tryCorrelationByDriveId(visualOrder, ubaGroups, logFunction);
  if (driveIdMapping.success && driveIdMapping.mappedCount >= visualOrder.length * 0.8) {
    return {
      method: "drive_id",
      mapping: driveIdMapping.mapping,
      confidence: "alta"
    };
  }
  
  // Método 2: Correlación por orden secuencial (si las cantidades coinciden)
  if (visualOrder.length === ubaGroups.length) {
    logFunction("  ✓ Método 2: Correlación secuencial (igual cantidad)");
    
    var sequentialMapping = [];
    for (var i = 0; i < visualOrder.length; i++) {
      sequentialMapping.push({
        visualPosition: i + 1,
        targetGroup: ubaGroups[i],
        sourceImage: visualOrder[i]
      });
    }
    
    return {
      method: "secuencial",
      mapping: sequentialMapping,
      confidence: "alta"
    };
  }
  
  // Método 3: Correlación por análisis visual de imágenes
  logFunction("  → Método 3: Análisis visual de similitudes");
  var visualMapping = correlateByVisualSimilarity(visualOrder, ubaGroups, slides, logFunction);
  
  if (visualMapping.success) {
    return {
      method: "visual",
      mapping: visualMapping.mapping,
      confidence: visualMapping.confidence
    };
  }
  
  // Método fallido
  logFunction("  ⚠️ Correlación automática fallida");
  
  return {
    method: "failed",
    reason: "No se pudo determinar correlación automáticamente"
  };
}

/**
 * Intenta correlación por títulos BOOGIE_GROUP_X
 */
function tryCorrelationByTitles(visualOrder, ubaGroups, logFunction) {
  logFunction("  → Método 0: Correlación por títulos BOOGIE_GROUP_X");
  
  var mapping = [];
  var mappedCount = 0;
  
  for (var i = 0; i < visualOrder.length; i++) {
    var indexImage = visualOrder[i];
    
    if (indexImage.groupFromTitle !== null && indexImage.groupFromTitle < ubaGroups.length) {
      var targetGroup = ubaGroups[indexImage.groupFromTitle];
      
      mapping.push({
        visualPosition: i + 1,
        targetGroup: targetGroup,
        sourceImage: indexImage
      });
      
      mappedCount++;
      logFunction("    ✓ Imagen " + (i + 1) + " → Grupo " + (indexImage.groupFromTitle + 1) + " (por título)");
    } else {
      mapping.push(null); // Placeholder
    }
  }
  
  logFunction("    → Títulos mapeados: " + mappedCount + "/" + visualOrder.length);
  
  // Si no todos tienen título, rellenar con orden secuencial
  if (mappedCount < visualOrder.length) {
    var usedGroups = new Set();
    for (var i = 0; i < mapping.length; i++) {
      if (mapping[i]) {
        usedGroups.add(mapping[i].targetGroup.index);
      }
    }
    
    var nextGroupIndex = 0;
    for (var i = 0; i < mapping.length; i++) {
      if (!mapping[i]) {
        // Buscar siguiente grupo no usado
        while (nextGroupIndex < ubaGroups.length && usedGroups.has(nextGroupIndex)) {
          nextGroupIndex++;
        }
        
        if (nextGroupIndex < ubaGroups.length) {
          mapping[i] = {
            visualPosition: i + 1,
            targetGroup: ubaGroups[nextGroupIndex],
            sourceImage: visualOrder[i]
          };
          usedGroups.add(nextGroupIndex);
          nextGroupIndex++;
        }
      }
    }
  }
  
  return {
    success: mappedCount > 0,
    mapping: mapping,
    mappedCount: mappedCount
  };
}

/**
 * Intenta correlación por Drive ID
 */
function tryCorrelationByDriveId(visualOrder, ubaGroups, logFunction) {
  logFunction("  → Método 1: Correlación por Drive ID");
  
  var mapping = [];
  var mappedCount = 0;
  var usedGroups = new Set();
  
  for (var i = 0; i < visualOrder.length; i++) {
    var indexImage = visualOrder[i];
    var mapped = false;
    
    if (indexImage.driveId) {
      // Buscar grupo con el mismo Drive ID
      for (var j = 0; j < ubaGroups.length; j++) {
        if (!usedGroups.has(j) && ubaGroups[j].ubaInfo.driveId === indexImage.driveId) {
          mapping.push({
            visualPosition: i + 1,
            targetGroup: ubaGroups[j],
            sourceImage: indexImage
          });
          
          usedGroups.add(j);
          mappedCount++;
          mapped = true;
          break;
        }
      }
    }
    
    if (!mapped) {
      mapping.push(null); // Placeholder para mantener orden
    }
  }
  
  logFunction("    → Drive IDs mapeados: " + mappedCount + "/" + visualOrder.length);
  
  // Rellenar huecos con grupos restantes en orden
  var remainingGroups = [];
  for (var j = 0; j < ubaGroups.length; j++) {
    if (!usedGroups.has(j)) {
      remainingGroups.push(ubaGroups[j]);
    }
  }
  
  var remainingIndex = 0;
  for (var i = 0; i < mapping.length; i++) {
    if (mapping[i] === null && remainingIndex < remainingGroups.length) {
      mapping[i] = {
        visualPosition: i + 1,
        targetGroup: remainingGroups[remainingIndex],
        sourceImage: visualOrder[i]
      };
      remainingIndex++;
    }
  }
  
  return {
    success: mappedCount > 0,
    mapping: mapping,
    mappedCount: mappedCount
  };
}

/**
 * Correlación por similitud visual básica
 */
function correlateByVisualSimilarity(visualOrder, ubaGroups, slides, logFunction) {
  logFunction("    🔍 Analizando similitudes visuales...");
  
  try {
    var mapping = [];
    var usedGroups = [];
    var successfulMatches = 0;
    
    for (var i = 0; i < Math.min(visualOrder.length, ubaGroups.length); i++) {
      var indexImage = visualOrder[i];
      var bestMatch = null;
      var bestScore = 0;
      
      // Comparar con grupos UBA no usados
      for (var j = 0; j < ubaGroups.length; j++) {
        if (usedGroups.indexOf(j) === -1) {
          var ubaGroup = ubaGroups[j];
          var score = calculateVisualSimilarity(indexImage, ubaGroup, slides);
          
          if (score > bestScore && score > 0.3) { // Umbral mínimo
            bestScore = score;
            bestMatch = j;
          }
        }
      }
      
      if (bestMatch !== null) {
        mapping.push({
          visualPosition: i + 1,
          targetGroup: ubaGroups[bestMatch],
          sourceImage: indexImage,
          similarity: bestScore
        });
        
        usedGroups.push(bestMatch);
        successfulMatches++;
      }
    }
    
    var confidence = successfulMatches / visualOrder.length;
    logFunction("    → Matches exitosos: " + successfulMatches + "/" + visualOrder.length + 
                " (confianza: " + (confidence * 100).toFixed(1) + "%)");
    
    if (confidence >= 0.6) { // 60% de confianza mínima
      return {
        success: true,
        mapping: mapping,
        confidence: confidence >= 0.8 ? "alta" : "media"
      };
    }
    
  } catch (e) {
    logFunction("    ✗ Error en análisis visual: " + e.toString());
  }
  
  return { success: false };
}

/**
 * Calcula similitud visual básica entre imagen de índice y grupo UBA
 */
function calculateVisualSimilarity(indexImage, ubaGroup, slides) {
  try {
    var ubaInfo = ubaGroup.ubaInfo;
    
    if (!ubaInfo || !indexImage.width || !indexImage.height) {
      return 0;
    }
    
    // Comparar aspect ratios
    var indexRatio = indexImage.width / indexImage.height;
    var ubaRatio = ubaInfo.aspectRatio || (ubaInfo.width / ubaInfo.height);
    
    if (!ubaRatio || ubaRatio === 0) {
      return 0;
    }
    
    var ratioDiff = Math.abs(indexRatio - ubaRatio);
    var ratioScore = Math.max(0, 1 - (ratioDiff / 2)); // 0-1
    
    // Bonus por Drive ID coincidente
    if (indexImage.driveId && ubaInfo.driveId && indexImage.driveId === ubaInfo.driveId) {
      return 1.0; // Match perfecto
    }
    
    return ratioScore;
    
  } catch (e) {
    return 0;
  }
}

/**
 * Crea plan de reordenamiento universal
 */
function createUniversalReorderPlan(slides, mapping, indexSlideCount, logFunction) {
  var plan = { operations: [] };
  
  logFunction("📋 Creando plan universal...");
  logFunction("  📊 Mapeos: " + mapping.length);
  
  for (var i = 0; i < mapping.length; i++) {
    var mapItem = mapping[i];
    var newBasePosition = indexSlideCount + (i * 3);
    
    logFunction("  📍 Visual " + mapItem.visualPosition + " → Grupo en slides " + 
                mapItem.targetGroup.slides.join("-") + " → Base " + (newBasePosition + 1));
    
    // Crear operaciones para mover los 3 slides
    for (var slideIndex = 0; slideIndex < 3; slideIndex++) {
      var currentPos = mapItem.targetGroup.slides[slideIndex];
      var newPos = newBasePosition + slideIndex;
      
      if (currentPos !== newPos) {
        plan.operations.push({
          slideId: slides[currentPos].getObjectId(),
          currentPosition: currentPos,
          newPosition: newPos,
          visualOrder: mapItem.visualPosition,
          slideType: slideIndex === 0 ? "UBA" : (slideIndex === 1 ? "Eclipse" : "Others")
        });
      }
    }
  }
  
  logFunction("  📊 Operaciones: " + plan.operations.length);
  return plan;
}

/**
 * Ejecuta plan universal
 */
function executeUniversalPlan(presentation, plan, logFunction) {
  var result = { operationsExecuted: 0, slidesMoved: 0, errors: [] };
  
  if (plan.operations.length === 0) {
    logFunction("  ✓ No se requieren movimientos");
    return result;
  }
  
  var sortedOps = plan.operations.slice();
  sortedOps.sort(function(a, b) { return b.currentPosition - a.currentPosition; });
  
  logFunction("  🚀 Ejecutando " + sortedOps.length + " operaciones...");
  
  for (var i = 0; i < sortedOps.length; i++) {
    var op = sortedOps[i];
    
    try {
      var currentSlides = presentation.getSlides();
      var slideToMove = null;
      
      for (var j = 0; j < currentSlides.length; j++) {
        if (currentSlides[j].getObjectId() === op.slideId) {
          slideToMove = currentSlides[j];
          break;
        }
      }
      
      if (slideToMove) {
        slideToMove.move(op.newPosition);
        result.operationsExecuted++;
        result.slidesMoved++;
        
        if (i % 5 === 0 && i > 0) {
          Utilities.sleep(100);
        }
      }
      
    } catch (e) {
      result.errors.push("Error: " + e.toString());
    }
  }
  
  return result;
}

// Funciones auxiliares reutilizadas
function loadMappingFromPresentation() {
  try {
    var properties = PropertiesService.getDocumentProperties();
    var mappingJson = properties.getProperty('BOOGIE_THUMBNAIL_MAPPING');
    return mappingJson ? JSON.parse(mappingJson) : null;
  } catch (e) {
    return null;
  }
}

function slideHasLargeImage(slide, minWidth, minHeight) {
  try {
    var elements = slide.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
        var width = elements[i].getWidth();
        var height = elements[i].getHeight();
        if (width >= minWidth && height >= minHeight) {
          return true;
        }
      }
    }
  } catch (e) {}
  return false;
}