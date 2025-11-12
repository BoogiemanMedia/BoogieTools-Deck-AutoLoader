/**
 * GenerateIndexFromUBAs_DebugEnhanced.js
 * 
 * Versión con debugging exhaustivo para identificar problemas en la inserción de imágenes
 */

function generateIndexFromUBAsModule() {
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
  
  log("=== GENERANDO ÍNDICE AUTOMÁTICO DESDE UBAs (DEBUG ENHANCED) ===");
  log("Fecha/Hora: " + new Date().toString());
  
  try {
    var presentation = SlidesApp.getActivePresentation();
    var slides = presentation.getSlides();
    log("✓ Presentación obtenida: " + slides.length + " slides");
    
    // 1. Detectar todos los grupos UBA existentes EN ORDEN CORRECTO
    log("\n1. DETECTANDO GRUPOS UBA EXISTENTES");
    var ubaGroups = detectAllUBAGroups(slides);
    log("Total de grupos UBA encontrados: " + ubaGroups.length);
    
    if (ubaGroups.length === 0) {
      return returnResult(false, "No se encontraron grupos UBA en la presentación");
    }
    
    // Log detallado de grupos encontrados
    for (var g = 0; g < ubaGroups.length; g++) {
      var group = ubaGroups[g];
      log("  Grupo " + group.sequentialOrder + " detectado en slides " + 
          (group.startPos + 1) + "-" + (group.startPos + 3));
    }
    
    // 2. Extraer thumbnails de cada UBA EN ORDEN SECUENCIAL
    log("\n2. EXTRAYENDO THUMBNAILS DE UBAs EN ORDEN SECUENCIAL");
    var thumbnails = [];
    for (var i = 0; i < ubaGroups.length; i++) {
      var group = ubaGroups[i];
      var thumbnail = extractUBAThumbnailDebug(group, slides[group.startPos], log);
      if (thumbnail) {
        thumbnail.groupIndex = i;
        thumbnail.originalOrder = i + 1;
        thumbnail.slidePosition = group.startPos;
        thumbnails.push(thumbnail);
        log("  ✓ Thumbnail " + (i + 1) + " (slide " + (group.startPos + 1) + "): " + 
            thumbnail.width.toFixed(0) + "x" + thumbnail.height.toFixed(0) + 
            (thumbnail.driveId ? " (Drive ID: " + thumbnail.driveId.substring(0, 8) + "...)" : " (sin Drive ID)"));
      } else {
        log("  ✗ No se pudo extraer thumbnail del grupo " + (i + 1) + " (slide " + (group.startPos + 1) + ")");
      }
    }
    
    log("Total de thumbnails extraídos: " + thumbnails.length);
    
    // 3. Calcular layout del índice
    log("\n3. CALCULANDO LAYOUT DEL ÍNDICE");
    var indexLayout = calculateIndexLayout(thumbnails.length);
    log("Se crearán " + indexLayout.totalSlides + " slides de índice");
    log("Layout: " + indexLayout.imagesPerSlide + " imágenes por slide (" + 
        indexLayout.cols + "x" + indexLayout.rows + ")");
    
    // 4. Crear slides de índice con debugging enhanced
    log("\n4. CREANDO SLIDES DE ÍNDICE (DEBUG ENHANCED)");
    var indexSlides = createIndexSlidesDebugEnhanced(presentation, indexLayout, thumbnails, log);
    
    // 5. Crear mapeo para referencia futura
    log("\n5. CREANDO MAPEO DE REFERENCIA");
    var mappingData = createThumbnailMapping(thumbnails, indexSlides, ubaGroups);
    
    // 6. Guardar mapeo en propiedades de la presentación
    log("\n6. GUARDANDO MAPEO EN PRESENTACIÓN");
    saveMappingToPresentation(mappingData);
    
    // Resumen final
    var endTime = new Date().getTime();
    var totalTime = endTime - startTime;
    
    log("\n=== ÍNDICE GENERADO EXITOSAMENTE ===");
    log("Tiempo total: " + totalTime + "ms");
    log("Slides de índice creados: " + indexLayout.totalSlides);
    log("Thumbnails insertados: " + thumbnails.length);
    log("Grupos UBA mapeados: " + ubaGroups.length);
    
    return returnResult(true, "Índice automático generado correctamente");
    
  } catch (error) {
    log("\n✗ ERROR GENERAL: " + error.toString());
    log("  Stack: " + error.stack);
    return returnResult(false, "Error: " + error.toString());
  }
}

/**
 * Versión debug enhanced para extraer thumbnail con más logging
 */
function extractUBAThumbnailDebug(group, ubaSlide, logFunction) {
  try {
    logFunction("    → Extrayendo thumbnail del grupo " + group.sequentialOrder);
    var elements = ubaSlide.getPageElements();
    var largestImage = null;
    var largestArea = 0;
    
    logFunction("    → Elementos en slide: " + elements.length);
    
    // Encontrar la imagen más grande (UBA)
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
        var element = elements[i];
        var area = element.getWidth() * element.getHeight();
        
        logFunction("    → Imagen " + (i + 1) + ": " + element.getWidth().toFixed(0) + "x" + element.getHeight().toFixed(0) + " (área: " + area.toFixed(0) + ")");
        
        if (area > largestArea) {
          largestArea = area;
          largestImage = element;
        }
      }
    }
    
    if (!largestImage) {
      logFunction("    ✗ No se encontró imagen UBA en el slide");
      return null;
    }
    
    logFunction("    ✓ Imagen UBA más grande encontrada: " + largestImage.getWidth().toFixed(0) + "x" + largestImage.getHeight().toFixed(0));
    
    var image = largestImage.asImage();
    var thumbnail = {
      width: largestImage.getWidth(),
      height: largestImage.getHeight(),
      sourceUrl: null,
      driveId: null,
      blob: null,
      fileSize: null
    };
    
    // MÉTODO 1: Extraer Source URL
    try {
      thumbnail.sourceUrl = image.getSourceUrl();
      logFunction("    → Source URL: " + (thumbnail.sourceUrl ? "✓ extraída" : "✗ no disponible"));
      
      // Extraer Drive ID si existe
      if (thumbnail.sourceUrl) {
        var match = thumbnail.sourceUrl.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
        if (match) {
          thumbnail.driveId = match[1];
          logFunction("    → Drive ID extraído: " + thumbnail.driveId);
          
          // Verificar tamaño de archivo
          try {
            var file = DriveApp.getFileById(thumbnail.driveId);
            thumbnail.fileSize = file.getSize();
            logFunction("    → Tamaño de archivo: " + (thumbnail.fileSize / 1024 / 1024).toFixed(2) + " MB");
            
            // Verificar permisos
            var sharing = file.getSharingAccess();
            logFunction("    → Permisos de archivo: " + sharing);
            
          } catch (e) {
            logFunction("    ⚠️ No se pudo verificar archivo: " + e.toString());
          }
        }
      }
      
    } catch (e) {
      logFunction("    ✗ Error extrayendo Source URL: " + e.toString());
    }
    
    // MÉTODO 2: Intentar obtener blob (mejorado)
    try {
      thumbnail.blob = image.getBlob();
      if (thumbnail.blob) {
        try {
          var blobSize = thumbnail.blob.getSize();
          logFunction("    → Blob extraído: " + (blobSize / 1024).toFixed(1) + " KB");
          if (!thumbnail.fileSize) {
            thumbnail.fileSize = blobSize;
          }
        } catch (blobSizeError) {
          logFunction("    → Blob extraído (tamaño no disponible): " + blobSizeError.toString());
          // El blob existe pero no podemos obtener su tamaño, aún es válido
        }
      }
    } catch (e) {
      logFunction("    ✗ Error extrayendo blob: " + e.toString());
    }
    
    // Validar que tenemos al menos UN método para obtener la imagen
    if (!thumbnail.driveId && !thumbnail.sourceUrl && !thumbnail.blob) {
      logFunction("    ✗ No se pudo extraer ningún método de acceso a la imagen");
      return null;
    }
    
    logFunction("    ✓ Thumbnail extraído exitosamente");
    return thumbnail;
    
  } catch (e) {
    logFunction("    ✗ Error general extrayendo thumbnail: " + e.toString());
    return null;
  }
}

/**
 * Versión debug enhanced para crear slides de índice
 */
function createIndexSlidesDebugEnhanced(presentation, layout, thumbnails, logFunction) {
  var presentationId = presentation.getId();
  var LAYOUT_ID = "g175f42bfc06_0_153"; // UBA Rectangle Grid: 9
  var indexSlides = [];
  
  logFunction("🚀 INICIANDO INSERCIÓN DE IMÁGENES CON DEBUG ENHANCED");
  logFunction("📊 Estadísticas iniciales:");
  logFunction("  - Total thumbnails: " + thumbnails.length);
  logFunction("  - Slides a crear: " + layout.totalSlides);
  logFunction("  - Layout ID: " + LAYOUT_ID);
  
  // Posiciones optimizadas para el layout
  var positions = [
    {x: 58,  y: 15},  {x: 274, y: 15},  {x: 490, y: 15},     // Fila 1
    {x: 58,  y: 135}, {x: 274, y: 135}, {x: 490, y: 135},    // Fila 2  
    {x: 58,  y: 256}, {x: 274, y: 256}, {x: 490, y: 256}     // Fila 3
  ];
  
  // Procesar thumbnails EN ORDEN SECUENCIAL
  var orderedThumbnails = thumbnails.slice();
  orderedThumbnails.sort(function(a, b) {
    return a.originalOrder - b.originalOrder;
  });
  
  // Filtrar solo thumbnails válidos
  var validThumbnails = orderedThumbnails.filter(function(thumb) {
    return thumb.driveId || thumb.sourceUrl || thumb.blob;
  });
  
  logFunction("📋 Thumbnails válidos para inserción: " + validThumbnails.length + "/" + orderedThumbnails.length);
  
  // Estadísticas de métodos disponibles
  var driveIdCount = 0, sourceUrlCount = 0, blobCount = 0;
  for (var i = 0; i < validThumbnails.length; i++) {
    if (validThumbnails[i].driveId) driveIdCount++;
    if (validThumbnails[i].sourceUrl) sourceUrlCount++;
    if (validThumbnails[i].blob) blobCount++;
  }
  logFunction("📈 Métodos disponibles: Drive ID: " + driveIdCount + ", Source URL: " + sourceUrlCount + ", Blob: " + blobCount);
  
  var globalThumbnailIndex = 0;
  var totalInsertedSuccessfully = 0;
  var totalInsertionFailures = 0;
  
  for (var slideNum = 0; slideNum < layout.totalSlides; slideNum++) {
    logFunction("\n🎯 === CREANDO SLIDE DE ÍNDICE " + (slideNum + 1) + "/" + layout.totalSlides + " ===");
    
    // Crear slide
    var slideId = Utilities.getUuid();
    logFunction("🆔 UUID generado: " + slideId);
    
    var success = createSlideWithRetryDebug(presentationId, slideId, LAYOUT_ID, slideNum, logFunction);
    
    if (!success) {
      logFunction("❌ Error creando slide de índice " + (slideNum + 1));
      continue;
    }
    
    logFunction("✅ Slide " + (slideNum + 1) + " creado exitosamente: " + slideId);
    
    var indexSlide = {
      slideId: slideId,
      slideNumber: slideNum + 1,
      thumbnails: []
    };
    
    var thumbnailsInThisSlide = 0;
    
    logFunction("🖼️ Comenzando inserción de thumbnails en slide " + (slideNum + 1));
    logFunction("📍 Progreso global: desde thumbnail " + (globalThumbnailIndex + 1) + "/" + validThumbnails.length);
    
    while (globalThumbnailIndex < validThumbnails.length && thumbnailsInThisSlide < layout.imagesPerSlide) {
      var thumbnail = validThumbnails[globalThumbnailIndex];
      var position = positions[thumbnailsInThisSlide];
      
      logFunction("\n🔄 === PROCESANDO THUMBNAIL " + (globalThumbnailIndex + 1) + "/" + validThumbnails.length + " ===");
      logFunction("📍 Thumbnail orden original: " + thumbnail.originalOrder);
      logFunction("📍 Posición en slide: " + thumbnailsInThisSlide + " -> [" + position.x + "," + position.y + "]");
      logFunction("📏 Tamaño original: " + thumbnail.width.toFixed(0) + "x" + thumbnail.height.toFixed(0));
      if (thumbnail.fileSize) {
        logFunction("📦 Tamaño archivo: " + (thumbnail.fileSize / 1024 / 1024).toFixed(2) + " MB");
      }
      
      // Verificar límites de tamaño antes de insertar
      if (thumbnail.fileSize && thumbnail.fileSize > 25 * 1024 * 1024) { // 25MB
        logFunction("⚠️ Archivo muy grande (" + (thumbnail.fileSize / 1024 / 1024).toFixed(2) + " MB), omitiendo...");
        globalThumbnailIndex++;
        totalInsertionFailures++;
        continue;
      }
      
      // Determinar URL de imagen con múltiples métodos
      var imageUrlResult = determineImageUrl(thumbnail, logFunction);
      
      if (!imageUrlResult.success) {
        logFunction("❌ No se pudo determinar URL válida para thumbnail " + (globalThumbnailIndex + 1));
        globalThumbnailIndex++;
        totalInsertionFailures++;
        continue;
      }
      
      logFunction("🌐 URL determinada: " + imageUrlResult.method);
      logFunction("🔗 URL (primeros 60 chars): " + imageUrlResult.url.substring(0, 60) + "...");
      
      // Crear e insertar imagen con debugging enhanced
      var imageObjectId = Utilities.getUuid();
      var insertResult = insertSingleImageDebugEnhanced(
        presentationId, 
        slideId, 
        imageObjectId, 
        imageUrlResult.url, 
        position, 
        layout, 
        logFunction,
        globalThumbnailIndex + 1,
        validThumbnails.length
      );
      
      // Si falla, intentar método de fallback con blob
      if (!insertResult.success && thumbnail.blob) {
        logFunction("🔄 INTENTANDO FALLBACK CON BLOB...");
        
        try {
          var fallbackUrl = createTemporaryFileFromBlob(thumbnail.blob, logFunction);
          if (fallbackUrl) {
            logFunction("🔄 Reintentando con archivo temporal de fallback...");
            
            // Nuevo object ID para el reintento
            var fallbackObjectId = Utilities.getUuid();
            var fallbackResult = insertSingleImageDebugEnhanced(
              presentationId, 
              slideId, 
              fallbackObjectId, 
              fallbackUrl, 
              position, 
              layout, 
              logFunction,
              globalThumbnailIndex + 1,
              validThumbnails.length
            );
            
            if (fallbackResult.success) {
              insertResult = fallbackResult;
              imageObjectId = fallbackObjectId;
              logFunction("🎉 FALLBACK EXITOSO: Imagen insertada usando blob temporal");
            }
          }
        } catch (fallbackError) {
          logFunction("❌ Error en fallback: " + fallbackError.toString());
        }
      }
      
      if (insertResult.success) {
        // Guardar información del thumbnail
        indexSlide.thumbnails.push({
          imageObjectId: imageObjectId,
          position: thumbnailsInThisSlide,
          groupIndex: thumbnail.groupIndex,
          originalOrder: thumbnail.originalOrder,
          slidePosition: thumbnail.slidePosition,
          driveId: thumbnail.driveId,
          urlMethod: imageUrlResult.method
        });
        
        // NUEVO: Agregar título a la imagen para mantener referencia
        try {
          // Esperar un momento para que la imagen se cree completamente
          Utilities.sleep(100);
          
          // Buscar la imagen recién creada y agregarle el título
          var currentSlide = getSlideById(slideId);
          if (currentSlide) {
            var elements = currentSlide.getPageElements();
            for (var e = 0; e < elements.length; e++) {
              if (elements[e].getObjectId() === imageObjectId) {
                var imageElement = elements[e].asImage();
                var groupTitle = "BOOGIE_GROUP_" + (thumbnail.originalOrder);
                imageElement.setTitle(groupTitle);
                logFunction("🏷️ Título asignado: " + groupTitle);
                break;
              }
            }
          }
        } catch (titleError) {
          logFunction("⚠️ No se pudo asignar título: " + titleError.toString());
          // No es crítico, continuamos
        }
        
        logFunction("🎉 ÉXITO: Thumbnail " + (globalThumbnailIndex + 1) + " insertado en posición " + thumbnailsInThisSlide + " (" + imageUrlResult.method + ")");
        thumbnailsInThisSlide++;
        totalInsertedSuccessfully++;
      } else {
        logFunction("💥 FALLO: Thumbnail " + (globalThumbnailIndex + 1) + " falló después de todos los intentos");
        totalInsertionFailures++;
      }
            
      globalThumbnailIndex++;
      
      // Pausa entre inserciones
      Utilities.sleep(200);
      
      // Log de progreso cada 5 imágenes
      if ((globalThumbnailIndex) % 5 === 0) {
        logFunction("📊 PROGRESO INTERMEDIO:");
        logFunction("  - Procesados: " + globalThumbnailIndex + "/" + validThumbnails.length);
        logFunction("  - Exitosos: " + totalInsertedSuccessfully);
        logFunction("  - Fallidos: " + totalInsertionFailures);
        logFunction("  - Tasa éxito: " + (totalInsertedSuccessfully / (totalInsertedSuccessfully + totalInsertionFailures) * 100).toFixed(1) + "%");
      }
    }
    
    logFunction("✅ Slide " + (slideNum + 1) + " completado con " + thumbnailsInThisSlide + " thumbnails exitosos");
    logFunction("📊 Progreso global: " + globalThumbnailIndex + "/" + validThumbnails.length + " thumbnails procesados");
    indexSlides.push(indexSlide);
    
    if (globalThumbnailIndex >= validThumbnails.length) {
      logFunction("🏁 Todos los thumbnails han sido procesados");
      break;
    }
  }
  
  logFunction("\n🎊 === RESUMEN FINAL DE INSERCIÓN ===");
  logFunction("📊 Slides creados: " + indexSlides.length);
  logFunction("📊 Total thumbnails procesados: " + globalThumbnailIndex + "/" + validThumbnails.length);
  logFunction("📊 Inserciones exitosas: " + totalInsertedSuccessfully);
  logFunction("📊 Inserciones fallidas: " + totalInsertionFailures);
  logFunction("📊 Tasa de éxito final: " + (totalInsertedSuccessfully / (totalInsertedSuccessfully + totalInsertionFailures) * 100).toFixed(1) + "%");
  
  return indexSlides;
}

/**
 * Determina la mejor URL para usar con debugging mejorado
 */
function determineImageUrl(thumbnail, logFunction) {
  logFunction("🔍 Determinando URL óptima...");
  
  // Método 1: Drive ID directo (solo si es accesible)
  if (thumbnail.driveId) {
    logFunction("  → Intentando método Drive ID directo");
    try {
      // Verificar que el archivo existe y es accesible
      var file = DriveApp.getFileById(thumbnail.driveId);
      var url = 'https://drive.google.com/uc?export=view&id=' + thumbnail.driveId;
      
      logFunction("  ✅ Drive ID válido, archivo encontrado: " + file.getName());
      return {
        success: true,
        url: url,
        method: "Drive ID directo"
      };
    } catch (e) {
      logFunction("  ❌ Error con Drive ID: " + e.toString());
      // Continuar con otros métodos
    }
  }
  
  // Método 2: Source URL (filtrar URLs problemáticas)
  if (thumbnail.sourceUrl) {
    logFunction("  → Analizando Source URL");
    
    // Detectar URLs problemáticas de lh3.google.com
    if (thumbnail.sourceUrl.includes("lh3.google.com")) {
      logFunction("  ⚠️ URL lh3.google.com detectada - convirtiendo a formato público");
      
      // Si tenemos Drive ID, intentar crear URL pública
      if (thumbnail.driveId) {
        var publicUrl = 'https://drive.google.com/uc?export=view&id=' + thumbnail.driveId;
        logFunction("  → Usando Drive ID para URL pública: " + thumbnail.driveId);
        return {
          success: true,
          url: publicUrl,
          method: "Drive ID (convertido desde lh3)"
        };
      } else {
        logFunction("  ❌ URL lh3.google.com sin Drive ID disponible");
        // Continuar con método blob
      }
    } else {
      // URL normal, usar directamente
      logFunction("  ✅ Source URL válida");
      return {
        success: true,
        url: thumbnail.sourceUrl,
        method: "Source URL"
      };
    }
  }
  
  // Método 3: Crear archivo temporal desde blob (mejorado)
  if (thumbnail.blob) {
    logFunction("  → Intentando método archivo temporal desde blob");
    try {
      // Verificar que el blob es válido
      var blobSize;
      try {
        blobSize = thumbnail.blob.getSize();
        logFunction("  → Blob válido detectado: " + (blobSize / 1024).toFixed(1) + " KB");
      } catch (blobError) {
        logFunction("  ⚠️ Error obteniendo tamaño de blob: " + blobError.toString());
        // Continuar de todos modos
        blobSize = 0;
      }
      
      var tempFileName = "temp_thumbnail_" + new Date().getTime() + ".png";
      var tempFile = DriveApp.createFile(thumbnail.blob.setName(tempFileName));
      
      // Hacer el archivo público
      tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      var url = 'https://drive.google.com/uc?export=view&id=' + tempFile.getId();
      logFunction("  ✅ Archivo temporal creado: " + tempFile.getId());
      
      return {
        success: true,
        url: url,
        method: "Archivo temporal",
        tempFileId: tempFile.getId()
      };
    } catch (e) {
      logFunction("  ❌ Error creando archivo temporal: " + e.toString());
    }
  }
  
  // Método 4: Último recurso - convertir Drive ID si lo tenemos
  if (thumbnail.driveId) {
    logFunction("  → Último recurso: URL pública desde Drive ID");
    try {
      var fallbackUrl = 'https://drive.google.com/thumbnail?id=' + thumbnail.driveId + '&sz=w800';
      logFunction("  → Intentando URL de thumbnail: " + fallbackUrl);
      
      return {
        success: true,
        url: fallbackUrl,
        method: "Drive thumbnail (último recurso)"
      };
    } catch (e) {
      logFunction("  ❌ Error en último recurso: " + e.toString());
    }
  }
  
  logFunction("  ❌ No se pudo determinar URL válida con ningún método");
  return { success: false };
}

/**
 * Inserta una sola imagen con debugging exhaustivo
 */
function insertSingleImageDebugEnhanced(presentationId, slideId, imageObjectId, imageUrl, position, layout, logFunction, currentIndex, totalImages) {
  logFunction("🔧 === INICIANDO INSERCIÓN DE IMAGEN ===");
  logFunction("🆔 Object ID: " + imageObjectId);
  logFunction("🔗 URL: " + imageUrl.substring(0, 80) + "...");
  logFunction("📍 Posición: [" + position.x + ", " + position.y + "]");
  logFunction("📏 Tamaño objetivo: " + layout.thumbnailWidth + "x" + layout.thumbnailHeight + " pt");
  logFunction("📊 Progreso: " + currentIndex + "/" + totalImages);
  
  var maxAttempts = 3;
  var currentUrl = imageUrl;
  
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    logFunction("\n🎯 Intento " + attempt + "/" + maxAttempts);
    logFunction("🔗 URL actual: " + currentUrl.substring(0, 60) + "...");
    
    try {
      var request = {
        createImage: {
          objectId: imageObjectId,
          url: currentUrl,
          elementProperties: {
            pageObjectId: slideId,
            size: { 
              height: { magnitude: layout.thumbnailHeight, unit: 'PT' }, 
              width: { magnitude: layout.thumbnailWidth, unit: 'PT' } 
            },
            transform: { 
              scaleX: 1, 
              scaleY: 1, 
              translateX: position.x, 
              translateY: position.y, 
              unit: 'PT' 
            }
          }
        }
      };
      
      logFunction("🚀 Ejecutando batchUpdate (intento " + attempt + ")...");
      var beforeTime = new Date().getTime();
      
      Slides.Presentations.batchUpdate({ 
        requests: [request] 
      }, presentationId);
      
      var afterTime = new Date().getTime();
      var duration = afterTime - beforeTime;
      
      logFunction("🎉 ¡ÉXITO! Imagen insertada en " + duration + "ms (intento " + attempt + ")");
      return { success: true, attempt: attempt, duration: duration };
      
    } catch (e) {
      var errorMsg = e.toString();
      logFunction("❌ Error en intento " + attempt + ": " + errorMsg);
      
      if (e.details) {
        logFunction("📋 Detalles del error: " + JSON.stringify(e.details));
      }
      
      // Analizar tipo de error y sugerir solución
      if (errorMsg.includes("There was a problem retrieving the image")) {
        logFunction("🔍 Diagnóstico: Problema accediendo a la imagen");
        
        if (errorMsg.includes("publicly accessible")) {
          logFunction("🔐 Causa probable: Permisos de acceso");
        } else if (errorMsg.includes("size limit")) {
          logFunction("📏 Causa probable: Imagen muy grande");
        } else if (errorMsg.includes("supported formats")) {
          logFunction("📄 Causa probable: Formato no soportado");
        }
        
        // Intentar URL alternativa en el siguiente intento
        if (attempt < maxAttempts && currentUrl.includes("uc?export=view")) {
          currentUrl = currentUrl.replace("uc?export=view", "thumbnail?sz=w800");
          logFunction("🔄 Probando URL alternativa en siguiente intento: thumbnail");
        }
        
      } else if (errorMsg.includes("Rate Limit Exceeded") || errorMsg.includes("quota")) {
        logFunction("⏰ Diagnóstico: Límite de velocidad excedido");
        if (attempt < maxAttempts) {
          logFunction("⏳ Esperando 2 segundos antes del siguiente intento...");
          Utilities.sleep(2000);
        }
      }
      
      if (attempt === maxAttempts) {
        logFunction("💀 FALLO DEFINITIVO después de " + maxAttempts + " intentos");
        return { 
          success: false, 
          lastError: errorMsg,
          finalUrl: currentUrl,
          totalAttempts: maxAttempts
        };
      } else {
        logFunction("🔄 Preparando siguiente intento...");
        Utilities.sleep(500);
      }
    }
  }
  
  return { success: false };
}

/**
 * Crea un archivo temporal desde blob como método de fallback
 */
function createTemporaryFileFromBlob(blob, logFunction) {
  try {
    logFunction("  → Creando archivo temporal desde blob...");
    
    var tempFileName = "fallback_temp_" + new Date().getTime() + ".png";
    var tempFile = DriveApp.createFile(blob.setName(tempFileName));
    
    // Hacer el archivo público
    tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var url = 'https://drive.google.com/uc?export=view&id=' + tempFile.getId();
    logFunction("  ✅ Archivo temporal de fallback creado: " + tempFile.getId());
    
    return url;
    
  } catch (e) {
    logFunction("  ❌ Error creando archivo temporal de fallback: " + e.toString());
    return null;
  }
}

/**
 * Crea slide con debugging mejorado
 */
function createSlideWithRetryDebug(presentationId, slideId, layoutId, slideIndex, logFunction, maxRetries) {
  maxRetries = maxRetries || 3;
  
  logFunction("🏗️ Creando slide con ID: " + slideId);
  logFunction("🎨 Layout ID: " + layoutId);
  logFunction("📍 Posición: " + slideIndex);
  
  for (var i = 0; i < maxRetries; i++) {
    try {
      logFunction("🎯 Intento " + (i + 1) + "/" + maxRetries + " de crear slide");
      
      var request = {
        createSlide: {
          objectId: slideId,
          slideLayoutReference: {
            layoutId: layoutId
          },
          insertionIndex: slideIndex
        }
      };
      
      var beforeTime = new Date().getTime();
      
      Slides.Presentations.batchUpdate({
        requests: [request]
      }, presentationId);
      
      var afterTime = new Date().getTime();
      var duration = afterTime - beforeTime;
      
      logFunction("✅ Slide creado exitosamente en " + duration + "ms (intento " + (i + 1) + ")");
      return true;
      
    } catch (e) {
      logFunction("❌ Intento " + (i + 1) + " falló: " + e.toString());
      if (i < maxRetries - 1) {
        logFunction("⏳ Esperando 500ms antes del siguiente intento...");
        Utilities.sleep(500);
      }
    }
  }
  
  logFunction("💀 FALLO: No se pudo crear slide después de " + maxRetries + " intentos");
  return false;
}

// Las demás funciones auxiliares permanecen igual...
function detectAllUBAGroups(slides) {
  var groups = [];
  var MIN_UBA_WIDTH = 400;
  var MIN_UBA_HEIGHT = 200;
  
  for (var i = 0; i < slides.length - 2; i++) {
    var slide = slides[i];
    
    if (slideHasUBAImage(slide, MIN_UBA_WIDTH, MIN_UBA_HEIGHT)) {
      var group = {
        startPos: i,
        slides: [i, i + 1, i + 2],
        ubaSlide: slide,
        eclipseSlide: i + 1 < slides.length ? slides[i + 1] : null,
        othersSlide: i + 2 < slides.length ? slides[i + 2] : null,
        sequentialOrder: groups.length + 1
      };
      
      groups.push(group);
      i += 2;
    }
  }
  
  return groups;
}

function slideHasUBAImage(slide, minWidth, minHeight) {
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
  } catch (e) {
    Logger.log("Error verificando imagen UBA: " + e.toString());
  }
  
  return false;
}

function calculateIndexLayout(totalThumbnails) {
  var IMAGES_PER_SLIDE = 9;
  var COLS = 3;
  var ROWS = 3;
  
  var totalSlides = Math.ceil(totalThumbnails / IMAGES_PER_SLIDE);
  
  return {
    totalSlides: totalSlides,
    imagesPerSlide: IMAGES_PER_SLIDE,
    cols: COLS,
    rows: ROWS,
    thumbnailWidth: 202,
    thumbnailHeight: 106
  };
}

function createThumbnailMapping(thumbnails, indexSlides, ubaGroups) {
  var mapping = {
    version: "1.0",
    created: new Date().toISOString(),
    totalThumbnails: thumbnails.length,
    totalGroups: ubaGroups.length,
    slides: []
  };
  
  for (var i = 0; i < indexSlides.length; i++) {
    var slide = indexSlides[i];
    var slideMapping = {
      slideId: slide.slideId,
      slideNumber: slide.slideNumber,
      thumbnails: []
    };
    
    for (var j = 0; j < slide.thumbnails.length; j++) {
      var thumb = slide.thumbnails[j];
      var group = ubaGroups[thumb.groupIndex];
      
      slideMapping.thumbnails.push({
        imageObjectId: thumb.imageObjectId,
        position: thumb.position,
        groupIndex: thumb.groupIndex,
        groupSlides: group.slides,
        originalOrder: thumb.originalOrder,
        slidePosition: thumb.slidePosition,
        driveId: thumb.driveId
      });
    }
    
    mapping.slides.push(slideMapping);
  }
  
  return mapping;
}

function saveMappingToPresentation(mappingData) {
  try {
    var properties = PropertiesService.getDocumentProperties();
    properties.setProperty('BOOGIE_THUMBNAIL_MAPPING', JSON.stringify(mappingData));
    Logger.log("✓ Mapeo guardado en propiedades de la presentación");
  } catch (e) {
    Logger.log("✗ Error guardando mapeo: " + e.toString());
  }
}