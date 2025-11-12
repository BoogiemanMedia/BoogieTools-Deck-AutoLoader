/******************************
 *  Sistema de Logging Visual
 ******************************/
var visualLogs = []; // Array para almacenar logs

function addVisualLog(message) {
  visualLogs.push(message);
  Logger.log(message); // También mantener logs tradicionales
}

function getVisualLogs() {
  return visualLogs.join('\n');
}

function clearVisualLogs() {
  visualLogs = [];
}

/******************************
 *  onOpen() y Menú Personalizado
 ******************************/
function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('BoogieTools')
    .addItem('Abrir Panel', 'showSidebar')
    .addToUi();
}

/******************************
 *  Mostrar Barra Lateral
 ******************************/
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar_v27')
    .setTitle('BoogieTools - Panel de Control');
  SlidesApp.getUi().showSidebar(html);
}

/******************************
 *  1) Script Original - BRANDED
 ******************************/
 
  
function importImagesInBatch(folderId) {
  clearVisualLogs();
  addVisualLog("=== INICIANDO PROCESO BRANDED ===");

  if (!folderId) {
    folderId = '1R41ETZVRrscNvwaoBGwER6CrQYYXFZED'; // ID por defecto
  }
  var presentationId = SlidesApp.getActivePresentation().getId();
  var mainFolder = DriveApp.getFolderById(folderId);
  Logger.log('Folder ID: ' + folderId);
  Logger.log('Folder name: ' + mainFolder.getName());

  // Contadores globales
  var totalImagesProcessed = 0;
  var totalImagesSkipped = 0;

  // Ordenar subcarpetas por número (o alfabéticamente)
  var subFoldersIterator = mainFolder.getFolders();
  var subFoldersArray = [];
  while (subFoldersIterator.hasNext()) {
    subFoldersArray.push(subFoldersIterator.next());
  }
  subFoldersArray.sort(function(a, b) {
    var aNum = parseInt(a.getName(), 10);
    var bNum = parseInt(b.getName(), 10);
    if (isNaN(aNum) || isNaN(bNum)) {
      return a.getName().localeCompare(b.getName());
    }
    return aNum - bNum;
  });

  // Layout IDs
  var layoutIdForUniversalBase       = 'g175f42bfc06_0_36';
  var layoutIdForUniversalBaseSquare = 'g1904e63b9c6_0_8';
  var layoutIdForEclipse             = 'g21c0d753fff_0_116';
  var layoutIdForOthers              = 'g1f79e657f07_0_122';

  // Palabras clave SafeZone
  var safeKeys = ['safezone','genfill','genai'];

  // Procesar subcarpetas
  for (var idx = 0; idx < subFoldersArray.length; idx++) {
    var subFolder = subFoldersArray[idx];
    addVisualLog("\n---  Procesando subcarpeta " + (idx+1) + "/" + subFoldersArray.length + ": " + subFolder.getName());
    var files = subFolder.getFiles();
    var requests = [];

    // Variables locales
    var universalBaseFile = null;
    var universalSafeFile = null;
    var isSquareBase = false;
    var isSquareSafe = false;

    // Buscar archivo base y archivo SafeZone
    var tempFiles = subFolder.getFiles();
    while (tempFiles.hasNext()) {
      var f = tempFiles.next();
      var name = f.getName().toLowerCase();
      if (name.endsWith('.psd')) continue;
      if (name.indexOf('universal') === -1) continue;

      var isSafe = safeKeys.some(function(k){ return name.indexOf(k) !== -1; });
      var isSquare = name.indexOf('square') !== -1;

      if (isSafe && !universalSafeFile) {
        universalSafeFile = f;
        isSquareSafe = isSquare;
      } else if (!isSafe && !universalBaseFile) {
        universalBaseFile = f;
        isSquareBase = isSquare;
      }
    }

    // Función auxiliar para insertar UBA
    function insertUBA(file, isSquare, label) {
      var layoutId = isSquare ? layoutIdForUniversalBaseSquare : layoutIdForUniversalBase;
      var slideId = Utilities.getUuid();
      var success = createSlideWithRetry(presentationId, slideId, layoutId);
      if (!success) {
        addVisualLog("✗ No se pudo crear slide para " + label);
        return;
      }
      var imageUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
      var cfg = isSquare
        ? { w: 451, h: 451, x: 58, y: -48 }
        : { w: 635, h: 635, x: 56, y: -135 };

      requests = requests.concat(
        createImageRequest(slideId, imageUrl, cfg.h, cfg.w, cfg.x, cfg.y)
      );
      totalImagesProcessed++;
      addVisualLog("✓ Slide " + label + " creado (" + file.getName() + ")");
    }

    // Insertar UBA base
    if (universalBaseFile) insertUBA(universalBaseFile, isSquareBase, "Universal Base");
    // Insertar UBA SafeZone
    if (universalSafeFile) insertUBA(universalSafeFile, isSquareSafe, "Universal SafeZone");

    // 2) Slide Eclipse - ACTUALIZADO A 4 IMÁGENES
    if (layoutIdForEclipse) {
      var slideIdE = Utilities.getUuid();
      if (createSlideWithRetry(presentationId, slideIdE, layoutIdForEclipse)) {
        var eclipseImages = {
          "RTL":            { x:370,  y:  13, width:297, height:166 },
          "Large":          { x: 56,  y:  13, width:297, height:166 },
          "Focused":        { x: 56,  y:196, width:297, height:166 },
          "Unfocused":      { x:370,  y:196, width:297, height:166 }
        };
        files = subFolder.getFiles();
        while (files.hasNext()) {
          var file2 = files.next();
          var fn2 = file2.getName();
          if (fn2.toLowerCase().endsWith('.psd')) continue;
          for (var key in eclipseImages) {
            if (fn2.indexOf(key) !== -1) {
              var e = eclipseImages[key];
              var u = 'https://drive.google.com/uc?export=view&id=' + file2.getId();
              requests = requests.concat(
                createImageRequest(slideIdE, u, e.height, e.width, e.x, e.y)
              );
              totalImagesProcessed++;
              break;
            }
          }
        }
      }
    }

    // 3) Slide Others - ACTUALIZADO A 5 IMÁGENES
    var slideIdO = Utilities.getUuid();
    if (createSlideWithRetry(presentationId, slideIdO, layoutIdForOthers)) {
      var imagesSetup = {
        "Boxshot":     { x: 56,  y: 13, width: 82,  height:113 },
        "Horizontal":           { x:150,  y: 13, width:206, height:118 },
        "Postplay":            { x: 56,  y:141, width:300, height:170 },
        "Email":                { x:370,  y:13, width: 122, height:298 },
        "Mobile":               { x:506,  y:13, width: 122, height:268 }
      };
      files = subFolder.getFiles();
      while (files.hasNext()) {
        var f3 = files.next();
        var nm = f3.getName();
        if (nm.toLowerCase().endsWith('.psd')) continue;
        for (var kw in imagesSetup) {
          if (nm.indexOf(kw) !== -1) {
            var s = imagesSetup[kw];
            var link = 'https://drive.google.com/uc?export=view&id=' + f3.getId();
            requests = requests.concat(
              createImageRequest(slideIdO, link, s.height, s.width, s.x, s.y)
            );
            totalImagesProcessed++;
            break;
          }
        }
      }
    }

    // Ejecutar batchUpdate
    if (requests.length) {
      try {
        Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
        addVisualLog("✓ Imágenes insertadas para subcarpeta: " + subFolder.getName());
      } catch(e) {
        addVisualLog("✗ Error insertando imágenes en " + subFolder.getName() + ": " + e);
      }
    }
  }

  addVisualLog("\n=== PROCESO COMPLETADO ===");
  addVisualLog("Total de imágenes procesadas: " + totalImagesProcessed);
  addVisualLog("Total de imágenes omitidas: " + totalImagesSkipped);
  return { success: true, log: getVisualLogs() };
}

/******************************
/******************************


/******************************
 *  2) Script "UnBranded"
 ******************************/
function importImagesInBatchUnbranded(folderId) {
  clearVisualLogs();
  addVisualLog("=== INICIANDO PROCESO UNBRANDED ===");
  
  try {
    if (!folderId) {
      folderId = '1R41ETZVRrscNvwaoBGwER6CrQYYXFZED';
      addVisualLog("Usando folder ID por defecto: " + folderId);
    } else {
      addVisualLog("Usando folder ID proporcionado: " + folderId);
    }
    
    var presentationId = SlidesApp.getActivePresentation().getId();
    addVisualLog("Presentation ID: " + presentationId);
    
    // Verificar acceso a la carpeta
    var mainFolder;
    try {
      mainFolder = DriveApp.getFolderById(folderId);
      addVisualLog("Carpeta principal: " + mainFolder.getName());
    } catch (folderError) {
      addVisualLog("✗ ERROR al acceder a la carpeta:");
      addVisualLog("  - ID proporcionado: " + folderId);
      addVisualLog("  - Posibles causas:");
      addVisualLog("    • El ID de carpeta es incorrecto");
      addVisualLog("    • No tienes permisos para acceder a esta carpeta");
      addVisualLog("    • La carpeta no existe o fue eliminada");
      addVisualLog("");
      addVisualLog("Por favor verifica:");
      addVisualLog("  1. Que el ID sea correcto");
      addVisualLog("  2. Que tengas acceso a la carpeta");
      addVisualLog("  3. Que la carpeta contenga las imágenes esperadas");
      return { success: false, log: getVisualLogs() };
    }
    
    // Obtener subcarpetas
    var subFoldersIterator = mainFolder.getFolders();
    var subFoldersArray = [];
    
    while (subFoldersIterator.hasNext()) {
      subFoldersArray.push(subFoldersIterator.next());
    }
    
    if (subFoldersArray.length === 0) {
      addVisualLog("✗ No se encontraron subcarpetas en: " + mainFolder.getName());
      addVisualLog("Verifica que la carpeta contenga subcarpetas con las imágenes");
      return { success: false, log: getVisualLogs() };
    }
    
    addVisualLog("Encontradas " + subFoldersArray.length + " subcarpetas");
    
    // Ordenar subcarpetas
    subFoldersArray.sort(function(a, b) {
      var aNum = parseInt(a.getName(), 10);
      var bNum = parseInt(b.getName(), 10);
      if (isNaN(aNum) || isNaN(bNum)) {
        return a.getName().localeCompare(b.getName());
      }
      return aNum - bNum;
    });
    
    var layoutIdForArtwork = 'g1e96a94ded4_0_0';
    var totalImagesProcessed = 0;
    
    // Procesar cada subcarpeta
    for (var idx = 0; idx < subFoldersArray.length; idx++) {
      var subFolder = subFoldersArray[idx];
      addVisualLog("\n--- Procesando subcarpeta " + (idx + 1) + "/" + subFoldersArray.length + ": " + subFolder.getName());
      
      try {
        var slideId = Utilities.getUuid();
        
        if (!createSlideWithRetry(presentationId, slideId, layoutIdForArtwork)) {
          addVisualLog("✗ No se pudo crear slide para: " + subFolder.getName());
          continue;
        }
        
        var imagesConfig = {
          "Vertical":   { x: 56,  y: 13,  width: 75,  height: 106 },
          "Horizontal": { x: 146, y: 13,  width: 188, height: 106 },
          "Story Art":  { x: 56,  y: 135, width: 188, height: 106 },
          "Email":      { x: 260, y: 255, width: 93,  height: 106 },
          "Mobile":     { x: 465, y: 255, width: 93,  height: 106 }
        };
        
        var requests = [];
        var filesIterator = subFolder.getFiles();
        var filesProcessedInFolder = 0;
        
        while (filesIterator.hasNext()) {
          var file = filesIterator.next();
          var fileName = file.getName();
          
          if (fileName.toLowerCase().endsWith('.psd')) {
            continue;
          }
          
          for (var keyword in imagesConfig) {
            if (fileName.indexOf(keyword) !== -1) {
              var config = imagesConfig[keyword];
              var imageUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
              
              requests = requests.concat(
                createImageRequest(slideId, imageUrl, config.height, config.width, config.x, config.y)
              );
              
              filesProcessedInFolder++;
              totalImagesProcessed++;
              addVisualLog("  ✓ " + keyword + ": " + fileName);
              break;
            }
          }
        }
        
        if (requests.length > 0) {
          try {
            Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
            addVisualLog("  → " + filesProcessedInFolder + " imágenes insertadas");
          } catch (batchError) {
            addVisualLog("  ✗ Error al insertar imágenes: " + batchError.toString());
          }
        } else {
          addVisualLog("  ! No se encontraron imágenes válidas");
        }
        
      } catch (folderError) {
        addVisualLog("✗ Error procesando subcarpeta: " + folderError.toString());
      }
    }
    
    addVisualLog("\n=== PROCESO COMPLETADO ===");
    addVisualLog("Total de subcarpetas procesadas: " + subFoldersArray.length);
    addVisualLog("Total de imágenes insertadas: " + totalImagesProcessed);
    
    return { success: true, log: getVisualLogs() };
    
  } catch (mainError) {
    addVisualLog("\n✗ ERROR GENERAL:");
    addVisualLog(mainError.toString());
    addVisualLog("\nStack trace:");
    addVisualLog(mainError.stack || "No disponible");
    return { success: false, log: getVisualLogs() };
  }
}

/******************************
 *  Funciones Auxiliares
 ******************************/
function createSlideWithRetry(presentationId, slideId, layoutId, maxRetries) {
  if (typeof maxRetries === 'undefined') maxRetries = 3;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      var requests = [{
        createSlide: {
          objectId: slideId,
          slideLayoutReference: {
            predefinedLayout: layoutId
          }
        }
      }];
      
      Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
      return true;
      
    } catch (e) {
      Logger.log("Intento " + attempt + " fallido: " + e.toString());
      if (attempt === maxRetries) {
        Logger.log("Error definitivo creando slide: " + e.toString());
        return false;
      }
      Utilities.sleep(1000);
    }
  }
  return false;
}

function createImageRequest(slideId, imageUrl, height, width, x, y) {
  var EMU_PER_POINT = 12700;
  
  return [{
    createImage: {
      objectId: Utilities.getUuid(),
      url: imageUrl,
      elementProperties: {
        pageObjectId: slideId,
        size: {
          height: { magnitude: height * EMU_PER_POINT, unit: 'EMU' },
          width:  { magnitude: width  * EMU_PER_POINT, unit: 'EMU' }
        },
        transform: {
          scaleX: 1,
          scaleY: 1,
          translateX: x * EMU_PER_POINT,
          translateY: y * EMU_PER_POINT,
          unit: 'EMU'
        }
      }
    }
  }];
}

/******************************
 *  3) Reordenar Slides (sin cambios)
 ******************************/
function reorderSlidesInPresentation() {
  clearVisualLogs();
  addVisualLog("=== INICIANDO REORDENAMIENTO DE SLIDES ===\n");
  
  try {
    var presentation = SlidesApp.getActivePresentation();
    var slides = presentation.getSlides();
    
    addVisualLog("Total de slides en presentación: " + slides.length);
    
    if (slides.length < 3) {
      addVisualLog("✗ Error: Se necesitan al menos 3 slides (índice + grupos)");
      return { success: false, log: getVisualLogs() };
    }
    
    // 1. Detectar slides de índice (grid 3x3)
    addVisualLog("\n1. Detectando slides de índice...");
    var indexSlides = detectIndexSlides(slides);
    
    if (indexSlides.length === 0) {
      addVisualLog("✗ No se encontraron slides de índice con grid 3x3");
      return { success: false, log: getVisualLogs() };
    }
    
    addVisualLog("✓ Encontrados " + indexSlides.length + " slides de índice:");
    for (var i = 0; i < indexSlides.length; i++) {
      addVisualLog("  - Slide " + (indexSlides[i].position + 1) + " (Índice #" + indexSlides[i].indexNum + ")");
    }
    
    // 2. Detectar grupos UBA
    addVisualLog("\n2. Detectando grupos de artwork (UBA + Eclipse + Others)...");
    var groups = detectAllGroups(slides, indexSlides);
    
    if (groups.length === 0) {
      addVisualLog("✗ No se encontraron grupos válidos");
      return { success: false, log: getVisualLogs() };
    }
    
    addVisualLog("✓ Encontrados " + groups.length + " grupos:");
    for (var g = 0; g < groups.length; g++) {
      var group = groups[g];
      addVisualLog("  - Grupo " + (g + 1) + ": Slides " + 
        (group.slides[0] + 1) + "-" + (group.slides[2] + 1));
    }
    
    // 3. Asociar grupos con índices
    addVisualLog("\n3. Asociando grupos con slides de índice...");
    var associations = associateGroupsWithIndexes(indexSlides, groups);
    
    for (var key in associations) {
      addVisualLog("  - Índice #" + key + ": " + associations[key].length + " grupos asociados");
    }
    
    // 4. Reordenar
    addVisualLog("\n4. Reordenando slides...");
    var movesMade = reorderSlidesByAssociation(presentation, indexSlides, associations);
    
    addVisualLog("\n✓ Reordenamiento completado");
    addVisualLog("Total de movimientos realizados: " + movesMade);
    
    addVisualLog("\n=== PROCESO FINALIZADO CON ÉXITO ===");
    return { success: true, log: getVisualLogs(), movesMade: movesMade };
    
  } catch (e) {
    addVisualLog("\n✗ ERROR CRÍTICO:");
    addVisualLog(e.toString());
    addVisualLog("\nStack trace:");
    addVisualLog(e.stack || "No disponible");
    return { success: false, log: getVisualLogs() };
  }
}

function detectIndexSlides(slides) {
  var indexSlides = [];
  
  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var elements = slide.getPageElements();
    var imagePositions = [];
    
    for (var j = 0; j < elements.length; j++) {
      if (elements[j].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
        imagePositions.push({
          left: elements[j].getLeft(),
          top: elements[j].getTop(),
          width: elements[j].getWidth(),
          height: elements[j].getHeight()
        });
      }
    }
    
    if (imagePositions.length === 9) {
      var gridInfo = analyzeGrid(imagePositions);
      
      if (gridInfo.isGrid && gridInfo.rows === 3 && gridInfo.cols === 3) {
        var indexNum = extractIndexNumber(slide);
        
        indexSlides.push({
          slide: slide,
          position: i,
          indexNum: indexNum,
          gridInfo: gridInfo
        });
        
        Logger.log("Slide índice detectado en posición " + i + " (Índice #" + indexNum + ")");
      }
    }
  }
  
  indexSlides.sort(function(a, b) {
    return a.indexNum - b.indexNum;
  });
  
  return indexSlides;
}

function analyzeGrid(positions) {
  if (positions.length !== 9) {
    return { isGrid: false };
  }
  
  var sortedByTop = positions.slice().sort(function(a, b) {
    return a.top - b.top;
  });
  
  var rows = [
    sortedByTop.slice(0, 3),
    sortedByTop.slice(3, 6),
    sortedByTop.slice(6, 9)
  ];
  
  for (var i = 0; i < rows.length; i++) {
    rows[i].sort(function(a, b) {
      return a.left - b.left;
    });
  }
  
  var avgRowTop = [0, 0, 0];
  for (var r = 0; r < 3; r++) {
    var sum = 0;
    for (var c = 0; c < 3; c++) {
      sum += rows[r][c].top;
    }
    avgRowTop[r] = sum / 3;
  }
  
  var rowSpacing1 = Math.abs(avgRowTop[1] - avgRowTop[0]);
  var rowSpacing2 = Math.abs(avgRowTop[2] - avgRowTop[1]);
  var rowSpacingDiff = Math.abs(rowSpacing1 - rowSpacing2);
  
  var avgColLeft = [0, 0, 0];
  for (var col = 0; col < 3; col++) {
    var sumLeft = 0;
    for (var row = 0; row < 3; row++) {
      sumLeft += rows[row][col].left;
    }
    avgColLeft[col] = sumLeft / 3;
  }
  
  var colSpacing1 = Math.abs(avgColLeft[1] - avgColLeft[0]);
  var colSpacing2 = Math.abs(avgColLeft[2] - avgColLeft[1]);
  var colSpacingDiff = Math.abs(colSpacing1 - colSpacing2);
  
  var SPACING_TOLERANCE = 20;
  var isGrid = (rowSpacingDiff < SPACING_TOLERANCE) && (colSpacingDiff < SPACING_TOLERANCE);
  
  return {
    isGrid: isGrid,
    rows: 3,
    cols: 3,
    rowSpacing: (rowSpacing1 + rowSpacing2) / 2,
    colSpacing: (colSpacing1 + colSpacing2) / 2
  };
}

function extractIndexNumber(slide) {
  var shapes = slide.getShapes();
  
  for (var i = 0; i < shapes.length; i++) {
    var shape = shapes[i];
    
    if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
      var textRange = shape.getText();
      var text = textRange.asString().trim();
      
      var match = text.match(/\b(\d+)\b/);
      if (match) {
        return parseInt(match[1]);
      }
    }
  }
  
  return 0;
}

function associateGroupsWithIndexes(indexSlides, groups) {
  var associations = {};
  
  for (var i = 0; i < indexSlides.length; i++) {
    associations[indexSlides[i].indexNum] = [];
  }
  
  for (var g = 0; g < groups.length; g++) {
    var group = groups[g];
    var ubaSlidePos = group.slides[0];
    
    var closestIndex = null;
    var closestDistance = Infinity;
    
    for (var i = 0; i < indexSlides.length; i++) {
      var indexPos = indexSlides[i].position;
      var distance = Math.abs(ubaSlidePos - indexPos);
      
      if (distance < closestDistance && ubaSlidePos > indexPos) {
        closestDistance = distance;
        closestIndex = indexSlides[i].indexNum;
      }
    }
    
    if (closestIndex !== null) {
      group.associatedIndex = closestIndex;
      associations[closestIndex].push(group);
    }
  }
  
  return associations;
}

function reorderSlidesByAssociation(presentation, indexSlides, associations) {
  var movesMade = 0;
  var slides = presentation.getSlides();
  
  for (var i = indexSlides.length - 1; i >= 0; i--) {
    var indexSlide = indexSlides[i];
    var indexNum = indexSlide.indexNum;
    var groupsForThisIndex = associations[indexNum];
    
    if (!groupsForThisIndex || groupsForThisIndex.length === 0) {
      continue;
    }
    
    groupsForThisIndex.sort(function(a, b) {
      return a.originalOrder - b.originalOrder;
    });
    
    var insertPosition = indexSlide.position + 1;
    
    for (var g = groupsForThisIndex.length - 1; g >= 0; g--) {
      var group = groupsForThisIndex[g];
      
      for (var s = 2; s >= 0; s--) {
        var currentPos = group.slides[s];
        
        if (currentPos !== insertPosition) {
          slides = presentation.getSlides();
          var slideToMove = slides[currentPos];
          slideToMove.move(insertPosition);
          movesMade++;
          
          Logger.log("Moviendo slide de posición " + currentPos + " a " + insertPosition);
        }
        
        insertPosition++;
      }
    }
  }
  
  return movesMade;
}

function getImagesFromSlide(slide) {
  var images = [];
  var elements = slide.getPageElements();
  
  for (var i = 0; i < elements.length; i++) {
    if (elements[i].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      images.push({
        element: elements[i],
        left: elements[i].getLeft(),
        top: elements[i].getTop(),
        width: elements[i].getWidth(),
        height: elements[i].getHeight()
      });
    }
  }
  
  return images;
}

function findImageWithoutDriveId(images) {
  var imagesWithoutDriveId = [];
  
  for (var i = 0; i < images.length; i++) {
    var img = images[i].element.asImage();
    
    try {
      var sourceUrl = img.getSourceUrl();
      
      if (!sourceUrl || sourceUrl.indexOf('/d/') === -1) {
        imagesWithoutDriveId.push(images[i]);
      }
    } catch (e) {
      imagesWithoutDriveId.push(images[i]);
    }
  }
  
  if (imagesWithoutDriveId.length === 0) {
    return null;
  }
  
  return imagesWithoutDriveId[0];
}

function applyAdditionalFiltering(images) {
  var expectedPositions = generateExpectedGrid(images);
  var filteredImages = [];
  
  for (var i = 0; i < expectedPositions.length; i++) {
    var expectedPos = expectedPositions[i];
    var closestImage = findClosestImage(images, expectedPos);
    
    if (closestImage && filteredImages.indexOf(closestImage) === -1) {
      filteredImages.push(closestImage);
    }
  }
  
  Logger.log("Filtro adicional: " + images.length + " → " + filteredImages.length + " imágenes");
  return filteredImages;
}

function generateExpectedGrid(images) {
  var expectedPositions = [];
  
  var indexGroups = {};
  for (var i = 0; i < images.length; i++) {
    var img = images[i];
    if (!indexGroups[img.indexNum]) {
      indexGroups[img.indexNum] = [];
    }
    indexGroups[img.indexNum].push(img);
  }
  
  for (var indexNum in indexGroups) {
    var imagesInIndex = indexGroups[indexNum];
    
    var minLeft = Math.min.apply(Math, imagesInIndex.map(function(img) { return img.left; }));
    var maxLeft = Math.max.apply(Math, imagesInIndex.map(function(img) { return img.left; }));
    var minTop = Math.min.apply(Math, imagesInIndex.map(function(img) { return img.top; }));
    var maxTop = Math.max.apply(Math, imagesInIndex.map(function(img) { return img.top; }));
    
    var horizontalSpacing = (maxLeft - minLeft) / 2;
    var verticalSpacing = (maxTop - minTop) / 2;
    
    for (var row = 0; row < 3; row++) {
      for (var col = 0; col < 3; col++) {
        expectedPositions.push({
          indexNum: parseInt(indexNum),
          left: minLeft + (col * horizontalSpacing),
          top: minTop + (row * verticalSpacing),
          tolerance: 50
        });
      }
    }
  }
  
  return expectedPositions;
}

function findClosestImage(images, expectedPos) {
  var closestImage = null;
  var minDistance = Infinity;
  
  for (var i = 0; i < images.length; i++) {
    var img = images[i];
    
    if (img.indexNum !== expectedPos.indexNum) continue;
    
    var distance = Math.sqrt(
      Math.pow(img.left - expectedPos.left, 2) + 
      Math.pow(img.top - expectedPos.top, 2)
    );
    
    if (distance <= expectedPos.tolerance && distance < minDistance) {
      minDistance = distance;
      closestImage = img;
    }
  }
  
  return closestImage;
}

function extractAllImageInfo(slide) {
  var images = [];
  
  try {
    var elements = slide.getPageElements();
    
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
        var element = elements[i];
        var image = element.asImage();
        
        var imageInfo = {
          position: i,
          left: element.getLeft(),
          top: element.getTop(),
          width: element.getWidth(),
          height: element.getHeight(),
          driveId: null
        };
        
        try {
          var sourceUrl = image.getSourceUrl();
          if (sourceUrl) {
            var match = sourceUrl.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
            if (match) {
              imageInfo.driveId = match[1];
            }
          }
        } catch (e) {
        }
        
        try {
          var title = image.getTitle();
          if (title) {
            imageInfo.title = title;
          }
        } catch (e) {}
        
        images.push(imageInfo);
      }
    }
    
  } catch (e) {
    Logger.log("Error extrayendo info de imágenes: " + e.toString());
  }
  
  return images;
}

function detectAllGroups(slides, indexSlides) {
  var allGroups = [];
  var firstGroupPos = indexSlides[indexSlides.length - 1].position + 1;
  
  for (var s = firstGroupPos; s < slides.length - 2; s++) {
    if (slideHasLargeImage(slides[s])) {
      var groupInfo = {
        startPos: s,
        slides: [s, s + 1, s + 2],
        ubaInfo: extractUBAInfo(slides[s]),
        originalOrder: allGroups.length + 1
      };
      allGroups.push(groupInfo);
      s += 2;
    }
  }
  
  return allGroups;
}

function extractUBAInfo(slide) {
  var info = {
    driveId: null,
    width: 0,
    height: 0
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
      } catch (e) {}
    }
    
  } catch (e) {
    Logger.log("Error extrayendo info UBA: " + e.toString());
  }
  
  return info;
}

function slideHasLargeImage(slide) {
  var MIN_WIDTH = 400;
  var MIN_HEIGHT = 200;

  try {
    var elements = slide.getPageElements();

    for (var i = 0; i < elements.length; i++) {
      if (elements[i].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
        var width = elements[i].getWidth();
        var height = elements[i].getHeight();

        if (width > MIN_WIDTH && height > MIN_HEIGHT) {
          return true;
        }
      }
    }
  } catch (e) {}

  return false;
}

/******************************
 *  WRAPPERS PARA OTROS MÓDULOS
 ******************************/

/**
 * Genera índice desde UBAs (módulo generate_index_from_ubas.js)
 */
function generateIndexFromUBAs() {
  if (typeof generateIndexFromUBAsModule === 'function') {
    return generateIndexFromUBAsModule();
  }
  return { success: false, log: "Error: Módulo de índice no encontrado" };
}

/**
 * Reordena slides desde índice linkeado (módulo reorder_universal.js)
 */
function reorderFromLinkedIndex() {
  if (typeof reorderFromLinkedIndexModule === 'function') {
    return reorderFromLinkedIndexModule();
  }
  return { success: false, log: "Error: Módulo de reordenamiento no encontrado" };
}
