/**
 * SafezonesOverlayModule.js
 * 
 * Este módulo implementa la función toggleSafezonesOverlay() que:
 *   - Obtiene la diapositiva activa.
 *   - Determina el layout usando la Advanced Slides API (se obtiene el layoutObjectId).
 *   - Consulta en la configuración SAFEZONE_OVERLAY_CONFIG los parámetros para ese layout:
 *       • fileName: nombre del archivo PNG del overlay.
 *       • x, y, width, height: posición y tamaño (en pulgadas) de referencia.
 *   - Busca en una carpeta de Drive (OVERLAY_FOLDER_ID) el archivo cuyo nombre coincida.
 *   - Si ya existe en la diapositiva una imagen con el título "safezone_overlay", se elimina (toggle off).
 *   - De lo contrario, carga el overlay y lo inserta en el slide:
 *       • Para UBA Rectangle se usa el posicionamiento original.
 *       • Para UBA Square se usan los valores configurados (se aplican directamente).
 *       • Para los demás layouts, se escala el overlay para que ocupe el 100% del slide completo (ancho y alto).
 *
 * NOTA:
 *   - Personalizá OVERLAY_FOLDER_ID con el ID de la carpeta en Drive donde subas tus PNGs.
 *   - Ajustá los IDs de layout y los nombres de archivo en SAFEZONE_OVERLAY_CONFIG según tu proyecto.
 * 
 * Requiere que la Advanced Slides API esté habilitada.
 */

var OVERLAY_FOLDER_ID = "1KMk9sU3OvPDdmKrWJNLsC_Jtm6vpcbN8"; // ID de la carpeta de overlays

// Configuración: clave = layout ID (layoutObjectId) obtenido de la Advanced Slides API.
var SAFEZONE_OVERLAY_CONFIG = {
  // Para Branded:
  "g175f42bfc06_0_36": {   // UBA Rectangle
    fileName: "SAFEZONE_UBARectangle.PNG",
    x: 0.8,
    y: 0.2,
    width: 8.8,
    height: 4.64
  },
  "g1904e63b9c6_0_8": {    // UBA Square
    // Se aplica la configuración directamente; ajustá "height" en la configuración para lograr el tamaño deseado.
    fileName: "SAFEZONE_UBASquare.PNG",
    x: 0.81,    // Aproximadamente 58 pt
    y: 0.15,    // Aproximadamente -48 pt (ajustado)
    width: 6.26, // Aproximadamente 451 pt
    height: 4.64 // Valor configurado; ajustalo si necesitás una reducción diferente
  },
  "g21c0d753fff_0_116": {  // Mockups Eclipse
    fileName: "SAFEZONE_MockupsEclipse.PNG",
    x: 0.8,
    y: 0.2,
    width: 8.8,
    height: 4.64
  },
  "g1f79e657f07_0_122": {  // Mockups Darwin – REEMPLAZAR si corresponde
    fileName: "SAFEZONE_MockupsDarwin.PNG",
    x: 0.8,
    y: 0.2,
    width: 8.8,
    height: 4.64
  },
  // Para UnBranded (Mockups UnBranded)
  "g340af34de75_1_1880": { // Este layout se usa solo en UnBranded mockups
    fileName: "SAFEZONE_MockupsUnBranded.PNG",
    x: 0.8,
    y: 0.2,
    width: 8.8,
    height: 4.64
  },
    // ID nuevo de UnBranded que detectaste
  "g3524dc21236_0_1915": {
    fileName: "SAFEZONE_MockupsUnBranded.PNG",
    x: 0.8, y: 0.2, width: 8.8, height: 4.64
  }
};

/**
 * Convierte pulgadas a puntos (1 in = 72 pt)
 */
function inchesToPoints(inches) {
  return inches * 72;
}

/**
 * Retorna la diapositiva activa. Si no hay selección, devuelve la primera.
 */
function getActiveSlide() {
  var presentation = SlidesApp.getActivePresentation();
  var selection = presentation.getSelection();
  if (selection) {
    var currentPage = selection.getCurrentPage();
    if (currentPage) return currentPage;
  }
  var slides = presentation.getSlides();
  return slides[0];
}

/**
 * Utilizando la Advanced Slides API, retorna el layoutObjectId del slide (usando su objectId).
 */
function getSlideLayoutId(slide) {
  var presentationId = SlidesApp.getActivePresentation().getId();
  var pres = Slides.Presentations.get(presentationId);
  var slideId = slide.getObjectId();
  for (var i = 0; i < pres.slides.length; i++) {
    var s = pres.slides[i];
    if (s.objectId === slideId) {
      if (s.slideProperties && s.slideProperties.layoutObjectId) {
        return s.slideProperties.layoutObjectId;
      }
    }
  }
  return "";
}

/**
 * Busca en la carpeta de overlays un archivo cuyo nombre (sin distinguir mayúsculas) coincida.
 */
function getOverlayFile(fileName) {
  var folder = DriveApp.getFolderById(OVERLAY_FOLDER_ID);
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().toLowerCase() === fileName.toLowerCase()) {
      return file;
    }
  }
  return null;
}

/**
 * Función de debug que togglea (muestra/oculta) el overlay de SafeZone en el slide activo 
 * y acumula mensajes de debug para mostrarlos (por ejemplo, en el Sidebar).
 */
function debugToggleSafezonesOverlay() {
  var debugLog = "";
  function log(msg) {
    debugLog += msg + "\n";
    Logger.log(msg);
  }
  try {
    var slide = getActiveSlide();
    log("Slide activo: " + slide.getObjectId());
    
    // Obtener el layoutObjectId usando la Advanced Slides API.
    var layoutId = getSlideLayoutId(slide);
    log("Layout ID detectado mediante Advanced Slides API: '" + layoutId + "'");
    
    if (!layoutId || layoutId === "") {
      log("No se pudo detectar el layout del slide.");
      return debugLog + "Proceso completado (con errores: no se detectó layout).";
    }
    
    var config = SAFEZONE_OVERLAY_CONFIG[layoutId];
    log("Configuración para layout '" + layoutId + "': " + JSON.stringify(config));
    if (!config) {
      log("No hay configurado un overlay para el layout: " + layoutId);
      return debugLog + "Proceso completado (con errores: overlay no configurado).";
    }
    
    var images = slide.getImages();
    for (var i = 0; i < images.length; i++) {
      var img = images[i];
      if (img.getTitle && img.getTitle() === "safezone_overlay") {
        log("Overlay ya existente encontrado, se remueve.");
        img.remove();
        log("Overlay removido.");
        return debugLog + "Proceso completado (overlay removido).";
      }
    }
    
    var overlayFile = getOverlayFile(config.fileName);
    if (!overlayFile) {
      log("No se encontró el archivo de overlay: " + config.fileName + " en la carpeta de overlays.");
      return debugLog + "Proceso completado (con errores: archivo no encontrado).";
    }
    log("Archivo overlay encontrado: " + overlayFile.getName());
    
    var blob = overlayFile.getBlob();
    var overlayImage = slide.insertImage(blob);
    overlayImage.setTitle("safezone_overlay");
    
    // Para UBA Rectangle, aplicar posicionamiento original.
    if (layoutId === "g175f42bfc06_0_36") {
      log("Aplicando posicionamiento original para UBA Rectangle (" + layoutId + ")");
      overlayImage.setLeft(inchesToPoints(config.x));
      overlayImage.setTop(inchesToPoints(config.y));
      overlayImage.setWidth(inchesToPoints(config.width));
      overlayImage.setHeight(inchesToPoints(config.height));
    }
    // Para UBA Square, usar los valores configurados (se aplican directamente).
    else if (layoutId === "g1904e63b9c6_0_8") {
      log("Aplicando posicionamiento para UBA Square (" + layoutId + ")");
      overlayImage.setLeft(inchesToPoints(config.x));
      overlayImage.setTop(inchesToPoints(config.y));
      overlayImage.setWidth(inchesToPoints(config.width));
      overlayImage.setHeight(inchesToPoints(config.height));
    }
    // Para los demás layouts, se escala el overlay para ocupar el 100% del slide (ancho y alto).
    else {
      log("Aplicando escalado para ocupar el slide completo para " + layoutId);
      // Utilizamos getPageWidth() y getPageHeight() para evitar el error.
      var presentation = SlidesApp.getActivePresentation();
      var slideWidth = presentation.getPageWidth();
      var slideHeight = presentation.getPageHeight();
      overlayImage.setLeft(0);
      overlayImage.setTop(0);
      overlayImage.setWidth(slideWidth);
      overlayImage.setHeight(slideHeight);
    }
    
    log("Overlay insertado en el slide " + slide.getObjectId());
    return debugLog + "Proceso completado (SafeZone Overlay).";
  } catch (error) {
    log("Error en debugToggleSafezonesOverlay: " + error);
    return debugLog + "Proceso completado (con errores: " + error + ").";
  }
}

/**
 * Función principal que togglea (muestra/oculta) el overlay de SafeZone en el slide activo.
 * Invoca debugToggleSafezonesOverlay() para realizar la operación y retorna el debug log.
 */
function toggleSafezonesOverlay() {
  return debugToggleSafezonesOverlay();
}