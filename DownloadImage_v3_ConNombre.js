/**
 * downloadSelectedImages()
 *
 * Recorre los elementos seleccionados en la diapositiva activa y devuelve un array de objetos,
 * donde cada objeto tiene:
 *   • dataUrl: la imagen en formato Data URL (manteniendo su formato original).
 *   • fileName: el nombre original de la imagen (si existe); de lo contrario, se asigna un nombre genérico.
 *
 * Si ocurre algún error o no se encuentra ninguna imagen, devuelve un objeto con la propiedad error.
 */
function downloadSelectedImages() {
  try {
    var presentation = SlidesApp.getActivePresentation();
    var selection = presentation.getSelection();
    if (!selection) {
      throw "No se ha realizado ninguna selección.";
    }
    
    var pageElements;
    if (typeof selection.getPageElements === 'function') {
      pageElements = selection.getPageElements();
    } else if (typeof selection.getPageElementRange === 'function') {
      var range = selection.getPageElementRange();
      pageElements = range ? range.getPageElements() : [];
    } else {
      throw "La selección no contiene elementos válidos.";
    }
    
    if (!pageElements || pageElements.length === 0) {
      throw "No se seleccionó ningún elemento.";
    }
    
    var imagesArray = [];
    for (var i = 0; i < pageElements.length; i++){
      try {
        var image = pageElements[i].asImage();
        var blob = image.getBlob();
        var base64data = Utilities.base64Encode(blob.getBytes());
        var contentType = blob.getContentType();
        // Intentar obtener el título; si no existe o está vacío, usar el nombre del blob; si tampoco, asignar un nombre genérico.
        var fileName = image.getTitle();
        if (!fileName || fileName.trim() === "") {
          fileName = blob.getName();
          if (!fileName || fileName.trim() === "") {
            fileName = "Imagen " + (i + 1);
          }
        }
        var dataUrl = "data:" + contentType + ";base64," + base64data;
        imagesArray.push({ dataUrl: dataUrl, fileName: fileName });
      } catch(e) {
        // Si el elemento no es una imagen, continuar.
      }
    }
    if (imagesArray.length === 0) {
      throw "No se encontró ninguna imagen en la selección.";
    }
    return imagesArray;
  } catch (err) {
    return { error: "Error en downloadSelectedImages: " + err };
  }
}

/**
 * debugDownloadAllImagesFromSlide()
 *
 * Recorre TODOS los elementos del slide activo y filtra aquellos que sean imágenes (usando asImage()).
 * Devuelve un objeto con:
 *   • images: array de objetos { dataUrl, fileName } de cada imagen encontrada.
 *   • debug: cadena con los mensajes de debug.
 *
 * Si ocurre algún error o no se encuentran imágenes, se incluye la propiedad error.
 */
function debugDownloadAllImagesFromSlide() {
  var debugLog = "";
  function log(msg) {
    debugLog += msg + "\n";
    Logger.log(msg);
  }
  try {
    var slide = getActiveSlide();
    log("Slide activo: " + slide.getObjectId());
    
    var pageElements = slide.getPageElements();
    log("Se encontraron " + pageElements.length + " elementos en el slide.");
    
    var images = [];
    for (var i = 0; i < pageElements.length; i++) {
      try {
        var img = pageElements[i].asImage();
        images.push(img);
      } catch(e) {
        // No es imagen, se ignora.
      }
    }
    log("Se encontraron " + images.length + " elementos de imagen en el slide.");
    
    var imagesArray = [];
    for (var i = 0; i < images.length; i++){
      try {
        var blob = images[i].getBlob();
        var base64data = Utilities.base64Encode(blob.getBytes());
        var contentType = blob.getContentType();
        // Intenta obtener el título; si no existe, usar blob.getName(); si tampoco, asignar "Imagen X".
        var fileName = images[i].getTitle();
        if (!fileName || fileName.trim() === "") {
          fileName = blob.getName();
          if (!fileName || fileName.trim() === "") {
            fileName = "Imagen " + (i + 1);
          }
        }
        var dataUrl = "data:" + contentType + ";base64," + base64data;
        imagesArray.push({ dataUrl: dataUrl, fileName: fileName });
        log("Imagen " + (i+1) + ": " + fileName);
      } catch(e) {
        log("No se pudo procesar el elemento " + i + " como imagen.");
      }
    }
    if (imagesArray.length === 0) {
      log("No se encontraron imágenes en el slide.");
      return { error: "No se encontraron imágenes en el slide.", debug: debugLog };
    }
    log("Proceso completado (downloadAllImages).");
    return { images: imagesArray, debug: debugLog };
  } catch(err) {
    log("Error en debugDownloadAllImagesFromSlide: " + err);
    return { error: "Error en downloadAllImagesFromSlide: " + err, debug: debugLog };
  }
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
