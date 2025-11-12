/**
 * BorderModule.js - VERSIÓN OPTIMIZADA
 *
 * Funciones para agregar un borde de color a las imágenes seleccionadas en el slide.
 * * addBorderToSelected(colorHex, weightPt)
 * – colorHex: string con el color en hexadecimal, p.ej. "#00FF00"
 * – weightPt: grosor del borde en puntos
 *
 * addGreenBorder(), addYellowBorder(), addRedBorder()
 * – Llaman a addBorderToSelected() con los colores correspondientes.
 */

function addBorderToSelected(colorHex, weightPt) {
  var pres = SlidesApp.getActivePresentation();
  var sel  = pres.getSelection();
  var elems = [];

  // Obtener los elementos seleccionados (sin cambios aquí)
  if (sel.getSelectionType() === SlidesApp.SelectionType.PAGE_ELEMENT) {
    elems = sel.getPageElementRange().getPageElements();
  }
  
  if (elems.length === 0) {
    SlidesApp.getUi().alert("Por favor, selecciona al menos una imagen.");
    return;
  }
  
  // Para cada elemento: VERIFICAR si es una imagen y LUEGO aplicarle el borde.
  for (var i = 0; i < elems.length; i++) {
    var elem = elems[i];
    
    // ⭐ LA OPTIMIZACIÓN CLAVE ESTÁ AQUÍ ⭐
    // Comprobamos el tipo de elemento ANTES de hacer nada más.
    if (elem.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      // Como ya sabemos que es una imagen, podemos usar .asImage() de forma segura.
      var img = elem.asImage();
      var border = img.getBorder();
      
      border.setWeight(weightPt);
      border.getLineFill().setSolidFill(colorHex);
    }
    // Si no es una imagen, simplemente lo ignoramos y continuamos el bucle.
  }
}

/** Aplica un borde verde de 3pt */
function addGreenBorder() {
  addBorderToSelected("#00FF00", 2);
}

/** Aplica un borde amarillo de 3pt */
function addYellowBorder() {
  addBorderToSelected("#FFFF00", 2);
}

/** Aplica un borde rojo de 3pt */
function addRedBorder() {
  addBorderToSelected("#FF0000", 2);
}