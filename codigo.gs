/**
 * GENERADOR DE GALERÍA FOTOGRÁFICA DINÁMICO (VERSIÓN CORREGIDA)
 * Elizabeth - Docente de Informática (Uruguay)
 * * Este script inserta imágenes de Drive en una grilla de 5 columnas,
 * toma los nombres de una lista en la planilla y ajusta el tamaño a 140px.
 */
function generarGaleriaConfigurable() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Preguntar al docente cuántas fotos quiere insertar
  var respuesta = ui.prompt('Configuración de Galería', '¿Cuántas fotos deseás insertar?', ui.ButtonSet.OK_CANCEL);
  
  if (respuesta.getSelectedButton() != ui.Button.OK) return;
  
  var cantidadFotos = parseInt(respuesta.getResponseText());
  
  if (isNaN(cantidadFotos) || cantidadFotos <= 0) {
    ui.alert('Por favor, ingresá un número válido de fotos.');
    return;
  }

  // ==========================================
  // CONFIGURACIÓN (Ajustá estos valores)
  // ==========================================
  var folderId = 'TU_ID_DE_CARPETA_AQUÍ';    // <--- PEGÁ ACÁ EL ID DE TU CARPETA
  var columnaNombres = "H";                 // Letra de la columna con nombres
  var filaInicioNombres = 2;                // Fila donde empieza el primer nombre
  var altoImagen = 140;                     // Tamaño en píxeles
  
  // ==========================================
  // 1. OBTENCIÓN DE DATOS
  // ==========================================
  var sheet = SpreadsheetApp.getActiveSheet();
  var celdaActiva = sheet.getActiveCell();
  var filaInicial = celdaActiva.getRow();
  var colInicial = celdaActiva.getColumn();
  
  // Obtener lista de nombres de la columna
  var ultimaFila = sheet.getLastRow();
  var rangoNombres = sheet.getRange(columnaNombres + filaInicioNombres + ":" + columnaNombres + (filaInicioNombres + cantidadFotos - 1));
  var listaNombres = rangoNombres.getValues();
  
  // Obtener y ordenar fotos de Drive
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var listaFotos = [];
  
  while (files.hasNext()) {
    var file = files.next();
    listaFotos.push({
      id: file.getId(),
      nombreArchivo: file.getName()
    });
  }

  // Ordenar alfanuméricamente (1, 2, 3... 10, 11)
  listaFotos.sort(function(a, b) {
    return a.nombreArchivo.localeCompare(b.nombreArchivo, undefined, {numeric: true, sensitivity: 'base'});
  });

  // ==========================================
  // 2. PROCESO DE INSERCIÓN EN GRILLA
  // ==========================================
  var colRelativa = 0;
  var filaRelativa = 0;

  for (var i = 0; i < listaFotos.length && i < cantidadFotos; i++) {
    var fActual = filaInicial + filaRelativa;
    var cActual = colInicial + colRelativa;
    
    // CONSTRUCCIÓN DE URL CORREGIDA (Sin paréntesis extra)
    var urlDrive = "https://docs.google.com/uc?export=view&id=" + listaFotos[i].id;
    
    // A. Insertar la imagen
    sheet.getRange(fActual, cActual).setFormula('=IMAGE("' + urlDrive + '")');
    sheet.setRowHeight(fActual, altoImagen);
    sheet.setColumnWidth(cActual, altoImagen);
    
    // B. Insertar el nombre de la lista (fila de abajo)
    var nombreEstudiante = (listaNombres[i] && listaNombres[i][0]) ? listaNombres[i][0] : "Sin nombre";
    sheet.getRange(fActual + 1, cActual).setValue(nombreEstudiante);
    
    // C. Control de matriz (5 columnas)
    colRelativa++;
    if (colRelativa > 4) {
      colRelativa = 0;
      filaRelativa += 3; // Salto: Fila Foto + Fila Nombre + Fila Espacio
    }
  }

  // Formato estético final
  var rangoFinal = sheet.getRange(filaInicial, colInicial, filaRelativa + 2, 5);
  rangoFinal.setHorizontalAlignment("center").setVerticalAlignment("middle");
  rangoFinal.setWrap(true);
  
  ui.alert("✅ Proceso completado. Se han organizado " + i + " fotos.");
}
