/**
 * GENERADOR DE GALERÍA FOTOGRÁFICA DINÁMICO
 * Permite elegir la cantidad de fotos al inicio.
 * Los nombres se toman de una columna de la planilla.
 */
function generarGaleriaConfigurable() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Preguntar al docente cuántas fotos quiere insertar
  var respuesta = ui.prompt('Configuración de Galería', '¿Cuántas fotos deseás insertar?', ui.ButtonSet.OK_CANCEL);
  
  // Si el usuario cancela o cierra la ventana, salimos del script
  if (respuesta.getSelectedButton() != ui.Button.OK) return;
  
  var cantidadFotos = parseInt(respuesta.getResponseText());
  
  // Validar que sea un número válido
  if (isNaN(cantidadFotos) || cantidadFotos <= 0) {
    ui.alert('Por favor, ingresá un número válido de fotos.');
    return;
  }

  // ==========================================
  // CONFIGURACIÓN (Ajustar según tu planilla)
  // ==========================================
  var folderId = 'TU_ID_DE_CARPETA_AQUÍ';    // <--- PEGÁ ACÁ EL ID DE TU CARPETA
  var columnaNombres = "H";                 // Columna con nombres de alumnos
  var filaInicioNombres = 2;                // Fila donde empieza el primer nombre
  var altoImagen = 140; 
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var celdaActiva = sheet.getActiveCell();
  var filaInicial = celdaActiva.getRow();
  var colInicial = celdaActiva.getColumn();
  
  // Obtener lista de nombres
  var ultimaFila = sheet.getLastRow();
  var rangoNombres = sheet.getRange(columnaNombres + filaInicioNombres + ":" + columnaNombres + (filaInicioNombres + cantidadFotos));
  var listaNombres = rangoNombres.getValues();
  
  // Obtener y ordenar fotos de Drive
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var listaFotos = [];
  
  while (files.hasNext()) {
    var file = files.next();
    listaFotos.push({ id: file.getId(), nombreArchivo: file.getName() });
  }

  // Ordenar alfanuméricamente
  listaFotos.sort(function(a, b) {
    return a.nombreArchivo.localeCompare(b.nombreArchivo, undefined, {numeric: true, sensitivity: 'base'});
  });

  // ==========================================
  // 2. PROCESO DE INSERCIÓN
  // ==========================================
  var colRelativa = 0;
  var filaRelativa = 0;

  for (var i = 0; i < listaFotos.length && i < cantidadFotos; i++) {
    var fActual = filaInicial + filaRelativa;
    var cActual = colInicial + colRelativa;
    var url = "https://docs.google.com/uc?export=view&id=" + listaFotos[i].id;
    
    // Insertar Imagen
    sheet.getRange(fActual, cActual).setFormula('=IMAGE("' + url + '")');
    sheet.setRowHeight(fActual, altoImagen);
    sheet.setColumnWidth(cActual, altoImagen);
    
    // Insertar Nombre de la lista
    var nombreEstudiante = (listaNombres[i] && listaNombres[i][0]) ? listaNombres[i][0] : "Sin nombre";
    sheet.getRange(fActual + 1, cActual).setValue(nombreEstudiante);
    
    colRelativa++;
    if (colRelativa > 4) {
      colRelativa = 0;
      filaRelativa += 3; 
    }
  }

  // Formato final
  var rangoFinal = sheet.getRange(filaInicial, colInicial, filaRelativa + 2, 5);
  rangoFinal.setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  ui.alert("✅ Éxito: Se han organizado " + i + " fotos correctamente.");
}
