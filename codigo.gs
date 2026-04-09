/**
 * GENERADOR DE GALERÍA FOTOGRÁFICA DINÁMICO
 * Elizabeth - Docente de Informática (Uruguay)
 */
function generarGaleriaConfigurable() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('Configuración de Galería', '¿Cuántas fotos deseás insertar?', ui.ButtonSet.OK_CANCEL);
  
  if (respuesta.getSelectedButton() != ui.Button.OK) return;
  var cantidadFotos = parseInt(respuesta.getResponseText());
  
  if (isNaN(cantidadFotos) || cantidadFotos <= 0) {
    ui.alert('Por favor, ingresá un número válido.');
    return;
  }

  // --- CONFIGURACIÓN DE VARIABLES ---
  var folderId = '1LMctoQq-IfD3LfsEI9DqNU0a8RuV7E_H'; 
  var columnaNombres = "H"; 
  var filaInicioNombres = 2; 
  var altoImagen = 140; 
  // ---------------------------------

  var sheet = SpreadsheetApp.getActiveSheet();
  var celdaActiva = sheet.getActiveCell(); // PUNTO DE INICIO
  var filaInicial = celdaActiva.getRow();
  var colInicial = celdaActiva.getColumn();
  
  var ultimaFila = sheet.getLastRow();
  var rangoNombres = sheet.getRange(columnaNombres + filaInicioNombres + ":" + columnaNombres + (filaInicioNombres + cantidadFotos - 1));
  var listaNombres = rangoNombres.getValues();
  
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var listaFotos = [];
  
  while (files.hasNext()) {
    var file = files.next();
    listaFotos.push({ id: file.getId(), nombreArchivo: file.getName() });
  }

  listaFotos.sort(function(a, b) {
    return a.nombreArchivo.localeCompare(b.nombreArchivo, undefined, {numeric: true, sensitivity: 'base'});
  });

  var colRelativa = 0;
  var filaRelativa = 0;

  for (var i = 0; i < listaFotos.length && i < cantidadFotos; i++) {
    var fActual = filaInicial + filaRelativa;
    var cActual = colInicial + colRelativa;
    var urlDrive = "https://lh3.googleusercontent.com/d/" + listaFotos[i].id;
    
    sheet.getRange(fActual, cActual).setFormula('=IMAGE("' + urlDrive + '")');
    sheet.setRowHeight(fActual, altoImagen);
    sheet.setColumnWidth(cActual, altoImagen);
    
    var nombreEstudiante = (listaNombres[i] && listaNombres[i][0]) ? listaNombres[i][0] : "Sin nombre";
    sheet.getRange(fActual + 1, cActual).setValue(nombreEstudiante);
    
    colRelativa++;
    if (colRelativa > 4) {
      colRelativa = 0;
      filaRelativa += 3; // Fila Foto + Fila Nombre + Fila Espacio
    }
  }

  var rangoFinal = sheet.getRange(filaInicial, colInicial, filaRelativa + 2, 5);
  rangoFinal.setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  ui.alert("✅ Proceso completado con éxito.");
}
