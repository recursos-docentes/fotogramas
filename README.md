# 📸 Automatización de Galería Fotográfica para Google Sheets (Edición Docente)

Este proyecto proporciona un script de **Google Apps Script** diseñado para organizar registros fotográficos de forma masiva en una hoja de cálculo. Es ideal para gestionar grupos de alumnos, proyectos de robótica (LEGO, ESP32, Micro:bit) o evidencias de talleres técnicos en el ámbito de **EBI y UTU**.

## 🚀 Funcionalidades Principales

* **Interfaz Interactiva:** Al ejecutar el script, se abre una ventana preguntando cuántas fotos deseás insertar.
* **Inserción Masiva:** Organiza las imágenes en una grilla de 5 columnas con tamaño uniforme (140px).
* **Sincronización de Nombres:** Lee los nombres de los alumnos desde una columna de tu planilla y los coloca automáticamente debajo de cada foto.
* **Orden Alfanumérico:** Ordena los archivos de Drive (1, 2, 3... 10, 11) para que coincidan con el orden de tu lista de nombres.
* **Ubicación Personalizada:** La galería se construye a partir de la **celda que tengas seleccionada** en ese momento.

---

## ⚠️ Configuración Crítica antes de Empezar

Para que el script funcione correctamente, debés definir dos ubicaciones clave en tu hoja de cálculo:

1.  **¿Dónde empieza la Galería?**: El script comenzará a insertar las fotos en la **celda donde tengas el cursor posicionado (Celda Activa)** al momento de darle "Ejecutar" o presionar el botón. **Asegurate de hacer clic en la celda inicial (ej. `A1`) antes de empezar.**
2.  **¿Dónde está tu lista de nombres?**: Debés tener una columna con los nombres de los estudiantes (ej. Columna `H`). El script tomará el primer nombre de esa lista para la primera foto, el segundo para la segunda, y así sucesivamente.
3.  **Asegurate de que la carpeta en Drive  tenga el acceso configurado como "Cualquier persona con el enlace" en modo Lector**

---

## 🛠️ Instrucciones de Instalación

1.  En tu Google Sheets, ve a **Extensiones > Apps Script**.
2.  Borra cualquier código existente y pega el contenido del archivo `codigo.gs`.
3.  **Configura las variables iniciales** en el código:
    * `folderId`: El ID de tu carpeta de Drive (se obtiene de la URL de la carpeta).
    * `columnaNombres`: La letra de la columna donde están los nombres (ej: `"H"`).
    * `filaInicioNombres`: El número de fila donde empieza el primer alumno (ej: `2`).
4.  Guarda el proyecto con el icono del disco.

---

## 💻 El Código (codigo.gs)

```javascript
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
  var folderId = 'TU_ID_DE_CARPETA_AQUÍ'; 
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
    var urlDrive = "https://docs.google.com/uc?export=view&id=" + listaFotos[i].id;
    
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
```

## 🔐 Recomendaciones de Seguridad y Acceso

* **Ventana de Incógnito:** Si usás varias cuentas de Google (personal y @ceibal o @utu), es **muy recomendable** abrir la planilla en una ventana de incógnito. Esto evita errores de permisos al ejecutar el script.
* **ID de la Carpeta:** Es el código alfanumérico que aparece al final de la URL de tu carpeta en Drive.
    * *Ejemplo de URL:* `https://drive.google.com/drive/folders/1A2b3C4d5E6f7G8h9I0jK1lM2nO3pQ4rS`
    * *El ID a copiar es:* `1A2b3C4d5E6f7G8h9I0jK1lM2nO3pQ4rS`

## 📋 Tip para el Portafolio Docente
Si necesitás la lista de nombres actualizada para que coincida con tus fotos:
1. Entrá a tu **Portafolio Docente** de UTU.
2. Descargá la lista de alumnos del grupo correspondiente (generalmente en formato Excel/CSV).
3. Copiá la columna de nombres y pegala en la columna de tu planilla (ej. Columna `H`).
4. ¡Listo! El script asociará cada foto al nombre oficial del estudiante de forma automática.
