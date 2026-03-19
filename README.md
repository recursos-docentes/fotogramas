# Tutorial: Automatización de Galería de Imágenes en Google Sheets

Este recurso permite insertar y organizar masivamente imágenes desde una carpeta de Google Drive hacia una Hoja de Cálculo de Google, manteniendo un orden alfanumérico, dimensiones uniformes y etiquetas automáticas.

## 📋 Requisitos Previos

1.  **Carpeta en Drive:** Todas las fotos (máximo 27) deben estar en una sola carpeta.
2.  **Nomenclatura:** Los archivos deben estar numerados (ej. `1.jpg`, `2.jpg` o `Estudiante_01.png`) para asegurar el orden correcto.
3.  **Permisos:** La carpeta de Drive debe tener el acceso configurado como **"Cualquier persona con el enlace"** (lector) para que las imágenes sean visibles en la planilla.

---

## 🛠️ Instrucciones de Configuración

### 1. Obtener el ID de la carpeta
Entrá a tu carpeta de Google Drive y copiá el código que aparece al final de la URL.
* *Ejemplo de URL:* `https://drive.google.com/drive/folders/1abc123_identificador_de_tu_carpeta`
* *ID a copiar:* `1abc123_identificador_de_tu_carpeta`

### 2. Vincular el Script a la Hoja de Cálculo
1.  Abrí tu Google Sheets.
2.  Andá al menú superior: **Extensiones** > **Apps Script**.
3.  Borrá cualquier código existente y pegá el código que se encuentra al final de este tutorial.
4.  En la línea 5 del código, reemplazá `'TU_ID_DE_CARPETA_AQUÍ'` por el ID que copiaste en el paso anterior.
5.  Hacé clic en el icono del disco (**Guardar**) y ponéle un nombre al proyecto (ej. "Generador de Galería").

### 3. Ejecución del Script
1.  **Importante:** Hacé clic en la celda de la Hoja de Cálculo donde querés que comience la inserción (por ejemplo, la celda `A1`). El script empezará a trabajar desde esa posición exacta.
2.  Volvé a la pestaña de Apps Script y hacé clic en **Ejecutar**.
3.  **Autorización:** La primera vez, Google te pedirá permisos. Seleccioná tu cuenta, hacé clic en "Configuración avanzada" y luego en "Ir a [Nombre del proyecto] (no seguro)" para permitir el acceso a Drive y Sheets.

---

## 💻 Código para Apps Script

```javascript
/**
 * Genera una grilla de 27 fotos ordenadas a partir de la celda activa.
 * Organiza en filas de 5 imágenes, con nombre debajo y una fila libre de separación.
 */
function generarGrillaDesdeCeldaActiva() {
  // CONFIGURACIÓN
  var folderId = 'TU_ID_DE_CARPETA_AQUÍ'; 
  var altoImagen = 140; // Alto en píxeles
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var celdaInicial = sheet.getActiveCell();
  var filaInicial = celdaInicial.getRow();
  var colInicial = celdaInicial.getColumn();
  
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var listaFotos = [];

  // 1. Recolectar archivos
  while (files.hasNext()) {
    var file = files.next();
    listaFotos.push({
      nombre: file.getName(),
      id: file.getId()
    });
  }

  // 2. Ordenar alfanuméricamente
  listaFotos.sort(function(a, b) {
    return a.nombre.localeCompare(b.nombre, undefined, {numeric: true, sensitivity: 'base'});
  });

  var colRelativa = 0;
  var filaRelativa = 0;

  // 3. Insertar fotos (Máximo 27)
  for (var i = 0; i < listaFotos.length && i < 27; i++) {
    var fActual = filaInicial + filaRelativa;
    var cActual = colInicial + colRelativa;
    var url = "[https://docs.google.com/uc?export=view&id=](https://docs.google.com/uc?export=view&id=)" + listaFotos[i].id;
    
    // Insertar Imagen y configurar fila
    sheet.getRange(fActual, cActual).setFormula('=IMAGE("' + url + '")');
    sheet.setRowHeight(fActual, altoImagen);
    sheet.setColumnWidth(cActual, altoImagen);
    
    // Insertar Nombre en la fila de abajo
    sheet.getRange(fActual + 1, cActual).setValue(listaFotos[i].nombre);
    
    colRelativa++;

    // Control de matriz: 5 columnas
    if (colRelativa > 4) {
      colRelativa = 0;
      filaRelativa += 3; // Salto: imagen + nombre + espacio
    }
  }

  // 4. Formato estético
  var rangoFinal = sheet.getRange(filaInicial, colInicial, filaRelativa + 2, 5);
  rangoFinal.setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  Logger.log("Galería generada con éxito.");
}

## 💡 Notas para el Docente

* **Ajuste Dinámico:** Si necesitás que las fotos sean más grandes o más chicas, podés modificar el valor de la variable `altoImagen` (establecido por defecto en `140`) en la **línea 6** del script. Esto ajustará automáticamente tanto el alto de la fila como el ancho de la columna para mantener la proporción.
* **Uso en Proyectos:** Esta herramienta es ideal para diversos escenarios pedagógicos y técnicos:
    * **Portafolios de estudiantes:** Seguimiento visual de proyectos de robótica y programación.
    * **Registros de Huertas Institucionales:** Documentación fotográfica del crecimiento y automatización de cultivos.
    * **Catálogos de Componentes:** Organización de inventarios técnicos (sensores, placas ESP32, actuadores).
    * **Evaluaciones visuales:** Creación rápida de planillas de reconocimiento de errores en código o circuitos.
