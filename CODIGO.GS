// Renderiza la App Web
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Control de Inventario - Bodega')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Actualiza el estado en Google Sheets (Columna E)
 */
function actualizarEstadoProducto(codigo, modelo, nuevoEstado) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() == codigo.toString() && data[i][2].toString() == modelo.toString()) {
        sheet.getRange(i + 1, 5).setValue(nuevoEstado); // Columna E (Estado)
        return { exito: true, mensaje: "Estado actualizado correctamente" };
      }
    }
    return { exito: false, mensaje: "Producto no encontrado en la hoja" };
  } catch (error) {
    return { exito: false, mensaje: error.toString() };
  }
}

/**
 * Obtiene todos los productos de la hoja de cálculo
 */
function obtenerProductos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var productos = [];
  
  for (var i = 1; i < data.length; i++) {
    productos.push({
      codigo: data[i][0],      // Columna A
      producto: data[i][1],    // Columna B (Categoría / Tipo de Producto)
      modelo: data[i][2],      // Columna C
      ubicacion: data[i][3],   // Columna D
      estado: data[i][4]       // Columna E
    });
  }
  return productos;
}
