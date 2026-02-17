/**
 * EJECUTA ESTA FUNCIÓN PARA GENERAR LAS HOJAS
 */
function ejecutarConfiguracionInicial() {
  setupDatabase();
  console.log("Proceso de creación de hojas finalizado.");
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SCHEMA).forEach(key => {
    const table = SCHEMA[key];
    let sheet = ss.getSheetByName(table.sheetName);
    
    // Si la hoja no existe, la creamos
    if (!sheet) {
      sheet = ss.insertSheet(table.sheetName);
    }
    
    // Configuramos las cabeceras
    const headerRange = sheet.getRange(1, 1, 1, table.columns.length);
    headerRange.setValues([table.columns]);
    
    // Aplicamos formato Profesional (Clean Code UX)
    headerRange.setFontWeight("bold")
               .setBackground("#444444") // Gris oscuro profesional
               .setFontColor("#FFFFFF") // Texto blanco
               .setHorizontalAlignment("center");
    
    sheet.setFrozenRows(1); // Congelar cabecera
    sheet.autoResizeColumns(1, table.columns.length); // Ajustar ancho
  });
}

