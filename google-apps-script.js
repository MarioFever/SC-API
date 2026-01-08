/**
 * Google Apps Script para Google Sheets "SC verify"
 * 
 * Funcionalidades:
 * 1. addTimestamp: Añade fecha automática en columna K cuando se marca checkbox en columna E
 * 2. captureUserEmail: Captura el email del usuario que edita en columna O cuando se añade un dominio en columna A
 * 
 * Configuración:
 * - Hoja objetivo: "SC verify"
 * - Fila inicio: 368
 * - Columna trigger timestamp: E (5) - Checkbox "Verified SC?"
 * - Columna fecha: K (11) - "Date Created"
 * - Columna email: O (15)
 * 
 * Instalación:
 * 1. Abrir Google Sheets
 * 2. Extensions → Apps Script
 * 3. Pegar este código
 * 4. Guardar (Ctrl+S)
 * 
 * Nota: La columna O debe estar desprotegida o el usuario debe tener permisos de edición
 */

function onEdit(e) {
  // 1. Ejecutar Timestamp original
  addTimestamp(e);
  
  // 2. Ejecutar captura de usuario
  captureUserEmail(e);
}

function addTimestamp(e) {
  var tabName = SpreadsheetApp.getActiveSheet().getName();
  var startRow = 368; 
  var triggerCol = 5;   // Columna E (Verified SC? - checkbox)
  var dateCol = 11;     // Columna K (Date Created)
  
  if (tabName === "Partner owned webs UA") {
    startRow = 1;
    triggerCol = 4;
    dateCol = 5;  // Ajustar según corresponda para esta hoja
  } else if (tabName !== "SC verify") {
    return;
  }

  var rowModified = e.range.getRow();
  var colModified = e.range.getColumn();

  // Cuando se edita columna E (checkbox), escribe fecha en columna K
  if (colModified === triggerCol && rowModified >= startRow && 
      e.source.getActiveSheet().getName() === tabName && 
      e.source.getActiveSheet().getRange(rowModified, dateCol).getValue() == "") {
    e.source.getActiveSheet().getRange(rowModified, dateCol).setValue(new Date());
  }
}

function captureUserEmail(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "SC verify") return;

  var range = e.range;

  // Si se edita la Columna A (1) y no está vacía
  if (range.getColumn() == 1 && range.getRow() > 1 && range.getValue() !== "") {
    
    var userEmail = Session.getActiveUser().getEmail();
    
    if (!userEmail && e.user) {
      userEmail = e.user.email;
    }

    // DEBUG: Si falla, escribimos un aviso para saber que el script corre
    if (!userEmail) {
      userEmail = "⚠️ Falta Permiso";
    }

    sheet.getRange(range.getRow(), 15).setValue(userEmail);
  }
}

