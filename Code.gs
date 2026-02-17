/**
 * ============================================================================
 * SISTEMA: Bitácora RNME (Registro Nacional de Medios de Examen Electrónicos)
 * ARQUITECTURA: Backend Serverless (Google Apps Script)
 * ============================================================================
 */

/**
 * ----------------------------------------------------------------------------
 * 1. CONFIGURACIÓN Y ESQUEMA (Integridad Referencial)
 * ----------------------------------------------------------------------------
 */
const SCHEMA = {
  T_USUARIOS: {
    sheetName: "APP_USUARIOS",
    columns: ["email", "rol", "estado", "ultimo_acceso"]
  },
  T_EMPRESAS: {
    sheetName: "APP_EMPRESAS",
    columns: ["id_empresa", "razon_social", "ruc", "representante", "email", "direccion", "tipo_entidad", "actividad_principal"]
  },
  T_EQUIPOS: {
    sheetName: "APP_EQUIPOS",
    columns: ["id_registro", "id_empresa", "descripcion","marca", "serial_psicométrico", "serial_sensométrico", "estado_homologacion"]
  },
  T_RESOLUCIONES: {
    sheetName: "APP_RESOLUCIONES",
    columns: ["id_resolucion", "tipo_acto", "afecta", "fecha_emision", "vencimiento", "estado", "url_drive", "qr", "id_equipo_vinculado"]
  },
  T_UBICACIONES: {
    sheetName: "APP_UBICACIONES",
    columns: ["id_ubicacion", "id_equipo", "departamento", "distrito", "competencia","lugar_especifico", "estado_actual"]
  },
  T_CONFIGURACION: {
    sheetName: "APP_CONFIGURACION",
    columns: ["clave", "valor", "descripcion", "ultima_modificacion"]
  },
  T_AUDITORIA: {
    sheetName: "APP_AUDITORIA",
    columns: ["id_log", "fecha_hora", "usuario_email", "accion", "tabla_afectada", "detalle_cambio"]
  },
  T_CATALOGOS: {
    sheetName: "APP_CATALOGOS",
    columns: ["id_catalogo", "categoria", "valor", "padre_id", "estado"]
  }
};

const SCHEMA_DEPENDENCIES = {
  APP_EMPRESAS: [{ table: "APP_EQUIPOS", foreignKeyIndex: 1 }], 
  APP_EQUIPOS: [
    { table: "APP_RESOLUCIONES", foreignKeyIndex: 8 }, 
    { table: "APP_UBICACIONES", foreignKeyIndex: 1 }   
  ]
};

function saveToAppTable(schemaKey, data) {
  if (data.length === 0) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEMA[schemaKey].sheetName);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
}

function extractEmpresaId(propietario) {
  if (/Touring/i.test(propietario)) return "TAC";
  if (/Prototipo/i.test(propietario)) return "PRO";
  if (/Itran/i.test(propietario)) return "ITR";
  if (/Transtecno/i.test(propietario)) return "TEC";
  return "OTR";
}

/**
 * ----------------------------------------------------------------------------
 * 2. CORE DE BASE DE DATOS (Idempotencia y Setup)
 * ----------------------------------------------------------------------------
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(key => {
    const table = SCHEMA[key];
    let sheet = ss.getSheetByName(table.sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(table.sheetName);
    }
    const headerRange = sheet.getRange(1, 1, 1, table.columns.length);
    headerRange.setValues([table.columns]);
    headerRange.setFontWeight("bold").setBackground("#444444").setFontColor("#FFFFFF").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, table.columns.length);
  });
}

/**
 * ----------------------------------------------------------------------------
 * 3. RESILIENCIA Y CONTROL DE CONCURRENCIA
 * ----------------------------------------------------------------------------
 */
function runWithRetry(fn, ...args) {
  const MAX_RETRIES = 3;
  let attempt = 0;
  while (attempt < MAX_RETRIES) {
    try {
      return fn(...args);
    } catch (e) {
      attempt++;
      if (attempt === MAX_RETRIES) throw e;
      Utilities.sleep(Math.pow(2, attempt) * 1000 + (Math.random() * 100));
    }
  }
}

/**
 * ----------------------------------------------------------------------------
 * 4. SEGURIDAD Y AUTENTICACIÓN (RBAC & Cache)
 * ----------------------------------------------------------------------------
 */
function authenticateUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return "PUBLIC";
  const cache = CacheService.getScriptCache();
  const cachedRole = cache.get(`AUTH_${email}`);
  if (cachedRole) return cachedRole;

  return runWithRetry(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("APP_USUARIOS");
    if (!sheet || sheet.getLastRow() < 2) return "PUBLIC";
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase() && 
          String(data[i][2]).trim().toUpperCase() === "ACTIVO") {
        const dbRol = String(data[i][1]).trim().toUpperCase();
        cache.put(`AUTH_${email}`, dbRol, 1800);
        updateLastAccessAsync(sheet, i + 2);
        return dbRol;
      }
    }
    return "PUBLIC";
  });
}

function updateLastAccessAsync(sheet, rowIndex) {
  try { sheet.getRange(rowIndex, 4).setValue(new Date().toISOString()); } catch (e) {}
}

function logAuditActivity(accion, tabla_afectada, detalle_cambio) {
  try {
    const email = Session.getActiveUser().getEmail() || "Sistema Público";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_AUDITORIA");
    if (sheet) {
      const id_log = "LOG-" + Utilities.getUuid();
      const fechaIso = new Date().toISOString();
      const detalleStr = typeof detalle_cambio === 'object' ? JSON.stringify(detalle_cambio) : detalle_cambio;
      sheet.appendRow([id_log, fechaIso, email, accion, tabla_afectada, detalleStr]);
    }
  } catch (e) {}
}

/**
 * ----------------------------------------------------------------------------
 * 5. CONTROLADOR FRONTEND (API para Vue 3)
 * ----------------------------------------------------------------------------
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Bitácora RNME - ANTSV')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialPayload() {
  return runWithRetry(() => {
    const userEmail = Session.getActiveUser().getEmail() || "Invitado";
    const userRole = authenticateUser();
    
    const rawConfig = getTableData("APP_CONFIGURACION");
    const appConfig = {};
    rawConfig.forEach(row => { appConfig[row.clave] = row.valor; });
    
    const db = {
      equipos: getTableData("APP_EQUIPOS"),
      ubicaciones: getTableData("APP_UBICACIONES"),
      resoluciones: getTableData("APP_RESOLUCIONES"),
      empresas: userRole !== "PUBLIC" ? getTableData("APP_EMPRESAS") : [],
      catalogos: getTableData("APP_CATALOGOS"),
      configuracion: appConfig,
      logoBase64: getLogoBase64()
    };

    let result;
    if (userRole === "PUBLIC") {
      result = { role: "PUBLIC", user: userEmail, data: buildPublicView(db) };
    } else {
      result = { role: userRole, user: userEmail, data: db };
    }
    return JSON.stringify(result);
  });
}

function getTableData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      let cellValue = row[i];
      if (cellValue instanceof Date) { cellValue = cellValue.toISOString().split('T')[0]; }
      obj[header] = cellValue;
    });
    return obj;
  });
}

function buildPublicView(db) {
  const equiposVigentes = db.equipos.filter(e => e.estado_homologacion === "Vigente");
  const vigentesIds = equiposVigentes.map(e => e.id_registro);
  return {
    equipos: equiposVigentes,
    ubicaciones: db.ubicaciones.filter(u => u.estado_actual === "Activo" && vigentesIds.includes(u.id_equipo))
  };
}

/**
 * ----------------------------------------------------------------------------
 * 6. MOTOR DE MIGRACIÓN (Ya no lo necesitas, pero lo dejo por si acaso)
 * ----------------------------------------------------------------------------
 */
function runDataMigration() {
  // Función original conservada para evitar romper código
}

/**
 * ----------------------------------------------------------------------------
 * 7. GESTOR DE RESOLUCIONES (Drive + Sheets + Multi-Equipo)
 * ----------------------------------------------------------------------------
 */
function processResolutionUpload(fileData, formData) {
  return runWithRetry(() => {
    const config = getTableData("APP_CONFIGURACION");
    const ubiConfig = config.find(c => c.clave === "UBI_RESOLUCIONES");
    if (!ubiConfig || !ubiConfig.valor) throw new Error("La variable UBI_RESOLUCIONES no está en APP_CONFIGURACION.");

    const folderIdMatch = ubiConfig.valor.match(/folders\/([a-zA-Z0-9_-]+)/);
    const folderId = folderIdMatch ? folderIdMatch[1] : null;
    if (!folderId) throw new Error("URL de carpeta UBI_RESOLUCIONES no tiene un formato válido.");

    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mimeType, fileData.nombre);
    const driveFile = folder.createFile(blob);
    const fileUrl = driveFile.getUrl();

    const sheetRes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_RESOLUCIONES");
    const fechaEmision = new Date();
    const fechaVencimiento = new Date();
    fechaVencimiento.setFullYear(fechaVencimiento.getFullYear() + 1); 
    const fechaEmiStr = fechaEmision.toISOString().split('T')[0];
    const fechaVenStr = fechaVencimiento.toISOString().split('T')[0];

    const sheetEq = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_EQUIPOS");
    const eqData = sheetEq.getDataRange().getValues();
    const equiposAfectados = Array.isArray(formData.id_equipo) ? formData.id_equipo : [formData.id_equipo];
    
    equiposAfectados.forEach(equipoId => {
      sheetRes.appendRow([formData.id_resolucion, formData.tipo_acto, "", fechaEmiStr, fechaVenStr, "Vigente", fileUrl, "", equipoId]);
      for (let i = 1; i < eqData.length; i++) {
        if (String(eqData[i][0]).trim().toUpperCase() === String(equipoId).trim().toUpperCase()) {
          sheetEq.getRange(i + 1, 7).setValue("Homologado"); 
          break;
        }
      }
    });

    logAuditActivity("CREATE", "APP_RESOLUCIONES", { id_resolucion: formData.id_resolucion, equipos: equiposAfectados });
    return { success: true, url: fileUrl };
  });
}

/**
 * ----------------------------------------------------------------------------
 * 8. GESTOR DE EMPRESAS (CRUD Seguro)
 * ----------------------------------------------------------------------------
 */
function saveEmpresaTransaction(payload) {
  return runWithRetry(() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_EMPRESAS");
    const data = sheet.getDataRange().getValues();
    const isUpdate = payload.isUpdate;
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id_empresa) { rowIndex = i + 1; break; }
    }
    const rowData = [
      payload.id_empresa.toUpperCase(), payload.razon_social, payload.ruc, payload.representante,
      payload.email, payload.direccion, payload.tipo_entidad, payload.actividad_principal
    ];
    if (isUpdate && rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      logAuditActivity("UPDATE", "APP_EMPRESAS", payload);
    } else {
      if (rowIndex > -1) throw new Error("El ID (Siglas) de la empresa ya existe.");
      sheet.appendRow(rowData);
      logAuditActivity("CREATE", "APP_EMPRESAS", payload);
    }
    return true;
  });
}

function deleteEmpresaTransaction(id_empresa) {
  return runWithRetry(() => {
    const eqSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_EQUIPOS");
    if (eqSheet && eqSheet.getLastRow() > 1) {
      const eqData = eqSheet.getDataRange().getValues();
      const hasChildren = eqData.some((row, i) => i > 0 && row[1] === id_empresa);
      if (hasChildren) throw new Error(`Integridad Protegida: No se puede eliminar la empresa ${id_empresa} porque tiene Equipos vinculados.`);
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_EMPRESAS");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id_empresa) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex);
      logAuditActivity("DELETE", "APP_EMPRESAS", { id_empresa: id_empresa });
    } else {
      throw new Error("La empresa no fue encontrada.");
    }
    return true;
  });
}

/**
 * ----------------------------------------------------------------------------
 * 9. GESTOR DE EQUIPOS (CRUD Seguro con Autonumeración)
 * ----------------------------------------------------------------------------
 */
function saveEquipoTransaction(payload) {
  return runWithRetry(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("APP_EQUIPOS");
    const isUpdate = payload.isUpdate;
    let finalId = payload.id_registro;

    if (!isUpdate) {
      const configSheet = ss.getSheetByName("APP_CONFIGURACION");
      const configData = configSheet.getDataRange().getValues();
      let configRow = -1;
      let currentNumberStr = "ESP000";
      for(let i = 1; i < configData.length; i++) {
        if(configData[i][0] === "NUMERACION_EQUIPOS") {
          currentNumberStr = configData[i][1];
          configRow = i + 1;
          break;
        }
      }
      let numPart = parseInt(currentNumberStr.replace("ESP", ""), 10) || 0;
      numPart++;
      finalId = "ESP" + String(numPart).padStart(3, '0');
      if(configRow > -1) {
        configSheet.getRange(configRow, 2).setValue(finalId);
      }
    }

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    if (isUpdate) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === finalId.trim().toUpperCase()) {
          rowIndex = i + 1; break;
        }
      }
    }

    const rowData = [
      finalId, payload.id_empresa, payload.descripcion, payload.marca,
      payload.serial_psicometrico, payload.serial_sensometrico, payload.estado_homologacion || "" 
    ];

    if (isUpdate && rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      logAuditActivity("UPDATE", "APP_EQUIPOS", { ...payload, id_registro: finalId });
    } else {
      sheet.appendRow(rowData);
      logAuditActivity("CREATE", "APP_EQUIPOS", { ...payload, id_registro: finalId });
    }
    return true;
  });
}

function deleteEquipoTransaction(id_registro) {
  return runWithRetry(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resSheet = ss.getSheetByName("APP_RESOLUCIONES");
    if (resSheet && resSheet.getLastRow() > 1) {
      const resData = resSheet.getDataRange().getValues();
      const hasResolutions = resData.some((row, i) => i > 0 && row[8] === id_registro);
      if (hasResolutions) throw new Error(`Integridad Protegida: El equipo ${id_registro} tiene Resoluciones. Páselo a Inactivo.`);
    }
    const ubiSheet = ss.getSheetByName("APP_UBICACIONES");
    if (ubiSheet && ubiSheet.getLastRow() > 1) {
      const ubiData = ubiSheet.getDataRange().getValues();
      const hasLocations = ubiData.some((row, i) => i > 0 && row[1] === id_registro);
      if (hasLocations) throw new Error(`Integridad Protegida: El equipo ${id_registro} tiene Historial de Ubicaciones.`);
    }
    const sheet = ss.getSheetByName("APP_EQUIPOS");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id_registro) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex);
      logAuditActivity("DELETE", "APP_EQUIPOS", { id_registro: id_registro });
    } else {
      throw new Error("El equipo no fue encontrado.");
    }
    return true;
  });
}

/**
 * ----------------------------------------------------------------------------
 * 10. GESTOR DE UBICACIONES (Historial Geográfico)
 * ----------------------------------------------------------------------------
 */
function saveUbicacionTransaction(payload) {
  return runWithRetry(() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_UBICACIONES");
    const data = sheet.getDataRange().getValues();
    
    if (payload.estado_actual === "Activo") {
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === payload.id_equipo && data[i][6] === "Activo") {
          sheet.getRange(i + 1, 7).setValue("Histórico");
        }
      }
    }

    const newId = "UBI-" + Utilities.getUuid().substring(0,8).toUpperCase();
    const rowData = [
      newId, payload.id_equipo, payload.departamento, payload.distrito,
      payload.competencia, payload.lugar_especifico, payload.estado_actual
    ];

    sheet.appendRow(rowData);
    logAuditActivity("CREATE", "APP_UBICACIONES", payload);
    return true;
  });
}