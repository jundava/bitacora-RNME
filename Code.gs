/**
 * ============================================================================
 * SISTEMA: Bitácora RNME (Registro Nacional de Medios de Examen Electrónicos)
 * ARQUITECTURA: Backend Serverless (Google Apps Script)
 * ============================================================================
 */

const SCHEMA = {
  T_USUARIOS: { sheetName: "APP_USUARIOS", columns: ["email", "rol", "permisos", "estado", "ultimo_acceso", "avatar"] },
  T_EMPRESAS: { sheetName: "APP_EMPRESAS", columns: ["id_empresa", "razon_social", "ruc", "representante", "email", "direccion", "tipo_entidad", "actividad_principal"] },
  T_EQUIPOS: { sheetName: "APP_EQUIPOS", columns: ["id_registro", "id_empresa", "descripcion","marca", "serial_psicométrico", "serial_sensométrico", "estado_homologacion"] },
  T_RESOLUCIONES: { sheetName: "APP_RESOLUCIONES", columns: ["id_resolucion", "tipo_acto", "afecta", "fecha_emision", "vencimiento", "estado", "url_drive", "qr", "id_equipo_vinculado"] },
  T_UBICACIONES: { sheetName: "APP_UBICACIONES", columns: ["id_ubicacion", "id_equipo", "departamento", "distrito", "competencia","lugar_especifico", "estado_actual", "fecha_cierre"] },
  T_CONFIGURACION: { sheetName: "APP_CONFIGURACION", columns: ["clave", "valor", "descripcion", "ultima_modificacion"] },
  T_AUDITORIA: { sheetName: "APP_AUDITORIA", columns: ["id_log", "fecha_hora", "usuario_email", "accion", "tabla_afectada", "detalle_cambio"] },
  T_CATALOGOS: { sheetName: "APP_CATALOGOS", columns: ["id_catalogo", "categoria", "valor", "padre_id", "estado"] }
};

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(key => {
    const table = SCHEMA[key];
    let sheet = ss.getSheetByName(table.sheetName);
    if (!sheet) sheet = ss.insertSheet(table.sheetName);
    const headerRange = sheet.getRange(1, 1, 1, table.columns.length);
    headerRange.setValues([table.columns]).setFontWeight("bold").setBackground("#444444").setFontColor("#FFFFFF").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, table.columns.length);
  });
}

function runWithRetry(fn, ...args) {
  const MAX_RETRIES = 3;
  let attempt = 0;
  while (attempt < MAX_RETRIES) {
    try { return fn(...args); } 
    catch (e) {
      attempt++;
      if (attempt === MAX_RETRIES) throw e;
      Utilities.sleep(Math.pow(2, attempt) * 1000 + (Math.random() * 100));
    }
  }
}

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
    
    // CORRECCIÓN 1: Traemos 6 columnas (hasta la F, donde está el avatar)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues(); 
    
    for (let i = 0; i < data.length; i++) {      
      // data[i][0] = email (Col A) | data[i][3] = estado (Col D)
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase() && String(data[i][3]).trim().toUpperCase() === "ACTIVO") {
        const dbRol = String(data[i][1]).trim().toUpperCase();
        cache.put(`AUTH_${email}`, dbRol, 1800);
        
        // CORRECCIÓN 2: Escribimos el último acceso en la columna 5 (Col E)
        try { sheet.getRange(i + 2, 5).setValue(new Date().toISOString()); } catch (e) {}
        
        return dbRol;
      }
    }
    return "PUBLIC";
  });
}

function logAuditActivity(accion, tabla_afectada, detalle_cambio) {
  try {
    const email = Session.getActiveUser().getEmail() || "Sistema Público";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_AUDITORIA");
    if (sheet) {
      const id_log = "LOG-" + Utilities.getUuid();
      const detalleStr = typeof detalle_cambio === 'object' ? JSON.stringify(detalle_cambio) : detalle_cambio;
      sheet.appendRow([id_log, new Date().toISOString(), email, accion, tabla_afectada, detalleStr]);
    }
  } catch (e) {}
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Bitácora RNME - ANTSV')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialPayload() {
  return runWithRetry(() => {
    const userRole = authenticateUser();
    const rawConfig = getTableData("APP_CONFIGURACION");
    const appConfig = {};
    rawConfig.forEach(row => { appConfig[row.clave] = row.valor; });
    
    const db = {
      equipos: getTableData("APP_EQUIPOS"),
      ubicaciones: getTableData("APP_UBICACIONES"),
      resoluciones: getTableData("APP_RESOLUCIONES"),
      empresas: userRole !== "PUBLIC" ? getTableData("APP_EMPRESAS") : [],
      usuarios: userRole !== "PUBLIC" ? getTableData("APP_USUARIOS") : [],
      catalogos: getTableData("APP_CATALOGOS"),
      configuracion: appConfig,
      configuracion_raw: rawConfig,
      logoBase64: getLogoBase64() 
    };

    if (userRole === "PUBLIC") {
      const eqVigentes = db.equipos.filter(e => e.estado_homologacion === "Vigente");
      const ids = eqVigentes.map(e => e.id_registro);
      return JSON.stringify({ 
        role: "PUBLIC", 
        user: Session.getActiveUser().getEmail() || "Invitado", 
        data: { 
          equipos: eqVigentes, 
          ubicaciones: db.ubicaciones.filter(u => u.estado_actual === "Activo" && ids.includes(u.id_equipo)),
          logoBase64: db.logoBase64 // También lo enviamos a la vista pública
        } 
      });
    }
    return JSON.stringify({ role: userRole, user: Session.getActiveUser().getEmail() || "Invitado", data: db });
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
      obj[header] = row[i] instanceof Date ? row[i].toISOString().split('T')[0] : row[i];
    });
    return obj;
  });
}

// ================= MOTOR DE SEGURIDAD BACKEND =================
function requerirEditor(modulo) {
  const email = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_USUARIOS");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
      const rol = String(data[i][1]).trim().toUpperCase();
      if (rol === "ADMIN") return true; // Administrador pasa directo
      
      try {
        const permisos = JSON.parse(data[i][2]);
        if (permisos.roles && permisos.roles.includes('Todos')) return true;
        if (permisos[modulo] === 'Editor') return true;
      } catch(e) {}
      break;
    }
  }
  
  // Si llega hasta aquí, no es Editor ni Admin
  throw new Error("ACCESO DENEGADO: No tienes permisos de Editor para el módulo " + modulo);
}

// ================= GESTOR DE RESOLUCIONES =================
function processResolutionUpload(fileData, formData) {
  return runWithRetry(() => {
    const config = getTableData("APP_CONFIGURACION");
    const ubiConfig = config.find(c => c.clave === "UBI_RESOLUCIONES");
    if (!ubiConfig || !ubiConfig.valor) throw new Error("Carpeta UBI_RESOLUCIONES no configurada.");

    const folderId = (ubiConfig.valor.match(/folders\/([a-zA-Z0-9_-]+)/) || [])[1];
    const driveFile = DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mimeType, fileData.nombre));
    const fileUrl = driveFile.getUrl();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRes = ss.getSheetByName("APP_RESOLUCIONES");
    const sheetEq = ss.getSheetByName("APP_EQUIPOS");
    
    const eqData = sheetEq.getDataRange().getValues();
    const resData = sheetRes.getDataRange().getValues();
    
    const dEmi = new Date();
    const fEmiStr = dEmi.toISOString().split('T')[0];
    let fVenStr = "";

    // LÓGICA DE VENCIMIENTOS
    if (formData.tipo_acto === "3-Cambio Ubicación") {
      // Hereda vencimiento de la resolución afectada
      const resMadre = resData.find(r => r[0] === formData.afecta);
      if (!resMadre) throw new Error("No se encontró la Resolución Anterior para heredar el vencimiento.");
      fVenStr = new Date(resMadre[4]).toISOString().split('T')[0];
    } else {
      // Actos 1 y 2: Un año exacto desde la firma (hoy)
      const dVen = new Date();
      dVen.setFullYear(dVen.getFullYear() + 1);
      fVenStr = dVen.toISOString().split('T')[0];
    }
    
    const equiposAfectados = Array.isArray(formData.id_equipo) ? formData.id_equipo : [formData.id_equipo];
    const afectaId = formData.afecta || ""; // Resolución anterior
    
    equiposAfectados.forEach(eqId => {
      // Inserción incluyendo el campo 'afecta' (índice 2)
      sheetRes.appendRow([formData.id_resolucion, formData.tipo_acto, afectaId, fEmiStr, fVenStr, "Vigente", fileUrl, "", eqId]);
      
      // Si es Homologación o Renovación, sostenemos el estado a Homologado
      if(formData.tipo_acto !== "3-Cambio Ubicación") {
        for (let i = 1; i < eqData.length; i++) if (String(eqData[i][0]).trim() === eqId) { sheetEq.getRange(i + 1, 7).setValue("Homologado"); break; }
      }
    });

    // LÓGICA DE UBICACIONES UNIVERSAL (Ahora aplica si el usuario llena las ubicaciones, sin importar el tipo de acto)
    if (formData.ubicaciones_equipos && Object.keys(formData.ubicaciones_equipos).length > 0) {
      const sheetUbi = ss.getSheetByName("APP_UBICACIONES");
      const ubiData = sheetUbi.getDataRange().getValues();
      equiposAfectados.forEach(eqId => {
        const p = formData.ubicaciones_equipos[eqId];
        if(p && p.departamento && p.distrito && p.lugar_especifico) {
            for (let i = 1; i < ubiData.length; i++) if (ubiData[i][1] === eqId && ubiData[i][6] === "Activo") sheetUbi.getRange(i + 1, 7).setValue("Histórico");
            sheetUbi.appendRow(["UBI-" + Utilities.getUuid().substring(0,8).toUpperCase(), eqId, p.departamento, p.distrito, p.competencia || "Distrital", p.lugar_especifico, "Activo"]);
        }
      });
    }

    logAuditActivity("CREATE", "APP_RESOLUCIONES", formData.id_resolucion);
    return { success: true };
  });
}

function updateResolucionTransaction(payload) {
  return runWithRetry(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRes = ss.getSheetByName("APP_RESOLUCIONES");
    const resData = sheetRes.getDataRange().getValues();
    let originalRes = null; const rowsToDelete = [];
    
    for (let i = 1; i < resData.length; i++) {
      if (String(resData[i][0]).trim() === payload.id_original) {
        if (!originalRes) originalRes = { fEmi: resData[i][3], fVen: resData[i][4], st: resData[i][5], url: resData[i][6] };
        rowsToDelete.push(i + 1);
      }
    }

    if (!originalRes) throw new Error("Resolución original no encontrada.");
    for (let i = rowsToDelete.length - 1; i >= 0; i--) sheetRes.deleteRow(rowsToDelete[i]);

    const equipos = Array.isArray(payload.id_equipo) ? payload.id_equipo : [payload.id_equipo];
    const fEmiStr = originalRes.fEmi instanceof Date ? originalRes.fEmi.toISOString().split('T')[0] : originalRes.fEmi;
    const fVenStr = originalRes.fVen instanceof Date ? originalRes.fVen.toISOString().split('T')[0] : originalRes.fVen;
    const afectaId = payload.afecta || "";
    
    equipos.forEach(eqId => sheetRes.appendRow([payload.id_nuevo, payload.tipo_acto, afectaId, fEmiStr, fVenStr, originalRes.st, originalRes.url, "", eqId]));

    if (payload.ubicaciones_equipos && Object.keys(payload.ubicaciones_equipos).length > 0) {
      const sheetUbi = ss.getSheetByName("APP_UBICACIONES");
      const ubiData = sheetUbi.getDataRange().getValues();
      equipos.forEach(eqId => {
        const p = payload.ubicaciones_equipos[eqId];
        if(p && p.departamento && p.distrito && p.lugar_especifico) {
            for (let i = 1; i < ubiData.length; i++) if (ubiData[i][1] === eqId && ubiData[i][6] === "Activo") sheetUbi.getRange(i + 1, 7).setValue("Histórico");
            sheetUbi.appendRow(["UBI-" + Utilities.getUuid().substring(0,8).toUpperCase(), eqId, p.departamento, p.distrito, p.competencia || "Distrital", p.lugar_especifico, "Activo"]);
        }
      });
    }

    logAuditActivity("UPDATE", "APP_RESOLUCIONES", payload.id_original);
    return true;
  });
}

// ================= GESTOR DE EMPRESAS =================
function saveEmpresaTransaction(p) {
  return runWithRetry(() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_EMPRESAS");
    const data = sheet.getDataRange().getValues();
    let idx = -1; for (let i = 1; i < data.length; i++) if (data[i][0] === p.id_empresa) { idx = i + 1; break; }
    const row = [p.id_empresa.toUpperCase(), p.razon_social, p.ruc, p.representante, p.email, p.direccion, p.tipo_entidad, p.actividad_principal];
    
    if (p.isUpdate && idx > -1) { sheet.getRange(idx, 1, 1, row.length).setValues([row]); logAuditActivity("UPDATE", "APP_EMPRESAS", p.id_empresa); } 
    else { if (idx > -1) throw new Error("ID ya existe."); sheet.appendRow(row); logAuditActivity("CREATE", "APP_EMPRESAS", p.id_empresa); }
    return true;
  });
}

function deleteEmpresaTransaction(id) {
  return runWithRetry(() => {
    const eqSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_EQUIPOS");
    if (eqSheet && eqSheet.getLastRow() > 1 && eqSheet.getDataRange().getValues().some((row, i) => i > 0 && row[1] === id)) throw new Error(`La empresa ${id} tiene Equipos vinculados.`);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_EMPRESAS");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) if (data[i][0] === id) { sheet.deleteRow(i + 1); logAuditActivity("DELETE", "APP_EMPRESAS", id); return true; }
    throw new Error("Empresa no encontrada.");
  });
}

// ================= GESTOR DE EQUIPOS =================
function saveEquipoTransaction(p) {
  return runWithRetry(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("APP_EQUIPOS");
    let fId = p.id_registro;

    if (!p.isUpdate) {
      const confSheet = ss.getSheetByName("APP_CONFIGURACION");
      const cData = confSheet.getDataRange().getValues();
      let cRow = -1, curr = "ESP000";
      for(let i = 1; i < cData.length; i++) if(cData[i][0] === "NUMERACION_EQUIPOS") { curr = cData[i][1]; cRow = i + 1; break; }
      fId = "ESP" + String((parseInt(curr.replace("ESP", "")) || 0) + 1).padStart(3, '0');
      if(cRow > -1) confSheet.getRange(cRow, 2).setValue(fId);
    }

    const data = sheet.getDataRange().getValues();
    let idx = -1; if (p.isUpdate) { for (let i = 1; i < data.length; i++) if (String(data[i][0]).toUpperCase() === fId.toUpperCase()) { idx = i + 1; break; } }
    const row = [fId, p.id_empresa, p.descripcion, p.marca, p.serial_psicometrico, p.serial_sensometrico, p.estado_homologacion || ""];

    if (p.isUpdate && idx > -1) { sheet.getRange(idx, 1, 1, row.length).setValues([row]); logAuditActivity("UPDATE", "APP_EQUIPOS", fId); } 
    else { sheet.appendRow(row); logAuditActivity("CREATE", "APP_EQUIPOS", fId); }
    return true;
  });
}

function deleteEquipoTransaction(id) {
  return runWithRetry(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSheetByName("APP_RESOLUCIONES").getDataRange().getValues().some((r, i) => i > 0 && r[8] === id)) throw new Error("El equipo tiene Resoluciones.");
    if (ss.getSheetByName("APP_UBICACIONES").getDataRange().getValues().some((r, i) => i > 0 && r[1] === id)) throw new Error("El equipo tiene Ubicaciones.");
    
    const sheet = ss.getSheetByName("APP_EQUIPOS");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) if (String(data[i][0]) === id) { sheet.deleteRow(i + 1); logAuditActivity("DELETE", "APP_EQUIPOS", id); return true; }
    throw new Error("Equipo no encontrado.");
  });
}

// ================= GESTOR DE UBICACIONES =================
function saveUbicacionTransaction(p) {
  return runWithRetry(() => {
    requerirEditor("Ubicaciones");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("APP_UBICACIONES");
    const data = sheet.getDataRange().getValues();
    const hoy = new Date().toISOString().split('T')[0];

    // Evitar duplicados idénticos
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === p.id_equipo && data[i][6] === "Activo") {
        if (data[i][2] === p.departamento && data[i][3] === p.distrito && data[i][5] === p.lugar_especifico) {
          throw new Error("El equipo ya se encuentra en esta ubicación exacta.");
        }
        // Cerrar ubicación anterior como Histórico
        sheet.getRange(i + 1, 7).setValue("Histórico");
        sheet.getRange(i + 1, 8).setValue(hoy);
      }
    }

    sheet.appendRow(["UBI-" + Utilities.getUuid().substring(0,8).toUpperCase(), p.id_equipo, p.departamento, p.distrito, p.competencia, p.lugar_especifico, p.estado_actual, ""]);
    return true;
  });
}
/**
 * Mantiene el frontend HTML limpio y centraliza los assets en el servidor.
 */

// ================= UTILIDADES =================
function getLogoBase64() {
  return "https://i.postimg.cc/SxmBF7N1/Bitacora-Logo.png";
}

/**
 * ----------------------------------------------------------------------------
 * 13. MOTOR DEL TIEMPO (Cron Job Diario a las 00:00 hs)
 * ----------------------------------------------------------------------------
 */
function cronJobControlDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRes = ss.getSheetByName("APP_RESOLUCIONES");
  const sheetEq = ss.getSheetByName("APP_EQUIPOS");
  const sheetUbi = ss.getSheetByName("APP_UBICACIONES");

  const resData = sheetRes.getDataRange().getValues();
  const eqData = sheetEq.getDataRange().getValues();
  const ubiData = sheetUbi.getDataRange().getValues();

  const today = new Date();
  today.setHours(0,0,0,0); 
  const fechaHoyStr = today.toISOString().split('T')[0];
  const vigentesPorEquipo = new Set();

  // 1. Sincronizar Resoluciones
  for (let i = 1; i < resData.length; i++) {
    let tipoActo = resData[i][1];
    let emision = new Date(resData[i][3]);
    let vencimiento = new Date(resData[i][4]);
    let debeEstarVigente = (today >= emision && today <= vencimiento);
    let nuevoEstadoRes = debeEstarVigente ? "Vigente" : "No Vigente";
    if (resData[i][5] !== nuevoEstadoRes) sheetRes.getRange(i + 1, 6).setValue(nuevoEstadoRes);
    if (debeEstarVigente && (tipoActo === "1-Homologación" || tipoActo === "2-Renovación")) vigentesPorEquipo.add(String(resData[i][8]).trim());
  }

  // 2. Sincronizar Equipos y Ubicaciones con FECHA DE CIERRE
  for (let i = 1; i < eqData.length; i++) {
    let idEq = String(eqData[i][0]).trim();
    let esHomologado = vigentesPorEquipo.has(idEq);
    let nuevoEstadoEq = esHomologado ? "Homologado" : "No Homologado";
    if (eqData[i][6] !== nuevoEstadoEq) sheetEq.getRange(i + 1, 7).setValue(nuevoEstadoEq);

    // Regla para Ubicaciones
    let nuevoEstadoUbi = esHomologado ? "Activo" : "Inactivo";
    for (let j = 1; j < ubiData.length; j++) {
      if (String(ubiData[j][1]).trim() === idEq && (ubiData[j][6] === "Activo" || ubiData[j][6] === "Inactivo")) {
        if (ubiData[j][6] !== nuevoEstadoUbi) {
          sheetUbi.getRange(j + 1, 7).setValue(nuevoEstadoUbi);
          // Si pasa a Inactivo ponemos fecha, si vuelve a Activo la borramos
          sheetUbi.getRange(j + 1, 8).setValue(nuevoEstadoUbi === "Inactivo" ? fechaHoyStr : "");
        }
      }
    }
  }
}

/**
 * ----------------------------------------------------------------------------
 * 14. GESTOR DE CONFIGURACIÓN DEL SISTEMA
 * ----------------------------------------------------------------------------
 */
function saveConfigTransaction(payloadArray) {
  return runWithRetry(() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_CONFIGURACION");
    const data = sheet.getDataRange().getValues();

    // Recorrer lo que envió el frontend y actualizar fila por fila
    payloadArray.forEach(item => {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === item.clave) {
          sheet.getRange(i + 1, 2).setValue(item.valor); // Columna B (valor)
          sheet.getRange(i + 1, 4).setValue(new Date().toISOString()); // Columna D (ultima_modificacion)
          break;
        }
      }
    });

    logAuditActivity("UPDATE", "APP_CONFIGURACION", "Ajuste de parámetros del sistema");
    return true;
  });
}

/**
 * ----------------------------------------------------------------------------
 * 15. GESTOR DE USUARIOS Y PERMISOS
 * ----------------------------------------------------------------------------
 */
function saveUsuarioTransaction(p) {
  return runWithRetry(() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_USUARIOS");
    const data = sheet.getDataRange().getValues();
    let idx = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === p.email.trim().toLowerCase()) { 
        idx = i + 1; break; 
      }
    }
    
    // Convertimos el objeto de permisos a texto JSON para guardarlo
    const permisosStr = JSON.stringify(p.permisos);
    const row = [p.email.trim().toLowerCase(), p.rol, permisosStr, p.estado, p.ultimo_acceso || "", p.avatar || ""];

    if (p.isUpdate && idx > -1) { 
      sheet.getRange(idx, 1, 1, row.length).setValues([row]); 
      logAuditActivity("UPDATE", "APP_USUARIOS", p.email); 
    } else { 
      if (idx > -1) throw new Error("El correo ya está registrado."); 
      sheet.appendRow(row); 
      logAuditActivity("CREATE", "APP_USUARIOS", p.email); 
    }
    return true;
  });
}

function deleteUsuarioTransaction(email) {
  return runWithRetry(() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APP_USUARIOS");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === email.trim().toLowerCase()) { 
        sheet.deleteRow(i + 1); 
        logAuditActivity("DELETE", "APP_USUARIOS", email); 
        return true; 
      }
    }
    throw new Error("Usuario no encontrado.");
  });
}

