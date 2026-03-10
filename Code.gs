/**
 * ============================================================================
 * SISTEMA: Bitácora RNME - ANTSV
 * ARQUITECTURA: Backend Serverless (Google Apps Script) + Firestore (NoSQL)
 * ============================================================================
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Bitácora RNME - ANTSV')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ================= UTILIDADES Y CONEXIÓN CORE =================
function getLogoBase64() {
  return "https://i.postimg.cc/SxmBF7N1/Bitacora-Logo.png";
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

// ================= CAPA DE ACCESO A DATOS (FIRESTORE) =================
function getCollectionData(collectionName) {
  const db = getFirestore();
  try {
    const documents = db.getDocuments(collectionName);
    return documents.map(doc => unwrapFirestoreDoc(doc));
  } catch(e) {
    console.warn(`Colección ${collectionName} vacía o no encontrada.`);
    return [];
  }
}

// ================= SEGURIDAD Y AUDITORÍA =================
function authenticateUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return "PUBLIC";
  
  const cache = CacheService.getScriptCache();
  const cachedRole = cache.get(`AUTH_${email}`);
  if (cachedRole) return cachedRole;

  return runWithRetry(() => {
    const db = getFirestore();
    try {
      const docId = email.trim().toLowerCase().replace(/\//g, '_'); 
      const doc = db.getDocument(`APP_USUARIOS/${docId}`);
      const user = unwrapFirestoreDoc(doc);
      
      if (user && String(user.estado).trim().toUpperCase() === "ACTIVO") {
        const dbRol = String(user.rol).trim().toUpperCase();
        cache.put(`AUTH_${email}`, dbRol, 1800); 
        
        user.ultimo_acceso = new Date().toISOString();
        db.updateDocument(`APP_USUARIOS/${docId}`, user);
        return dbRol;
      }
    } catch (e) {}
    return "PUBLIC";
  });
}

function requerirEditor(modulo) {
  const email = Session.getActiveUser().getEmail();
  try {
    const docId = email.trim().toLowerCase().replace(/\//g, '_'); 
    const db = getFirestore();
    const doc = db.getDocument(`APP_USUARIOS/${docId}`);
    const user = unwrapFirestoreDoc(doc);
    
    if (String(user.rol).trim().toUpperCase() === "ADMIN") return true; 
    
    let permisos = user.permisos;
    if (typeof permisos === 'string') permisos = JSON.parse(permisos);
    
    if (permisos.roles && permisos.roles.includes('Todos')) return true;
    if (permisos[modulo] === 'Editor') return true;
  } catch(e) {}
  
  throw new Error("ACCESO DENEGADO: No tienes permisos de Editor para el módulo " + modulo);
}

function logAuditActivity(accion, tabla_afectada, detalle_cambio) {
  try {
    const email = Session.getActiveUser().getEmail() || "Sistema Público";
    const db = getFirestore();
    const id_log = "LOG-" + Utilities.getUuid();
    const detalleStr = typeof detalle_cambio === 'object' ? JSON.stringify(detalle_cambio) : detalle_cambio;
    
    db.updateDocument(`APP_AUDITORIA/${id_log}`, {
      id_log: id_log,
      fecha_hora: new Date().toISOString(),
      usuario_email: email,
      accion: accion,
      tabla_afectada: tabla_afectada,
      detalle_cambio: detalleStr
    });
  } catch (e) {}
}

// ================= INICIALIZACIÓN DEL FRONTEND =================
function getInitialPayload() {
  return runWithRetry(() => {
    const userRole = authenticateUser();
    
    const rawConfig = getCollectionData("APP_CONFIGURACION");
    const appConfig = {};
    rawConfig.forEach(row => { appConfig[row.clave] = row.valor; });
    
    const db = {
      equipos: getCollectionData("APP_EQUIPOS"),
      ubicaciones: getCollectionData("APP_UBICACIONES"),
      resoluciones: getCollectionData("APP_RESOLUCIONES"),
      empresas: userRole !== "PUBLIC" ? getCollectionData("APP_EMPRESAS") : [],
      usuarios: userRole !== "PUBLIC" ? getCollectionData("APP_USUARIOS") : [],
      catalogos: getCollectionData("APP_CATALOGOS"),
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
          logoBase64: db.logoBase64
        } 
      });
    }
    
    return JSON.stringify({ role: userRole, user: Session.getActiveUser().getEmail() || "Invitado", data: db });
  });
}

// ================= GESTOR DE EMPRESAS =================
function saveEmpresaTransaction(p) {
  return runWithRetry(() => {
    requerirEditor("Empresas");
    const db = getFirestore();
    const docId = String(p.id_empresa).toUpperCase().trim();
    const docPath = `APP_EMPRESAS/${docId}`;
    
    const payload = {
      id_empresa: docId, razon_social: p.razon_social, ruc: p.ruc, 
      representante: p.representante, email: p.email, direccion: p.direccion, 
      tipo_entidad: p.tipo_entidad, actividad_principal: p.actividad_principal
    };

    if (p.isUpdate) { 
      db.updateDocument(docPath, payload); 
      logAuditActivity("UPDATE", "APP_EMPRESAS", docId); 
    } else { 
      try {
        db.getDocument(docPath);
        throw new Error("El ID de la empresa ya existe.");
      } catch (e) {
        if(e.message.includes("ya existe")) throw e;
        db.updateDocument(docPath, payload); 
        logAuditActivity("CREATE", "APP_EMPRESAS", docId); 
      }
    }
    return true;
  });
}

function deleteEmpresaTransaction(id) {
  return runWithRetry(() => {
    requerirEditor("Empresas");
    const db = getFirestore();
    const equipos = getCollectionData("APP_EQUIPOS");
    if (equipos.some(eq => eq.id_empresa === id)) throw new Error(`La empresa ${id} tiene Equipos vinculados.`);
    
    db.deleteDocument(`APP_EMPRESAS/${id}`);
    logAuditActivity("DELETE", "APP_EMPRESAS", id); 
    return true;
  });
}

// ================= GESTOR DE EQUIPOS =================
function saveEquipoTransaction(p) {
  return runWithRetry(() => {
    requerirEditor("Equipos");
    const db = getFirestore();
    let fId = p.id_registro;

    if (!p.isUpdate) {
      let configDoc;
      try {
        configDoc = unwrapFirestoreDoc(db.getDocument("APP_CONFIGURACION/NUMERACION_EQUIPOS"));
      } catch(e) {
        configDoc = { clave: "NUMERACION_EQUIPOS", valor: "ESP000", descripcion: "Secuencia Autonumérica" };
      }
      const curr = configDoc.valor || "ESP000";
      fId = "ESP" + String((parseInt(curr.replace("ESP", "")) || 0) + 1).padStart(3, '0');
      
      configDoc.valor = fId;
      configDoc.ultima_modificacion = new Date().toISOString();
      db.updateDocument("APP_CONFIGURACION/NUMERACION_EQUIPOS", configDoc);
    }

    const docPath = `APP_EQUIPOS/${fId}`;
    const descripcionEstatica = `${fId}-${p.id_empresa}:[${p.serial_psicometrico}-${p.serial_sensometrico}]`;

    const payload = {
      id_registro: fId, id_empresa: p.id_empresa, descripcion: descripcionEstatica, 
      marca: p.marca, serial_psicometrico: p.serial_psicometrico, 
      serial_sensometrico: p.serial_sensometrico, estado_homologacion: p.estado_homologacion || ""
    };

    if (p.isUpdate) { 
      db.updateDocument(docPath, payload); 
      logAuditActivity("UPDATE", "APP_EQUIPOS", fId); 
    } else { 
      try {
        db.getDocument(docPath);
        throw new Error("El ID de equipo ya existe.");
      } catch(e) {
        if(e.message.includes("ya existe")) throw e;
        db.updateDocument(docPath, payload); 
        logAuditActivity("CREATE", "APP_EQUIPOS", fId); 
      }
    }
    return true;
  });
}

function deleteEquipoTransaction(id) {
  return runWithRetry(() => {
    requerirEditor("Equipos");
    const db = getFirestore();
    const resoluciones = getCollectionData("APP_RESOLUCIONES");
    const ubicaciones = getCollectionData("APP_UBICACIONES");
    
    if (resoluciones.some(r => r.equipos_vinculados && r.equipos_vinculados.includes(id))) throw new Error("El equipo tiene Resoluciones.");
    if (ubicaciones.some(u => u.id_equipo === id)) throw new Error("El equipo tiene Ubicaciones.");
    
    db.deleteDocument(`APP_EQUIPOS/${id}`);
    logAuditActivity("DELETE", "APP_EQUIPOS", id); 
    return true;
  });
}

// ================= GESTOR DE UBICACIONES =================
function saveUbicacionTransaction(p) {
  return runWithRetry(() => {
    requerirEditor("Ubicaciones");
    const db = getFirestore();
    
    // MODO EDICIÓN: Solo actualizamos los datos de texto de ese registro exacto
    if (p.isUpdate) {
      const docPath = `APP_UBICACIONES/${p.id_ubicacion}`;
      const existingDoc = unwrapFirestoreDoc(db.getDocument(docPath));
      
      existingDoc.departamento = p.departamento;
      existingDoc.distrito = p.distrito;
      existingDoc.competencia = p.competencia;
      existingDoc.lugar_especifico = p.lugar_especifico;
      
      db.updateDocument(docPath, existingDoc);
      logAuditActivity("UPDATE", "APP_UBICACIONES", p.id_ubicacion);
      return true;
    } 
    // MODO CREACIÓN (Traslado): Archiva el activo anterior y crea uno nuevo
    else {
      const ubicaciones = getCollectionData("APP_UBICACIONES");
      const hoy = new Date().toISOString().split('T')[0];

      const activas = ubicaciones.filter(u => u.id_equipo === p.id_equipo && u.estado_actual === "Activo");
      
      activas.forEach(uActiva => {
        if (uActiva.departamento === p.departamento && uActiva.distrito === p.distrito && uActiva.lugar_especifico === p.lugar_especifico) {
          throw new Error("El equipo ya se encuentra en esta ubicación exacta.");
        }
        uActiva.estado_actual = "Histórico";
        uActiva.fecha_cierre = hoy;
        db.updateDocument(`APP_UBICACIONES/${uActiva.id_ubicacion}`, uActiva);
      });

      const newId = "UBI-" + Utilities.getUuid().substring(0,8).toUpperCase();
      const payload = {
        id_ubicacion: newId, id_equipo: p.id_equipo, departamento: p.departamento, 
        distrito: p.distrito, competencia: p.competencia, lugar_especifico: p.lugar_especifico, 
        estado_actual: p.estado_actual, fecha_cierre: ""
      };
      
      db.updateDocument(`APP_UBICACIONES/${newId}`, payload);
      logAuditActivity("CREATE", "APP_UBICACIONES", p.id_equipo); 
      return true;
    }
  });
}

function deleteUbicacionTransaction(id_ubicacion) {
  return runWithRetry(() => {
    requerirEditor("Ubicaciones");
    const db = getFirestore();
    db.deleteDocument(`APP_UBICACIONES/${id_ubicacion}`);
    logAuditActivity("DELETE", "APP_UBICACIONES", id_ubicacion); 
    return true;
  });
}

// ================= GESTOR DE RESOLUCIONES =================
function processResolutionUpload(fileData, formData) {
  return runWithRetry(() => {
    requerirEditor("Resoluciones");
    const db = getFirestore();
    const nuevoIdRes = String(formData.id_resolucion).trim().toUpperCase();
    const docPath = `APP_RESOLUCIONES/${nuevoIdRes.replace(/\//g, '_')}`;

    try {
      db.getDocument(docPath);
      throw new Error("La resolución " + formData.id_resolucion + " ya se encuentra registrada.");
    } catch(e) { if(e.message.includes("registrada")) throw e; }

    let folderId = "";
    try {
      const ubiConfig = unwrapFirestoreDoc(db.getDocument("APP_CONFIGURACION/UBI_RESOLUCIONES"));
      folderId = (ubiConfig.valor.match(/folders\/([a-zA-Z0-9_-]+)/) || [])[1];
    } catch(e) { throw new Error("Carpeta UBI_RESOLUCIONES no configurada."); }

    const driveFile = DriveApp.getFolderById(folderId).createFile(
      Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mimeType, fileData.nombre)
    );
    const fileUrl = driveFile.getUrl();

    const fEmiStr = formData.fecha_emision;
    let fVenStr = formData.vencimiento;    

    if (formData.tipo_acto === "3-Cambio Ubicación") {
      try {
        const resMadre = unwrapFirestoreDoc(db.getDocument(`APP_RESOLUCIONES/${String(formData.afecta).replace(/\//g, '_')}`));
        fVenStr = resMadre.vencimiento;
      } catch(e) { throw new Error("No se encontró la Resolución Anterior para heredar el vencimiento."); }
    } 
    
    const equiposAfectados = Array.isArray(formData.id_equipo) ? formData.id_equipo : [formData.id_equipo];
    
    const payloadRes = {
      id_resolucion: formData.id_resolucion, tipo_acto: formData.tipo_acto, afecta: formData.afecta || "", 
      fecha_emision: fEmiStr, vencimiento: fVenStr, estado: "Vigente", url_drive: fileUrl, qr: "", 
      equipos_vinculados: equiposAfectados
    };
    db.updateDocument(docPath, payloadRes);

    equiposAfectados.forEach(eqId => {
      if(formData.tipo_acto !== "3-Cambio Ubicación") {
        try {
          const eqDoc = unwrapFirestoreDoc(db.getDocument(`APP_EQUIPOS/${eqId}`));
          eqDoc.estado_homologacion = "Homologado";
          db.updateDocument(`APP_EQUIPOS/${eqId}`, eqDoc);
        } catch(e){}
      }
      if (formData.ubicaciones_equipos && formData.ubicaciones_equipos[eqId]) {
        const p = formData.ubicaciones_equipos[eqId];
        if(p && p.departamento && p.distrito && p.lugar_especifico) {
            getCollectionData("APP_UBICACIONES").filter(u => u.id_equipo === eqId && u.estado_actual === "Activo").forEach(u => {
              u.estado_actual = "Histórico"; db.updateDocument(`APP_UBICACIONES/${u.id_ubicacion}`, u);
            });
            const uId = "UBI-" + Utilities.getUuid().substring(0,8).toUpperCase();
            db.updateDocument(`APP_UBICACIONES/${uId}`, {
              id_ubicacion: uId, id_equipo: eqId, departamento: p.departamento, distrito: p.distrito, 
              competencia: p.competencia || "Distrital", lugar_especifico: p.lugar_especifico, estado_actual: "Activo", fecha_cierre: ""
            });
        }
      }
    });

    logAuditActivity("CREATE", "APP_RESOLUCIONES", formData.id_resolucion);
    return { success: true };
  });
}

function updateResolucionTransaction(payload) {
  return runWithRetry(() => {
    requerirEditor("Resoluciones");
    const db = getFirestore();
    const idSafeAnterior = String(payload.id_original).trim().toUpperCase().replace(/\//g, '_');
    const idSafeNuevo = String(payload.id_nuevo).trim().toUpperCase().replace(/\//g, '_');
    
    let originalRes;
    try {
      originalRes = unwrapFirestoreDoc(db.getDocument(`APP_RESOLUCIONES/${idSafeAnterior}`));
    } catch(e) { throw new Error("Resolución original no encontrada."); }

    if (idSafeAnterior !== idSafeNuevo) db.deleteDocument(`APP_RESOLUCIONES/${idSafeAnterior}`);

    originalRes.id_resolucion = payload.id_nuevo;
    originalRes.tipo_acto = payload.tipo_acto;
    originalRes.afecta = payload.afecta || "";
    originalRes.fecha_emision = payload.fecha_emision || originalRes.fecha_emision;
    originalRes.vencimiento = payload.vencimiento || originalRes.vencimiento;
    originalRes.equipos_vinculados = Array.isArray(payload.id_equipo) ? payload.id_equipo : [payload.id_equipo];

    db.updateDocument(`APP_RESOLUCIONES/${idSafeNuevo}`, originalRes);

    logAuditActivity("UPDATE", "APP_RESOLUCIONES", payload.id_original);
    return true;
  });
}

// ================= GESTORES DE USUARIO Y CONFIG =================
function saveConfigTransaction(payloadArray) {
  return runWithRetry(() => {
    requerirEditor("Configuracion");
    const db = getFirestore();
    payloadArray.forEach(item => {
      item.ultima_modificacion = new Date().toISOString();
      db.updateDocument(`APP_CONFIGURACION/${item.clave}`, item);
    });
    logAuditActivity("UPDATE", "APP_CONFIGURACION", "Ajuste de parámetros");
    return true;
  });
}

function saveUsuarioTransaction(p) {
  return runWithRetry(() => {
    if(Session.getActiveUser().getEmail() !== "jundanielvallejostaniwaki@gmail.com" && p.rol === "Admin") {
       throw new Error("Solo el Super Administrador puede crear otros Admins.");
    }
    
    const db = getFirestore();
    const docId = String(p.email).trim().toLowerCase().replace(/\//g, '_');
    
    if (!p.isUpdate) {
      try {
        db.getDocument(`APP_USUARIOS/${docId}`);
        throw new Error("El correo ya está registrado.");
      } catch(e) { if(e.message.includes("registrado")) throw e; }
    }

    const payload = {
      email: p.email.trim().toLowerCase(), rol: p.rol, 
      permisos: typeof p.permisos === 'object' ? JSON.stringify(p.permisos) : p.permisos, 
      estado: p.estado, ultimo_acceso: p.ultimo_acceso || "", avatar: p.avatar || ""
    };

    db.updateDocument(`APP_USUARIOS/${docId}`, payload);
    logAuditActivity(p.isUpdate ? "UPDATE" : "CREATE", "APP_USUARIOS", p.email); 
    return true;
  });
}

function deleteUsuarioTransaction(email) {
  return runWithRetry(() => {
    const docId = String(email).trim().toLowerCase().replace(/\//g, '_');
    getFirestore().deleteDocument(`APP_USUARIOS/${docId}`);
    logAuditActivity("DELETE", "APP_USUARIOS", email); 
    return true;
  });
}

// ================= CRON JOB DIARIO =================
function cronJobControlDiario() {
  const db = getFirestore();
  const resoluciones = getCollectionData("APP_RESOLUCIONES");
  const equipos = getCollectionData("APP_EQUIPOS");
  const ubicaciones = getCollectionData("APP_UBICACIONES");

  const today = new Date();
  today.setHours(0,0,0,0); 
  const fechaHoyStr = today.toISOString().split('T')[0];
  const vigentesPorEquipo = new Set();

  resoluciones.forEach(res => {
    let emision = new Date(res.fecha_emision);
    let vencimiento = new Date(res.vencimiento);
    let debeEstarVigente = (today >= emision && today <= vencimiento);
    let nuevoEstadoRes = debeEstarVigente ? "Vigente" : "No Vigente";
    
    if (res.estado !== nuevoEstadoRes) {
      res.estado = nuevoEstadoRes;
      db.updateDocument(`APP_RESOLUCIONES/${res.id_resolucion.replace(/\//g, '_')}`, res);
    }
    
    if (debeEstarVigente && (res.tipo_acto === "1-Homologación" || res.tipo_acto === "2-Renovación")) {
      (res.equipos_vinculados || []).forEach(eqId => vigentesPorEquipo.add(eqId.trim()));
    }
  });

  equipos.forEach(eq => {
    let esHomologado = vigentesPorEquipo.has(String(eq.id_registro).trim());
    let nuevoEstadoEq = esHomologado ? "Homologado" : "No Homologado";
    
    if (eq.estado_homologacion !== nuevoEstadoEq) {
      eq.estado_homologacion = nuevoEstadoEq;
      db.updateDocument(`APP_EQUIPOS/${eq.id_registro}`, eq);
    }

    let nuevoEstadoUbi = esHomologado ? "Activo" : "Inactivo";
    ubicaciones.filter(u => String(u.id_equipo).trim() === String(eq.id_registro).trim() && (u.estado_actual === "Activo" || u.estado_actual === "Inactivo")).forEach(u => {
      if (u.estado_actual !== nuevoEstadoUbi) {
        u.estado_actual = nuevoEstadoUbi;
        u.fecha_cierre = (nuevoEstadoUbi === "Inactivo") ? fechaHoyStr : "";
        db.updateDocument(`APP_UBICACIONES/${u.id_ubicacion}`, u);
      }
    });
  });
}