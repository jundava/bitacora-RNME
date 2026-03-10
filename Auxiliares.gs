/**
 * ============================================================================
 * CAPA DE ACCESO A DATOS (DAL) - FIRESTORE
 * ============================================================================
 */

function getFirestore() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('FIREBASE_CLIENT_EMAIL');
  const projectId = props.getProperty('FIREBASE_PROJECT_ID');
  let key = props.getProperty('FIREBASE_PRIVATE_KEY');

  if (!email || !projectId || !key) {
    throw new Error("CATASTRÓFICO: Faltan credenciales de Firebase en las Propiedades del Script.");
  }

  // Sanitización obligatoria de la llave RSA
  key = key.replace(/\\n/g, '\n');

  return FirestoreApp.getFirestore(email, key, projectId);
}

/**
 * Decodificador recursivo para la API REST de Firestore
 * Convierte { stringValue: "Hola" } en -> "Hola"
 */
function extractFirestoreValue(valObj) {
  if (valObj === null || valObj === undefined) return "";
  if (typeof valObj !== 'object') return valObj; 
  
  if ('stringValue' in valObj) return valObj.stringValue;
  if ('integerValue' in valObj) return Number(valObj.integerValue);
  if ('doubleValue' in valObj) return Number(valObj.doubleValue);
  if ('booleanValue' in valObj) return valObj.booleanValue;
  if ('arrayValue' in valObj) return (valObj.arrayValue.values || []).map(extractFirestoreValue);
  if ('mapValue' in valObj) {
    const mapObj = {};
    const fields = valObj.mapValue.fields || {};
    Object.keys(fields).forEach(k => { mapObj[k] = extractFirestoreValue(fields[k]); });
    return mapObj;
  }
  if ('timestampValue' in valObj) return valObj.timestampValue;
  if ('nullValue' in valObj) return null;
  return "";
}

function unwrapFirestoreDoc(doc) {
  if (!doc || !doc.fields) return {};
  const unwrapped = {};
  Object.keys(doc.fields).forEach(key => {
    unwrapped[key] = extractFirestoreValue(doc.fields[key]);
  });
  return unwrapped;
}

/**
 * Lector Universal NoSQL
 * Extrae todos los documentos de una colección y los devuelve como Array de Objetos.
 */
function getCollectionData(collectionName) {
  const db = getFirestore();
  try {
    const documents = db.getDocuments(collectionName);
    // Ahora decodificamos cada documento antes de mandarlo a Vue 3
    return documents.map(doc => unwrapFirestoreDoc(doc));
  } catch(e) {
    console.warn(`Colección ${collectionName} vacía o no encontrada: ${e.message}`);
    return [];
  }
}

function testFirestoreConnection() {
  try {
    const db = getFirestore();
    const testDocId = "TEST-" + Utilities.getUuid().substring(0, 5);
    const testPayload = { 
      mensaje: "Conexión bidireccional exitosa desde GAS", 
      fecha: new Date().toISOString(),
      estado: "ACTIVO"
    };
    
    console.log("1. Intentando escribir en Firestore con ID estricto...");
    // Usamos updateDocument: Si no existe, lo crea con este ID exacto. Si existe, lo actualiza.
    db.updateDocument("APP_TEST/" + testDocId, testPayload);
    console.log("✅ Escritura exitosa. ID forzado: " + testDocId);
    
    console.log("2. Intentando leer el documento...");
    const doc = db.getDocument("APP_TEST/" + testDocId);
    console.log("✅ Lectura exitosa. Datos recibidos: ", JSON.stringify(doc.fields));
    
  } catch (error) {
    console.error("❌ FALLO DE CONEXIÓN: ", error.message);
  }
}

/**
 * ============================================================================
 * MOTOR ETL: MIGRACIÓN DE GOOGLE SHEETS A FIRESTORE (ONE-OFF SCRIPT)
 * ============================================================================
 */
function ejecutarMigracionMasiva() {
  const db = getFirestore();
  
  // Mapeo estricto del esquema relacional (Sheets) a colecciones NoSQL (Firestore)
  // idKey define qué columna será el nombre exacto del Documento.
  const esquemaMigracion = [
    { coleccion: "APP_CONFIGURACION", idKey: "clave" },
    { coleccion: "APP_CATALOGOS", idKey: "id_catalogo" },
    { coleccion: "APP_EMPRESAS", idKey: "id_empresa" },
    { coleccion: "APP_USUARIOS", idKey: "email" },
    { coleccion: "APP_EQUIPOS", idKey: "id_registro" },
    { coleccion: "APP_RESOLUCIONES", idKey: "id_resolucion" },
    { coleccion: "APP_UBICACIONES", idKey: "id_ubicacion" }
  ];

  console.log("🚀 INICIANDO MIGRACIÓN MASIVA DE DATOS A FIRESTORE...");

  esquemaMigracion.forEach(tabla => {
    console.log(`⏳ Extrayendo datos de la hoja: ${tabla.coleccion}...`);
    
    // Leemos la hoja usando tu función actual que ya convierte fechas a ISOString
    const datos = getTableData(tabla.coleccion); 
    
    if (!datos || datos.length === 0) {
      console.warn(`⚠️ Hoja ${tabla.coleccion} vacía o no encontrada. Omitiendo.`);
      return; // Pasa a la siguiente colección
    }

    let contadorExito = 0;
    
    datos.forEach(fila => {
      try {
        // 1. Sanitización Crítica: Extraemos el ID y limpiamos barras '/' que Firestore interpreta como subcolecciones
        let docId = String(fila[tabla.idKey]).trim().replace(/\//g, '_'); 
        
        // Excepción de seguridad: Ignorar filas donde el ID principal esté vacío
        if (!docId || docId === "undefined" || docId === "") return;

        // 2. Operación Atómica (Upsert)
        db.updateDocument(`${tabla.coleccion}/${docId}`, fila);
        contadorExito++;
        
        // 3. Rate Limit Defender: Pausa táctica de 20ms para evitar abrumar la API (Jittering)
        Utilities.sleep(20); 

      } catch (error) {
        console.error(`❌ Fallo en ${tabla.coleccion} -> Registro: ${fila[tabla.idKey]} | Causa: ${error.message}`);
      }
    });
    
    console.log(`✅ COLECCIÓN ${tabla.coleccion}: ${contadorExito} documentos migrados y sincronizados.`);
  });
  
  console.log("🏆 MIGRACIÓN COMPLETADA AL 100%. Por favor, verifica la consola de Firebase.");
}

function migrarResolucionesNoSQL() {
  const db = getFirestore();
  console.log("🚀 Iniciando migración y agrupación NoSQL de APP_RESOLUCIONES...");
  
  const datos = getTableData("APP_RESOLUCIONES");
  if (!datos || datos.length === 0) return;

  // 1. DICCIONARIO DE AGRUPACIÓN
  const resolucionesAgrupadas = {};

  datos.forEach(fila => {
    const idOriginal = String(fila.id_resolucion).trim();
    if (!idOriginal || idOriginal === "undefined" || idOriginal === "") return;
    
    // Limpiamos el ID para usarlo como llave en Firestore (evitamos barras '/' y el símbolo '°' que rompen la API)
    const docId = idOriginal.replace(/\//g, '_').replace(/°/g, '').trim();
    
    // Si la resolución no existe en nuestro diccionario, la creamos
    if (!resolucionesAgrupadas[docId]) {
      resolucionesAgrupadas[docId] = {
        id_resolucion: idOriginal,
        tipo_acto: String(fila.tipo_acto || ""),
        afecta: String(fila.afecta || ""),
        fecha_emision: String(fila.fecha_emision || ""),
        vencimiento: String(fila.vencimiento || ""),
        estado: String(fila.estado || "Vigente"),
        url_drive: String(fila.url_drive || ""),
        equipos_vinculados: [] // <-- Arreglo mágico NoSQL
      };
    }
    
    // Inyectamos el equipo al arreglo de esta resolución
    const equipo = String(fila.id_equipo_vinculado || "").trim();
    if (equipo && !resolucionesAgrupadas[docId].equipos_vinculados.includes(equipo)) {
      resolucionesAgrupadas[docId].equipos_vinculados.push(equipo);
    }
  });

  // 2. INYECCIÓN A FIRESTORE
  let contador = 0;
  Object.keys(resolucionesAgrupadas).forEach(key => {
    try {
      db.updateDocument(`APP_RESOLUCIONES/${key}`, resolucionesAgrupadas[key]);
      contador++;
      // Jittering: Pausa de 100ms para garantizar que la API respire entre cada documento pesado
      Utilities.sleep(100); 
    } catch (e) {
      console.error(`❌ Error en ${key}: ${e.message}`);
    }
  });
  
  console.log(`✅ EXCELENTE: ${contador} Resoluciones maestras agrupadas y migradas exitosamente.`);
}

/** AQUI VERSION 1 */

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

