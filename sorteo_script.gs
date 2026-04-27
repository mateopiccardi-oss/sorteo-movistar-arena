// ============================================================
//  SORTEO MOVISTAR ARENA — Apps Script completo
//  
//  CONFIGURACIÓN INICIAL:
//  1. Abrí este Sheet: docs.google.com/spreadsheets/d/[ID]
//  2. Menú → Extensiones → Apps Script
//  3. Pegá todo este código
//  4. Completá la sección CONFIG abajo
//  5. Menú → Implementar → Nueva implementación
//     - Tipo: Aplicación web
//     - Ejecutar como: Yo
//     - Quién tiene acceso: Cualquier persona
//  6. Copiá la URL que te da → la vas a necesitar en tu app
// ============================================================

const CONFIG = {

  // ID del Sheet de empleados (el que ya tenés)
  // URL: docs.google.com/spreadsheets/d/[ESTE-ID]/edit
  SHEET_EMPLEADOS_ID: "15YojsGUsDfhFBLOkBGxkLXvC7kSr6HUsJQgf99RhpWQ",

  // Nombre de la pestaña con los empleados
  SHEET_EMPLEADOS_HOJA: "Hoja 1",

  // ID del Sheet donde van a caer las inscripciones y ganadores
  // Puede ser el mismo Sheet de empleados (agregamos pestañas nuevas)
  // o uno nuevo que crees vacío
  SHEET_SORTEO_ID: "15YojsGUsDfhFBLOkBGxkLXvC7kSr6HUsJQgf99RhpWQ",

  // Nombre del remitente en los mails
  MAIL_REMITENTE: "Movistar Arena — RRHH",

  // Asunto del mail (variables: {nombre}, {evento})
  MAIL_ASUNTO: "Ganaste entradas para {evento}!",

  // Cuerpo del mail (variables: {nombre}, {evento}, {venue}, {fecha})
  // Nota: evitar emojis en el asunto para compatibilidad maxima
  MAIL_CUERPO: `Hola {nombre},

Felicitaciones! Resultaste ganador/a del sorteo de entradas para {evento}.

Fecha: {fecha}
Lugar: {venue}

Te adjuntamos tu entrada al mail.

Nos vemos en el show!

Equipo Movistar Arena`,

};

// ============================================================
//  ROUTER — maneja todos los requests de la app
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {
      case "validarMail":       return resp(validarMail(body.mail));
      case "inscribir":         return resp(inscribir(body.mail, body.showId, body.showNombre, body.fecha));
      case "getInscriptos":     return resp(getInscriptos(body.showId));
      case "guardarGanadores":  return resp(guardarGanadores(body.ganadores, body.showId, body.showNombre, body.fecha, body.venue));
      case "enviarMails":       return resp(enviarMails(body.showId, body.entradasXGan));
      case "syncGanadores":     return resp(syncGanadores(body.ganadores));
      case "trackingGanadores": return resp(trackingGanadores(body.ganadores, body.showNombre, body.fecha));
      case "leerTracking":           return resp(leerTracking());
      case "getUltimasVictorias":    return resp(getUltimasVictorias(body.nombres));
      case "getShows":               return resp(getShows());
      default:                  return resp({ ok: false, error: "Acción desconocida: " + action });
    }
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

function doGet(e) {
  // Para testear que el script está activo
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: "Sorteo Arena API activa" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function resp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  VALIDAR MAIL — busca el mail en el Sheet de empleados
//  Devuelve: { ok, existe, nombre, mail }
// ============================================================
function validarMail(mail) {
  if (!mail) return { ok: false, error: "Mail vacío" };

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_EMPLEADOS_ID);
  const hoja = ss.getSheetByName(CONFIG.SHEET_EMPLEADOS_HOJA);
  if (!hoja) return { ok: false, error: "No encontré la hoja de empleados" };

  const datos = hoja.getDataRange().getValues();
  const mailNorm = mail.trim().toLowerCase();

  // Busca en todas las columnas la que tenga formato mail
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    for (let j = 0; j < fila.length; j++) {
      const cell = String(fila[j]).trim().toLowerCase();
      if (cell === mailNorm) {
        // Buscar nombre en la misma fila (primera columna con texto largo)
        let nombre = '';
        for (let k = 0; k < fila.length; k++) {
          const val = String(fila[k]).trim();
          if (k !== j && val && !val.includes('@') && val.length > 2) {
            nombre = val;
            break;
          }
        }
        return { ok: true, existe: true, nombre: nombre || mail, mail: mail.trim() };
      }
    }
  }

  return { ok: true, existe: false };
}

// ============================================================
//  INSCRIBIR — registra la inscripción en el Sheet de sorteos
//  Devuelve: { ok, mensaje, yaInscripto }
// ============================================================
function inscribir(mail, showId, showNombre, fecha) {
  if (!mail || !showId) return { ok: false, error: "Datos incompletos" };

  // Validar que existe en empleados
  const validacion = validarMail(mail);
  if (!validacion.existe) return { ok: false, error: "Mail no encontrado en el registro de empleados" };

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  const hojaName = "Inscripciones";
  let hoja = ss.getSheetByName(hojaName);

  // Crear hoja si no existe
  if (!hoja) {
    hoja = ss.insertSheet(hojaName);
    hoja.appendRow(["Timestamp", "Show ID", "Show", "Fecha Show", "Mail", "Nombre", "Fecha Inscripción"]);
    hoja.getRange(1, 1, 1, 7).setFontWeight("bold");
  }

  // Verificar si ya está inscripto
  const datos = hoja.getDataRange().getValues();
  const mailNorm = mail.trim().toLowerCase();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][4]).trim().toLowerCase() === mailNorm &&
        String(datos[i][1]).trim() === String(showId).trim()) {
      return { ok: true, yaInscripto: true, mensaje: "Ya estás inscripto para este show", nombre: validacion.nombre };
    }
  }

  // Registrar inscripción
  const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  hoja.appendRow([
    new Date(),
    showId,
    showNombre,
    fecha || "",
    mail.trim().toLowerCase(),
    validacion.nombre,
    fechaHora
  ]);

  return { ok: true, yaInscripto: false, mensaje: "¡Inscripción confirmada!", nombre: validacion.nombre };
}

// ============================================================
//  HELPERS
// ============================================================
function normalizarShowId(s) {
  return String(s).trim().toUpperCase().replace(/[^A-Z0-9]/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
}

// ============================================================
//  GET INSCRIPTOS — devuelve todos los inscriptos de un show
// ============================================================
function getInscriptos(showId) {
  if (!showId) return { ok: false, error: "Show ID requerido" };

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  const hoja = ss.getSheetByName("Inscripciones");
  if (!hoja) return { ok: true, inscriptos: [] };

  const datos = hoja.getDataRange().getValues();
  const inscriptos = [];

  for (let i = 1; i < datos.length; i++) {
    if (normalizarShowId(datos[i][1]) === normalizarShowId(showId)) {
      inscriptos.push({
        mail: String(datos[i][4]).trim(),
        nombre: String(datos[i][5]).trim(),
        fechaInscripcion: String(datos[i][6]).trim(),
      });
    }
  }

  return { ok: true, inscriptos };
}

// ============================================================
//  GUARDAR GANADORES — registra ganadores después del sorteo
// ============================================================
function guardarGanadores(ganadores, showId, showNombre, fecha, venue) {
  if (!ganadores || !ganadores.length) return { ok: false, error: "Sin ganadores" };

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  let hoja = ss.getSheetByName("Ganadores");

  if (!hoja) {
    hoja = ss.insertSheet("Ganadores");
    hoja.appendRow(["Timestamp", "Show ID", "Show", "Fecha Show", "Venue", "Mail", "Nombre", "Estado Mail", "PDF enviado"]);
    hoja.getRange(1, 1, 1, 9).setFontWeight("bold");
  }

  const fecha_reg = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  ganadores.forEach(g => {
    hoja.appendRow([
      new Date(),
      showId,
      showNombre,
      fecha || "",
      venue || "Movistar Arena",
      g.mail || g.email,
      g.nombre,
      "Pendiente",
      ""
    ]);
  });

  return { ok: true, mensaje: `${ganadores.length} ganadores guardados` };
}

// ============================================================
//  ENVIAR MAILS — envía mails a los ganadores de un show
//  Los PDFs deben estar en una carpeta de Drive nombrada igual
//  que el Show ID, dentro de una carpeta "Entradas"
//
//  Estructura de Drive:
//  📁 Sorteo Movistar Arena
//    📁 Entradas
//      📁 [SHOW_ID]        ← PDFs de este show, ordenados
//        📄 entrada_1.pdf
//        📄 entrada_2.pdf
// ============================================================
function enviarMails(showId, entradasXGan) {
  if (!showId) return { ok: false, error: "Show ID requerido" };
  entradasXGan = parseInt(entradasXGan) || 1;
  Logger.log("Entradas por ganador: " + entradasXGan);

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  const hoja = ss.getSheetByName("Ganadores");
  if (!hoja) return { ok: false, error: "No hay ganadores registrados" };

  const datos = hoja.getDataRange().getValues();
  const ganadores = [];
  const filas = [];

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() === String(showId).trim() && datos[i][7] === "Pendiente") {
      ganadores.push({
        mail: String(datos[i][5]).trim(),
        nombre: String(datos[i][6]).trim(),
        showNombre: String(datos[i][2]).trim(),
        fecha: String(datos[i][3]).trim(),
        venue: String(datos[i][4]).trim(),
        fila: i + 1,
      });
      filas.push(i + 1);
    }
  }

  if (!ganadores.length) return { ok: true, enviados: 0, mensaje: "No hay ganadores pendientes" };

  // Buscar PDFs en Drive (por nombre del show o por ID)
  const showNombre2 = ganadores.length > 0 ? ganadores[0].showNombre : "";
  const pdfs = buscarPDFs(showId, showNombre2);

  const errores = [];
  let enviados = 0;

  ganadores.forEach((g, i) => {
    try {
      const asunto = CONFIG.MAIL_ASUNTO
        .replace(/{nombre}/g, g.nombre)
        .replace(/{evento}/g, g.showNombre);

      const cuerpo = CONFIG.MAIL_CUERPO
        .replace(/{nombre}/g, g.nombre)
        .replace(/{evento}/g, g.showNombre)
        .replace(/{venue}/g, g.venue)
        .replace(/{fecha}/g, g.fecha);

      const opts = { name: CONFIG.MAIL_REMITENTE };

      // Adjuntar N PDFs por ganador según entradasXGan
      const pdfStart = i * entradasXGan;
      const pdfSlice = pdfs.slice(pdfStart, pdfStart + entradasXGan);
      if (pdfSlice.length > 0) {
        opts.attachments = pdfSlice.map(function(p) { return p.getAs(MimeType.PDF); });
        Logger.log("Ganador " + (i+1) + " (" + g.nombre + "): " + pdfSlice.length + " PDF(s) adjuntos");
      } else {
        Logger.log("Ganador " + (i+1) + " (" + g.nombre + "): sin PDFs disponibles (offset " + pdfStart + ")");
      }

      // Enviar como HTML para soportar UTF-8 y emojis correctamente
      opts.htmlBody = cuerpo.replace(/\n/g, "<br>");
      GmailApp.sendEmail(g.mail, asunto, cuerpo, opts);

      // Marcar como enviado en el Sheet
      hoja.getRange(g.fila, 8).setValue("Enviado");
      hoja.getRange(g.fila, 9).setValue(pdfSlice.length > 0 ? pdfSlice.map(function(p){return p.getName();}).join(", ") : "Sin PDF");

      enviados++;
      Utilities.sleep(1200);
    } catch (err) {
      errores.push(`${g.nombre} (${g.mail}): ${err.message}`);
    }
  });

  return {
    ok: true,
    enviados,
    errores,
    mensaje: `${enviados} mail${enviados !== 1 ? 's' : ''} enviado${enviados !== 1 ? 's' : ''}` +
             (errores.length ? ` · ${errores.length} con error` : '')
  };
}

// ============================================================
//  GET SHOWS — devuelve los shows únicos del Sheet
// ============================================================
function getShows() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  const hoja = ss.getSheetByName("Inscripciones");
  if (!hoja) return { ok: true, shows: [] };

  const datos = hoja.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < datos.length; i++) {
    const id = String(datos[i][1]).trim();
    if (id && !map[id]) {
      map[id] = {
        id,
        nombre: String(datos[i][2]).trim(),
        fecha: String(datos[i][3]).trim(),
      };
    }
  }

  return { ok: true, shows: Object.values(map) };
}

// ============================================================
//  BUSCAR PDFs EN DRIVE
//  Estructura esperada:
//  Mi Drive / Sorteo Movistar Arena / Entradas / [showId] /
// ============================================================
function buscarPDFs(showId, showNombre) {
  try {
    // Buscar carpeta raiz
    const todasCarpetas = DriveApp.getFoldersByName("Sorteo Movistar Arena");
    if (!todasCarpetas.hasNext()) {
      Logger.log("No encontre 'Sorteo Movistar Arena'. Listando carpetas en raiz...");
      const root = DriveApp.getRootFolder();
      const subs = root.getFolders();
      const nombres = [];
      while (subs.hasNext()) nombres.push(subs.next().getName());
      Logger.log("Carpetas en raiz: " + nombres.join(", "));
      return [];
    }

    const raiz = todasCarpetas.next();
    Logger.log("OK - Carpeta raiz: " + raiz.getName());

    const entradas = raiz.getFoldersByName("Entradas");
    if (!entradas.hasNext()) {
      const subs2 = raiz.getFolders();
      const n2 = [];
      while (subs2.hasNext()) n2.push(subs2.next().getName());
      Logger.log("No encontre 'Entradas'. Subcarpetas disponibles: " + n2.join(", "));
      return [];
    }

    const carpetaEntradas = entradas.next();

    // Listar shows disponibles
    const subs3 = carpetaEntradas.getFolders();
    const showsDisponibles = [];
    while (subs3.hasNext()) showsDisponibles.push(subs3.next().getName());
    Logger.log("Shows disponibles: " + showsDisponibles.join(", "));

    // Buscar carpeta del show: exacto > por ID > case-insensitive
    let carpetaShow = null;

    if (showNombre) {
      const m1 = carpetaEntradas.getFoldersByName(showNombre);
      if (m1.hasNext()) { carpetaShow = m1.next(); Logger.log("Match exacto: " + showNombre); }
    }
    if (!carpetaShow) {
      const m2 = carpetaEntradas.getFoldersByName(String(showId));
      if (m2.hasNext()) { carpetaShow = m2.next(); Logger.log("Match por ID: " + showId); }
    }
    if (!carpetaShow) {
      const buscar = (showNombre || showId).toString().toLowerCase().trim();
      const m3 = carpetaEntradas.getFolders();
      while (m3.hasNext()) {
        const f = m3.next();
        if (f.getName().toLowerCase().trim() === buscar) {
          carpetaShow = f;
          Logger.log("Match case-insensitive: " + f.getName());
          break;
        }
      }
    }

    if (!carpetaShow) {
      Logger.log("No encontre carpeta para: '" + showNombre + "'. Shows disponibles: " + showsDisponibles.join(", "));
      return [];
    }

    // getFilesByMimeType no disponible en todos los entornos — usamos getFiles() y filtramos
    const archivos = carpetaShow.getFiles();
    const pdfs = [];
    while (archivos.hasNext()) {
      const f = archivos.next();
      const mime = f.getMimeType();
      if (mime === MimeType.PDF || mime === "application/pdf") {
        pdfs.push(f);
      }
    }
    pdfs.sort(function(a, b) { return a.getName().localeCompare(b.getName()); });

    const nombres = pdfs.map(function(p) { return p.getName(); });
    Logger.log("PDFs encontrados (" + pdfs.length + "): " + nombres.join(", "));
    return pdfs;
  } catch (e) {
    Logger.log("Error en buscarPDFs: " + e.message);
    return [];
  }
}

function syncGanadores(ganadores) {
  try {
    if (!ganadores || !ganadores.length) return { ok: true, mensaje: "Sin datos para sincronizar" };

    var ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
    var hoja = ss.getSheetByName("Backup Ganadores") || ss.insertSheet("Backup Ganadores");

    // Clear and rewrite headers + data
    hoja.clearContents();
    var headers = ["ID", "Mail", "Nombre", "DNI", "Show ID", "Show Nombre", "Show", "Venue", "Fecha Show", "Fecha Sorteo", "Estado"];
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);

    var rows = ganadores.map(function(g) {
      return [
        g.id || "",
        g.mail || "",
        g.nombre || "",
        g.dni || "",
        g.evId || "",
        g.evNombre || "",
        g.show || "",
        g.venue || "",
        g.fecha || "",
        g.fechaGano || "",
        g.estado || "pendiente"
      ];
    });

    if (rows.length > 0) {
      hoja.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    // Format header row
    hoja.getRange(1, 1, 1, headers.length)
      .setBackground("#0D1220")
      .setFontColor("#00D4FF")
      .setFontWeight("bold");

    Logger.log("syncGanadores: " + rows.length + " registros guardados");
    return { ok: true, mensaje: rows.length + " ganadores sincronizados en Sheets" };
  } catch(e) {
    Logger.log("Error en syncGanadores: " + e.message);
    return { ok: false, error: "Error al sincronizar: " + e.message };
  }
}

function trackingGanadores(ganadores, showNombre, fecha) {
  try {
    if (!ganadores || !ganadores.length) return { ok: true, mensaje: "Sin ganadores para agregar" };

    var ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
    var hoja = ss.getSheetByName("Tracking Ganadores");
    if (!hoja) return { ok: false, error: "No se encontro la pestana Tracking Ganadores" };

    // Build column header: "SHOW - FECHA"
    var header = showNombre + (fecha ? " - " + fecha : "");

    // NEVER touch columns A and B (they have formulas)
    // Always insert new show at column C, pushing existing data to the right
    // First check if this show already has a column (to avoid duplicates)
    var lastCol = hoja.getLastColumn();
    var existingCol = -1;

    if (lastCol >= 3) {
      // Check row 2 for show names
      var showHeaders = hoja.getRange(2, 3, 1, lastCol - 2).getValues()[0];
      for (var c = 0; c < showHeaders.length; c++) {
        if (String(showHeaders[c]).trim().toUpperCase() === showNombre.trim().toUpperCase()) {
          existingCol = c + 3;
          break;
        }
      }
    }

    var col;
    if (existingCol > 0) {
      // Show already exists — append to that column
      col = existingCol;
      Logger.log("Columna existente para " + showNombre + ": " + col);
    } else {
      // Insert new column at position C (column 3), pushing everything right
      hoja.insertColumnBefore(3);
      col = 3;
      // Fila 1: fecha, Fila 2: nombre del show
      hoja.getRange(1, col).setValue(fecha || "");
      hoja.getRange(2, col).setValue(showNombre.toUpperCase());
      hoja.getRange(1, col)
        .setBackground("#0D1220")
        .setFontColor("#00D4FF")
        .setFontWeight("bold")
        .setWrap(true);
      hoja.getRange(2, col)
        .setBackground("#0D1220")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setWrap(true);
      hoja.setColumnWidth(col, 160);
      Logger.log("Nueva columna insertada en C para: " + showNombre);
    }

    // Find first empty row in this column (starting from row 3 — row1=fecha, row2=show)
    var lastRow = hoja.getLastRow();
    var firstEmpty = 3;
    if (lastRow >= 3) {
      var colVals = hoja.getRange(3, col, lastRow - 2, 1).getValues();
      for (var r = 0; r < colVals.length; r++) {
        if (colVals[r][0] !== "") {
          firstEmpty = r + 4;
        } else {
          firstEmpty = r + 3;
          break;
        }
      }
    }

    // Write ganadores uppercase to match tracking format
    var names = ganadores.map(function(g) { return [(g.nombre || g.mail).toUpperCase()]; });
    hoja.getRange(firstEmpty, col, names.length, 1).setValues(names);

    Logger.log("trackingGanadores: " + ganadores.length + " agregados en col " + col + " fila " + firstEmpty);
    return { ok: true, mensaje: ganadores.length + " ganadores agregados en " + showNombre + " (columna C)" };
  } catch(e) {
    Logger.log("Error en trackingGanadores: " + e.message);
    return { ok: false, error: "Error: " + e.message };
  }
}

function leerTracking() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
    var hoja = ss.getSheetByName("Tracking Ganadores");
    if (!hoja) return { ok: false, error: "No se encontro la pestana Tracking Ganadores" };

    var lastCol = hoja.getLastColumn();
    var lastRow = hoja.getLastRow();

    if (lastCol < 3 || lastRow < 3) return { ok: true, columnas: [] };

    // Read all data from column C onwards
    var data = hoja.getRange(1, 3, lastRow, lastCol - 2).getValues();
    var columnas = [];

    for (var c = 0; c < data[0].length; c++) {
      // Format fecha - could be Date object or string
      var rawFecha = data[0][c];
      var fecha = "";
      if (rawFecha instanceof Date) {
        var d = rawFecha.getDate();
        var m = rawFecha.getMonth() + 1;
        var y = rawFecha.getFullYear();
        fecha = String(d).padStart(2,"0") + "/" + String(m).padStart(2,"0") + "/" + y;
      } else if (rawFecha) {
        fecha = String(rawFecha).trim();
        // Try to parse if it's a date string like "Sat Apr 18 2026..."
        var dt = new Date(fecha);
        if (!isNaN(dt) && dt.getFullYear() > 1990) {
          var d2 = dt.getDate();
          var m2 = dt.getMonth() + 1;
          var y2 = dt.getFullYear();
          fecha = String(d2).padStart(2,"0") + "/" + String(m2).padStart(2,"0") + "/" + y2;
        }
      }
      var show  = String(data[1][c] || "").trim();   // Row 2 = show name

      if (!show) continue; // Skip empty columns

      var nombres = [];
      for (var r = 2; r < data.length; r++) { // Row 3+ = ganadores
        var nombre = String(data[r][c] || "").trim();
        if (nombre) nombres.push(nombre);
      }

      if (nombres.length > 0) {
        columnas.push({ show: show, fecha: fecha, nombres: nombres });
      }
    }

    // Also get total employees from Hoja 1
    var totalEmpleados = 0;
    try {
      var hojaEmp = ss.getSheetByName(CONFIG.SHEET_EMPLEADOS_HOJA);
      if (hojaEmp) {
        var lastRowEmp = hojaEmp.getLastRow();
        // Count non-empty rows (minus header)
        totalEmpleados = lastRowEmp > 1 ? lastRowEmp - 1 : 0;
      }
    } catch(e2) {
      Logger.log("Error leyendo empleados: " + e2.message);
    }

    // Read base tickets from B123
    var ticketsBase = 0;
    try {
      var valB123 = hoja.getRange("B123").getValue();
      if (valB123 && !isNaN(Number(valB123))) ticketsBase = parseInt(valB123);
    } catch(e2) {
      Logger.log("Error leyendo B123: " + e2.message);
    }

    // Read canonical colaboradores list from columns A (names) and B (win count)
    var colaboradores = [];
    try {
      if (lastRow >= 3) {
        var abData = hoja.getRange(3, 1, lastRow - 2, 2).getValues();
        for (var r = 0; r < abData.length; r++) {
          var colNombre = String(abData[r][0] || "").trim();
          var colVictorias = parseInt(abData[r][1]) || 0;
          if (colNombre && colVictorias > 0) {
            colaboradores.push({ nombre: colNombre, victorias: colVictorias });
          }
        }
        colaboradores.sort(function(a, b) { return b.victorias - a.victorias; });
      }
    } catch(e3) {
      Logger.log("Error leyendo cols A-B: " + e3.message);
    }

    Logger.log("leerTracking: " + columnas.length + " columnas, " + colaboradores.length + " colaboradores con victorias, ticketsBase: " + ticketsBase);
    return { ok: true, columnas: columnas, colaboradores: colaboradores, totalEmpleados: totalEmpleados, ticketsBase: ticketsBase };
  } catch(e) {
    Logger.log("Error en leerTracking: " + e.message);
    return { ok: false, error: e.message };
  }
}

// Devuelve, para cada nombre en la lista, la última fecha en que ganó según el tracking sheet.
// Los nombres se comparan en UPPERCASE para tolerar diferencias de formato.
function getUltimasVictorias(nombres) {
  try {
    if (!nombres || !nombres.length) return { ok: true, data: [] };

    var ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
    var hoja = ss.getSheetByName("Tracking Ganadores");
    if (!hoja) return { ok: false, error: "No se encontró la pestaña 'Tracking Ganadores'" };

    var lastRow = hoja.getLastRow();
    var lastCol = hoja.getLastColumn();
    if (lastCol < 3 || lastRow < 3) return { ok: true, data: nombres.map(function(n){ return {nombre:n, ultimaVictoria:null}; }) };

    // Leer todo: fila 1 = fechas, fila 2 = shows, fila 3+ = ganadores
    var allData = hoja.getRange(1, 3, lastRow, lastCol - 2).getValues();
    var numCols = allData[0].length;

    // Mapa: NOMBRE_UPPER -> { nombre original, ultimaVictoria YYYY-MM-DD }
    var mapa = {};
    nombres.forEach(function(n) { mapa[n.trim().toUpperCase()] = { nombre: n, ultimaVictoria: null }; });

    for (var c = 0; c < numCols; c++) {
      var raw = allData[0][c];
      var fechaISO = "";
      if (raw instanceof Date) {
        var dd = String(raw.getDate()).padStart(2,"0");
        var mm = String(raw.getMonth()+1).padStart(2,"0");
        fechaISO = raw.getFullYear() + "-" + mm + "-" + dd;
      } else if (raw) {
        var parts = String(raw).trim().split("/");
        if (parts.length === 3) fechaISO = parts[2] + "-" + parts[1].padStart(2,"0") + "-" + parts[0].padStart(2,"0");
      }
      if (!fechaISO) continue;

      for (var r = 2; r < allData.length; r++) {
        var cell = String(allData[r][c] || "").trim().toUpperCase();
        if (!cell) continue;
        if (mapa[cell] !== undefined) {
          if (!mapa[cell].ultimaVictoria || fechaISO > mapa[cell].ultimaVictoria) {
            mapa[cell].ultimaVictoria = fechaISO;
          }
        }
      }
    }

    return { ok: true, data: Object.values(mapa) };
  } catch(e) {
    Logger.log("Error en getUltimasVictorias: " + e.message);
    return { ok: false, error: e.message };
  }
}

function testDrive() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Test Drive", "Ingresa el nombre exacto del show:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var nombre = resp.getResponseText().trim();
  var pdfs = buscarPDFs(nombre, nombre);
  var logs = Logger.getLog();
  ui.alert("Resultado: " + pdfs.length + " PDF(s)\n\n" + logs);
}

// DEBUG: ejecutar manualmente para probar envío con PDF
function debugEnviarUno() {
  var ui = SpreadsheetApp.getUi();

  // Paso 1: pedir show ID
  var r1 = ui.prompt("Debug Envío", "Show ID (ej: COLDPLAY3):", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var showId = r1.getResponseText().trim();

  // Paso 2: pedir mail de prueba
  var r2 = ui.prompt("Debug Envío", "Mail destinatario de prueba:", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var mailPrueba = r2.getResponseText().trim();

  // Buscar PDFs
  Logger.log("Buscando PDFs para: " + showId);
  var pdfs = buscarPDFs(showId, showId);
  Logger.log("PDFs encontrados: " + pdfs.length);

  if (!pdfs.length) {
    var logs = Logger.getLog();
    ui.alert("Sin PDFs encontrados.\n\n" + logs);
    return;
  }

  // Intentar enviar con el primer PDF
  try {
    GmailApp.sendEmail(
      mailPrueba,
      "TEST - Sorteo Arena - " + showId,
      "Este es un mail de prueba para verificar el adjunto de PDFs.",
      {
        name: "Sorteo Arena TEST",
        attachments: [pdfs[0].getAs(MimeType.PDF)]
      }
    );
    ui.alert("✓ Mail enviado a " + mailPrueba + " con PDF adjunto: " + pdfs[0].getName() + "\n\nTotal PDFs disponibles: " + pdfs.length);
  } catch(e) {
    ui.alert("Error al enviar: " + e.message + "\n\nLogs:\n" + Logger.getLog());
  }
}


