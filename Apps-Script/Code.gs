/**
 * Inscripciones a campamentos - Jesucristo La Esperanza (Fase 1)
 * Pegar en: Google Sheets > Extensiones > Apps Script.
 * 1) Ejecutar setup() desde el editor y autorizar los permisos. 2) Implementar como aplicación web.
 */
const ADMIN_EMAIL = 'iglesiajle33@gmail.com';
const METODOS_PAGO = ['Transferencia', 'Efectivo en la iglesia'];

const HOJAS = {
  Campamentos: ['ID', 'Nombre', 'Inicio', 'Fin', 'Lugar', 'Publico', 'EdadMin', 'EdadMax', 'PrecioTotal', 'Sena', 'Cupo', 'Estado'],
  Inscripciones: ['ID', 'Fecha', 'CampamentoID', 'Responsable', 'Email', 'Telefono', 'Lider', 'Cantidad', 'Total', 'Pagado', 'Saldo', 'Estado', 'Metodo', 'Notas', 'VencePago'],
  Participantes: ['ID', 'InscripcionID', 'Nombre', 'Apellido', 'DNI', 'FechaNac', 'Alimentacion', 'AutorizacionTutor', 'CheckIn'],
  Pagos: ['ID', 'InscripcionID', 'Fecha', 'Monto', 'Metodo', 'Comprobante', 'AprobadoPor']
};

function instruccionesPago(metodo) {
  if (metodo === 'Transferencia') return 'Elegiste transferencia bancaria. Enviá el pago desde tu banco o billetera a esta cuenta de Cuenta DNI (Caja de ahorros en pesos). Titular: Lucas Samuel Castells. Alias: castells06. CBU: 0140304403659961460900. Luego compartí el comprobante respondiendo a este correo.';
  return 'Elegiste efectivo en la iglesia. El equipo te va a contactar para coordinar cuándo completar el pago.';
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Abrí este proyecto desde la hoja de Google Sheets antes de ejecutar setup().');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  Object.keys(HOJAS).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    sh.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const inscripciones = ss.getSheetByName('Inscripciones');
  inscripciones.getRange(2, 15, Math.max(inscripciones.getMaxRows() - 1, 1), 1).setNumberFormat('dd/MM/yyyy');
  const c = ss.getSheetByName('Campamentos');
  const idsExistentes = c.getLastRow() > 1 ? c.getRange(2, 1, c.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const iniciales = [
    ['activados-2027', 'Campamento Activados 2027', '2027-01-08', '2027-01-10', 'Complejo TAM', 'Jóvenes', 0, 0, 0, 0, 0, 'Abierto'],
    ['jle-kids-2027', 'Campamento JLE Kids', '2027-01-27', '2027-01-28', 'Complejo TAM', 'Niños', 5, 12, 0, 0, 0, 'Abierto']
  ];
  iniciales.forEach(fila => {
    const rowIndex = idsExistentes.indexOf(fila[0]);
    if (rowIndex < 0) {
      c.appendRow(fila);
    } else {
      // Actualiza solo los datos corregidos; conserva fechas, precios y estado que ya editó el equipo.
      const sheetRow = rowIndex + 2;
      c.getRange(sheetRow, 5).setValue(fila[4]); // Lugar
      c.getRange(sheetRow, 7, 1, 2).setValues([[fila[6], fila[7]]]); // Edades; 0/0 significa sin límite
      c.getRange(sheetRow, 11).setValue(0); // Sin límite de participantes
    }
  });
  instalarLimpiezaAutomatica();
}

function autorizarServicios() {
  const ss = obtenerHoja_();
  const cuotaMail = MailApp.getRemainingDailyQuota();
  return 'Permisos listos para la hoja ' + ss.getName() + '. Cuota de correo disponible: ' + cuotaMail;
}

function doGet() {
  const template = HtmlService.createTemplateFromFile('inscripcion');
  template.isAppsScript = true;
  return template.evaluate()
    .setTitle('Experiencias JLE · Campamentos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function obtenerCampamentosWeb() {
  return { ok: true, campamentos: listarCampamentos() };
}

function registrarInscripcionWeb(datos) {
  return guardarInscripcion_(datos);
}



function guardarInscripcion_(d) {
  const lock = LockService.getScriptLock();
  let lockTomado = false;
  try {
    lock.waitLock(20000);
    lockTomado = true;
    if (!d || typeof d !== 'object') return { ok: false, error: 'Solicitud no válida.' };
    if (d.web) return { ok: false, error: 'Solicitud no válida.' }; // honeypot anti-spam

    const camp = listarCampamentos().find(c => c.id === d.campamentoId);
    if (!camp || camp.estado !== 'Abierto') return { ok: false, error: 'Este campamento no está abierto.' };

    const resp = { nombre: limpiar(d.nombre), email: limpiar(d.email), tel: limpiar(d.telefono) };
    const metodo = limpiar(d.metodo);
    if (!METODOS_PAGO.includes(metodo)) return { ok: false, error: 'Elegí una forma de pago válida.' };
    const parts = Array.isArray(d.participantes) ? d.participantes.slice(0, 10) : [];
    if (!resp.nombre || !/^\S+@\S+\.\S+$/.test(resp.email) || !resp.tel) return { ok: false, error: 'Revisá los datos del responsable.' };
    if (!parts.length) return { ok: false, error: 'Agregá al menos un participante.' };

    let hayMenor = false;
    for (const p of parts) {
      if (!limpiar(p.nombre) || !limpiar(p.apellido) || !limpiar(p.dni) || !p.fechaNac) return { ok: false, error: 'Completá nombre, apellido, DNI y fecha de nacimiento de cada participante.' };
      const edad = edadEn(p.fechaNac, camp.inicio);
      if (edad < 18) hayMenor = true;
      if ((camp.edadMin && edad < camp.edadMin) || (camp.edadMax && edad > camp.edadMax)) {
        return { ok: false, error: limpiar(p.nombre) + ' no está en el rango de edad (' + camp.edadMin + ' a ' + camp.edadMax + ' años).' };
      }
    }
    if (hayMenor && !d.autorizacion) return { ok: false, error: 'Se necesita la autorización del padre, madre o tutor.' };

    const ss = obtenerHoja_();
    const ins = ss.getSheetByName('Inscripciones');
    const pa = ss.getSheetByName('Participantes');
    if (!ins || !pa) throw new Error('Faltan las hojas Inscripciones o Participantes. Ejecutá setup().');

    const id = 'INS-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const total = camp.precio * parts.length;
    const vencePago = fechaMasDiasBuenosAires_(new Date(), 14);
    ins.appendRow([id, new Date(), camp.id, resp.nombre, resp.email, resp.tel, limpiar(d.lider), parts.length, total, '', '', 'Pendiente', metodo, limpiar(d.notas), vencePago]);
    const r = ins.getLastRow();
    ins.getRange(r, 10).setFormula('=SUMIFS(Pagos!D:D,Pagos!B:B,A' + r + ',Pagos!G:G,"<>")');
    ins.getRange(r, 11).setFormula('=I' + r + '-J' + r);
    ins.getRange(r, 15).setNumberFormat('dd/MM/yyyy');

    const participantRows = parts.map((p, i) => [id + '-' + (i + 1), id, limpiar(p.nombre), limpiar(p.apellido), limpiar(p.dni), p.fechaNac, limpiar(p.alimentacion), hayMenor ? 'Sí' : 'No aplica', '']);
    pa.getRange(pa.getLastRow() + 1, 1, participantRows.length, HOJAS.Participantes.length).setValues(participantRows);
    SpreadsheetApp.flush();

    // La inscripción ya está guardada. Un problema de correo no debe hacerla parecer fallida.
    const totalTexto = total > 0 ? '$' + total : 'A confirmar';
    const venceTexto = Utilities.formatDate(vencePago, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
    const resumen = 'Inscripción ' + id + '\n' + camp.nombre + ' (' + camp.inicio + ' al ' + camp.fin + ')\nParticipantes: ' + parts.length +
      '\nForma de pago elegida: ' + metodo + '\nTotal: ' + totalTexto + '\nFecha límite de pago: ' + venceTexto +
      '\n\n' + instruccionesPago(metodo) + '\n\nSi el pago no queda registrado antes de la fecha límite, la inscripción y sus datos asociados se eliminarán automáticamente.\n\nQue Dios los bendiga. — Jesucristo La Esperanza';
    try {
      MailApp.sendEmail(resp.email, 'Recibimos tu inscripción · ' + camp.nombre, resumen, { name: 'Jesucristo La Esperanza', replyTo: ADMIN_EMAIL });
      MailApp.sendEmail(ADMIN_EMAIL, 'Nueva inscripción ' + id, resp.nombre + ' (' + resp.email + ') inscribió a ' + parts.length + ' persona(s) en ' + camp.nombre + '.\nForma de pago: ' + metodo + '\nTotal: ' + totalTexto);
    } catch (mailErr) {
      console.error('[correo] Inscripción guardada ' + id + '; no se enviaron los avisos: ' + (mailErr.stack || mailErr));
    }

    return { ok: true, id: id, total: total, vencePago: venceTexto };
  } catch (err) {
    console.error('[guardarInscripcion_] ' + (err.stack || err));
    return { ok: false, error: 'No pudimos guardar la inscripción en Google Sheets. Revisá que setup() se haya ejecutado y que las hojas estén creadas.' };
  } finally {
    if (lockTomado) lock.releaseLock();
  }
}

function obtenerHoja_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const activa = SpreadsheetApp.getActive();
  if (activa) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', activa.getId());
    return activa;
  }
  throw new Error('Ejecutá setup() desde el proyecto vinculado a Google Sheets antes de usar la Web App.');
}

function listarCampamentos() {
  const ss = obtenerHoja_();
  const rows = ss.getSheetByName('Campamentos').getDataRange().getValues().slice(1).filter(r => r[0]);
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  return rows.map(r => {
    return { id: r[0], nombre: r[1], inicio: fecha(r[2]), fin: fecha(r[3]), lugar: r[4], publico: r[5], edadMin: Number(r[6]) || 0, edadMax: Number(r[7]) || 0,
      precio: Number(r[8]) || 0, sena: Number(r[9]) || 0, estado: r[11] };
  }).filter(c => c.estado === 'Abierto' && c.inicio > hoy);
}

/** Crea un único disparador diario para quitar inscripciones que vencieron sin pago aprobado. */
function instalarLimpiezaAutomatica() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'limpiarInscripcionesVencidas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('limpiarInscripcionesVencidas')
    .timeBased().everyDays(1).atHour(3).inTimezone('America/Argentina/Buenos_Aires').create();
}

/** Elimina inscripciones impagas vencidas y sus participantes y pagos relacionados. */
function limpiarInscripcionesVencidas() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = obtenerHoja_();
    const ins = ss.getSheetByName('Inscripciones');
    const pa = ss.getSheetByName('Participantes');
    const pagos = ss.getSheetByName('Pagos');
    if (!ins || !pa || !pagos || ins.getLastRow() < 2) return 0;

    const tz = 'America/Argentina/Buenos_Aires';
    const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const rows = ins.getRange(2, 1, ins.getLastRow() - 1, HOJAS.Inscripciones.length).getValues();
    const pagosRows = pagos.getLastRow() > 1 ? pagos.getRange(2, 1, pagos.getLastRow() - 1, HOJAS.Pagos.length).getValues() : [];
    const aprobados = {};
    pagosRows.forEach(p => {
      const id = String(p[1] || '');
      if (id && String(p[6] || '').trim()) aprobados[id] = (aprobados[id] || 0) + Number(p[3] || 0);
    });

    const vencidas = [];
    rows.forEach((r, index) => {
      const id = String(r[0] || '');
      if (!id) return;
      const estado = String(r[11] || '').toLowerCase();
      const total = Number(r[8]) || 0;
      const pagoAprobado = aprobados[id] || 0;
      const pagada = estado === 'pagado' || estado === 'pago completo' || (pagoAprobado > 0 && pagoAprobado >= total);
      if (pagada) return;

      let vence = r[14] instanceof Date ? r[14] : null;
      if (!vence) {
        const fechaAlta = r[1] instanceof Date ? new Date(r[1]) : new Date(r[1]);
        if (isNaN(fechaAlta.getTime())) return;
        vence = fechaMasDiasBuenosAires_(fechaAlta, 14);
      }
      // Comparar días en Buenos Aires para respetar las dos semanas completas.
      if (Utilities.formatDate(vence, tz, 'yyyy-MM-dd') < hoy) {
        vencidas.push({ row: index + 2, id: id, email: String(r[4] || ''), responsable: String(r[3] || ''), campamentoId: String(r[2] || '') });
      }
    });
    if (!vencidas.length) return 0;

    const ids = new Set(vencidas.map(v => v.id));
    borrarFilasRelacionadas_(pa, 2, ids);
    borrarFilasRelacionadas_(pagos, 2, ids);
    vencidas.map(v => v.row).sort((a, b) => b - a).forEach(row => ins.deleteRow(row));
    vencidas.forEach(v => {
      if (!v.email || !/^\S+@\S+\.\S+$/.test(v.email)) return;
      try {
        MailApp.sendEmail(v.email, 'Inscripción vencida · Experiencias JLE',
          'Hola ' + v.responsable + ': venció el plazo de 14 días para completar el pago de tu inscripción ' + v.id + '. La inscripción fue eliminada junto con sus datos asociados. Si todavía querés participar, podés volver a inscribirte mientras el evento siga abierto.',
          { name: 'Jesucristo La Esperanza', replyTo: ADMIN_EMAIL });
      } catch (mailErr) { console.error('[correo vencimiento] ' + v.id + ': ' + mailErr); }
    });
    return vencidas.length;
  } finally {
    lock.releaseLock();
  }
}

function borrarFilasRelacionadas_(sheet, idColumn, ids) {
  if (sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, idColumn, sheet.getLastRow() - 1, 1).getValues();
  const rows = [];
  values.forEach((r, i) => { if (ids.has(String(r[0] || ''))) rows.push(i + 2); });
  rows.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));
}

function fechaMasDiasBuenosAires_(fechaBase, dias) {
  const tz = 'America/Argentina/Buenos_Aires';
  const partes = Utilities.formatDate(fechaBase, tz, 'yyyy-MM-dd').split('-').map(Number);
  // Guardar al mediodía UTC mantiene el mismo día calendario de Argentina al leer la celda.
  return new Date(Date.UTC(partes[0], partes[1] - 1, partes[2] + dias, 12, 0, 0));
}

function fecha(v) { return v instanceof Date ? Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : String(v); }
function edadEn(nac, ref) {
  const n = String(nac).slice(0, 10).split('-').map(Number);
  const r = String(ref).slice(0, 10).split('-').map(Number);
  let edad = r[0] - n[0];
  if (r[1] < n[1] || (r[1] === n[1] && r[2] < n[2])) edad--;
  return edad;
}
function limpiar(v) { const s = String(v == null ? '' : v).trim().slice(0, 150); return /^[=+\-@]/.test(s) ? "'" + s : s; } // evita fórmulas inyectadas

