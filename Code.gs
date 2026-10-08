/**
 * Inscripciones a campamentos - Jesucristo La Esperanza (Fase 1)
 * Pegar en: Google Sheets > Extensiones > Apps Script.
 * 1) Ejecutar setup() una vez. 2) Implementar > Aplicación web (Ejecutar como: yo / Acceso: cualquiera).
 */
const ADMIN_EMAIL = 'iglesiajle33@gmail.com';
const METODOS_PAGO = ['Transferencia', 'Efectivo en la iglesia'];

const HOJAS = {
  Campamentos: ['ID', 'Nombre', 'Inicio', 'Fin', 'Lugar', 'Publico', 'EdadMin', 'EdadMax', 'PrecioTotal', 'Sena', 'Cupo', 'Estado'],
  Inscripciones: ['ID', 'Fecha', 'CampamentoID', 'Responsable', 'Email', 'Telefono', 'Lider', 'Cantidad', 'Total', 'Pagado', 'Saldo', 'Estado', 'Metodo', 'Notas'],
  Participantes: ['ID', 'InscripcionID', 'Nombre', 'Apellido', 'DNI', 'FechaNac', 'Alimentacion', 'AutorizacionTutor', 'CheckIn'],
  Pagos: ['ID', 'InscripcionID', 'Fecha', 'Monto', 'Metodo', 'Comprobante', 'AprobadoPor']
};

function instruccionesPago(metodo) {
  if (metodo === 'Transferencia') return 'Elegiste transferencia bancaria. El equipo te va a enviar por correo el alias/CBU y cómo compartir el comprobante. Podés responder a este correo si tenés alguna duda.';
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
  const c = ss.getSheetByName('Campamentos');
  const idsExistentes = c.getLastRow() > 1 ? c.getRange(2, 1, c.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const iniciales = [
    ['activados-2027', 'Campamento Activados 2027', '2027-01-08', '2027-01-10', 'Complejo TAM', 'Jóvenes', 12, 30, 0, 0, 100, 'Abierto'],
    ['jle-kids-2027', 'Campamento JLE Kids', '2027-01-27', '2027-01-28', 'A confirmar', 'Niños', 5, 12, 0, 0, 50, 'Abierto']
  ];
  iniciales.forEach(fila => {
    if (idsExistentes.indexOf(fila[0]) < 0) c.appendRow(fila);
  });
}

function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'campamentos') {
    return json({ ok: true, campamentos: listarCampamentos() });
  }
  const template = HtmlService.createTemplateFromFile('inscripcion');
  template.isAppsScript = true;
  return template.evaluate().setTitle('Experiencias JLE · Campamentos');
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

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: 'La solicitud llegó vacía.' });
    }
    return json(guardarInscripcion_(JSON.parse(e.postData.contents)));
  } catch (err) {
    console.error('[doPost] ' + (err.stack || err));
    return json({ ok: false, error: 'No pudimos procesar la inscripción.' });
  }
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
    if (parts.length > camp.disponibles) return { ok: false, error: 'Quedan ' + camp.disponibles + ' lugares disponibles.' };

    const inicio = new Date(camp.inicio);
    let hayMenor = false;
    for (const p of parts) {
      if (!limpiar(p.nombre) || !limpiar(p.apellido) || !limpiar(p.dni) || !p.fechaNac) return { ok: false, error: 'Completá nombre, apellido, DNI y fecha de nacimiento de cada participante.' };
      const edad = edadEn(p.fechaNac, inicio);
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
    ins.appendRow([id, new Date(), camp.id, resp.nombre, resp.email, resp.tel, limpiar(d.lider), parts.length, total, '', '', 'Pendiente', metodo, limpiar(d.notas)]);
    const r = ins.getLastRow();
    ins.getRange(r, 10).setFormula('=SUMIF(Pagos!B:B,A' + r + ',Pagos!D:D)');
    ins.getRange(r, 11).setFormula('=I' + r + '-J' + r);

    const participantRows = parts.map((p, i) => [id + '-' + (i + 1), id, limpiar(p.nombre), limpiar(p.apellido), limpiar(p.dni), p.fechaNac, limpiar(p.alimentacion), hayMenor ? 'Sí' : 'No aplica', '']);
    pa.getRange(pa.getLastRow() + 1, 1, participantRows.length, HOJAS.Participantes.length).setValues(participantRows);
    SpreadsheetApp.flush();

    // La inscripción ya está guardada. Un problema de correo no debe hacerla parecer fallida.
    const totalTexto = total > 0 ? '$' + total : 'A confirmar';
    const resumen = 'Inscripción ' + id + '\n' + camp.nombre + ' (' + camp.inicio + ' al ' + camp.fin + ')\nParticipantes: ' + parts.length +
      '\nForma de pago elegida: ' + metodo + '\nTotal: ' + totalTexto + '\n\n' + instruccionesPago(metodo) + '\n\nQue Dios los bendiga. — Jesucristo La Esperanza';
    try {
      MailApp.sendEmail(resp.email, 'Recibimos tu inscripción · ' + camp.nombre, resumen, { name: 'Jesucristo La Esperanza', replyTo: ADMIN_EMAIL });
      MailApp.sendEmail(ADMIN_EMAIL, 'Nueva inscripción ' + id, resp.nombre + ' (' + resp.email + ') inscribió a ' + parts.length + ' persona(s) en ' + camp.nombre + '.\nForma de pago: ' + metodo + '\nTotal: ' + totalTexto);
    } catch (mailErr) {
      console.error('[correo] Inscripción guardada ' + id + '; no se enviaron los avisos: ' + (mailErr.stack || mailErr));
    }

    return { ok: true, id: id, total: total };
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
  const ins = ss.getSheetByName('Inscripciones').getDataRange().getValues().slice(1);
  return rows.map(r => {
    const usados = ins.filter(i => i[2] === r[0] && i[11] !== 'Cancelado').reduce((s, i) => s + Number(i[7] || 0), 0);
    return { id: r[0], nombre: r[1], inicio: fecha(r[2]), fin: fecha(r[3]), lugar: r[4], publico: r[5], edadMin: Number(r[6]) || 0, edadMax: Number(r[7]) || 0,
      precio: Number(r[8]) || 0, sena: Number(r[9]) || 0, cupo: Number(r[10]) || 0, disponibles: Math.max(0, (Number(r[10]) || 0) - usados), estado: r[11] };
  });
}

function fecha(v) { return v instanceof Date ? Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : String(v); }
function edadEn(nac, ref) { const n = new Date(nac); let e = ref.getFullYear() - n.getFullYear(); if (ref < new Date(ref.getFullYear(), n.getMonth(), n.getDate())) e--; return e; }
function limpiar(v) { const s = String(v == null ? '' : v).trim().slice(0, 150); return /^[=+\-@]/.test(s) ? "'" + s : s; } // evita fórmulas inyectadas
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

