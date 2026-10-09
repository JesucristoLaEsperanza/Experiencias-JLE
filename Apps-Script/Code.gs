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
  Pagos: ['ID', 'InscripcionID', 'Fecha', 'Monto', 'Metodo', 'Comprobante', 'AprobadoPor', 'AvisoEnviado'],
  Tarifas: ['CampamentoID', 'Desde', 'Precio', 'Detalle']
};

function instruccionesPago(metodo) {
  if (metodo === 'Transferencia') return 'Elegiste transferencia bancaria. Enviá el pago desde tu banco o billetera al Banco de la Provincia de Buenos Aires. Titular: Natalia Ruiz. Alias: jle.jovenes. Luego compartí el comprobante respondiendo a este correo.';
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
    ['activados-2027', 'Campamento Activados 2027', '2027-01-08', '2027-01-10', 'Complejo TAM', 'Jóvenes', 12, 0, 0, 0, 0, 'Abierto'],
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
      c.getRange(sheetRow, 7, 1, 2).setValues([[fila[6], fila[7]]]); // Activados desde 12 años; 0 como EdadMax significa sin máximo.
      c.getRange(sheetRow, 11).setValue(0); // Sin límite de participantes
    }
  });
  sembrarTarifas_(ss);
  instalarLimpiezaAutomatica(ss);
  instalarAvisoPago_(ss);
}

function autorizarServicios() {
  const ss = obtenerHoja_();
  const cuotaMail = MailApp.getRemainingDailyQuota();
  return 'Permisos listos para la hoja ' + ss.getName() + '. Cuota de correo disponible: ' + cuotaMail;
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('inscripcion');
  template.isAppsScript = true;
  template.isEmbedded = Boolean(e && e.parameter && e.parameter.campamento);
  // Apps Script sirve el HTML dentro de userCodeAppPanel, cuya URL interna no
  // conserva los parámetros de /exec. Pasamos el ID desde la solicitud original.
  template.initialCampIdJson = jsonLiteral_(e && e.parameter && e.parameter.campamento ? String(e.parameter.campamento) : '');
  let resultadoInicial;
  try {
    resultadoInicial = { ok: true, campamentos: listarCampamentos() };
  } catch (error) {
    console.error('[doGet] No se pudieron cargar los campamentos: ' + (error.stack || error));
    resultadoInicial = { ok: false, campamentos: [], error: 'No se pudo leer la hoja Campamentos. Revisá que setup() se haya ejecutado y que las hojas Campamentos y Tarifas existan.' };
  }
  template.initialCampDataJson = jsonLiteral_(resultadoInicial);
  return template.evaluate()
    .setTitle('Experiencias JLE · Campamentos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonLiteral_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
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
        const rango = camp.edadMax ? 'tener entre ' + camp.edadMin + ' y ' + camp.edadMax + ' años' : 'tener al menos ' + camp.edadMin + ' años';
        return { ok: false, error: limpiar(p.nombre) + ' debe ' + rango + '.' };
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
    ins.getRange(r, 10).setFormula('=IF(L' + r + '="Pagado",I' + r + ',SUMIFS(Pagos!D:D,Pagos!B:B,A' + r + ',Pagos!G:G,"<>"))');
    ins.getRange(r, 11).setFormula('=I' + r + '-J' + r);
    ins.getRange(r, 15).setNumberFormat('dd/MM/yyyy');

    const participantRows = parts.map((p, i) => [id + '-' + (i + 1), id, limpiar(p.nombre), limpiar(p.apellido), limpiar(p.dni), p.fechaNac, limpiar(p.alimentacion), hayMenor ? 'Sí' : 'No aplica', '']);
    pa.getRange(pa.getLastRow() + 1, 1, participantRows.length, HOJAS.Participantes.length).setValues(participantRows);
    SpreadsheetApp.flush();

    // La inscripción ya está guardada. Un problema de correo no debe hacerla parecer fallida.
    const totalTexto = total > 0 ? moneda_(total) : (camp.detallePrecio || 'A confirmar');
    const venceTexto = Utilities.formatDate(vencePago, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
    const resumen = 'Inscripción ' + id + '\n' + camp.nombre + ' (' + camp.inicio + ' al ' + camp.fin + ')\nParticipantes: ' + parts.length +
      '\nForma de pago elegida: ' + metodo + '\nTotal: ' + totalTexto + '\nFecha límite de pago: ' + venceTexto +
      '\n\n' + instruccionesPago(metodo) + '\n\nPodés pagar en partes. Si el equipo registra y aprueba un primer pago parcial antes del vencimiento, la inscripción continúa sin fecha límite para el saldo; el precio puede actualizarse según el mes en que pagues. Si no se registra ningún pago en 14 días, se elimina la inscripción.\n\nWhatsApp del equipo: https://wa.me/5493364289558 o https://wa.me/5493364510510\n\nQue Dios los bendiga. — Jesucristo La Esperanza';
    try {
      MailApp.sendEmail(resp.email, 'Recibimos tu inscripción · ' + camp.nombre, resumen, { name: 'Jesucristo La Esperanza', replyTo: ADMIN_EMAIL });
      MailApp.sendEmail(ADMIN_EMAIL, 'Nueva inscripción ' + id, resp.nombre + ' (' + resp.email + ') inscribió a ' + parts.length + ' persona(s) en ' + camp.nombre + '.\nForma de pago: ' + metodo + '\nTotal: ' + totalTexto);
    } catch (mailErr) {
      console.error('[correo] Inscripción guardada ' + id + '; no se enviaron los avisos: ' + (mailErr.stack || mailErr));
    }

    return { ok: true, id: id, total: total, vencePago: venceTexto, detallePrecio: camp.detallePrecio || '' };
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
  const tarifas = leerTarifas_(ss);
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  return rows.map(r => {
    const vigente = precioVigente_(r[0], Number(r[8]) || 0, tarifas, new Date());
    return { id: r[0], nombre: r[1], inicio: fecha(r[2]), fin: fecha(r[3]), lugar: r[4], publico: r[5], edadMin: Number(r[6]) || 0, edadMax: Number(r[7]) || 0,
      precio: vigente.precio, detallePrecio: vigente.detalle, mesPrecio: vigente.mes, sena: Number(r[9]) || 0, estado: r[11] };
  }).filter(c => c.estado === 'Abierto' && c.inicio > hoy);
}

function sembrarTarifas_(ss) {
  const sh = ss.getSheetByName('Tarifas');
  const base = [
    ['activados-2027', '2026-10', 120000, ''],
    ['activados-2027', '2026-11', 135000, ''],
    ['activados-2027', '2026-12', 150000, ''],
    ['activados-2027', '2027-01', 0, 'En enero, consultá el precio con el equipo por WhatsApp: 336 428-9558 o 336 451-0510.'],
    ['jle-kids-2027', '2026-10', 75000, ''],
    ['jle-kids-2027', '2026-11', 75000, ''],
    ['jle-kids-2027', '2026-12', 75000, ''],
    ['jle-kids-2027', '2027-01', 100000, '']
  ];
  const existentes = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  base.forEach(f => {
    const indice = existentes.findIndex(r => {
      const mes = r[1] instanceof Date ? Utilities.formatDate(r[1], 'America/Argentina/Buenos_Aires', 'yyyy-MM') : String(r[1]).slice(0, 7);
      return String(r[0]) === f[0] && mes === f[1];
    });
    if (indice < 0) {
      sh.appendRow(f);
      return;
    }
    const fila = indice + 2;
    // Las tarifas acordadas prevalecen al ejecutar setup. Se conserva un precio
    // de enero de Activados si el equipo ya lo confirmó manualmente.
    if (f[0] === 'activados-2027' && f[1] === '2027-01') {
      if (!(Number(sh.getRange(fila, 3).getValue()) > 0)) sh.getRange(fila, 4).setValue(f[3]);
    } else {
      sh.getRange(fila, 3, 1, 2).setValues([[f[2], f[3]]]);
    }
  });
}

function leerTarifas_(ss) {
  const sh = ss.getSheetByName('Tarifas');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().filter(r => r[0] && r[1]).map(r => ({
    campamentoId: String(r[0]), mes: r[1] instanceof Date ? Utilities.formatDate(r[1], 'America/Argentina/Buenos_Aires', 'yyyy-MM') : String(r[1]).slice(0, 7),
    precio: Number(r[2]) || 0, detalle: String(r[3] || '')
  }));
}

function precioVigente_(campamentoId, precioBase, tarifas, fechaReferencia) {
  const mesActual = Utilities.formatDate(fechaReferencia || new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM');
  const aplicable = tarifas.filter(t => t.campamentoId === String(campamentoId) && t.mes <= mesActual).sort((a, b) => a.mes.localeCompare(b.mes));
  if (aplicable.length) {
    const tarifa = aplicable[aplicable.length - 1];
    return { precio: tarifa.precio, detalle: tarifa.detalle || (tarifa.precio ? '' : 'A confirmar'), mes: tarifa.mes };
  }
  return { precio: Number(precioBase) || 0, detalle: Number(precioBase) ? '' : 'A confirmar', mes: '' };
}

/** Crea un único disparador diario para quitar inscripciones que vencieron sin pago aprobado. */
function instalarLimpiezaAutomatica(ss) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'limpiarInscripcionesVencidas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('limpiarInscripcionesVencidas')
    .timeBased().everyDays(1).atHour(3).inTimezone('America/Argentina/Buenos_Aires').create();
}

function instalarAvisoPago_(ss) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alEditarPago')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEditarPago').forSpreadsheet(ss).onEdit().create();
}

/** Mantiene los saldos al precio vigente y comunica cambios mensuales por correo. */
function actualizarMontosInscripciones_(ss, avisarCambio) {
  const ins = ss.getSheetByName('Inscripciones');
  const campRows = ss.getSheetByName('Campamentos').getDataRange().getValues().slice(1).filter(r => r[0]);
  const camps = {};
  campRows.forEach(r => camps[String(r[0])] = { base: Number(r[8]) || 0 });
  const tarifas = leerTarifas_(ss);
  const pagados = sumarPagosAprobados_(ss);
  if (ins.getLastRow() < 2) return;
  const filas = ins.getRange(2, 1, ins.getLastRow() - 1, HOJAS.Inscripciones.length).getValues();
  filas.forEach((r, i) => {
    const id = String(r[0] || ''), campId = String(r[2] || ''), cantidad = Number(r[7]) || 0;
    if (!id || !camps[campId]) return;
    const estadoPrevio = String(r[11] || '').toLowerCase();
    if (estadoPrevio === 'cancelado') return;
    const tarifa = precioVigente_(campId, camps[campId].base, tarifas, new Date());
    const previo = Number(r[8]) || 0;
    const aprobado = pagados[id] || 0;
    const yaEstabaPagada = estadoPrevio === 'pagado' || estadoPrevio === 'pago completo';
    const marcadoPagado = aprobado === 0 && yaEstabaPagada;
    const yaSaldada = yaEstabaPagada;
    // Solo se congela el importe que ya estaba marcado como pagado; un pago nuevo
    // se compara con la tarifa del mes actual, aunque cubra el precio del mes anterior.
    const total = yaSaldada ? previo : Math.max(0, tarifa.precio * cantidad);
    const montoPagado = marcadoPagado ? total : aprobado;
    const saldo = Math.max(0, total - montoPagado);
    const importeConfirmado = tarifa.precio > 0 || !tarifa.detalle;
    const pagoCompleto = aprobado > 0 && importeConfirmado && saldo === 0;
    const estado = marcadoPagado || pagoCompleto ? 'Pagado' : (aprobado > 0 ? 'Pago parcial' : 'Pendiente');
    const sheetRow = i + 2;
    if (previo !== total) {
      ins.getRange(sheetRow, 9).setValue(total).setNumberFormat('$ #,##0');
      if (avisarCambio && r[4]) {
        try { enviarAvisoSaldo_(String(r[4]), String(r[3] || ''), id, total, montoPagado, saldo, tarifa.detalle, r[14], marcadoPagado || aprobado > 0); }
        catch (mailErr) { console.error('[aviso cambio de precio] ' + id + ': ' + (mailErr.stack || mailErr)); }
      }
    }
    ins.getRange(sheetRow, 10).setFormula('=IF(L' + sheetRow + '="Pagado",I' + sheetRow + ',SUMIFS(Pagos!D:D,Pagos!B:B,A' + sheetRow + ',Pagos!G:G,"<>"))');
    ins.getRange(sheetRow, 11).setFormula('=MAX(0,I' + sheetRow + '-J' + sheetRow + ')');
    ins.getRange(sheetRow, 12).setValue(estado);
    if (aprobado > 0 || marcadoPagado) ins.getRange(sheetRow, 15).clearContent();
  });
  SpreadsheetApp.flush();
}

function sumarPagosAprobados_(ss) {
  const sh = ss.getSheetByName('Pagos');
  const suma = {};
  if (!sh || sh.getLastRow() < 2) return suma;
  sh.getRange(2, 1, sh.getLastRow() - 1, HOJAS.Pagos.length).getValues().forEach(p => {
    const id = String(p[1] || '');
    if (id && String(p[6] || '').trim()) suma[id] = (suma[id] || 0) + (Number(p[3]) || 0);
  });
  return suma;
}

function enviarAvisoSaldo_(email, nombre, id, total, pagado, saldo, detalle, vence, hayPago) {
  if (!email || !/^\S+@\S+\.\S+$/.test(email)) return;
  const textoTotal = total ? moneda_(total) : (detalle || 'A confirmar');
  const textoSaldo = total ? (saldo ? moneda_(saldo) : '$0') : (detalle ? 'A confirmar' : '$0');
  const textoPagado = moneda_(pagado);
  const estadoPlazo = hayPago
    ? 'Como ya registramos un pago parcial, el saldo no tiene fecha límite. El precio puede actualizarse según el mes en que pagues.'
    : 'Tenés hasta ' + (vence instanceof Date ? Utilities.formatDate(vence, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy') : '14 días desde tu inscripción') + ' para registrar el primer pago; después la inscripción se elimina.';
  const cuerpo = 'Hola ' + (nombre || '') + ':\n\nActualizamos tu inscripción ' + id + '.\nPrecio vigente: ' + textoTotal + '\nPagado y aprobado: ' + textoPagado + '\nSaldo pendiente: ' + textoSaldo + '\n\n' + estadoPlazo +
    '\n\nSi necesitás consultar o enviar el comprobante por WhatsApp, escribinos: https://wa.me/5493364289558 o https://wa.me/5493364510510';
  MailApp.sendEmail(email, 'Actualización de saldo · Experiencias JLE', cuerpo, { name: 'Jesucristo La Esperanza', replyTo: ADMIN_EMAIL });
}

/** Registra un pago aprobado en la hoja Pagos; completar A:F y escribir AprobadoPor al final. */
function alEditarPago(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Pagos' || e.range.getLastColumn() < 7 || e.range.getColumn() > 7) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const ss = obtenerHoja_();
    const inicio = Math.max(2, e.range.getRow()), fin = e.range.getLastRow();
    for (let row = inicio; row <= fin; row++) {
      const p = sh.getRange(row, 1, 1, HOJAS.Pagos.length).getValues()[0];
      if (!String(p[6] || '').trim() || p[7] || !p[1] || !p[2] || !(Number(p[3]) > 0) || !p[4]) continue;
      if (!p[0]) { p[0] = 'PAY-' + Utilities.getUuid().slice(0, 8).toUpperCase(); sh.getRange(row, 1).setValue(p[0]); }
      actualizarMontosInscripciones_(ss, false);
      const ins = ss.getSheetByName('Inscripciones');
      const idx = ins.getLastRow() > 1 ? ins.getRange(2, 1, ins.getLastRow() - 1, 1).getValues().findIndex(r => String(r[0]) === String(p[1])) : -1;
      if (idx < 0) continue;
      const insRow = idx + 2;
      const vals = ins.getRange(insRow, 1, 1, HOJAS.Inscripciones.length).getValues()[0];
      const suma = sumarPagosAprobados_(ss), pagado = suma[String(vals[0])] || 0;
      const total = Number(vals[8]) || 0, saldo = Math.max(0, total - pagado);
      const camp = ss.getSheetByName('Campamentos').getDataRange().getValues().slice(1).find(c => String(c[0]) === String(vals[2]));
      const tarifa = camp ? precioVigente_(camp[0], Number(camp[8]) || 0, leerTarifas_(ss), new Date()) : { detalle: '', precio: total };
      const importeConfirmado = Number(tarifa.precio) > 0 || !tarifa.detalle;
      ins.getRange(insRow, 10).setFormula('=IF(L' + insRow + '="Pagado",I' + insRow + ',SUMIFS(Pagos!D:D,Pagos!B:B,A' + insRow + ',Pagos!G:G,"<>"))');
      ins.getRange(insRow, 11).setFormula('=MAX(0,I' + insRow + '-J' + insRow + ')');
      ins.getRange(insRow, 12).setValue(saldo === 0 && importeConfirmado ? 'Pagado' : 'Pago parcial');
      ins.getRange(insRow, 15).clearContent(); // Un pago parcial confirmado quita el vencimiento de 14 días.
      const email = String(vals[4] || '');
      try {
        enviarAvisoSaldo_(email, String(vals[3] || ''), String(vals[0]), total, pagado, saldo, tarifa.detalle, '', true);
        sh.getRange(row, 8).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');
      } catch (mailErr) { console.error('[aviso pago] ' + p[0] + ': ' + (mailErr.stack || mailErr)); }
    }
  } finally { lock.releaseLock(); }
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

    actualizarMontosInscripciones_(ss, true);

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
      // Cualquier pago parcial aprobado conserva la inscripción y quita el vencimiento.
      if (estado === 'pagado' || estado === 'pago completo' || pagoAprobado > 0) return;

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

function moneda_(valor) { return '$' + Math.round(Number(valor) || 0).toLocaleString('es-AR'); }

function fecha(v) { return v instanceof Date ? Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : String(v); }
function edadEn(nac, ref) {
  const n = String(nac).slice(0, 10).split('-').map(Number);
  const r = String(ref).slice(0, 10).split('-').map(Number);
  let edad = r[0] - n[0];
  if (r[1] < n[1] || (r[1] === n[1] && r[2] < n[2])) edad--;
  return edad;
}
function limpiar(v) { const s = String(v == null ? '' : v).trim().slice(0, 150); return /^[=+\-@]/.test(s) ? "'" + s : s; } // evita fórmulas inyectadas

