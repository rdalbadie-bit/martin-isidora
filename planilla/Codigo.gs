/** @OnlyCurrentDoc */
/**
 * Invitación de Isidora y Martín: lista de invitados y confirmaciones.
 *
 * Hojas:
 *  - Invitados: una fila por invitación. Al escribir el nombre se crea su código.
 *    La planilla arma el link personal y el botón "Enviar" de WhatsApp.
 *  - Confirmaciones: llegan solas desde el formulario de la invitación.
 *  - Mensaje: el texto del WhatsApp (se puede editar).
 *  - Resumen: totales, alimentación y canciones.
 *
 * La invitación usa esta planilla como "aplicación web":
 *  - doGet(?id=código)  -> nombre, saludo y cupos del invitado (para personalizar el sobre).
 *  - doPost(respuesta)  -> guarda la confirmación.
 */

const LINK_BASE = 'https://isidoraymartin.cl/?i=';

const H_INV = ['Código', 'Nombre en el sobre', 'Saludo', 'Cupos', 'Teléfono', 'Link personal',
  'Enviar por WhatsApp', 'Respondió', 'Asiste', 'Personas confirmadas'];
const H_CONF = ['Fecha', 'Código', 'Invitado (lista)', 'Asiste', 'Nombre completo', 'Alimentación',
  'Pareja', 'Alimentación pareja', 'Canción infaltable'];

const MSJ_1 = '{saludo}: ¡nos casamos! Martín y yo queremos invitarte a celebrar con nosotros el sábado ' +
  '10 de abril de 2027 en Parque Lo Arcaya, Pirque. Abre tu invitación y confírmanos aquí: {link}';
const MSJ_2 = '{saludo}: ¡nos casamos! Martín y yo queremos invitarlos a celebrar con nosotros el sábado ' +
  '10 de abril de 2027 en Parque Lo Arcaya, Pirque. Abran su invitación y confírmennos aquí: {link}';

/* ---------- Menú ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invitación')
    .addItem('Crear códigos que falten', 'codigosFaltantes')
    .addToUi();
}

/* ---------- Armar la planilla (se corre una vez) ---------- */
function configurar() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Santiago');
  ss.setSpreadsheetLocale('en_US');   // las fórmulas de abajo usan comas como separador

  const inv = hoja_(ss, 'Invitados', H_INV);
  const conf = hoja_(ss, 'Confirmaciones', H_CONF);
  const men = ss.getSheetByName('Mensaje') || ss.insertSheet('Mensaje');
  const res = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen');
  ss.getSheets().forEach(s => {
    if (['Invitados', 'Confirmaciones', 'Mensaje', 'Resumen'].indexOf(s.getName()) < 0 && s.getLastRow() === 0) ss.deleteSheet(s);
  });

  // Mensaje de WhatsApp (editable)
  men.getRange('A1:B3').setValues([
    ['Mensaje (1 persona)', MSJ_1],
    ['Mensaje (2 o más)', MSJ_2],
    ['Cómo funciona', '{saludo} se cambia por la columna "Saludo" de Invitados (si está vacía: "Hola, " y el nombre). ' +
      '{link} se cambia por el link personal. Se usa el mensaje 2 cuando "Cupos" es 2 o más.']
  ]);
  men.getRange('A1:A3').setFontWeight('bold');
  men.setColumnWidth(1, 170); men.setColumnWidth(2, 640);
  men.getRange('B1:B3').setWrap(true);

  // Invitados: link personal y botón de WhatsApp
  inv.getRange('F1').setFormula('={"Link personal"; ARRAYFORMULA(IF(A2:A="",, "' + LINK_BASE + '"&A2:A))}');
  inv.getRange('G1').setFormula(
    '={"Enviar por WhatsApp"; ARRAYFORMULA(IF((A2:A="")+(E2:E=""),, HYPERLINK(' +
    '"https://wa.me/"&IF(LEN(REGEXREPLACE(TO_TEXT(E2:E),"[^0-9]",""))=9,"56","")&REGEXREPLACE(TO_TEXT(E2:E),"[^0-9]","")' +
    '&"?text="&ENCODEURL(SUBSTITUTE(SUBSTITUTE(IF(D2:D>1,Mensaje!$B$2,Mensaje!$B$1),"{saludo}",IF(C2:C="","Hola, "&B2:B,C2:C)),"{link}",F2:F)),' +
    '"Enviar")))}');
  inv.getRange('E2:E').setNumberFormat('@');
  inv.getRange('H2:H').setNumberFormat('dd-mm-yyyy hh:mm');
  inv.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 20)
    .setHelpText('Cuántas personas incluye esta invitación').build());
  [80, 190, 200, 60, 130, 330, 90, 120, 60, 100].forEach((w, i) => inv.setColumnWidth(i + 1, w));
  inv.getRange('A1:J1').setBackground('#6F1D2B').setFontColor('#ffffff');
  inv.getRange('A1:E1').setNote(null);
  inv.getRange('B1').setNote('Escribe aquí el nombre como irá en el sobre (ej.: Familia Pérez). El código se crea solo.');
  inv.getRange('C1').setNote('Opcional. Ej.: Querida tía Carmen. Si queda vacío, el mensaje dice "Hola, " y el nombre.');
  inv.getRange('E1').setNote('Con código de país, ej.: +56 9 1234 5678. Si pones solo 9 dígitos, se agrega el 56.');

  if (inv.getLastRow() < 2) {
    inv.getRange('B2:E3').setValues([
      ['Familia Pérez (ejemplo)', 'Querida familia Pérez', 4, ''],
      ['Tía Carmen (ejemplo)', 'Querida tía Carmen', 1, '']
    ]);
    codigosFaltantes();
  }

  // Confirmaciones
  conf.getRange('A2:A').setNumberFormat('dd-mm-yyyy hh:mm');
  [120, 70, 170, 60, 190, 150, 190, 150, 220].forEach((w, i) => conf.setColumnWidth(i + 1, w));
  conf.getRange('A1:I1').setBackground('#7B7D58').setFontColor('#ffffff');

  // Resumen
  res.clear();
  const dieta = t => '=COUNTIFS(Confirmaciones!D2:D,"Sí",Confirmaciones!F2:F,"' + t + '")+COUNTIFS(Confirmaciones!D2:D,"Sí",Confirmaciones!H2:H,"' + t + '")';
  res.getRange('A1:B14').setValues([
    ['Invitaciones', '=COUNTA(Invitados!B2:B)'],
    ['Lugares reservados (cupos)', '=SUM(Invitados!D2:D)'],
    ['Invitaciones con respuesta', '=COUNTA(Invitados!H2:H)'],
    ['Invitaciones sin respuesta', '=B1-B3'],
    ['Personas que asisten', '=COUNTIF(Confirmaciones!D2:D,"Sí")+COUNTIFS(Confirmaciones!D2:D,"Sí",Confirmaciones!G2:G,"<>")'],
    ['Personas que no asisten', '=COUNTIF(Confirmaciones!D2:D,"No")'],
    ['', ''],
    ['Alimentación de quienes asisten', 'Personas'],
    ['Ninguna', dieta('Ninguna')],
    ['Vegetariano', dieta('Vegetariano')],
    ['Vegano', dieta('Vegano')],
    ['Celíaco', dieta('Celíaco')],
    ['Otro', dieta('Otro*')],
    ['', '']
  ]);
  res.getRange('D1:F1').setValues([['Alergias y restricciones (detalle)', '', 'Canciones pedidas']]);
  res.getRange('D2').setFormula('=IFERROR(FILTER({Confirmaciones!E2:E, Confirmaciones!F2:F}, Confirmaciones!D2:D="Sí", ' +
    'Confirmaciones!F2:F<>"Ninguna", Confirmaciones!F2:F<>""), "")');
  res.getRange('D40').setFormula('=IFERROR(FILTER({Confirmaciones!G2:G, Confirmaciones!H2:H}, Confirmaciones!D2:D="Sí", ' +
    'Confirmaciones!G2:G<>"", Confirmaciones!H2:H<>"Ninguna", Confirmaciones!H2:H<>""), "")');
  res.getRange('C40').setValue('Parejas:');
  res.getRange('F2').setFormula('=IFERROR(FILTER(Confirmaciones!I2:I, Confirmaciones!I2:I<>""), "")');
  res.getRange('A1:A14').setFontWeight('bold');
  res.getRange('A8:B8').setBackground('#F1EDE4');
  res.getRange('D1:F1').setFontWeight('bold').setBackground('#F1EDE4');
  res.setColumnWidth(1, 260); res.setColumnWidth(2, 90); res.setColumnWidth(3, 70);
  res.setColumnWidth(4, 200); res.setColumnWidth(5, 220); res.setColumnWidth(6, 260);

  ss.setActiveSheet(inv);
  return 'Planilla lista';
}

function hoja_(ss, nombre, encabezados) {
  const sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  sh.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/* ---------- Códigos de invitado ---------- */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Invitados' || e.range.getRow() < 2) return;
  const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  if (c1 <= 5 && c2 >= 5) arreglaTelefonos_(sh.getRange(e.range.getRow(), 5, e.range.getNumRows(), 1));
  if (c1 <= 2 && c2 >= 2) codigosFaltantes();
}

// Un teléfono escrito como "+56 9 ..." se lee como fórmula: lo dejamos como texto.
function arreglaTelefonos_(rg) {
  const f = rg.getFormulas();
  if (!f.some(r => r[0])) return;
  const v = rg.getDisplayValues();
  rg.setNumberFormat('@').setValues(f.map((r, i) => [r[0] ? r[0].replace(/^=/, '') : v[i][0]]));
}

function codigosFaltantes() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Invitados');
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const vals = sh.getRange(2, 1, n, 2).getValues();
  const usados = new Set(vals.map(v => String(v[0])).filter(String));
  let cambio = false;
  const out = vals.map(v => {
    if (String(v[0]).trim() || !String(v[1]).trim()) return [v[0]];
    let c;
    do { c = nuevoCodigo_(); } while (usados.has(c));
    usados.add(c);
    cambio = true;
    return [c];
  });
  if (cambio) sh.getRange(2, 1, n, 1).setValues(out);
}

function nuevoCodigo_() {
  const a = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}

/* ---------- Aplicación web (la usa la invitación) ---------- */
function doGet(e) {
  const id = String((e && e.parameter && e.parameter.id) || '').trim().toLowerCase();
  let out = { ok: false };
  if (id) {
    const fila = buscaInvitado_(id);
    if (fila) out = { ok: true, nombre: String(fila.datos[1]), saludo: String(fila.datos[2]), cupos: Number(fila.datos[3]) || 0 };
  }
  return json_(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const ss = SpreadsheetApp.getActive();
    const conf = ss.getSheetByName('Confirmaciones');
    const codigo = limpia_(d.codigo, 20).toLowerCase();
    const nombre = limpia_(d.nombre, 120);
    if (!nombre) return json_({ ok: false, error: 'Falta el nombre' });

    const inv = codigo ? buscaInvitado_(codigo) : null;
    const asiste = d.asiste === 'no' ? 'No' : 'Sí';
    const pareja = asiste === 'Sí' ? limpia_(d.pareja, 120) : '';
    const fila = [new Date(), codigo, inv ? String(inv.datos[1]) : '', asiste, nombre,
      asiste === 'Sí' ? limpia_(d.dieta, 80) : '', pareja, pareja ? limpia_(d.parejaDieta, 80) : '',
      asiste === 'Sí' ? limpia_(d.cancion, 150) : ''];

    // Si la misma persona vuelve a responder con el mismo link, se reemplaza su respuesta anterior.
    const m = conf.getLastRow() - 1;
    const prev = m > 0 ? conf.getRange(2, 1, m, 5).getValues() : [];
    const clave = (codigo + '|' + nombre).toLowerCase();
    const j = prev.findIndex(v => (String(v[1]) + '|' + String(v[4])).toLowerCase() === clave);
    if (j >= 0) conf.getRange(j + 2, 1, 1, fila.length).setValues([fila]);
    else conf.getRange(conf.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);

    if (inv) actualizaInvitado_(ss, codigo, inv.fila);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function buscaInvitado_(codigo) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Invitados');
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const vals = sh.getRange(2, 1, n, 4).getValues();
  const i = vals.findIndex(v => String(v[0]).trim().toLowerCase() === codigo);
  return i < 0 ? null : { fila: i + 2, datos: vals[i] };
}

function actualizaInvitado_(ss, codigo, filaInv) {
  const conf = ss.getSheetByName('Confirmaciones');
  const m = conf.getLastRow() - 1;
  const vals = m > 0 ? conf.getRange(2, 1, m, 7).getValues() : [];
  let ultima = null, si = false, personas = 0;
  vals.forEach(v => {
    if (String(v[1]).toLowerCase() !== codigo) return;
    if (!ultima || v[0] > ultima) ultima = v[0];
    if (v[3] === 'Sí') { si = true; personas += 1 + (String(v[6]).trim() ? 1 : 0); }
  });
  ss.getSheetByName('Invitados').getRange(filaInv, 8, 1, 3)
    .setValues([[ultima || '', ultima ? (si ? 'Sí' : 'No') : '', ultima ? personas : '']]);
}

function limpia_(s, max) {
  s = String(s == null ? '' : s).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;   // evita que un texto se lea como fórmula
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
