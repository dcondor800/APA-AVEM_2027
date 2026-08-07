// ============================================================
// AVEM 2027 - Google Apps Script (backend del formulario)
// ============================================================
//
// Que hace:
//   1. Recibe el POST del formulario de la landing
//   2. Agrega una fila a la hoja "Asistente" o "Empresa"
//   3. Envia un correo de confirmacion a la persona que se registro
//
// ------------------------------------------------------------
// PASO 1: Crear la hoja de calculo
//   - Ve a https://sheets.google.com
//   - Crea una nueva hoja y nombrala: "AVEM 2027 Registros"
//
// PASO 2: Crear las pestanas
//   - Renombra "Hoja 1" a: Asistente
//   - Crea una segunda pestana llamada: Empresa
//   (Los nombres deben ser EXACTOS, respetando mayusculas)
//
// PASO 3: Encabezados de la hoja "Asistente" (fila 1)
//   A1: Fecha                  J1: Talento y empleabilidad
//   B1: Nombre                 K1: Feria mas dinamica
//   C1: Empresa                L1: AVEM mas digital
//   D1: Actividad              M1: Networking y conexion empresarial
//   E1: Cargo                  N1: Mejor experiencia para asistentes y marcas
//   F1: Correo                 O1: Recinto Ferial
//   G1: Celular                P1: Evento multiproteina
//   H1: Acepto Datos           Q1: Que no puede faltar
//   I1: Nuevos formatos de contenido
//
// PASO 4: Encabezados de la hoja "Empresa" (fila 1)
//   A1: Fecha                  K1: Talento y empleabilidad
//   B1: Nombre                 L1: Feria mas dinamica
//   C1: Empresa                M1: AVEM mas digital
//   D1: Actividad              N1: Networking y conexion empresarial
//   E1: Cargo                  O1: Mejor experiencia para asistentes y marcas
//   F1: Correo                 P1: Recinto Ferial
//   G1: Celular                Q1: Evento multiproteina
//   H1: Interes de Participacion   R1: Que no puede faltar
//   I1: Acepto Datos
//   J1: Nuevos formatos de contenido
//
// PASO 5: Pegar este codigo
//   - Dentro de la hoja: Extensiones > Apps Script
//   - Borra el codigo de ejemplo y pega desde "function doPost" hasta el final
//
// PASO 6: Desplegar como aplicacion web
//   - Implementar > Nueva implementacion > Tipo: "Aplicacion web"
//   - Ejecutar como: "Yo"
//   - Quien tiene acceso: "Cualquier persona"
//   - Implementar y autorizar los permisos que pida
//     (ahora pedira tambien permiso para ENVIAR CORREO en tu nombre)
//
// PASO 7: Copiar la URL y pegarla en la landing
//   - Copia la URL https://script.google.com/macros/s/.../exec
//   - Pegala en index.html, linea ~780, dentro de sendToSheet()
//
// IMPORTANTE: si modificas este codigo despues del primer deploy, debes crear
//   una NUEVA implementacion (no editar la existente). Eso genera una URL nueva
//   que hay que volver a pegar en index.html.
//
// ------------------------------------------------------------
// SOBRE EL CORREO DE CONFIRMACION
//
//   - Se envia DESPUES de guardar la fila. Si el correo falla, el registro
//     igual queda guardado en la hoja.
//   - El pie de pagina se toma de la landing publicada en GitHub Pages y se
//     incrusta en el mensaje (inline), asi se ve aunque el cliente de correo
//     bloquee imagenes externas.
//   - CUOTA DIARIA: una cuenta Gmail gratuita permite ~100 correos al dia;
//     una cuenta Google Workspace, ~1500. Si se supera, el registro se sigue
//     guardando pero el correo no sale.
//   - El remitente sera la cuenta de Google que despliega el script.
// ============================================================

var FOOTER_URL = 'https://dcondor800.github.io/APA-AVEM_2027/screenshots/email-footer.png';
var ASUNTO = 'Gracias por registrar tu interes en AVEM 2027';

// ---- REMITENTE -------------------------------------------------------
// Direccion desde la que salen los correos.
//
// OJO: Apps Script NO puede conectarse a un servidor SMTP. No tiene API de
// sockets; UrlFetchApp solo hace HTTP/HTTPS. Por eso aqui no se ponen host,
// puerto ni contrasena de aplicacion: eso va en la configuracion de Gmail.
//
// Google solo acepta esta direccion si es un ALIAS VERIFICADO de la cuenta
// que ejecuta el script. Para darla de alta:
//
//   Gmail > Configuracion > Cuentas e importacion > "Enviar como"
//         > Anadir otra direccion de correo
//
//   1. Escribe la direccion (ej. apaeventos@apa.org.pe)
//   2. Desmarca "Tratar como un alias"
//   3. Servidor SMTP del cliente + puerto 587 (TLS)
//      Usuario: la direccion completa
//      Contrasena: la contrasena de aplicacion del cliente
//      (16 caracteres en 4 grupos, ej. "jiab sppf msnh bfqm";
//       requiere verificacion en 2 pasos activa en esa cuenta)
//   4. Google envia un codigo a esa direccion; alguien con acceso lo confirma
//
// Enviar a traves del SMTP del cliente es lo recomendable: el correo se
// relaya por su servidor, asi SPF y DKIM alinean con su dominio y no hay
// problemas de DMARC ni de reputacion.
//
// Si se deja vacio, o si la direccion no es un alias valido, el correo sale
// igual desde la cuenta que desplego el script (no se pierde ningun envio).
var REMITENTE = '';                              // ej: 'apaeventos@apa.org.pe'

// A donde van las respuestas si el destinatario pulsa "Responder".
// Esto NO requiere alias ni configuracion: funciona desde ya.
var RESPONDER_A = 'apaeventos@apa.org.pe';

var NOMBRE_REMITENTE = 'AVEM 2027';

// Logger.log existe en los dos runtimes de Apps Script (V8 y el antiguo Rhino).
// console solo existe en V8, por eso no se usa directamente.
function avisar(texto) {
  try { Logger.log(texto); } catch (e) {}
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(data.tipo);

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Hoja no encontrada: ' + data.tipo}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var fecha = new Date().toLocaleString('es-PE', {timeZone: 'America/Lima'});
    var row;

    if (data.tipo === 'Asistente') {
      row = [
        fecha,
        data.nombre,
        data.empresa,
        data.actividad,
        data.cargo,
        data.correo,
        data.celular,
        data.acepto,
        data.sel1,
        data.sel2,
        data.sel3,
        data.sel4,
        data.sel5,
        data.sel6,
        data.sel7,
        data.sel8,
        data.textarea
      ];
    } else {
      row = [
        fecha,
        data.nombre,
        data.empresa,
        data.actividad,
        data.cargo,
        data.correo,
        data.celular,
        data.interes,
        data.acepto,
        data.sel1,
        data.sel2,
        data.sel3,
        data.sel4,
        data.sel5,
        data.sel6,
        data.sel7,
        data.sel8,
        data.textarea
      ];
    }

    sheet.appendRow(row);

    // El correo va despues de guardar y en su propio try/catch:
    // si falla el envio, el registro ya quedo a salvo en la hoja.
    try {
      enviarConfirmacion(data.correo);
    } catch (errMail) {
      avisar('Fallo el correo de confirmacion a ' + data.correo + ': ' + errMail);
    }

    return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function enviarConfirmacion(correo) {
  if (!correo) return;
  correo = String(correo).trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) {
    avisar('Correo con formato invalido, no se envia: ' + correo);
    return;
  }
  if (MailApp.getRemainingDailyQuota() < 1) {
    avisar('Cuota diaria de correo agotada. No se envio a ' + correo);
    return;
  }

  var opciones = {
    to: correo,
    subject: ASUNTO,
    name: NOMBRE_REMITENTE,
    body: textoPlano(),
    htmlBody: cuerpoHtml(true)
  };

  if (RESPONDER_A) opciones.replyTo = RESPONDER_A;

  // El pie se incrusta como imagen inline (cid) para que se vea aunque el
  // cliente bloquee imagenes remotas. Si no se puede descargar, se envia sin el.
  try {
    var blob = UrlFetchApp.fetch(FOOTER_URL).getBlob().setName('footer.png');
    opciones.inlineImages = {footer: blob};
  } catch (errImg) {
    avisar('No se pudo cargar el pie de pagina, se envia sin imagen: ' + errImg);
    opciones.htmlBody = cuerpoHtml(false);
  }

  enviarConRemitente(opciones);
}

// Intenta enviar desde REMITENTE. Si Google lo rechaza porque no es un alias
// verificado, reintenta desde la cuenta por defecto en vez de perder el envio.
function enviarConRemitente(opciones) {
  if (REMITENTE) {
    try {
      opciones.from = REMITENTE;
      MailApp.sendEmail(opciones);
      return;
    } catch (err) {
      avisar('No se pudo enviar como ' + REMITENTE + ' (revisa que sea alias verificado). ' +
             'Se reenvia desde la cuenta por defecto. Detalle: ' + err);
      delete opciones.from;
    }
  }
  MailApp.sendEmail(opciones);
}

function cuerpoHtml(conFooter) {
  var MORADO = '#5752A6';
  var NAVY = '#2E2C6E';

  var pie = conFooter
    ? '<img src="cid:footer" width="600" alt="AVEM 2027 - Asociacion Peruana de Avicultura"' +
      ' style="display:block; width:100%; max-width:600px; height:auto; border:0; outline:none; text-decoration:none;">'
    : '';

  var vinieta = function (texto) {
    return '<tr>' +
      '<td valign="top" style="width:18px; padding:0 0 10px; color:' + MORADO + '; font-size:15px; line-height:1.55;">&bull;</td>' +
      '<td style="padding:0 0 10px; color:' + MORADO + '; font-size:15px; line-height:1.55;">' + texto + '</td>' +
      '</tr>';
  };

  return '' +
  '<!DOCTYPE html><html><body style="margin:0; padding:0; background:#f4f4f6;">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f4f4f6;">' +
  '<tr><td align="center" style="padding:24px 12px;">' +
  '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0"' +
  ' style="width:100%; max-width:600px; background:#ffffff; border-radius:14px; overflow:hidden;' +
  ' font-family:Poppins, Helvetica, Arial, sans-serif;">' +

  '<tr><td style="padding:34px 34px 8px;">' +
  '<p style="margin:0 0 16px; color:' + NAVY + '; font-size:16px; font-weight:700; line-height:1.5;">Estimado/a participante:</p>' +
  '<p style="margin:0 0 14px; color:' + MORADO + '; font-size:15px; line-height:1.55;">' +
  'Gracias por registrar tu inter&eacute;s en formar parte del Congreso de Avicultura AVEM 2027.</p>' +
  '<p style="margin:0 0 14px; color:' + MORADO + '; font-size:15px; line-height:1.55;">' +
  'Una vez culminada esta etapa, te enviaremos informaci&oacute;n de acuerdo con la modalidad de participaci&oacute;n que hayas seleccionado:</p>' +

  '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; margin:0 0 4px;">' +
  vinieta('Si registraste tu inter&eacute;s como <strong>empresa</strong>, recibir&aacute;s el plano final de la zona de exposici&oacute;n y los precios correspondientes a las diferentes categor&iacute;as de participaci&oacute;n.') +
  vinieta('Si registraste tu inter&eacute;s como <strong>asistente</strong>, ser&aacute;s de los primeros en conocer las tarifas de inscripci&oacute;n al evento.') +
  '</table>' +

  '<p style="margin:14px 0 30px; color:' + MORADO + '; font-size:15px; line-height:1.55;">' +
  'Esperamos contar contigo en AVEM 2027.</p>' +
  '</td></tr>' +

  '<tr><td style="padding:0; font-size:0; line-height:0;">' + pie + '</td></tr>' +

  '</table></td></tr></table></body></html>';
}

function textoPlano() {
  return 'Estimado/a participante:\n\n' +
    'Gracias por registrar tu interes en formar parte del Congreso de Avicultura AVEM 2027.\n\n' +
    'Una vez culminada esta etapa, te enviaremos informacion de acuerdo con la modalidad ' +
    'de participacion que hayas seleccionado:\n\n' +
    '- Si registraste tu interes como empresa, recibiras el plano final de la zona de exposicion ' +
    'y los precios correspondientes a las diferentes categorias de participacion.\n\n' +
    '- Si registraste tu interes como asistente, seras de los primeros en conocer las tarifas ' +
    'de inscripcion al evento.\n\n' +
    'Esperamos contar contigo en AVEM 2027.\n';
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status: 'ok', message: 'AVEM 2027 API activa'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// DIAGNOSTICO
// ============================================================
// Cambia la direccion de abajo por la tuya, selecciona
// diagnosticoCorreo en el desplegable de arriba y pulsa Ejecutar.
// Despues abre "Registro de ejecucion" para ver el resultado de
// cada paso. Cada paso dice OK o FALLO con el motivo.
// ============================================================

// Cambia SOLO esta linea por tu direccion.
var CORREO_DE_PRUEBA = 'CAMBIAME@ejemplo.com';

function diagnosticoCorreo() {
  var L = [];
  function log(t) { L.push(t); Logger.log(t); }

  log('--- DIAGNOSTICO AVEM 2027 ---');
  log('Destinatario de prueba: ' + CORREO_DE_PRUEBA);

  // Se busca el marcador "CAMBIAME", no la direccion completa: asi un
  // buscar-y-reemplazar de la direccion no puede afectar a esta comprobacion.
  if (CORREO_DE_PRUEBA.indexOf('CAMBIAME') !== -1) {
    log('FALLO: no cambiaste CORREO_DE_PRUEBA por tu direccion real.');
    return L.join('\n');
  }

  // 1. Runtime
  try {
    log('1. Runtime V8: ' + (typeof console !== 'undefined' ? 'SI' : 'NO (Rhino antiguo)'));
  } catch (e) {
    log('1. Runtime: no se pudo determinar');
  }

  // 2. Cuenta y cuota
  try {
    log('2. Ejecutando como: ' + Session.getEffectiveUser().getEmail());
    log('   Cuota de correo restante hoy: ' + MailApp.getRemainingDailyQuota());
    log('   Remitente configurado: ' + (REMITENTE || '(ninguno, sale desde la cuenta de arriba)'));
    log('   Responder a: ' + (RESPONDER_A || '(sin configurar)'));
  } catch (e) {
    log('2. FALLO al leer cuenta/cuota: ' + e);
    log('   >> Suele significar que faltan permisos. Vuelve a autorizar el script.');
    return L.join('\n');
  }

  // 3. Envio minimo, sin HTML ni imagen: aisla si el problema es MailApp
  try {
    MailApp.sendEmail(CORREO_DE_PRUEBA, 'AVEM 2027 - prueba 1 de 2 (texto simple)',
      'Si recibes este mensaje, MailApp funciona correctamente.');
    log('3. Envio simple: OK (enviado)');
  } catch (e) {
    log('3. FALLO en el envio simple: ' + e);
    log('   >> El problema es MailApp: permisos o cuota. No sigas al paso 4.');
    return L.join('\n');
  }

  // 4. Descarga del pie de pagina
  var blob = null;
  try {
    blob = UrlFetchApp.fetch(FOOTER_URL).getBlob().setName('footer.png');
    log('4. Descarga del pie: OK (' + blob.getBytes().length + ' bytes)');
  } catch (e) {
    log('4. FALLO al descargar el pie: ' + e);
    log('   >> El correo se enviara igual, pero sin el banner.');
  }

  // 5. Envio completo, igual al que recibe quien se registra
  try {
    var opciones = {
      to: CORREO_DE_PRUEBA,
      subject: 'AVEM 2027 - prueba 2 de 2 (correo real)',
      name: NOMBRE_REMITENTE,
      body: textoPlano(),
      htmlBody: cuerpoHtml(blob !== null)
    };
    if (RESPONDER_A) opciones.replyTo = RESPONDER_A;
    if (blob) opciones.inlineImages = {footer: blob};
    enviarConRemitente(opciones);
    log('5. Envio completo: OK (enviado)');
    log('   Comprueba en el correo recibido que el remitente sea el esperado.');
  } catch (e) {
    log('5. FALLO en el envio completo: ' + e);
    return L.join('\n');
  }

  log('--- FIN: se enviaron 2 correos. REVISA TAMBIEN LA CARPETA DE SPAM. ---');
  return L.join('\n');
}
