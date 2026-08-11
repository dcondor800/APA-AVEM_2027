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
// PASO 1: La hoja de calculo
//   Ya existe: "INTERES INSCRIPCION AVEM 2027", en la cuenta de David, y
//   contiene registros reales. Su ID esta abajo en SHEET_ID.
//   (Si alguna vez hubiera que rehacerla desde cero: hoja nueva en
//    https://sheets.google.com con las pestanas y encabezados de abajo.)
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
// PASO 5: Compartir la hoja y crear el script INDEPENDIENTE
//   La hoja vive en la cuenta de David y el script tiene que ejecutarse desde
//   la de APA, para que los correos salgan del dominio del cliente. Como no
//   pueden estar en la misma cuenta, el script NO va contenido en la hoja
//   (nada de Extensiones > Apps Script): va suelto y la abre por ID.
//
//   5a. Desde la cuenta de David: compartir la hoja con
//       avem.inscripciones@apa.org.pe con permiso de EDITOR.
//   5b. Copiar el ID de la hoja de su URL y pegarlo arriba en SHEET_ID.
//   5c. Entrar en https://script.google.com CON LA CUENTA DE APA
//       > Nuevo proyecto, y pegar este codigo entero.
//
//   Alternativa descartada: transferir la propiedad de la hoja a APA. Google
//   suele bloquear las transferencias entre organizaciones distintas, y
//   copiarla duplicaria la fuente de datos y dejaria atras los registros ya
//   guardados.
//
// PASO 6: Desplegar como aplicacion web
//   IMPORTANTE: hacerlo DESDE LA CUENTA avem.inscripciones@apa.org.pe, que es
//   la que APA facilito para esto. La cuenta que despliega es la que acaba
//   figurando como remitente de los correos.
//   - Implementar > Nueva implementacion > Tipo: "Aplicacion web"
//   - Ejecutar como: "Yo"  (es decir, la cuenta de APA)
//   - Quien tiene acceso: "Cualquier persona"
//   - Implementar y autorizar los permisos que pida
//     (ahora pedira tambien permiso para ENVIAR CORREO en tu nombre)
//   Para comprobarlo despues: ejecutar diagnosticoCorreo(), que avisa si la
//   cuenta no pertenece al dominio del cliente.
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

// ---- HOJA DE CALCULO -------------------------------------------------
// ID de la hoja. Se saca de su URL, entre /d/ y /edit:
//   https://docs.google.com/spreadsheets/d/ESTE_TROZO_ES_EL_ID/edit
//
// Hace falta porque la hoja y el script viven en cuentas distintas: la hoja
// esta en la cuenta de David y el script se ejecuta desde la de APA, para que
// los correos salgan del dominio del cliente. Al no ser un script contenido en
// la hoja, getActiveSpreadsheet() devolveria null y hay que abrirla por ID.
//
// Requisito: la hoja debe estar COMPARTIDA CON avem.inscripciones@apa.org.pe
// con permiso de Editor, o el script no podra escribir en ella.
//
// Si se deja vacio, el script asume que esta contenido en la hoja y usa
// getActiveSpreadsheet(), como antes.
var SHEET_ID = '1kc7vNzGQJoemRef3RzIOaJ1dwM8G_xCQmWjSjXZv6QI';

var FOOTER_URL = 'https://dcondor800.github.io/APA-AVEM_2027/screenshots/email-footer.png';
var ASUNTO = 'Gracias por registrar tu interes en AVEM 2027';

// ---- REMITENTE -------------------------------------------------------
// EL REMITENTE LO DETERMINA LA CUENTA QUE DESPLIEGA EL SCRIPT, no esta
// constante. APA facilito avem.inscripciones@apa.org.pe para este proyecto:
// desplegando desde ella, los correos salen a su nombre y esto queda vacio.
//
// Por que es la mejor via, y no un alias:
//   - El correo sale de servidores de Google autorizados por apa.org.pe (su
//     SPF incluye _spf.google.com y su MX es smtp.google.com), asi que SPF y
//     DKIM alinean de forma nativa y el correo pasa DMARC.
//   - No aparece la anotacion "enviado por / via gmail.com".
//   - La cuota diaria sube de ~100 correos (Gmail gratuito) a ~1500
//     (Google Workspace, que es lo que usa APA).
//
// Solo hay que rellenar esta constante si se quiere que el remitente visible
// sea DISTINTO de la cuenta que ejecuta el script; en ese caso la direccion
// tiene que ser un alias verificado de esa cuenta ("Enviar como" en Gmail).
// Si se rellena con una direccion no valida, enviarConRemitente() reintenta
// desde la cuenta por defecto y no se pierde ningun envio.
var REMITENTE = '';

// A donde van las respuestas si el destinatario pulsa "Responder".
// Se deja en apaeventos@apa.org.pe, que es la direccion de contacto publicada
// en el pie de la landing y la que el equipo de eventos ya atiende. El correo
// sale de avem.inscripciones@ pero las respuestas caen en el buzon conocido.
// Si se prefiere que lleguen a la propia cuenta de inscripciones, cambiar por
// 'avem.inscripciones@apa.org.pe' o dejar vacio.
var RESPONDER_A = 'apaeventos@apa.org.pe';

var NOMBRE_REMITENTE = 'AVEM 2027';

// ---- PRUEBAS ---------------------------------------------------------
// Direccion a la que diagnosticoCorreo() envia los correos de prueba.
// DEJALO VACIO salvo que quieras probar contra otra direccion: vacio
// significa "a la cuenta que ejecuta el script", que es lo habitual.
// Solo se usa al ejecutar el diagnostico a mano; no interviene en los
// registros reales del formulario.
var CORREO_DE_PRUEBA = '';

// Logger.log existe en los dos runtimes de Apps Script (V8 y el antiguo Rhino).
// console solo existe en V8, por eso no se usa directamente.
function avisar(texto) {
  try { Logger.log(texto); } catch (e) {}
}

// Abre la hoja por ID cuando el script es independiente, o la activa cuando
// esta contenido en ella.
function abrirHoja() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = abrirHoja();
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

// Envia el correo y devuelve la direccion desde la que acabo saliendo, para que
// el diagnostico pueda confirmarlo sin tener que abrir la bandeja de entrada.
// Si REMITENTE esta vacio, sale de la cuenta que ejecuta el script.
function enviarConRemitente(opciones) {
  if (REMITENTE) {
    try {
      opciones.from = REMITENTE;
      MailApp.sendEmail(opciones);
      return REMITENTE;
    } catch (err) {
      avisar('No se pudo enviar como ' + REMITENTE + ' (revisa que sea alias verificado ' +
             'de la cuenta que ejecuta el script). Se reenvia desde la cuenta por ' +
             'defecto. Detalle: ' + err);
      delete opciones.from;
    }
  }
  MailApp.sendEmail(opciones);
  return Session.getEffectiveUser().getEmail();
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
// No hay nada que rellenar: selecciona diagnosticoCorreo en el
// desplegable de la barra superior y pulsa Ejecutar. Los correos de
// prueba llegan a la cuenta que ejecuta el script. Despues abre
// "Registro de ejecucion": cada paso dice OK o FALLO con el motivo.
// ============================================================

function diagnosticoCorreo() {
  var L = [];
  function log(t) { L.push(t); Logger.log(t); }

  // Por defecto se prueba contra la propia cuenta que ejecuta, que es quien
  // esta haciendo la prueba. Evita tener que editar nada y, sobre todo, evita
  // que un buscar-y-reemplazar toque las referencias a la variable.
  var destino = CORREO_DE_PRUEBA || Session.getEffectiveUser().getEmail();

  log('--- DIAGNOSTICO AVEM 2027 ---');
  log('Destinatario de prueba: ' + destino);

  // 1. Runtime
  try {
    log('1. Runtime V8: ' + (typeof console !== 'undefined' ? 'SI' : 'NO (Rhino antiguo)'));
  } catch (e) {
    log('1. Runtime: no se pudo determinar');
  }

  // 2. Cuenta y cuota
  try {
    var cuenta = Session.getEffectiveUser().getEmail();
    var cuota = MailApp.getRemainingDailyQuota();
    log('2. Ejecutando como: ' + cuenta);
    log('   Cuota de correo restante hoy: ' + cuota);
    log('   Remitente configurado: ' + (REMITENTE || '(ninguno, sale desde la cuenta de arriba)'));
    log('   Responder a: ' + (RESPONDER_A || '(sin configurar)'));
    // El fallo tipico es desplegar desde la cuenta equivocada; se avisa aqui.
    if (cuenta.indexOf('@apa.org.pe') === -1) {
      log('   >> AVISO: esta cuenta NO es del dominio apa.org.pe.');
      log('      Los correos saldran a su nombre, con anotacion "via gmail.com",');
      log('      sin alinear DMARC y con la cuota reducida. Vuelve a desplegar');
      log('      el script desde la cuenta de APA.');
    } else {
      log('   >> Cuenta del dominio del cliente: correcto.');
      if (cuota < 500) {
        log('      Ojo: la cuota sugiere que no es una cuenta Workspace de pago.');
      }
    }
  } catch (e) {
    log('2. FALLO al leer cuenta/cuota: ' + e);
    log('   >> Suele significar que faltan permisos. Vuelve a autorizar el script.');
    return L.join('\n');
  }

  // 3. Acceso a la hoja. Con el script separado de la hoja, que falte el
  //    permiso de Editor es el fallo mas probable, y romperia los registros
  //    sin dar ninguna senal en la landing.
  try {
    var ss = abrirHoja();
    log('3. Hoja: "' + ss.getName() + '" (' + (SHEET_ID ? 'abierta por ID' : 'contenedora') + ')');
    var faltan = [];
    ['Asistente', 'Empresa'].forEach(function (n) {
      var h = ss.getSheetByName(n);
      if (h) log('   Pestana "' + n + '": OK, ' + h.getLastRow() + ' filas');
      else faltan.push(n);
    });
    if (faltan.length) {
      log('   >> FALTAN pestanas: ' + faltan.join(', ') + '. Los nombres deben ser exactos.');
    }
    // Comprobacion real de escritura: leer no garantiza permiso de Editor.
    var prueba = ss.getSheetByName('Asistente');
    if (prueba) {
      var fila = prueba.getLastRow() + 1;
      prueba.getRange(fila, 1).setValue('__prueba__');
      SpreadsheetApp.flush();
      prueba.deleteRow(fila);
      log('   Permiso de escritura: OK (se escribio y borro una fila de prueba)');
    }
  } catch (e) {
    log('3. FALLO al acceder a la hoja: ' + e);
    log('   >> Si el script corre desde la cuenta de APA, comprueba que SHEET_ID');
    log('      este puesto y que la hoja este compartida con esa cuenta como Editor.');
    return L.join('\n');
  }

  // 4. Envio minimo, sin HTML ni imagen: aisla si el problema es MailApp
  try {
    MailApp.sendEmail(destino, 'AVEM 2027 - prueba 1 de 2 (texto simple)',
      'Si recibes este mensaje, MailApp funciona correctamente.');
    log('4. Envio simple: OK (enviado)');
  } catch (e) {
    log('4. FALLO en el envio simple: ' + e);
    log('   >> El problema es MailApp: permisos o cuota. No sigas al paso 5.');
    return L.join('\n');
  }

  // 5. Descarga del pie de pagina
  var blob = null;
  try {
    blob = UrlFetchApp.fetch(FOOTER_URL).getBlob().setName('footer.png');
    log('5. Descarga del pie: OK (' + blob.getBytes().length + ' bytes)');
  } catch (e) {
    log('5. FALLO al descargar el pie: ' + e);
    log('   >> El correo se enviara igual, pero sin el banner.');
  }

  // 6. Envio completo, igual al que recibe quien se registra
  try {
    var opciones = {
      to: destino,
      subject: 'AVEM 2027 - prueba 2 de 2 (correo real)',
      name: NOMBRE_REMITENTE,
      body: textoPlano(),
      htmlBody: cuerpoHtml(blob !== null)
    };
    if (RESPONDER_A) opciones.replyTo = RESPONDER_A;
    if (blob) opciones.inlineImages = {footer: blob};
    var usado = enviarConRemitente(opciones);
    log('6. Envio completo: OK (enviado)');
    log('   Salio desde: ' + usado);
  } catch (e) {
    log('6. FALLO en el envio completo: ' + e);
    return L.join('\n');
  }

  log('--- FIN: se enviaron 2 correos. REVISA TAMBIEN LA CARPETA DE SPAM. ---');
  return L.join('\n');
}
