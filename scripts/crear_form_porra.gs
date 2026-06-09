/**
 * PORRA MUNDIAL 2026 — Sistema con pestañas personales
 *
 * FLUJO:
 *   1. Participante rellena el form (solo alias + email)
 *   2. Apps Script crea una pestaña con sus 72 partidos
 *   3. Le manda un email con el link directo a su pestaña
 *   4. Rellena sus marcadores cómodamente en la hoja
 *   5. El scorer lee todas las pestañas y actualiza el dashboard
 *
 * SETUP (una sola vez):
 *   1. Ejecuta crearFormularioPorra  → crea Form + Spreadsheet nuevos
 *   2. Ejecuta configurarTrigger     → activa el trigger automático
 */

var NETLIFY_TOKEN = 'nfp_kp6cBifMAgrs9KkZdAkargBP4ryJacLm52ff';
var NETLIFY_SITE  = 'porra-mundial-sdx8u';
var FECHA_LIMITE  = '11 jun 2026 a las 17:00h';

// Partidos fase de grupos: [id, local, golesL, golesV, visitante, grupo]
var PARTIDOS = [
    ['G-A1-A2', 'México 🇲🇽', '', '', '🇿🇦 Sudáfrica', 'Grupo A'],
    ['G-A3-A4', 'Corea del Sur 🇰🇷', '', '', '🇨🇿 Chequia', 'Grupo A'],
    ['G-A1-A3', 'México 🇲🇽', '', '', '🇰🇷 Corea del Sur', 'Grupo A'],
    ['G-A2-A4', 'Sudáfrica 🇿🇦', '', '', '🇨🇿 Chequia', 'Grupo A'],
    ['G-A1-A4', 'México 🇲🇽', '', '', '🇨🇿 Chequia', 'Grupo A'],
    ['G-A2-A3', 'Sudáfrica 🇿🇦', '', '', '🇰🇷 Corea del Sur', 'Grupo A'],
    ['G-B1-B2', 'Canadá 🇨🇦', '', '', '🇨🇭 Suiza', 'Grupo B'],
    ['G-B3-B4', 'Catar 🇶🇦', '', '', '🇧🇦 Bosnia y Herzegovina', 'Grupo B'],
    ['G-B1-B3', 'Canadá 🇨🇦', '', '', '🇶🇦 Catar', 'Grupo B'],
    ['G-B2-B4', 'Suiza 🇨🇭', '', '', '🇧🇦 Bosnia y Herzegovina', 'Grupo B'],
    ['G-B1-B4', 'Canadá 🇨🇦', '', '', '🇧🇦 Bosnia y Herzegovina', 'Grupo B'],
    ['G-B2-B3', 'Suiza 🇨🇭', '', '', '🇶🇦 Catar', 'Grupo B'],
    ['G-C1-C2', 'Brasil 🇧🇷', '', '', '🇲🇦 Marruecos', 'Grupo C'],
    ['G-C3-C4', 'Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿', '', '', '🇭🇹 Haití', 'Grupo C'],
    ['G-C1-C3', 'Brasil 🇧🇷', '', '', '🏴󠁧󠁢󠁳󠁣󠁴󠁿 Escocia', 'Grupo C'],
    ['G-C2-C4', 'Marruecos 🇲🇦', '', '', '🇭🇹 Haití', 'Grupo C'],
    ['G-C1-C4', 'Brasil 🇧🇷', '', '', '🇭🇹 Haití', 'Grupo C'],
    ['G-C2-C3', 'Marruecos 🇲🇦', '', '', '🏴󠁧󠁢󠁳󠁣󠁴󠁿 Escocia', 'Grupo C'],
    ['G-D1-D2', 'Estados Unidos 🇺🇸', '', '', '🇵🇾 Paraguay', 'Grupo D'],
    ['G-D3-D4', 'Australia 🇦🇺', '', '', '🇹🇷 Turquía', 'Grupo D'],
    ['G-D1-D3', 'Estados Unidos 🇺🇸', '', '', '🇦🇺 Australia', 'Grupo D'],
    ['G-D2-D4', 'Paraguay 🇵🇾', '', '', '🇹🇷 Turquía', 'Grupo D'],
    ['G-D1-D4', 'Estados Unidos 🇺🇸', '', '', '🇹🇷 Turquía', 'Grupo D'],
    ['G-D2-D3', 'Paraguay 🇵🇾', '', '', '🇦🇺 Australia', 'Grupo D'],
    ['G-E1-E2', 'Alemania 🇩🇪', '', '', '🇨🇼 Curazao', 'Grupo E'],
    ['G-E3-E4', 'Costa Rica 🇨🇷', '', '', '🇪🇨 Ecuador', 'Grupo E'],
    ['G-E1-E3', 'Alemania 🇩🇪', '', '', '🇨🇷 Costa Rica', 'Grupo E'],
    ['G-E2-E4', 'Curazao 🇨🇼', '', '', '🇪🇨 Ecuador', 'Grupo E'],
    ['G-E1-E4', 'Alemania 🇩🇪', '', '', '🇪🇨 Ecuador', 'Grupo E'],
    ['G-E2-E3', 'Curazao 🇨🇼', '', '', '🇨🇷 Costa Rica', 'Grupo E'],
    ['G-F1-F2', 'Países Bajos 🇳🇱', '', '', '🇯🇵 Japón', 'Grupo F'],
    ['G-F3-F4', 'Túnez 🇹🇳', '', '', '🇸🇪 Suecia', 'Grupo F'],
    ['G-F1-F3', 'Países Bajos 🇳🇱', '', '', '🇹🇳 Túnez', 'Grupo F'],
    ['G-F2-F4', 'Japón 🇯🇵', '', '', '🇸🇪 Suecia', 'Grupo F'],
    ['G-F1-F4', 'Países Bajos 🇳🇱', '', '', '🇸🇪 Suecia', 'Grupo F'],
    ['G-F2-F3', 'Japón 🇯🇵', '', '', '🇹🇳 Túnez', 'Grupo F'],
    ['G-G1-G2', 'Bélgica 🇧🇪', '', '', '🇪🇬 Egipto', 'Grupo G'],
    ['G-G3-G4', 'Irán 🇮🇷', '', '', '🇳🇿 Nueva Zelanda', 'Grupo G'],
    ['G-G1-G3', 'Bélgica 🇧🇪', '', '', '🇮🇷 Irán', 'Grupo G'],
    ['G-G2-G4', 'Egipto 🇪🇬', '', '', '🇳🇿 Nueva Zelanda', 'Grupo G'],
    ['G-G1-G4', 'Bélgica 🇧🇪', '', '', '🇳🇿 Nueva Zelanda', 'Grupo G'],
    ['G-G2-G3', 'Egipto 🇪🇬', '', '', '🇮🇷 Irán', 'Grupo G'],
    ['G-H1-H2', 'España 🇪🇸', '', '', '🇨🇻 Cabo Verde', 'Grupo H'],
    ['G-H3-H4', 'Arabia Saudí 🇸🇦', '', '', '🇺🇾 Uruguay', 'Grupo H'],
    ['G-H1-H3', 'España 🇪🇸', '', '', '🇸🇦 Arabia Saudí', 'Grupo H'],
    ['G-H2-H4', 'Cabo Verde 🇨🇻', '', '', '🇺🇾 Uruguay', 'Grupo H'],
    ['G-H1-H4', 'España 🇪🇸', '', '', '🇺🇾 Uruguay', 'Grupo H'],
    ['G-H2-H3', 'Cabo Verde 🇨🇻', '', '', '🇸🇦 Arabia Saudí', 'Grupo H'],
    ['G-I1-I2', 'Francia 🇫🇷', '', '', '🇸🇳 Senegal', 'Grupo I'],
    ['G-I3-I4', 'Noruega 🇳🇴', '', '', '🇮🇶 Irak', 'Grupo I'],
    ['G-I1-I3', 'Francia 🇫🇷', '', '', '🇳🇴 Noruega', 'Grupo I'],
    ['G-I2-I4', 'Senegal 🇸🇳', '', '', '🇮🇶 Irak', 'Grupo I'],
    ['G-I1-I4', 'Francia 🇫🇷', '', '', '🇮🇶 Irak', 'Grupo I'],
    ['G-I2-I3', 'Senegal 🇸🇳', '', '', '🇳🇴 Noruega', 'Grupo I'],
    ['G-J1-J2', 'Argentina 🇦🇷', '', '', '🇩🇿 Argelia', 'Grupo J'],
    ['G-J3-J4', 'Austria 🇦🇹', '', '', '🇯🇴 Jordania', 'Grupo J'],
    ['G-J1-J3', 'Argentina 🇦🇷', '', '', '🇦🇹 Austria', 'Grupo J'],
    ['G-J2-J4', 'Argelia 🇩🇿', '', '', '🇯🇴 Jordania', 'Grupo J'],
    ['G-J1-J4', 'Argentina 🇦🇷', '', '', '🇯🇴 Jordania', 'Grupo J'],
    ['G-J2-J3', 'Argelia 🇩🇿', '', '', '🇦🇹 Austria', 'Grupo J'],
    ['G-K1-K2', 'Portugal 🇵🇹', '', '', '🇨🇴 Colombia', 'Grupo K'],
    ['G-K3-K4', 'Uzbekistán 🇺🇿', '', '', '🇨🇩 RD Congo', 'Grupo K'],
    ['G-K1-K3', 'Portugal 🇵🇹', '', '', '🇺🇿 Uzbekistán', 'Grupo K'],
    ['G-K2-K4', 'Colombia 🇨🇴', '', '', '🇨🇩 RD Congo', 'Grupo K'],
    ['G-K1-K4', 'Portugal 🇵🇹', '', '', '🇨🇩 RD Congo', 'Grupo K'],
    ['G-K2-K3', 'Colombia 🇨🇴', '', '', '🇺🇿 Uzbekistán', 'Grupo K'],
    ['G-L1-L2', 'Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿', '', '', '🇭🇷 Croacia', 'Grupo L'],
    ['G-L3-L4', 'Ghana 🇬🇭', '', '', '🇵🇦 Panamá', 'Grupo L'],
    ['G-L1-L3', 'Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿', '', '', '🇬🇭 Ghana', 'Grupo L'],
    ['G-L2-L4', 'Croacia 🇭🇷', '', '', '🇵🇦 Panamá', 'Grupo L'],
    ['G-L1-L4', 'Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿', '', '', '🇵🇦 Panamá', 'Grupo L'],
    ['G-L2-L3', 'Croacia 🇭🇷', '', '', '🇬🇭 Ghana', 'Grupo L']
];

// Premios especiales
var PREMIOS = [
  ['mvp',       'MVP del Torneo',      ''],
  ['goleador',  'Goleador del Torneo', ''],
  ['portero',   'Portero del Torneo',  ''],
];

function getSpreadsheetId() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Primero ejecuta crearFormularioPorra.');
  return id;
}

// ══════════════════════════════════════════════════════════════════
//  1. CREAR FORMULARIO (solo alias + email)
// ══════════════════════════════════════════════════════════════════

function crearFormularioPorra() {
  // Crear Spreadsheet maestro
  var ss = SpreadsheetApp.create('Porra Mundial 2026 — Pronósticos');

  // Crear hoja de instrucciones como portada
  var portada = ss.getActiveSheet();
  portada.setName('ℹ️ Info');
  portada.getRange('A1').setValue('Porra Mundial 2026').setFontSize(16).setFontWeight('bold');
  portada.getRange('A2').setValue('Cada participante tiene su propia pestaña. No edites las pestañas de otros.');
  portada.setColumnWidth(1, 500);

  // Guardar IDs
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  // Crear formulario de registro (solo alias + email)
  var form = FormApp.create('⚽ Porra Mundial 2026 — Registro');
  form.setDescription(
    'Regístrate para recibir tu hoja de pronósticos personalizada.\n' +
    'Te llegará un email con el link a tu pestaña en Google Sheets.\n\n' +
    'Fecha límite para rellenar pronósticos: ' + FECHA_LIMITE
  );
  form.setCollectEmail(false);
  form.setConfirmationMessage('✅ ¡Registrado! Recibirás un email con el link a tu hoja de pronósticos en breve.');

  form.addTextItem()
    .setTitle('Tu alias')
    .setHelpText('El nombre con el que apareceréis en el ranking')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Tu email')
    .setHelpText('Te enviaremos el link a tu hoja personal de pronósticos')
    .setValidation(
      FormApp.createTextValidation()
        .requireTextIsEmail()
        .setHelpText('Introduce un email válido')
        .build()
    )
    .setRequired(true);

  // Vincular form al spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  PropertiesService.getScriptProperties().setProperty('FORM_ID', form.getId());

  Logger.log('✅ Todo creado desde cero');
  Logger.log('📋 Form de registro: ' + form.getPublishedUrl());
  Logger.log('📊 Spreadsheet: ' + ss.getUrl());
  Logger.log('');
  Logger.log('➡️  Ahora ejecuta configurarTrigger');
}

// ══════════════════════════════════════════════════════════════════
//  2. TRIGGER AUTOMÁTICO
// ══════════════════════════════════════════════════════════════════

function configurarTrigger() {
  var ssId = getSpreadsheetId();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ssId)
    .onFormSubmit()
    .create();
  Logger.log('✅ Trigger instalado. Ya puedes compartir el formulario.');
}

// ══════════════════════════════════════════════════════════════════
//  3. CREAR PESTAÑA CUANDO ALGUIEN SE REGISTRA
// ══════════════════════════════════════════════════════════════════

function onFormSubmit(e) {
  try {
    var respuestas = e.namedValues;
    var alias = (respuestas['Tu alias'] || [''])[0].trim();
    var email = (respuestas['Tu email'] || [''])[0].trim().toLowerCase();

    if (!alias || !email) {
      Logger.log('⚠️ Respuesta sin alias o email, ignorada.');
      return;
    }

    var ss  = SpreadsheetApp.openById(getSpreadsheetId());
    var url = crearPestanaParticipante(ss, alias, email);

    // Enviar email con el link
    enviarEmailRegistro(alias, email, url);

    Logger.log('✅ ' + alias + ' (' + email + ') — pestaña creada y email enviado');
  } catch(err) {
    Logger.log('❌ Error en onFormSubmit: ' + err.toString());
  }
}

function crearPestanaParticipante(ss, alias, email) {
  // Si ya existe la pestaña, devolver su URL sin duplicar
  var existente = ss.getSheetByName(alias);
  if (existente) {
    return ss.getUrl() + '#gid=' + existente.getSheetId();
  }

  var sheet = ss.insertSheet(alias);

  // ── Encabezado ──────────────────────────────────────────────────
  var headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setValues([['ID Partido', 'Equipo Local', 'Goles Local ⬅', 'Goles Visitante ➡', 'Equipo Visitante', 'Grupo']]);
  headerRange.setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);

  // ── Datos: los 72 partidos ───────────────────────────────────────
  var data = PARTIDOS.map(function(p) {
    return [p[0], p[1], p[2], p[3], p[4], p[5]];
  });
  sheet.getRange(2, 1, data.length, 6).setValues(data);

  // ── Columnas de goles: fondo amarillo, editable ─────────────────
  var golesRange = sheet.getRange(2, 3, data.length, 2);
  golesRange.setBackground('#fff9c4').setHorizontalAlignment('center').setFontSize(12);

  // ── Validación numérica en goles ─────────────────────────────────
  var rule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 20)
    .setHelpText('Introduce un número entre 0 y 20')
    .setAllowInvalid(false)
    .build();
  golesRange.setDataValidation(rule);

  // ── Separadores de grupo: fila con color por grupo ───────────────
  var grupos = {};
  data.forEach(function(row, i) {
    var g = row[5];
    if (!grupos[g]) grupos[g] = [];
    grupos[g].push(i + 2); // fila en la hoja (1-indexed, +1 header)
  });
  var groupColors = [
    '#e3f2fd','#fce4ec','#f3e5f5','#e8f5e9',
    '#fff3e0','#e0f7fa','#f9fbe7','#fbe9e7',
    '#ede7f6','#e0f2f1','#fff8e1','#fafafa'
  ];
  var groupNames = Object.keys(grupos);
  groupNames.forEach(function(g, gi) {
    var color = groupColors[gi % groupColors.length];
    grupos[g].forEach(function(row) {
      sheet.getRange(row, 1, 1, 6).setBackground(color);
    });
    // Restaurar amarillo en columnas de goles
    grupos[g].forEach(function(row) {
      sheet.getRange(row, 3, 1, 2).setBackground('#fff9c4');
    });
  });

  // ── Sección premios (debajo de los partidos) ─────────────────────
  var premioStart = data.length + 3;
  sheet.getRange(premioStart - 1, 1, 1, 3)
    .setValues([['PREMIOS ESPECIALES (+15 pts c/u)', '', '']])
    .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(premioStart, 1, PREMIOS.length, 3)
    .setValues(PREMIOS.map(function(p) { return [p[0], p[1], p[2]]; }));
  sheet.getRange(premioStart, 3, PREMIOS.length, 1)
    .setBackground('#fff9c4').setHorizontalAlignment('left');

  // ── Columnas anchas ──────────────────────────────────────────────
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 90);

  // ── Proteger: solo el participante (y el owner) puede editar ─────
  try {
    var protection = sheet.protect().setDescription('Hoja de ' + alias);
    protection.removeEditors(protection.getEditors());
    protection.addEditor(email);
    // Desproteger solo las celdas editables (goles + premios)
    var editableRanges = [
      sheet.getRange(2, 3, data.length, 2),
      sheet.getRange(premioStart, 3, PREMIOS.length, 1)
    ];
    protection.setUnprotectedRanges(editableRanges);
  } catch(e) {
    Logger.log('⚠️ Protección omitida: ' + e.toString());
  }

  return ss.getUrl() + '#gid=' + sheet.getSheetId();
}

// ══════════════════════════════════════════════════════════════════
//  4. EMAIL DE BIENVENIDA
// ══════════════════════════════════════════════════════════════════

function enviarEmailRegistro(alias, email, url) {
  var asunto = '⚽ Porra Mundial 2026 — Tu hoja de pronósticos, ' + alias;
  var cuerpo =
    'Hola ' + alias + ',\n\n' +
    '¡Ya estás registrado en la Porra del Mundial 2026!\n\n' +
    '📊 Aquí tienes tu hoja personal para rellenar los pronósticos:\n' +
    url + '\n\n' +
    '🕔 Fecha límite: ' + FECHA_LIMITE + '\n\n' +
    'Instrucciones:\n' +
    '• Rellena los goles en las columnas amarillas\n' +
    '• Para los Premios Especiales, escribe el nombre del jugador\n' +
    '• Guarda los cambios con Cmd+S / Ctrl+S\n\n' +
    '¡Mucha suerte!\n' +
    'Organización Porra Mundial 2026';

  MailApp.sendEmail(email, asunto, cuerpo);
}

// ══════════════════════════════════════════════════════════════════
//  5. SCORER + DEPLOY NETLIFY (se añadirá en el siguiente paso)
// ══════════════════════════════════════════════════════════════════

function calcularYPublicar() {
  Logger.log('calcularYPublicar: scorer pendiente de implementar');
}
