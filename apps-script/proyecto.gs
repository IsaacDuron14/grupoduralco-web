// ==== Duralco Constructores — "¿Tiene un proyecto?" (www.grupoduralco.com/proyecto) ====
// Correos que reciben cada proyecto (separados por coma)
const CORREOS = 'gerente.operaciones@grupoduralco.com, ventas@grupoduralco.com, Ever.duron@grupoduralco.com';
const CARPETA = 'Proyectos Duralco Constructores';   // carpeta en Google Drive
const PESTAÑA = 'Proyectos';
const ESTADOS = ['Nuevo', 'Contactado', 'Visita agendada', 'Cotizado', 'Ganado', 'Perdido'];
const COLORES = {'Nuevo':'#FFF2CC','Contactado':'#CFE2F3','Visita agendada':'#D9D2E9','Cotizado':'#FCE5CD','Ganado':'#D9EAD3','Perdido':'#EFEFEF'};

function doPost(e) {
  return procesar(JSON.parse(e.postData.contents));
}

function procesar(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const campos = d.campos || [];
    const fecha = new Date();
    const v = k => { const c = campos.find(x => x.k === k); return c ? c.v : ''; };
    const folio = v('Folio') || d.folio || ('DC-' + fecha.getTime());
    const quien = v('Empresa') || v('Nombre') || 'sin nombre';

    // 1. Archivos (planos, fotos) a Google Drive
    const blobs = [], enlaces = [];
    let peso = 0;
    if ((d.archivos || []).length) {
      const raiz = carpeta(DriveApp.getRootFolder(), CARPETA);
      const dir = carpeta(raiz, folio + ' – ' + quien);
      d.archivos.forEach(a => {
        const b = Utilities.newBlob(Utilities.base64Decode(a.datos), a.tipo, a.nombre);
        const f = dir.createFile(b);
        peso += b.getBytes().length;
        blobs.push(b);
        enlaces.push(f.getUrl());
      });
    }

    // 2. Registro en la hoja (columnas por nombre; Estado para seguimiento)
    if (!d.prueba) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let hoja = ss.getSheetByName(PESTAÑA);
      if (!hoja) hoja = ss.insertSheet(PESTAÑA, 0);
      const claves = ['Fecha', 'Estado', ...campos.map(c => c.k), 'Archivos'];
      let enc = hoja.getLastRow() ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0] : [];
      const estilo = r => r.setFontWeight('bold').setBackground('#071626').setFontColor('#ffffff');
      if (!enc.length) {
        enc = claves;
        estilo(hoja.getRange(1, 1, 1, enc.length).setValues([enc]));
        hoja.setFrozenRows(1);
        prepararEstado(hoja);
      } else {
        claves.forEach(k => {
          if (enc.indexOf(k) < 0) { enc.push(k); estilo(hoja.getRange(1, enc.length).setValue(k)); }
        });
      }
      hoja.appendRow(enc.map(h =>
        h === 'Fecha' ? fecha :
        h === 'Estado' ? 'Nuevo' :
        h === 'Archivos' ? enlaces.join('\n') : limpio(v(h))));
    }

    // 3. Correo con PDF de resumen + planos adjuntos
    if (CORREOS) {
      const visita = v('Visita técnica') === 'Sí';
      const asunto = (visita ? '🔶 VISITA · ' : '') + (v('Tipo de proyecto') || 'Proyecto') + ' · ' + quien +
                     (v('Ciudad o municipio') ? ' · ' + v('Ciudad o municipio') : '') +
                     (v('Inicio') ? ' · Inicio: ' + v('Inicio') : '') + ' · ' + folio;
      const pdf = crearPdf(campos, fecha, enlaces, folio, quien, visita);
      const adjuntar = peso < 18 * 1024 * 1024;   // Gmail admite hasta 25 MB por correo
      MailApp.sendEmail({
        to: CORREOS,
        subject: asunto,
        body: 'Nuevo proyecto recibido desde www.grupoduralco.com/proyecto\n\n' +
              campos.filter(c => c.v).map(c => c.k + ': ' + c.v).join('\n') +
              (enlaces.length ? '\n\nArchivos en Drive:\n' + enlaces.join('\n') : '') +
              (enlaces.length && !adjuntar ? '\n(Los archivos son pesados: ábralos desde los enlaces de Drive.)' : '') +
              '\n\nEl resumen completo va adjunto en PDF.',
        attachments: adjuntar ? [pdf, ...blobs] : [pdf]
      });
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Lista desplegable y colores en la columna Estado
function prepararEstado(hoja) {
  const col = 2, filas = hoja.getMaxRows() - 1;
  const rango = hoja.getRange(2, col, filas, 1);
  rango.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(false).build());
  const reglas = ESTADOS.map(e => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(e).setBackground(COLORES[e]).setRanges([rango]).build());
  hoja.setConditionalFormatRules(hoja.getConditionalFormatRules().concat(reglas));
  hoja.setColumnWidth(col, 140);
}

function carpeta(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}
function zona() { return Session.getScriptTimeZone(); }
function limpio(x) {
  x = (x == null ? '' : String(x));
  return /^[=+\-@]/.test(x) ? "'" + x : x;
}

function crearPdf(campos, fecha, enlaces, folio, quien, visita) {
  const esc = x => String(x == null ? '' : x)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  const filas = campos.filter(c => c.k !== 'Folio').map(c =>
    '<tr><td class="k">' + esc(c.k) + '</td><td>' + (esc(c.v) || '—') + '</td></tr>').join('');
  const html = `
  <html><head><style>
    body{font-family:Arial,sans-serif;color:#1a2530;font-size:10.5pt;margin:0}
    .top{background:#071626;color:#fff;padding:22px 28px}
    .top h1{margin:0;font-size:20pt;letter-spacing:1px}
    .top p{margin:4px 0 0;color:#C99A2E;font-size:10pt;letter-spacing:2px;text-transform:uppercase}
    .box{padding:20px 28px}
    .vis{background:#FFF2CC;border-left:4px solid #C99A2E;padding:10px 12px;margin:0 0 12px;font-weight:bold}
    h2{font-size:12pt;color:#071626;border-bottom:2px solid #C99A2E;padding-bottom:4px;margin:6px 0 8px}
    table{width:100%;border-collapse:collapse}
    td{padding:7px 8px;border-bottom:1px solid #dde2e6;vertical-align:top}
    td.k{width:30%;color:#5b6772;font-weight:bold}
    .pie{margin-top:24px;font-size:8.5pt;color:#8c99a4;text-align:center}
  </style></head><body>
    <div class="top"><h1>DURALCO CONSTRUCTORES</h1><p>Solicitud de proyecto · ${esc(folio)}</p></div>
    <div class="box">
      ${visita ? '<div class="vis">Solicita visita técnica sin costo</div>' : ''}
      <h2>${esc(quien)} · ${Utilities.formatDate(fecha, zona(), 'dd/MM/yyyy HH:mm')}</h2>
      <table>${filas}</table>
      ${enlaces.length ? '<h2 style="margin-top:18px">Planos y archivos</h2><p>' + enlaces.map(esc).join('<br>') + '</p>' : ''}
      <div class="pie">Grupo Duralco · Duralco Constructores · www.grupoduralco.com/proyecto</div>
    </div>
  </body></html>`;
  return Utilities.newBlob(html, 'text/html', 'resumen.html')
    .getAs('application/pdf')
    .setName('Proyecto ' + folio + ' - ' + quien + '.pdf');
}

// Ejecútela una vez desde el editor para dar permisos y probar (no escribe en la hoja)
function probar() {
  procesar({
    prueba: true,
    campos: [{k: 'Folio', v: 'DC-PRUEBA'}, {k: 'Tipo de proyecto', v: 'Estructuras metálicas'},
             {k: 'Visita técnica', v: 'Sí'}, {k: 'Nombre', v: 'PRUEBA'}, {k: 'Ciudad o municipio', v: 'San Pedro Sula'}],
    archivos: [{nombre: 'plano-prueba.txt', tipo: 'text/plain', datos: Utilities.base64Encode('Plano de prueba')}]
  });
}
