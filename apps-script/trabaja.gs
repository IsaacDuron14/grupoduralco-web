// ==== Grupo Duralco — Trabaje con nosotros (Personal y Contratistas) ====
// Correos que reciben cada formulario (separados por coma)
const CORREOS = {
  Personal:    'gerente.operaciones@grupoduralco.com, ventas@grupoduralco.com, Ever.duron@grupoduralco.com',
  Contratista: 'gerente.operaciones@grupoduralco.com, ventas@grupoduralco.com, Ever.duron@grupoduralco.com'
};
const CARPETA = 'Postulaciones Grupo Duralco';   // carpeta en Google Drive

function doPost(e) {
  return procesar(JSON.parse(e.postData.contents));
}

function procesar(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tipo = d.tipo === 'Contratista' ? 'Contratista' : 'Personal';
    const pestaña = tipo === 'Contratista' ? 'Contratistas' : 'Personal';
    const campos = d.campos || [];
    const fecha = new Date();
    const v = k => { const c = campos.find(x => x.k === k); return c ? c.v : ''; };
    const quien = tipo === 'Contratista'
      ? (v('Empresa o razón social') || v('Representante o contacto'))
      : v('Nombre completo');

    // 1. Archivos a Google Drive
    const blobs = [], enlaces = [];
    if ((d.archivos || []).length) {
      const sub = carpeta(carpeta(DriveApp.getRootFolder(), CARPETA), pestaña);
      const persona = carpeta(sub, Utilities.formatDate(fecha, zona(), 'yyyy-MM-dd') + ' ' + (quien || 'sin nombre'));
      d.archivos.forEach(a => {
        const b = Utilities.newBlob(Utilities.base64Decode(a.datos), a.tipo, a.nombre);
        const f = persona.createFile(b);
        blobs.push(b);
        enlaces.push(f.getUrl());
      });
    }

    // 2. Registro en la hoja (pestaña Personal o Contratistas)
    //    Las columnas se buscan por nombre: si el formulario trae un campo nuevo,
    //    se agrega una columna al final sin desordenar las anteriores.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = d.prueba ? null : (ss.getSheetByName(pestaña) || ss.insertSheet(pestaña));
    if (hoja) {
      const claves = ['Fecha', ...campos.map(c => c.k), 'Archivos'];
      let enc = hoja.getLastRow() ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0] : [];
      const estilo = r => r.setFontWeight('bold').setBackground('#071626').setFontColor('#ffffff');
      if (!enc.length) {
        enc = claves;
        estilo(hoja.getRange(1, 1, 1, enc.length).setValues([enc]));
        hoja.setFrozenRows(1);
      } else {
        claves.forEach(k => {
          if (enc.indexOf(k) < 0) { enc.push(k); estilo(hoja.getRange(1, enc.length).setValue(k)); }
        });
      }
      hoja.appendRow(enc.map(h => h === 'Fecha' ? fecha : h === 'Archivos' ? enlaces.join('\n') : limpio(v(h))));
    }

    // 3. Correo con PDF de resumen + archivos adjuntos
    const para = CORREOS[tipo];
    if (para) {
      const titulo = tipo === 'Contratista' ? 'Registro de contratista' : 'Postulación de personal';
      const extra = tipo === 'Personal' && v('Puesto u oficio') ? ' (' + v('Puesto u oficio') + ')' : '';
      const pdf = crearPdf(titulo, campos, fecha, enlaces, quien);
      MailApp.sendEmail({
        to: para,
        subject: (tipo === 'Contratista' ? 'Contratista – ' : 'Postulación – ') + (quien || 'sin nombre') + extra,
        body: titulo + ' recibida desde www.grupoduralco.com/trabaja\n\n' +
              campos.filter(c => c.v).map(c => c.k + ': ' + c.v).join('\n') +
              (enlaces.length ? '\n\nArchivos en Drive:\n' + enlaces.join('\n') : '') +
              '\n\nEl resumen completo va adjunto en PDF.',
        attachments: [pdf, ...blobs]
      });
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
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

function crearPdf(titulo, campos, fecha, enlaces, quien) {
  const esc = x => String(x == null ? '' : x)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  const filas = campos.map(c =>
    '<tr><td class="k">' + esc(c.k) + '</td><td>' + (esc(c.v) || '—') + '</td></tr>').join('');
  const html = `
  <html><head><style>
    body{font-family:Arial,sans-serif;color:#1a2530;font-size:10.5pt;margin:0}
    .top{background:#071626;color:#fff;padding:22px 28px}
    .top h1{margin:0;font-size:20pt;letter-spacing:1px}
    .top p{margin:4px 0 0;color:#C99A2E;font-size:10pt;letter-spacing:2px;text-transform:uppercase}
    .box{padding:20px 28px}
    h2{font-size:12pt;color:#071626;border-bottom:2px solid #C99A2E;padding-bottom:4px;margin:6px 0 8px}
    table{width:100%;border-collapse:collapse}
    td{padding:7px 8px;border-bottom:1px solid #dde2e6;vertical-align:top}
    td.k{width:32%;color:#5b6772;font-weight:bold}
    .pie{margin-top:24px;font-size:8.5pt;color:#8c99a4;text-align:center}
  </style></head><body>
    <div class="top"><h1>GRUPO DURALCO</h1><p>${esc(titulo)}</p></div>
    <div class="box">
      <h2>${esc(quien || '')} · ${Utilities.formatDate(fecha, zona(), 'dd/MM/yyyy HH:mm')}</h2>
      <table>${filas}</table>
      ${enlaces.length ? '<h2 style="margin-top:18px">Archivos</h2><p>' + enlaces.map(esc).join('<br>') + '</p>' : ''}
      <div class="pie">Grupo Duralco · San Pedro Sula, Honduras · www.grupoduralco.com/trabaja</div>
    </div>
  </body></html>`;
  return Utilities.newBlob(html, 'text/html', 'resumen.html')
    .getAs('application/pdf')
    .setName(titulo + ' - ' + (quien || 'sin nombre') + '.pdf');
}

// Ejecútela una vez desde el editor para dar permisos y probar (no escribe en la hoja)
function probar() {
  procesar({
    prueba: true,
    tipo: 'Personal',
    campos: [{k: 'Nombre completo', v: 'PRUEBA'}, {k: 'Puesto u oficio', v: 'Soldador'}],
    archivos: [{nombre: 'prueba.txt', tipo: 'text/plain',
                datos: Utilities.base64Encode('Archivo de prueba')}]
  });
}
