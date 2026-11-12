// ==== Grupo Duralco — Cotizadores por división (www.grupoduralco.com/corebolt, etc.) ====
// Una sola hoja "Solicitudes Grupo Duralco": una pestaña por división (resumen) y otra con el detalle de productos.
const CORREOS = 'ventas@grupoduralco.com, Ever.duron@grupoduralco.com';
const CARPETA = 'Solicitudes Grupo Duralco';          // carpeta en Google Drive para planos
const ESTADOS = ['Nuevo', 'Contactado', 'Cotizado', 'Ganado', 'Perdido'];
const COLORES = {'Nuevo':'#FFF2CC','Contactado':'#CFE2F3','Cotizado':'#FCE5CD','Ganado':'#D9EAD3','Perdido':'#EFEFEF'};
const MARCA = {Corebolt: '#1E6A98', Mecatechn: '#071626', Alidpart: '#071626', Guacamaya: '#071626'};

function doPost(e) {
  return procesar(JSON.parse(e.postData.contents));
}

function procesar(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const div = d.division || 'General';
    const campos = d.campos || [];
    const items = d.items || [];
    const fecha = new Date();
    const v = k => { const c = campos.find(x => x.k === k); return c ? c.v : ''; };
    const folio = v('Folio') || d.folio || (div.slice(0, 2).toUpperCase() + '-' + fecha.getTime());
    const quien = v('Nombre o empresa') || 'sin nombre';
    const via = d.via || v('Enviado por') || 'Web';   // WhatsApp o Correo

    // 1. Planos a Drive
    const blobs = [], enlaces = [];
    let peso = 0;
    if ((d.archivos || []).length) {
      const dir = carpeta(carpeta(carpeta(DriveApp.getRootFolder(), CARPETA), div), folio + ' – ' + quien);
      d.archivos.forEach(a => {
        const b = Utilities.newBlob(Utilities.base64Decode(a.datos), a.tipo, a.nombre);
        peso += b.getBytes().length;
        blobs.push(b);
        enlaces.push(dir.createFile(b).getUrl());
      });
    }

    // 2. Hoja: pestaña de la división (una fila por solicitud) + pestaña de detalle (una fila por producto)
    if (!d.prueba) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const resumen = items.map((it, i) => (i + 1) + '. ' + it.producto + ' × ' + it.cantidad).join('\n');
      fila(ss, div, ['Fecha', 'Estado', ...campos.map(c => c.k), 'Resumen', 'Archivos'], true, h =>
        h === 'Fecha' ? fecha : h === 'Estado' ? 'Nuevo' : h === 'Resumen' ? resumen :
        h === 'Archivos' ? enlaces.join('\n') : limpio(v(h)));
      items.forEach((it, i) => fila(ss, div + ' · detalle',
        ['Fecha', 'Folio', 'Cliente', '#', 'Categoría', 'Producto', 'Descripción técnica', 'Cantidad'], false, h =>
        h === 'Fecha' ? fecha : h === 'Folio' ? folio : h === 'Cliente' ? limpio(quien) : h === '#' ? i + 1 :
        h === 'Categoría' ? limpio(it.categoria) : h === 'Producto' ? limpio(it.producto) :
        h === 'Descripción técnica' ? limpio(it.descripcion) : Number(it.cantidad) || limpio(it.cantidad)));
    }

    // 3. Correo con PDF + planos
    if (CORREOS) {
      const asunto = (via === 'WhatsApp' ? '[WhatsApp] ' : '[Correo] ') + 'Cotización ' + div.toUpperCase() + ' · ' + quien +
                     (v('Ciudad de entrega') ? ' · ' + v('Ciudad de entrega') : '') +
                     ' · ' + items.length + ' producto(s) · ' + folio;
      const pdf = crearPdf(div, campos, items, fecha, enlaces, folio, quien);
      const adjuntar = peso < 18 * 1024 * 1024;
      MailApp.sendEmail({
        to: CORREOS,
        replyTo: /@/.test(v('Correo')) ? v('Correo') : 'ventas@grupoduralco.com',
        subject: asunto,
        body: 'Nueva solicitud de cotización ' + div + ' desde www.grupoduralco.com (enviada por ' + via + ')\n' +
              (via === 'WhatsApp' ? 'El cliente también la envió por WhatsApp' + (v('Enviado a') ? ' al ' + v('Enviado a') : '') + '.\n' : '') + '\n' +
              items.map((it, i) => (i + 1) + '. ' + it.producto + ' × ' + it.cantidad + '\n   ' + it.descripcion).join('\n') +
              '\n\n' + campos.filter(c => c.v).map(c => c.k + ': ' + c.v).join('\n') +
              (enlaces.length ? '\n\nArchivos en Drive:\n' + enlaces.join('\n') : '') +
              '\n\nEl resumen completo va adjunto en PDF.',
        attachments: adjuntar ? [pdf, ...blobs] : [pdf]
      });
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Agrega una fila mapeando columnas por nombre (crea la pestaña y columnas que falten)
function fila(ss, nombre, claves, conEstado, valor) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = ss.insertSheet(nombre);
  const estilo = r => r.setFontWeight('bold').setBackground('#071626').setFontColor('#ffffff');
  let enc = hoja.getLastRow() ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0] : [];
  if (!enc.length) {
    enc = claves;
    estilo(hoja.getRange(1, 1, 1, enc.length).setValues([enc]));
    hoja.setFrozenRows(1);
    if (conEstado) prepararEstado(hoja);
  } else {
    claves.forEach(k => { if (enc.indexOf(k) < 0) { enc.push(k); estilo(hoja.getRange(1, enc.length).setValue(k)); } });
  }
  hoja.appendRow(enc.map(valor));
}

function prepararEstado(hoja) {
  const rango = hoja.getRange(2, 2, hoja.getMaxRows() - 1, 1);
  rango.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(false).build());
  hoja.setConditionalFormatRules(hoja.getConditionalFormatRules().concat(ESTADOS.map(e =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(e).setBackground(COLORES[e]).setRanges([rango]).build())));
  hoja.setColumnWidth(2, 120);
}

function carpeta(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}
function limpio(x) {
  x = (x == null ? '' : String(x));
  return /^[=+\-@]/.test(x) ? "'" + x : x;
}

function crearPdf(div, campos, items, fecha, enlaces, folio, quien) {
  const esc = x => String(x == null ? '' : x)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  const color = MARCA[div] || '#071626';
  const prods = items.map((it, i) => '<tr><td class="n">' + (i + 1) + '</td><td><b>' + esc(it.producto) + '</b></td><td>' +
    esc(it.descripcion) + '</td><td class="q">' + esc(Number(it.cantidad).toLocaleString('es-HN')) + '</td></tr>').join('');
  const datos = campos.filter(c => c.k !== 'Folio' && c.k !== 'Productos').map(c =>
    '<tr><td class="k">' + esc(c.k) + '</td><td>' + (esc(c.v) || '—') + '</td></tr>').join('');
  const html = `
  <html><head><style>
    body{font-family:Arial,sans-serif;color:#1a2530;font-size:10pt;margin:0}
    .top{background:#071626;color:#fff;padding:20px 28px;border-bottom:5px solid ${color}}
    .top h1{margin:0;font-size:20pt;letter-spacing:2px}
    .top p{margin:4px 0 0;color:#A9C6D4;font-size:9.5pt;letter-spacing:2px;text-transform:uppercase}
    .box{padding:18px 28px}
    h2{font-size:12pt;color:#071626;border-bottom:2px solid ${color};padding-bottom:4px;margin:14px 0 8px}
    table{width:100%;border-collapse:collapse}
    th{text-align:left;font-size:8.5pt;color:#5b6772;border-bottom:2px solid #071626;padding:6px}
    td{padding:7px 6px;border-bottom:1px solid #dde2e6;vertical-align:top}
    td.k{width:30%;color:#5b6772;font-weight:bold} td.n{width:22px;color:#8c99a4} td.q{text-align:right;white-space:nowrap;font-weight:bold}
    .pie{margin-top:22px;font-size:8.5pt;color:#8c99a4;text-align:center}
  </style></head><body>
    <div class="top"><h1>${esc(div.toUpperCase())}</h1><p>Solicitud de cotización · ${esc(folio)} · ${Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}</p></div>
    <div class="box">
      <h2>Productos solicitados (${items.length})</h2>
      <table><tr><th></th><th>Producto</th><th>Descripción técnica</th><th style="text-align:right">Cantidad</th></tr>${prods}</table>
      <h2>Cliente · ${esc(quien)}</h2>
      <table>${datos}</table>
      ${enlaces.length ? '<h2>Planos y archivos</h2><p>' + enlaces.map(esc).join('<br>') + '</p>' : ''}
      <div class="pie">Grupo Duralco · ${esc(div)} · www.grupoduralco.com</div>
    </div>
  </body></html>`;
  return Utilities.newBlob(html, 'text/html', 'resumen.html').getAs('application/pdf')
    .setName('Cotización ' + div + ' ' + folio + ' - ' + quien + '.pdf');
}

// Ejecútela una vez desde el editor para dar permisos y probar (no escribe en la hoja)
function probar() {
  procesar({
    prueba: true, division: 'Corebolt', via: 'WhatsApp',
    campos: [{k: 'Folio', v: 'CB-PRUEBA'}, {k: 'Nombre o empresa', v: 'PRUEBA'}, {k: 'Ciudad de entrega', v: 'San Pedro Sula'}],
    items: [{categoria: 'Pernos', producto: 'Perno J', descripcion: 'Ø 3/4" · L = 24" · rosca 6" · UNC · ASTM F1554 Grado 36', cantidad: 40}],
    archivos: [{nombre: 'plano-prueba.txt', tipo: 'text/plain', datos: Utilities.base64Encode('Plano de prueba')}]
  });
}
