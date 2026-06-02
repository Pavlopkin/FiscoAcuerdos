// Genera un .docx mínimo válido usando JSZip (sin depender de docx.js)
// Requiere en el HTML:
//   <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
//   <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>

function descargarAcuerdoWord() {

  if (typeof JSZip === 'undefined') {
    alert('Error: JSZip no está cargado.');
    return;
  }

  // ── Datos del formulario ─────────────────────────────────

  var apoderadoVal = document.getElementById('ac-apoderado').value;
  var apoderadoMap = { natalia: 'Natalia', mauricio: 'Mauricio', maximina: 'Maximina' };
  var apoderado    = apoderadoMap[apoderadoVal] || apoderadoVal;

  var juzgadoNum = document.getElementById('ac-juzgado').value;
  var juzgado    = 'Juzgado N\u00b0 ' + juzgadoNum;

  var expediente = document.getElementById('ac-expediente').value.trim() || '_______________';
  var tipoId     = ((document.querySelector('input[name="identificacion"]:checked') || {}).value || 'DNI').toUpperCase();
  var numeroId   = document.getElementById('ac-dni').value.trim() || '_______________';
  var nombre     = document.getElementById('ac-contribuyente').value.trim() || '_______________';
  var domicilio  = document.getElementById('ac-domicilio').value.trim() || '_______________';

  var caracterVal = ((document.querySelector('input[name="caracter"]:checked') || {}).value || '');
  var caracterMap = { demandado: 'Demandado', apoderado: 'Apoderado', representante: 'Representante legal' };
  var caracter    = caracterMap[caracterVal] || '_______________';

  var cantTitulos = ((document.querySelector('input[name="cantTitulos"]:checked') || {}).value || '0');
  var titulos = [];
  for (var i = 1; i <= parseInt(cantTitulos); i++) {
    var elP = document.getElementById('titulo-periodo-' + i);
    var elM = document.getElementById('titulo-monto-' + i);
    var per = elP ? elP.value.trim() : '';
    var mon = elM ? elM.value.trim() : '';
    if (per || mon) titulos.push({ n: i, periodo: per, monto: mon });
  }

  var hoy   = new Date();
  var fecha = hoy.toLocaleDateString('es-AR', { day: '2-digit', month: 'long', year: 'numeric' });
  var elTexto = document.getElementById('acuerdo-texto');
  var textoAcuerdo = elTexto ? elTexto.innerText.trim() : '';

  // ── Helpers XML ──────────────────────────────────────────

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Párrafo: texto, negrita, tamaño (en halfpoints), alineación, color
  function parXml(texto, opts) {
    opts = opts || {};
    var sz   = opts.sz   || 22;
    var bold = opts.bold ? '<w:b/>' : '';
    var jc   = opts.jc   ? '<w:jc w:val="' + opts.jc + '"/>' : '';
    var col  = opts.color ? '<w:color w:val="' + opts.color + '"/>' : '';
    var spb  = opts.spb  != null ? opts.spb  : 80;
    var spa  = opts.spa  != null ? opts.spa  : 80;
    return '<w:p>' +
      '<w:pPr>' + jc + '<w:spacing w:before="' + spb + '" w:after="' + spa + '"/></w:pPr>' +
      '<w:r><w:rPr>' + bold + '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>' + col + '</w:rPr>' +
      '<w:t xml:space="preserve">' + esc(texto) + '</w:t></w:r>' +
    '</w:p>';
  }

  // Celda de tabla
  function celdaXml(texto, opts) {
    opts = opts || {};
    var w     = opts.w || 4500;
    var bold  = opts.bold  ? '<w:b/>' : '';
    var fill  = opts.fill  || 'FFFFFF';
    var col   = opts.color || '1A1A2E';
    var sz    = opts.sz    || 20;
    var bTop  = opts.bTop  || 'C8D0DB';
    var bSz   = opts.bSzTop || 4;
    var borderDef = function(color, sz) {
      return 'w:val="single" w:sz="' + (sz||4) + '" w:space="0" w:color="' + color + '"';
    };
    var borderColor = opts.borderColor || 'C8D0DB';
    return '<w:tc>' +
      '<w:tcPr>' +
        '<w:tcW w:w="' + w + '" w:type="dxa"/>' +
        '<w:tcBorders>' +
          '<w:top '    + borderDef(opts.bTop || borderColor, opts.bSzTop || 4) + '/>' +
          '<w:left '   + borderDef(borderColor, 4) + '/>' +
          '<w:bottom ' + borderDef(borderColor, 4) + '/>' +
          '<w:right '  + borderDef(borderColor, 4) + '/>' +
        '</w:tcBorders>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>' +
        '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>' +
      '</w:tcPr>' +
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>' +
      '<w:r><w:rPr>' + bold + '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/><w:color w:val="' + col + '"/></w:rPr>' +
      '<w:t xml:space="preserve">' + esc(texto) + '</w:t></w:r></w:p>' +
    '</w:tc>';
  }

  function filaXml(celdas) {
    return '<w:tr>' + celdas + '</w:tr>';
  }

  function tablaXml(totalW, colWidths, filas) {
    var cols = colWidths.map(function(w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('');
    return '<w:tbl>' +
      '<w:tblPr>' +
        '<w:tblW w:w="' + totalW + '" w:type="dxa"/>' +
        '<w:tblBorders>' +
          '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="C8D0DB"/>' +
          '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="C8D0DB"/>' +
        '</w:tblBorders>' +
      '</w:tblPr>' +
      '<w:tblGrid>' + cols + '</w:tblGrid>' +
      filas.join('') +
    '</w:tbl>';
  }

  // ── Construir contenido ──────────────────────────────────

  var CONT = 9638; // A4 con márgenes 2cm aprox
  var c1 = Math.round(CONT * 0.40), c2 = CONT - c1;

  var partes = [];

  // Encabezado
  partes.push(parXml('FISCO DE LA PROVINCIA DE BUENOS AIRES', { sz: 26, bold: true, jc: 'center', spb: 0, spa: 40 }));
  partes.push(parXml('Fuero Comercial de la Naci\u00f3n', { sz: 20, jc: 'center', color: '555577', spb: 0, spa: 200 }));
  partes.push('<w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D0D7E3"/></w:pBdr></w:pPr><w:r><w:t></w:t></w:r></w:p>');
  partes.push(parXml('ACUERDO DE PAGO', { sz: 28, bold: true, jc: 'center', spb: 240, spa: 60 }));
  partes.push(parXml('Buenos Aires, ' + fecha, { sz: 20, jc: 'center', color: '666666', spb: 0, spa: 280 }));

  // Tabla de datos
  partes.push(parXml('Datos del acuerdo', { sz: 22, bold: true, spb: 160, spa: 100 }));
  partes.push(tablaXml(CONT, [c1, c2], [
    filaXml(
      celdaXml('Campo',     { w: c1, bold: true, fill: 'E8EDF5', color: '2C3A5A', sz: 20 }) +
      celdaXml('Valor',     { w: c2, bold: true, fill: 'E8EDF5', color: '2C3A5A', sz: 20 })
    ),
    filaXml(celdaXml('Apoderado/a',                     { w: c1 }) + celdaXml(apoderado,   { w: c2, bold: true })),
    filaXml(celdaXml('Juzgado',                         { w: c1 }) + celdaXml(juzgado,     { w: c2 })),
    filaXml(celdaXml('Expediente N\u00b0',              { w: c1 }) + celdaXml(expediente,  { w: c2 })),
    filaXml(celdaXml('N\u00b0 de ' + tipoId,            { w: c1 }) + celdaXml(numeroId,   { w: c2 })),
    filaXml(celdaXml('Apellido y nombre',               { w: c1 }) + celdaXml(nombre,      { w: c2, bold: true })),
    filaXml(celdaXml('Domicilio',                       { w: c1 }) + celdaXml(domicilio,   { w: c2 })),
    filaXml(celdaXml('Car\u00e1cter del firmante',       { w: c1 }) + celdaXml(caracter,   { w: c2 })),
    filaXml(celdaXml('Cantidad de t\u00edtulos',         { w: c1 }) + celdaXml(cantTitulos, { w: c2 })),
  ]));

  // Títulos ejecutivos
  if (titulos.length > 0) {
    var ct1 = Math.round(CONT * 0.12), ct2 = Math.round(CONT * 0.50), ct3 = CONT - ct1 - ct2;
    partes.push(parXml('T\u00edtulos ejecutivos', { sz: 22, bold: true, spb: 240, spa: 100 }));
    var filasTit = [filaXml(
      celdaXml('N\u00b0',       { w: ct1, bold: true, fill: 'E8EDF5', color: '2C3A5A' }) +
      celdaXml('Per\u00edodo',  { w: ct2, bold: true, fill: 'E8EDF5', color: '2C3A5A' }) +
      celdaXml('Monto',         { w: ct3, bold: true, fill: 'E8EDF5', color: '2C3A5A' })
    )];
    titulos.forEach(function(t) {
      filasTit.push(filaXml(
        celdaXml(String(t.n),                   { w: ct1 }) +
        celdaXml(t.periodo,                     { w: ct2 }) +
        celdaXml(t.monto ? '$ ' + t.monto : '', { w: ct3 })
      ));
    });
    partes.push(tablaXml(CONT, [ct1, ct2, ct3], filasTit));
  }

  // Texto del acuerdo
  if (textoAcuerdo) {
    partes.push(parXml('Texto del acuerdo', { sz: 22, bold: true, spb: 240, spa: 100 }));
    textoAcuerdo.split('\n').filter(function(l) { return l.trim(); }).forEach(function(l) {
      partes.push(parXml(l, { sz: 22, spb: 100, spa: 60 }));
    });
  }

  // Firmas: tabla con 2 columnas, borde superior en cada celda
  partes.push(parXml('', { spb: 400, spa: 0 }));
  var cf = Math.floor(CONT / 2);
  var nombreFirma  = nombre !== '_______________' ? nombre : ' ';
  var pieDemandado = caracter + ' \u2014 ' + tipoId + ' ' + numeroId;

  function celdaFirmaXml(w, l1, l2) {
    return '<w:tc>' +
      '<w:tcPr>' +
        '<w:tcW w:w="' + w + '" w:type="dxa"/>' +
        '<w:tcBorders>' +
          '<w:top    w:val="single" w:sz="12" w:space="0" w:color="333333"/>' +
          '<w:left   w:val="none"   w:sz="0"  w:space="0" w:color="FFFFFF"/>' +
          '<w:bottom w:val="none"   w:sz="0"  w:space="0" w:color="FFFFFF"/>' +
          '<w:right  w:val="none"   w:sz="0"  w:space="0" w:color="FFFFFF"/>' +
        '</w:tcBorders>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>' +
        '<w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="200" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tcMar>' +
      '</w:tcPr>' +
      '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="60"/></w:pPr>' +
        '<w:r><w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="1A1A2E"/></w:rPr>' +
        '<w:t xml:space="preserve">' + esc(l1) + '</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>' +
        '<w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="666666"/></w:rPr>' +
        '<w:t xml:space="preserve">' + esc(l2) + '</w:t></w:r></w:p>' +
    '</w:tc>';
  }

  partes.push(tablaXml(CONT, [cf, CONT - cf], [
    filaXml(celdaFirmaXml(cf, apoderado, 'Apoderado/a fiscal') + celdaFirmaXml(CONT - cf, nombreFirma, pieDemandado))
  ]));

  // ── Armar archivos del ZIP ───────────────────────────────

  var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"' +
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' +
    ' xmlns:v="urn:schemas-microsoft-com:vml"' +
    ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:w10="urn:schemas-microsoft-com:office:word"' +
    ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' +
    ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' +
    ' xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"' +
    ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"' +
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"' +
    ' mc:Ignorable="w14 wp14">' +
    '<w:body>' +
    partes.join('') +
    '<w:sectPr>' +
      '<w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="709" w:footer="709" w:gutter="0"/>' +
    '</w:sectPr>' +
    '</w:body></w:document>';

  var relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
    '</w:rPr></w:rPrDefault></w:docDefaults>' +
    '</w:styles>';

  var workbookRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  var contentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml"  ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>';

  // ── Empaquetar y descargar ───────────────────────────────

  var zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml);
  zip.file('_rels/.rels', workbookRelsXml);
  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/_rels/document.xml.rels', relsXml);

  zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    .then(function(blob) {
      var nombreLimpio = nombre !== '_______________'
        ? nombre.replace(/[^a-zA-Z\u00c0-\u024f\s]/g, '').trim().replace(/\s+/g, '_')
        : 'contribuyente';
      var expLimpio = expediente !== '_______________'
        ? expediente.replace(/\//g, '-')
        : 'sin_expediente';
      saveAs(blob, 'Acuerdo_' + nombreLimpio + '_' + expLimpio + '.docx');
    })
    .catch(function(err) {
      console.error('Error al generar Word:', err);
      alert('No se pudo generar el archivo.');
    });
}
