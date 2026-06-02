// Genera el acuerdo en .docx con el mismo formato que el PDF impreso
// Requiere en el HTML (antes de </body>):
//   <script src="https://unpkg.com/jszip@3.10.1/dist/jszip.min.js"></script>
//   <script src="https://unpkg.com/file-saver@2.0.5/dist/FileSaver.min.js"></script>

function descargarAcuerdoWord() {

  if (typeof JSZip === 'undefined') {
    alert('Error: JSZip no está cargado.');
    return;
  }

  // ── Leer el texto exacto que muestra la pantalla ─────────
  var elTexto = document.getElementById('acuerdo-texto');
  if (!elTexto) { alert('No hay acuerdo generado.'); return; }

  // Extraer párrafos del HTML generado por generarAcuerdo()
  // El contenido tiene <p> con texto, algunos con <strong> (negrita)
  var parrafos = [];
  var nodos = elTexto.childNodes;
  for (var i = 0; i < nodos.length; i++) {
    var nodo = nodos[i];
    if (nodo.nodeType === Node.ELEMENT_NODE) {
      // Cada <p> o <div> es un párrafo
      parrafos.push({ html: nodo.innerHTML, tag: nodo.tagName });
    } else if (nodo.nodeType === Node.TEXT_NODE && nodo.textContent.trim()) {
      parrafos.push({ html: nodo.textContent, tag: 'P' });
    }
  }

  // Datos para nombre de archivo
  var nombre     = (document.getElementById('ac-contribuyente') || {}).value || '';
  var expediente = (document.getElementById('ac-expediente')    || {}).value || '';
  nombre     = nombre.trim()     || 'contribuyente';
  expediente = expediente.trim() || 'sin_expediente';

  // ── Helpers XML ──────────────────────────────────────────

  function esc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Convierte el innerHTML de un párrafo en runs de Word
  // Soporta <strong>, <b>, <em>, <i>, texto plano
  function htmlARuns(html) {
    // Parsear el HTML en un DOM temporal
    var tmp = document.createElement('div');
    tmp.innerHTML = html;

    var runs = '';
    function procesarNodo(nodo, bold, italic) {
      if (nodo.nodeType === Node.TEXT_NODE) {
        var txt = nodo.textContent;
        if (!txt) return;
        runs += '<w:r><w:rPr>';
        if (bold)   runs += '<w:b/><w:bCs/>';
        if (italic) runs += '<w:i/><w:iCs/>';
        runs += '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>';
        runs += '<w:t xml:space="preserve">' + esc(txt) + '</w:t></w:r>';
      } else if (nodo.nodeType === Node.ELEMENT_NODE) {
        var tag = nodo.tagName.toUpperCase();
        var esBold   = bold   || tag === 'STRONG' || tag === 'B';
        var esItalic = italic || tag === 'EM'     || tag === 'I';
        for (var j = 0; j < nodo.childNodes.length; j++) {
          procesarNodo(nodo.childNodes[j], esBold, esItalic);
        }
      }
    }
    for (var k = 0; k < tmp.childNodes.length; k++) {
      procesarNodo(tmp.childNodes[k], false, false);
    }
    return runs;
  }

  // Párrafo Word con runs mixtos (bold/normal dentro del mismo párrafo)
  function parMixto(html, opts) {
    opts = opts || {};
    var spb = opts.spb != null ? opts.spb : 0;
    var spa = opts.spa != null ? opts.spa : 160;
    var jc  = opts.jc  || 'both'; // justificado
    var runs = htmlARuns(html);
    if (!runs) runs = '<w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t></w:t></w:r>';
    return '<w:p>' +
      '<w:pPr>' +
        '<w:jc w:val="' + jc + '"/>' +
        '<w:spacing w:before="' + spb + '" w:after="' + spa + '" w:line="276" w:lineRule="auto"/>' +
      '</w:pPr>' +
      runs +
    '</w:p>';
  }

  // Párrafo simple (sin HTML interno)
  function parSimple(texto, opts) {
    opts = opts || {};
    var spb  = opts.spb  != null ? opts.spb  : 0;
    var spa  = opts.spa  != null ? opts.spa  : 160;
    var jc   = opts.jc   || 'both';
    var sz   = opts.sz   || 22;
    var bold = opts.bold || false;
    return '<w:p>' +
      '<w:pPr>' +
        '<w:jc w:val="' + jc + '"/>' +
        '<w:spacing w:before="' + spb + '" w:after="' + spa + '" w:line="276" w:lineRule="auto"/>' +
      '</w:pPr>' +
      '<w:r><w:rPr>' +
        (bold ? '<w:b/><w:bCs/>' : '') +
        '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>' +
      '</w:rPr><w:t xml:space="preserve">' + esc(texto) + '</w:t></w:r>' +
    '</w:p>';
  }

  // Celda de firma
  function celdaFirma(w, texto) {
    return '<w:tc>' +
      '<w:tcPr>' +
        '<w:tcW w:w="' + w + '" w:type="dxa"/>' +
        '<w:tcBorders>' +
          '<w:top    w:val="single" w:sz="8"  w:space="0" w:color="333333"/>' +
          '<w:left   w:val="none"   w:sz="0"  w:space="0" w:color="auto"/>' +
          '<w:bottom w:val="none"   w:sz="0"  w:space="0" w:color="auto"/>' +
          '<w:right  w:val="none"   w:sz="0"  w:space="0" w:color="auto"/>' +
        '</w:tcBorders>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>' +
        '<w:tcMar>' +
          '<w:top    w:w="120" w:type="dxa"/>' +
          '<w:left   w:w="0"   w:type="dxa"/>' +
          '<w:bottom w:w="0"   w:type="dxa"/>' +
          '<w:right  w:w="0"   w:type="dxa"/>' +
        '</w:tcMar>' +
      '</w:tcPr>' +
      '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>' +
        '<w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="444444"/></w:rPr>' +
        '<w:t xml:space="preserve">' + esc(texto) + '</w:t></w:r>' +
      '</w:p>' +
    '</w:tc>';
  }

  // ── Construir el documento ───────────────────────────────

  var CONT = 9638; // A4 con márgenes ~2.5cm
  var kids = [];

  // Título centrado
  kids.push(parSimple('Acuerdo de Pago', { jc: 'center', sz: 26, bold: true, spb: 0, spa: 320 }));

  // Párrafos del acuerdo (texto generado por la app)
  for (var i = 0; i < parrafos.length; i++) {
    kids.push(parMixto(parrafos[i].html, { spb: 0, spa: 160, jc: 'both' }));
  }

  // Espacio antes de firmas
  kids.push(parSimple('', { spb: 0, spa: 0 }));
  kids.push(parSimple('', { spb: 0, spa: 0 }));

  // Tabla de firmas (2 columnas sin bordes laterales, borde top como línea)
  var cf = Math.floor(CONT / 2);
  kids.push(
    '<w:tbl>' +
      '<w:tblPr>' +
        '<w:tblW w:w="' + CONT + '" w:type="dxa"/>' +
        '<w:tblBorders>' +
          '<w:top    w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
          '<w:left   w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
          '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
          '<w:right  w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
          '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
          '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '</w:tblBorders>' +
        '<w:tblCellMar>' +
          '<w:left  w:w="0" w:type="dxa"/>' +
          '<w:right w:w="0" w:type="dxa"/>' +
        '</w:tblCellMar>' +
      '</w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="' + cf + '"/><w:gridCol w:w="' + (CONT - cf) + '"/></w:tblGrid>' +
      '<w:tr>' +
        celdaFirma(cf,          'Firma del/la apoderado/a fiscal') +
        celdaFirma(CONT - cf,   'Firma del demandado') +
      '</w:tr>' +
    '</w:tbl>'
  );

  // ── Ensamblar document.xml ───────────────────────────────

  var documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"' +
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' +
    ' xmlns:v="urn:schemas-microsoft-com:vml"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:w10="urn:schemas-microsoft-com:office:word"' +
    ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' +
    ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"' +
    ' mc:Ignorable="w14">' +
    '<w:body>' +
    kids.join('') +
    '<w:sectPr>' +
      '<w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/>' +
    '</w:sectPr>' +
    '</w:body></w:document>';

  var relsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  var stylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' +
      '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
      '<w:lang w:val="es-AR"/>' +
    '</w:rPr></w:rPrDefault></w:docDefaults>' +
    '</w:styles>';

  var dotRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  var contentTypesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml"  ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>';

  // ── Empaquetar y descargar ───────────────────────────────

  var zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml);
  zip.file('_rels/.rels',         dotRelsXml);
  zip.file('word/document.xml',   documentXml);
  zip.file('word/styles.xml',     stylesXml);
  zip.file('word/_rels/document.xml.rels', relsXml);

  zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  }).then(function(blob) {
    var nombreLimpio = nombre.replace(/[^a-zA-Z\u00c0-\u024f\s]/g, '').trim().replace(/\s+/g, '_') || 'contribuyente';
    var expLimpio    = expediente.replace(/\//g, '-') || 'sin_expediente';
    saveAs(blob, 'Acuerdo_' + nombreLimpio + '_' + expLimpio + '.docx');
  }).catch(function(err) {
    console.error('Error al generar Word:', err);
    alert('No se pudo generar el archivo. Ver consola.');
  });
}
