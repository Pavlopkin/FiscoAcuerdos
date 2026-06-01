function descargarAcuerdoWord() {

  var Document      = window.Document      || (window.docx && window.docx.Document);
  var Packer        = window.Packer        || (window.docx && window.docx.Packer);
  var Paragraph     = window.Paragraph     || (window.docx && window.docx.Paragraph);
  var TextRun       = window.TextRun       || (window.docx && window.docx.TextRun);
  var Table         = window.Table         || (window.docx && window.docx.Table);
  var TableRow      = window.TableRow      || (window.docx && window.docx.TableRow);
  var TableCell     = window.TableCell     || (window.docx && window.docx.TableCell);
  var AlignmentType = window.AlignmentType || (window.docx && window.docx.AlignmentType);
  var BorderStyle   = window.BorderStyle   || (window.docx && window.docx.BorderStyle);
  var WidthType     = window.WidthType     || (window.docx && window.docx.WidthType);

  if (!Document || !Packer) {
    alert('Error: la librería docx no se cargó correctamente.');
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

  // ── Layout ───────────────────────────────────────────────

  var PW = 11906, ML = 1134, CONT = PW - ML * 2;

  var bVis = { style: BorderStyle.SINGLE, size: 1, color: 'C8D0DB' };
  var bInv = { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' };
  var bordesVis = { top: bVis, bottom: bVis, left: bVis, right: bVis };
  var bordesInv = { top: bInv, bottom: bInv, left: bInv, right: bInv };

  // ── Helpers ──────────────────────────────────────────────

  function mkPar(texto, opts) {
    opts = opts || {};
    var runCfg = { text: texto, font: 'Arial', size: opts.size || 22, color: opts.color || '1A1A2E' };
    if (opts.bold) runCfg.bold = true;
    var parCfg = {
      spacing: { before: opts.antes != null ? opts.antes : 80, after: opts.despues != null ? opts.despues : 80 },
      children: [new TextRun(runCfg)]
    };
    if (opts.center) parCfg.alignment = AlignmentType.CENTER;
    return new Paragraph(parCfg);
  }

  function sep() {
    return new Paragraph({
      spacing: { before: 0, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D0D7E3', space: 1 } },
      children: [new TextRun({ text: '', font: 'Arial', size: 22 })]
    });
  }

  function mkCell(texto, w, opts) {
    opts = opts || {};
    var fill = opts.header ? 'E8EDF5' : 'FFFFFF';
    var runCfg = { text: texto, font: 'Arial', size: 20, color: opts.header ? '2C3A5A' : '1A1A2E' };
    if (opts.bold || opts.header) runCfg.bold = true;
    return new TableCell({
      borders: bordesVis,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      shading: { fill: fill },
      children: [new Paragraph({ children: [new TextRun(runCfg)] })]
    });
  }

  // Celda de firma: línea superior como borde de celda (top), no de párrafo
  function mkFirmaCell(w, linea1, linea2) {
    var topBorder = { style: BorderStyle.SINGLE, size: 8, color: '333333' };
    var borders = {
      top:    topBorder,
      bottom: bInv,
      left:   bInv,
      right:  bInv
    };
    return new TableCell({
      borders: borders,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: linea1, font: 'Arial', size: 20, bold: true, color: '1A1A2E' })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: linea2, font: 'Arial', size: 18, color: '666666' })]
        })
      ]
    });
  }

  function mkRow() { return new TableRow({ children: Array.from(arguments) }); }

  // ── Tabla de datos ───────────────────────────────────────

  var c1 = Math.round(CONT * 0.40), c2 = CONT - c1;
  var tablaDatos = new Table({
    width: { size: CONT, type: WidthType.DXA },
    columnWidths: [c1, c2],
    rows: [
      mkRow(mkCell('Campo', c1, { header: true }),           mkCell('Valor', c2, { header: true })),
      mkRow(mkCell('Apoderado/a', c1),                       mkCell(apoderado, c2, { bold: true })),
      mkRow(mkCell('Juzgado', c1),                           mkCell(juzgado, c2)),
      mkRow(mkCell('Expediente N\u00b0', c1),                mkCell(expediente, c2)),
      mkRow(mkCell('N\u00b0 de ' + tipoId, c1),              mkCell(numeroId, c2)),
      mkRow(mkCell('Apellido y nombre', c1),                 mkCell(nombre, c2, { bold: true })),
      mkRow(mkCell('Domicilio', c1),                         mkCell(domicilio, c2)),
      mkRow(mkCell('Car\u00e1cter del firmante', c1),         mkCell(caracter, c2)),
      mkRow(mkCell('Cantidad de t\u00edtulos', c1),           mkCell(cantTitulos.toString(), c2)),
    ]
  });

  // ── Tabla de títulos ejecutivos ──────────────────────────

  var tablaTitulos = null;
  if (titulos.length > 0) {
    var ct1 = Math.round(CONT * 0.12), ct2 = Math.round(CONT * 0.50), ct3 = CONT - ct1 - ct2;
    var rowsTit = [mkRow(
      mkCell('N\u00b0', ct1, { header: true }),
      mkCell('Per\u00edodo', ct2, { header: true }),
      mkCell('Monto', ct3, { header: true })
    )];
    titulos.forEach(function(t) {
      rowsTit.push(mkRow(
        mkCell(String(t.n), ct1),
        mkCell(t.periodo, ct2),
        mkCell(t.monto ? ('$ ' + t.monto) : '', ct3)
      ));
    });
    tablaTitulos = new Table({
      width: { size: CONT, type: WidthType.DXA },
      columnWidths: [ct1, ct2, ct3],
      rows: rowsTit
    });
  }

  // ── Texto del acuerdo ────────────────────────────────────

  var parsAcuerdo = textoAcuerdo.split('\n').filter(function(l) { return l.trim(); }).map(function(l) {
    return mkPar(l, { antes: 100, despues: 60 });
  });

  // ── Tabla de firmas ──────────────────────────────────────
  // La línea superior se pone como borde TOP de la celda (no del párrafo)
  // y se agrega espacio vacío arriba con una fila invisible de relleno

  var cf = Math.floor(CONT / 2);
  var nombreFirma  = nombre !== '_______________' ? nombre : ' ';
  var pieDemandado = caracter + ' \u2014 ' + tipoId + ' ' + numeroId;

  // Fila de espacio antes de las firmas
  var filaEspacio = new TableRow({
    children: [
      new TableCell({
        borders: bordesInv,
        width: { size: cf, type: WidthType.DXA },
        children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '', font: 'Arial', size: 22 })] })]
      }),
      new TableCell({
        borders: bordesInv,
        width: { size: CONT - cf, type: WidthType.DXA },
        children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '', font: 'Arial', size: 22 })] })]
      })
    ]
  });

  var tablaFirmas = new Table({
    width: { size: CONT, type: WidthType.DXA },
    columnWidths: [cf, CONT - cf],
    rows: [
      filaEspacio,
      mkRow(
        mkFirmaCell(cf, apoderado, 'Apoderado/a fiscal'),
        mkFirmaCell(CONT - cf, nombreFirma, pieDemandado)
      )
    ]
  });

  // ── Armar documento ──────────────────────────────────────

  var kids = [
    mkPar('FISCO DE LA PROVINCIA DE BUENOS AIRES', { bold: true, center: true, size: 26, antes: 0, despues: 40 }),
    mkPar('Fuero Comercial de la Naci\u00f3n',      { center: true, size: 20, color: '555577', despues: 180 }),
    sep(),
    mkPar('ACUERDO DE PAGO',        { bold: true, center: true, size: 28, antes: 220, despues: 60 }),
    mkPar('Buenos Aires, ' + fecha, { center: true, size: 20, color: '666666', despues: 280 }),
    mkPar('Datos del acuerdo',      { bold: true, size: 22, antes: 160, despues: 100 }),
    tablaDatos,
  ];

  if (tablaTitulos) {
    kids.push(mkPar('T\u00edtulos ejecutivos', { bold: true, size: 22, antes: 240, despues: 100 }));
    kids.push(tablaTitulos);
  }

  if (parsAcuerdo.length) {
    kids.push(mkPar('Texto del acuerdo', { bold: true, size: 22, antes: 240, despues: 100 }));
    parsAcuerdo.forEach(function(p) { kids.push(p); });
  }

  kids.push(mkPar('', { antes: 400, despues: 0 }));
  kids.push(tablaFirmas);

  var doc = new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: PW, height: 16838 },
          margin: { top: ML, right: ML, bottom: ML, left: ML }
        }
      },
      children: kids
    }]
  });

  Packer.toBlob(doc).then(function(blob) {
    var nombreLimpio = nombre !== '_______________'
      ? nombre.replace(/[^a-zA-Z\u00c0-\u024f\s]/g, '').trim().replace(/\s+/g, '_')
      : 'contribuyente';
    var expLimpio = expediente !== '_______________'
      ? expediente.replace(/\//g, '-')
      : 'sin_expediente';
    saveAs(blob, 'Acuerdo_' + nombreLimpio + '_' + expLimpio + '.docx');
  }).catch(function(err) {
    console.error('Error al generar Word:', err);
    alert('No se pudo generar el archivo. Detalle en consola.');
  });
}
