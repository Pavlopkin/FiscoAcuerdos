function descargarAcuerdoWord() {

  // docx UMD expone las clases directamente en window
  const Document      = window.Document      || (window.docx && window.docx.Document);
  const Packer        = window.Packer        || (window.docx && window.docx.Packer);
  const Paragraph     = window.Paragraph     || (window.docx && window.docx.Paragraph);
  const TextRun       = window.TextRun       || (window.docx && window.docx.TextRun);
  const Table         = window.Table         || (window.docx && window.docx.Table);
  const TableRow      = window.TableRow      || (window.docx && window.docx.TableRow);
  const TableCell     = window.TableCell     || (window.docx && window.docx.TableCell);
  const AlignmentType = window.AlignmentType || (window.docx && window.docx.AlignmentType);
  const BorderStyle   = window.BorderStyle   || (window.docx && window.docx.BorderStyle);
  const WidthType     = window.WidthType     || (window.docx && window.docx.WidthType);

  if (!Document || !Packer) {
    alert('Error: la librería docx no se cargó correctamente.');
    return;
  }

  // ── Datos del formulario ─────────────────────────────────

  const apoderadoVal = document.getElementById('ac-apoderado').value;
  const apoderadoMap = { natalia: 'Natalia', mauricio: 'Mauricio', maximina: 'Maximina' };
  const apoderado    = apoderadoMap[apoderadoVal] || apoderadoVal;

  const juzgadoNum = document.getElementById('ac-juzgado').value;
  const juzgado    = 'Juzgado N\u00b0 ' + juzgadoNum;

  const expediente = document.getElementById('ac-expediente').value.trim() || '_______________';
  const tipoId     = (document.querySelector('input[name="identificacion"]:checked')?.value || 'DNI').toUpperCase();
  const numeroId   = document.getElementById('ac-dni').value.trim() || '_______________';
  const nombre     = document.getElementById('ac-contribuyente').value.trim() || '_______________';
  const domicilio  = document.getElementById('ac-domicilio').value.trim() || '_______________';

  const caracterVal = document.querySelector('input[name="caracter"]:checked')?.value || '';
  const caracterMap = { demandado: 'Demandado', apoderado: 'Apoderado', representante: 'Representante legal' };
  const caracter    = caracterMap[caracterVal] || '_______________';

  const cantTitulos = document.querySelector('input[name="cantTitulos"]:checked')?.value || '0';
  const titulos = [];
  for (var i = 1; i <= parseInt(cantTitulos); i++) {
    var periodo = (document.getElementById('titulo-periodo-' + i) || {}).value || '';
    var monto   = (document.getElementById('titulo-monto-' + i)   || {}).value || '';
    periodo = periodo.trim(); monto = monto.trim();
    if (periodo || monto) titulos.push({ n: i, periodo: periodo, monto: monto });
  }

  var hoy   = new Date();
  var fecha = hoy.toLocaleDateString('es-AR', { day: '2-digit', month: 'long', year: 'numeric' });

  var textoAcuerdo = (document.getElementById('acuerdo-texto') || {}).innerText || '';
  textoAcuerdo = textoAcuerdo.trim();

  // ── Constantes de layout ─────────────────────────────────

  var PW = 11906, ML = 1134, CONT = PW - ML * 2; // A4, márgenes 2cm

  // ── Estilos de borde ─────────────────────────────────────

  var bVis = { style: BorderStyle.SINGLE, size: 1, color: 'C8D0DB' };
  var bVis6 = { style: BorderStyle.SINGLE, size: 6, color: '4A5568' };
  var bInv = { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' };
  var bordesVis = { top: bVis, bottom: bVis, left: bVis, right: bVis };
  var bordesInv = { top: bInv, bottom: bInv, left: bInv, right: bInv };

  // ── Helpers ──────────────────────────────────────────────

  function txt(t, opts) {
    opts = opts || {};
    var cfg = { text: t, font: 'Arial', size: opts.size || 22, color: opts.color || '1A1A2E' };
    if (opts.bold)   cfg.bold   = true;
    if (opts.italic) cfg.italics = true;
    return new TextRun(cfg);
  }

  function par(t, opts) {
    opts = opts || {};
    var cfg = {
      spacing: { before: opts.antes != null ? opts.antes : 80, after: opts.despues != null ? opts.despues : 80 },
      children: [txt(t, opts)]
    };
    if (opts.center) cfg.alignment = AlignmentType.CENTER;
    if (opts.borde_top) cfg.border = { top: bVis6 };
    return new Paragraph(cfg);
  }

  function sep() {
    return new Paragraph({
      spacing: { before: 0, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D0D7E3', space: 1 } },
      children: [txt('')]
    });
  }

  function tc(texto, w, opts) {
    opts = opts || {};
    var fill = opts.header ? 'E8EDF5' : 'FFFFFF';
    return new TableCell({
      borders: bordesVis,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      shading: { fill: fill },
      children: [new Paragraph({ children: [txt(texto, { size: 20, bold: !!(opts.bold || opts.header), color: opts.header ? '2C3A5A' : '1A1A2E' })] })]
    });
  }

  function tcFirma(w, t1, t2) {
    return new TableCell({
      borders: bordesInv,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 700, bottom: 120, left: 200, right: 200 },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 8, color: '333333', space: 60 } },
          children: [txt(t1, { size: 20, bold: true })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [txt(t2, { size: 18, color: '666666' })]
        })
      ]
    });
  }

  function row() { return new TableRow({ children: Array.from(arguments) }); }

  // ── Tabla de datos ───────────────────────────────────────

  var c1 = Math.round(CONT * 0.40), c2 = CONT - c1;
  var tablaDatos = new Table({
    width: { size: CONT, type: WidthType.DXA },
    columnWidths: [c1, c2],
    rows: [
      row(tc('Campo', c1, { header: true }),               tc('Valor', c2, { header: true })),
      row(tc('Apoderado/a', c1),                           tc(apoderado, c2, { bold: true })),
      row(tc('Juzgado', c1),                               tc(juzgado, c2)),
      row(tc('Expediente N\u00b0', c1),                    tc(expediente, c2)),
      row(tc('N\u00b0 de ' + tipoId, c1),                  tc(numeroId, c2)),
      row(tc('Apellido y nombre', c1),                     tc(nombre, c2, { bold: true })),
      row(tc('Domicilio', c1),                             tc(domicilio, c2)),
      row(tc('Car\u00e1cter del firmante', c1),             tc(caracter, c2)),
      row(tc('Cantidad de t\u00edtulos', c1),               tc(cantTitulos.toString(), c2)),
    ]
  });

  // ── Tabla de títulos ejecutivos ──────────────────────────

  var tablaTitulos = null;
  if (titulos.length > 0) {
    var ct1 = Math.round(CONT * 0.12), ct2 = Math.round(CONT * 0.50), ct3 = CONT - ct1 - ct2;
    var rowsTit = [row(tc('N\u00b0', ct1, { header: true }), tc('Per\u00edodo', ct2, { header: true }), tc('Monto', ct3, { header: true }))];
    titulos.forEach(function(t) {
      rowsTit.push(row(tc(String(t.n), ct1), tc(t.periodo, ct2), tc(t.monto ? ('$ ' + t.monto) : '', ct3)));
    });
    tablaTitulos = new Table({ width: { size: CONT, type: WidthType.DXA }, columnWidths: [ct1, ct2, ct3], rows: rowsTit });
  }

  // ── Texto del acuerdo ────────────────────────────────────

  var parsAcuerdo = textoAcuerdo.split('\n').filter(function(l) { return l.trim(); }).map(function(l) {
    return par(l, { antes: 100, despues: 60 });
  });

  // ── Tabla de firmas ──────────────────────────────────────

  var cf = Math.floor(CONT / 2);
  var nombreFirma  = nombre !== '_______________' ? nombre : ' ';
  var pieDemandado = caracter + ' \u2014 ' + tipoId + ' ' + numeroId;

  var tablaFirmas = new Table({
    width: { size: CONT, type: WidthType.DXA },
    columnWidths: [cf, CONT - cf],
    rows: [new TableRow({ children: [tcFirma(cf, apoderado, 'Apoderado/a fiscal'), tcFirma(CONT - cf, nombreFirma, pieDemandado)] })]
  });

  // ── Armar contenido ──────────────────────────────────────

  var kids = [
    par('FISCO DE LA PROVINCIA DE BUENOS AIRES', { bold: true, center: true, size: 26, antes: 0, despues: 40 }),
    par('Fuero Comercial de la Naci\u00f3n',      { center: true, size: 20, color: '555577', despues: 180 }),
    sep(),
    par('ACUERDO DE PAGO',   { bold: true, center: true, size: 28, antes: 220, despues: 60 }),
    par('Buenos Aires, ' + fecha, { center: true, size: 20, color: '666666', despues: 280 }),
    par('Datos del acuerdo', { bold: true, size: 22, antes: 160, despues: 100 }),
    tablaDatos,
  ];

  if (tablaTitulos) {
    kids.push(par('T\u00edtulos ejecutivos', { bold: true, size: 22, antes: 240, despues: 100 }));
    kids.push(tablaTitulos);
  }

  if (parsAcuerdo.length) {
    kids.push(par('Texto del acuerdo', { bold: true, size: 22, antes: 240, despues: 100 }));
    parsAcuerdo.forEach(function(p) { kids.push(p); });
  }

  kids.push(par('', { antes: 500, despues: 0 }));
  kids.push(tablaFirmas);

  // ── Documento ────────────────────────────────────────────

  var doc = new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
    sections: [{
      properties: { page: { size: { width: PW, height: 16838 }, margin: { top: ML, right: ML, bottom: ML, left: ML } } },
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
