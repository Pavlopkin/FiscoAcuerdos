// ============================================================
//  DESCARGA ACUERDO EN WORD — FiscoAcuerdos
//  Requiere docx y FileSaver.js en el HTML (ver instrucciones)
// ============================================================
//
//  1) Agregá en index.html, antes de </body>:
//
//     <script src="https://unpkg.com/docx@8.5.0/build/index.js"></script>
//     <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
//     <script src="./js/descargar-acuerdo-word.js"></script>
//
//  2) En el resultado del acuerdo, junto al botón "Imprimir", agregá:
//
//     <button class="btn" onclick="descargarAcuerdoWord()">⬇ Descargar Word</button>
//
// ============================================================

function descargarAcuerdoWord() {

  const {
    Document, Packer, Paragraph, TextRun,
    Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType
  } = docx;

  // ── Leer datos del formulario ────────────────────────────

  const apoderadoVal = document.getElementById('ac-apoderado').value;
  const apoderadoMap = { natalia: 'Natalia', mauricio: 'Mauricio', maximina: 'Maximina' };
  const apoderado    = apoderadoMap[apoderadoVal] || apoderadoVal;

  const juzgadoNum   = document.getElementById('ac-juzgado').value;
  const juzgado      = `Juzgado N° ${juzgadoNum}`;

  const expediente   = document.getElementById('ac-expediente').value.trim() || '_______________';

  const tipoId       = document.querySelector('input[name="identificacion"]:checked')?.value?.toUpperCase() || 'DNI/CUIT';
  const numeroId     = document.getElementById('ac-dni').value.trim() || '_______________';
  const nombre       = document.getElementById('ac-contribuyente').value.trim() || '_______________';
  const domicilio    = document.getElementById('ac-domicilio').value.trim() || '_______________';

  const caracterVal  = document.querySelector('input[name="caracter"]:checked')?.value || '';
  const caracterMap  = { demandado: 'Demandado', apoderado: 'Apoderado', representante: 'Representante legal' };
  const caracter     = caracterMap[caracterVal] || '_______________';

  const cantTitulos  = document.querySelector('input[name="cantTitulos"]:checked')?.value || '—';

  // Títulos ejecutivos: leer los campos generados dinámicamente
  const titulos = [];
  for (let i = 1; i <= parseInt(cantTitulos); i++) {
    const periodo  = document.getElementById(`titulo-periodo-${i}`)?.value?.trim() || '';
    const monto    = document.getElementById(`titulo-monto-${i}`)?.value?.trim() || '';
    if (periodo || monto) titulos.push({ n: i, periodo, monto });
  }

  // Fecha actual
  const hoy   = new Date();
  const fecha = hoy.toLocaleDateString('es-AR', { day: '2-digit', month: 'long', year: 'numeric' });

  // Texto del acuerdo (igual al que genera generarAcuerdo())
  const textoAcuerdo = document.getElementById('acuerdo-texto')?.innerText?.trim() || '';

  // ── Helpers ──────────────────────────────────────────────

  const borde = { style: BorderStyle.SINGLE, size: 1, color: 'C8D0DB' };
  const bordes = { top: borde, bottom: borde, left: borde, right: borde };
  const sinBorde = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const sinBordes = { top: sinBorde, bottom: sinBorde, left: sinBorde, right: sinBorde };

  // Ancho de página A4 con márgenes de 2 cm
  const PW = 11906;
  const ML = 1134; // ~2 cm en DXA
  const CONTENIDO = PW - ML * 2; // ≈ 9638

  function p(texto, opts = {}) {
    return new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { before: opts.antes ?? 80, after: opts.despues ?? 80 },
      children: [new TextRun({
        text: texto,
        bold: !!opts.bold,
        size: opts.size ?? 22,
        font: 'Arial',
        color: opts.color ?? '1A1A2E',
      })]
    });
  }

  function separador() {
    return new Paragraph({
      spacing: { before: 0, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D0D7E3', space: 1 } },
      children: []
    });
  }

  function celda(texto, col, opts = {}) {
    return new TableCell({
      borders: bordes,
      width: { size: col, type: WidthType.DXA },
      margins: { top: 90, bottom: 90, left: 140, right: 140 },
      shading: opts.header
        ? { fill: 'E8EDF5', type: ShadingType.CLEAR }
        : { fill: 'FFFFFF', type: ShadingType.CLEAR },
      children: [new Paragraph({
        children: [new TextRun({
          text: texto,
          bold: !!opts.bold || !!opts.header,
          size: opts.size ?? 20,
          font: 'Arial',
          color: opts.header ? '2C3A5A' : '1A1A2E',
        })]
      })]
    });
  }

  function fila(...celdas) { return new TableRow({ children: celdas }); }

  // ── Tabla de datos ───────────────────────────────────────
  const c1 = Math.round(CONTENIDO * 0.38);
  const c2 = CONTENIDO - c1;

  const rows = [
    fila(celda('Campo', c1, { header: true }), celda('Valor', c2, { header: true })),
    fila(celda('Apoderado/a', c1),      celda(apoderado, c2, { bold: true })),
    fila(celda('Juzgado', c1),          celda(juzgado, c2)),
    fila(celda('Expediente N°', c1),    celda(expediente, c2)),
    fila(celda(`N° de ${tipoId}`, c1),  celda(numeroId, c2)),
    fila(celda('Apellido y nombre', c1),celda(nombre, c2, { bold: true })),
    fila(celda('Domicilio', c1),        celda(domicilio, c2)),
    fila(celda('Carácter del firmante', c1), celda(caracter, c2)),
    fila(celda('Cantidad de títulos', c1),   celda(cantTitulos.toString(), c2)),
  ];

  const tablaDatos = new Table({
    width: { size: CONTENIDO, type: WidthType.DXA },
    columnWidths: [c1, c2],
    rows
  });

  // ── Tabla de títulos ejecutivos ──────────────────────────
  let tablaTitulos = null;
  if (titulos.length > 0) {
    const ct1 = Math.round(CONTENIDO * 0.12);
    const ct2 = Math.round(CONTENIDO * 0.50);
    const ct3 = CONTENIDO - ct1 - ct2;
    const filasTit = [
      fila(
        celda('N°',      ct1, { header: true }),
        celda('Período', ct2, { header: true }),
        celda('Monto',   ct3, { header: true })
      ),
      ...titulos.map(t => fila(
        celda(t.n.toString(), ct1),
        celda(t.periodo, ct2),
        celda(t.monto ? `$ ${t.monto}` : '', ct3)
      ))
    ];
    tablaTitulos = new Table({
      width: { size: CONTENIDO, type: WidthType.DXA },
      columnWidths: [ct1, ct2, ct3],
      rows: filasTit
    });
  }

  // ── Texto del acuerdo ────────────────────────────────────
  // Convertir el texto generado en párrafos (split por salto de línea)
  const parrafosAcuerdo = textoAcuerdo
    .split('\n')
    .filter(l => l.trim() !== '')
    .map(linea => p(linea, { antes: 120, despues: 60, size: 22 }));

  // ── Tabla de firmas ──────────────────────────────────────
  const cf = Math.floor(CONTENIDO / 2);
  const tablafirmas = new Table({
    width: { size: CONTENIDO, type: WidthType.DXA },
    columnWidths: [cf, CONTENIDO - cf],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: sinBordes,
            width: { size: cf, type: WidthType.DXA },
            margins: { top: 600, bottom: 100, left: 200, right: 200 },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 8, color: '4A5568', space: 60 } },
              children: [new TextRun({ text: apoderado, size: 20, font: 'Arial', bold: true })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: 'Apoderado/a fiscal', size: 18, font: 'Arial', color: '666666' })]
            })]
          }),
          new TableCell({
            borders: sinBordes,
            width: { size: CONTENIDO - cf, type: WidthType.DXA },
            margins: { top: 600, bottom: 100, left: 200, right: 200 },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 8, color: '4A5568', space: 60 } },
              children: [new TextRun({ text: nombre !== '_______________' ? nombre : ' ', size: 20, font: 'Arial', bold: true })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: `${caracter} — ${tipoId} ${numeroId}`, size: 18, font: 'Arial', color: '666666' })]
            })]
          }),
        ]
      })
    ]
  });

  // ── Armar documento ──────────────────────────────────────
  const children = [
    // Encabezado
    p('FISCO DE LA PROVINCIA DE BUENOS AIRES', { bold: true, center: true, size: 26, antes: 0, despues: 40 }),
    p('Fuero Comercial de la Nación', { center: true, size: 20, color: '555577', despues: 200 }),
    separador(),
    p('ACUERDO DE PAGO', { bold: true, center: true, size: 28, antes: 240, despues: 60 }),
    p(`Buenos Aires, ${fecha}`, { center: true, size: 20, color: '666666', despues: 300 }),

    // Tabla de datos
    p('Datos del acuerdo', { bold: true, size: 22, antes: 200, despues: 120 }),
    tablaDatos,

    // Títulos ejecutivos (si hay)
    ...(tablaTitulos ? [
      p('Títulos ejecutivos', { bold: true, size: 22, antes: 280, despues: 120 }),
      tablaTitulos,
    ] : []),

    // Texto del acuerdo
    ...(parrafosAcuerdo.length > 0 ? [
      p('Texto del acuerdo', { bold: true, size: 22, antes: 280, despues: 120 }),
      ...parrafosAcuerdo,
    ] : []),

    // Firmas
    p('', { antes: 400 }),
    tablafirmas,
  ];

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 22 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: PW, height: 16838 }, // A4
          margin: { top: ML, right: ML, bottom: ML, left: ML }
        }
      },
      children
    }]
  });

  // ── Generar y descargar ──────────────────────────────────
  Packer.toBlob(doc).then(blob => {
    const nombreLimpio = nombre !== '_______________'
      ? nombre.replace(/[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]/g, '').trim().replace(/\s+/g, '_')
      : 'contribuyente';
    const expLimpio = expediente !== '_______________'
      ? expediente.replace(/\//g, '-')
      : 'sin_expediente';
    saveAs(blob, `Acuerdo_${nombreLimpio}_${expLimpio}.docx`);
  });
}
