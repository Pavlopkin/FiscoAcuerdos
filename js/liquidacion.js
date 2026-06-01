/* ============================================================
   liquidacion.js — Generador de Liquidaciones y Acuerdos
   Fuero Comercial de la Nación
   ============================================================

   REGLAS DE NEGOCIO
   ─────────────────
   Tasa de justicia:  2,2 % del monto reclamado
   Sobre tasa:        5 % o 10 % de la tasa de justicia
   Honorarios %:      6, 9, 12, 15, 18 % (con mínimo aplicable)
   Honorario mínimo A: $ 149.250  →  gastos según opción del formulario
   Honorario mínimo B: $ 101.994  →  gastos SIEMPRE fijos en $ 33.998
   Gastos:            $ 24.875  |  $ 49.750  |  $ 33.998 (automático)
   Servicios reg.:    opcional, montos $ 84.000 / 168.000 / 252.000
   Aportes s/ hono:   10 % sobre honorarios
   ============================================================ */

'use strict';

// ── Constantes ───────────────────────────────────────────────
const MESES = [
  'enero','febrero','marzo','abril','mayo','junio',
  'julio','agosto','septiembre','octubre','noviembre','diciembre'
];

const APODERADOS = {
  natalia:  'Natalia',
  mauricio: 'Mauricio',
  maximina: 'Maximina'
};

const GASTOS_MINIMO_B = 33998;
const MINIMO_A        = 149250;
const MINIMO_B        = 101994;

// ── Helpers ──────────────────────────────────────────────────

function formatPeso(n) {
  return '$ ' + n.toLocaleString('es-AR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function getRadioVal(name) {
  const el = document.querySelector(`input[name="${name}"]:checked`);
  return el ? el.value : null;
}

function selectRadio(groupId, el, val) {
  document.querySelectorAll('#' + groupId + ' .radio-opt')
    .forEach(o => o.classList.remove('selected'));
  el.classList.add('selected');
  el.querySelector('input[type="radio"]').checked = true;
}
window.selectRadio = selectRadio;

// Convierte número a palabras en español (pesos)
function numeroALetras(n) {
  const partes = n.toFixed(2).split('.');
  const entero = parseInt(partes[0]);
  const cents  = parseInt(partes[1]);

  const u = ['','un','dos','tres','cuatro','cinco','seis','siete','ocho','nueve',
             'diez','once','doce','trece','catorce','quince','dieciséis',
             'diecisiete','dieciocho','diecinueve','veinte'];
  const d = ['','','veinti','treinta','cuarenta','cincuenta',
             'sesenta','setenta','ochenta','noventa'];
  const c = ['','cien','doscientos','trescientos','cuatrocientos','quinientos',
             'seiscientos','setecientos','ochocientos','novecientos'];

  function conv(num) {
    if (num === 0) return '';
    if (num <= 20) return u[num];
    if (num < 100) {
      const t = Math.floor(num / 10), r = num % 10;
      return r === 0 ? d[t] : (t === 2 ? 'veinti' + u[r] : d[t] + ' y ' + u[r]);
    }
    if (num < 1000) {
      const t = Math.floor(num / 100), r = num % 100;
      return r === 0 ? c[t] : (t === 1 ? 'ciento ' + conv(r) : c[t] + ' ' + conv(r));
    }
    if (num < 1000000) {
      const t = Math.floor(num / 1000), r = num % 1000;
      const m = (t === 1) ? 'mil' : conv(t) + ' mil';
      return r === 0 ? m : m + ' ' + conv(r);
    }
    const t = Math.floor(num / 1000000), r = num % 1000000;
    const m = (t === 1) ? 'un millón' : conv(t) + ' millones';
    return r === 0 ? m : m + ' ' + conv(r);
  }

  const enteroLetras = (conv(entero) || 'cero').toUpperCase();
  const centsLetras  = cents > 0 ? ' CON ' + (conv(cents) || 'cero').toUpperCase() + ' CENTAVOS' : '';
  return 'PESOS ' + enteroLetras + centsLetras;
}

// ── Tab switching ─────────────────────────────────────────────
function switchTab(id, btn) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  btn.classList.add('active');
}
window.switchTab = switchTab;

// ── Lógica de visibilidad dinámica ───────────────────────────

function onMinimoChange() {
  const minimo = getRadioVal('minimoHono');
  const gastosWrap = document.getElementById('gastos-wrap');
  const gastosHint = document.getElementById('gastos-hint');

  if (minimo === 'B') {
    if (gastosWrap) gastosWrap.style.display = 'none';
    if (gastosHint) {
      gastosHint.textContent = 'Gastos fijos: ' + formatPeso(GASTOS_MINIMO_B);
      gastosHint.style.display = 'block';
    }
  } else {
    if (gastosWrap) gastosWrap.style.display = 'block';
    if (gastosHint) gastosHint.style.display = 'none';
  }
}

function onServiciosChange() {
  const v = getRadioVal('servicios');
  const wrap = document.getElementById('selector-serv-wrap');
  if (wrap) wrap.style.display = (v === 'NO') ? 'none' : 'block';
}

function showTitulos(n) {
  const cont = document.getElementById('titulos-inputs');
  if (!cont) return;
  cont.innerHTML = '';
  for (let i = 1; i <= n; i++) {
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.id = 'titulo-' + i;
    inp.className = 'titulo-input';
    inp.placeholder = 'Título ejecutivo N° ' + i;
    inp.style.cssText = 'font-size:13px;padding:8px 10px;border:1px solid var(--color-border);border-radius:var(--radius-md);background:var(--color-accent-bg);color:var(--color-text);outline:none;width:100%';
    cont.appendChild(inp);
  }
}
window.showTitulos = showTitulos;

// ── CÁLCULO DE LIQUIDACIÓN ───────────────────────────────────

function calcularLiq() {
  const monto         = parseFloat(document.getElementById('liq-monto').value) || 0;
  const contribuyente = document.getElementById('liq-contribuyente').value || '(sin nombre)';
  const juicio        = document.getElementById('liq-juicio').value || '—';
  const sobretasaPct  = parseFloat(getRadioVal('sobretasa')) || 5;
  const pct           = parseFloat(document.getElementById('liq-honorarios').value) || 6;
  const minimoHono    = getRadioVal('minimoHono');   // 'A' | 'B'
  const servicios     = getRadioVal('servicios');    // 'SI' | 'NO'
  const servMonto     = (servicios === 'SI')
    ? parseFloat(document.getElementById('liq-selectorServicios').value) || 0
    : 0;

  // Gastos según mínimo elegido
  let gastos;
  if (minimoHono === 'B') {
    gastos = GASTOS_MINIMO_B;
  } else {
    gastos = parseFloat(getRadioVal('gastos')) || 0;
  }

  // Honorarios: max(calculado, mínimo)
  const minimoValor   = (minimoHono === 'B') ? MINIMO_B : MINIMO_A;
  const honoCalculado = monto * (pct / 100);
  const honorarios    = Math.max(honoCalculado, minimoValor);

  // Rubros — fórmulas corregidas
  const tasa      = monto * 0.022;                        // 2,2 % del monto
  const sobretasa = tasa * (sobretasaPct / 100);          // 5 % o 10 % de la tasa
  const aportes   = honorarios * 0.10;                    // 10 % s/ honorarios
  const honofisco = honorarios * 0.40;                    // interno, para el acuerdo
  const honoapod  = honorarios * 0.60;                    // interno, para el acuerdo

  // Planilla 1: costas (tasa + sobretasa + gastos + servicios)
  const totalCostas = tasa + sobretasa + gastos + servMonto;

  // Planilla 2: honorarios + aportes
  const subtotalHono = honorarios + aportes;

  // Total general
  const totalGeneral = totalCostas + subtotalHono;

  // Fecha
  const now = new Date();
  const fechaTxt = `${now.getDate()} de ${MESES[now.getMonth()]} de ${now.getFullYear()}`;

  // Encabezado
  document.getElementById('liq-res-fecha').textContent = 'Actualizado al ' + fechaTxt;
  document.getElementById('liq-res-titulo').textContent =
    contribuyente + (juicio !== '—' ? ' — Juicio ' + juicio : '');

  // Items en el mismo orden que el Excel
  const items = [
    ['Capital actualizado',                                    formatPeso(monto)],
    ['Tasa de justicia (2,2 %)',                               formatPeso(tasa)],
    [`Sobre tasa (${sobretasaPct} %)`,                         formatPeso(sobretasa)],
    ['Gastos de juicio',                                       gastos > 0 ? formatPeso(gastos) : '—'],
    ['Servicios registrales',                                  servMonto > 0 ? formatPeso(servMonto) : 'No aplica'],
    ['Total costas',                                           formatPeso(totalCostas)],
    [`Honorarios (${pct} % — mín. ${formatPeso(minimoValor)})`, formatPeso(honorarios)],
    ['Aportes s/ honorarios (10 %)',                           formatPeso(aportes)],
    ['Subtotal a cargo del deudor',                            formatPeso(subtotalHono)],
  ];

  const grid = document.getElementById('liq-items');
  grid.innerHTML = '';
  items.forEach(([label, val]) => {
    grid.innerHTML +=
      `<div class="liq-item">
         <span class="liq-item-label">${label}</span>
         <span class="liq-item-val">${val}</span>
       </div>`;
  });

  document.getElementById('liq-total-val').textContent = formatPeso(totalGeneral);

  // Guardar para uso en el acuerdo
  window._liqData = {
    monto, tasa, sobretasa, sobretasaPct,
    honorarios, honofisco, honoapod,
    aportes, gastos, servMonto,
    totalCostas, subtotalHono, totalGeneral,
    contribuyente, juicio, minimoValor, pct
  };

  // Mostrar resultado
  document.getElementById('form-liquidacion').style.display = 'none';
  document.getElementById('resultado-liq').classList.add('visible');
}
window.calcularLiq = calcularLiq;

function limpiarLiq() {
  document.getElementById('liq-contribuyente').value = '';
  document.getElementById('liq-juicio').value = '';
  document.getElementById('liq-monto').value = '';
  document.querySelectorAll(
    '#rg-sobretasa .radio-opt, #rg-gastos .radio-opt, #rg-servicios .radio-opt, #rg-minimo .radio-opt'
  ).forEach(o => o.classList.remove('selected'));
  document.querySelectorAll(
    'input[name="sobretasa"], input[name="gastos"], input[name="servicios"], input[name="minimoHono"]'
  ).forEach(i => i.checked = false);
  onMinimoChange();
  onServiciosChange();
}
window.limpiarLiq = limpiarLiq;

function ocultarResultado() {
  document.getElementById('resultado-liq').classList.remove('visible');
  document.getElementById('form-liquidacion').style.display = 'block';
}
window.ocultarResultado = ocultarResultado;

// ── GENERACIÓN DEL ACUERDO ───────────────────────────────────

function generarAcuerdo() {
  const now  = new Date();
  const dia  = now.getDate();
  const mes  = MESES[now.getMonth()];
  const anio = now.getFullYear();

  const contrib    = document.getElementById('ac-contribuyente').value || 'LA DEMANDADA';
  const tipoId     = (getRadioVal('identificacion') || 'DNI').toUpperCase();
  const dni        = document.getElementById('ac-dni').value || '—';
  const dom        = document.getElementById('ac-domicilio').value || '—';
  const expediente = document.getElementById('ac-expediente').value || '—';
  const juzgado    = document.getElementById('ac-juzgado').value || '2';
  const apodVal    = document.getElementById('ac-apoderado').value;
  
  // Mapeo dinámico para mantener el nombre formal del Dr. Luparia según tu plantilla
  let apodNombre = APODERADOS[apodVal] || apodVal;
  if (apodVal === 'mauricio') {
    apodNombre = 'Dr. Mauricio Julián Luparia de la Colina';
  } else {
    apodNombre = 'Dr./Dra. ' + apodNombre;
  }

  // Carácter de la representación
  const caracteres = { demandado: 'demandado', apoderado: 'apoderado', representante: 'representante legal' };
  const caracterVal = getRadioVal('caracter') || 'demandado';
  const caracter = caracteres[caracterVal];

  // Títulos ejecutivos
  const titInputs = document.querySelectorAll('[id^="titulo-"]');
  const titulosVals = [...titInputs].map(i => i.value.trim()).filter(Boolean);
  let titulosTexto;
  if (titulosVals.length === 0) {
    titulosTexto = 'el/los título/s ejecutivo/s correspondientes';
  } else if (titulosVals.length === 1) {
    titulosTexto = titulosVals[0];
  } else {
    titulosTexto = titulosVals.slice(0, -1).join(', ') + ' y ' + titulosVals[titulosVals.length - 1];
  }

  // Montos de la liquidación (si se calculó antes)
  const d = window._liqData || {};
  const fmonto   = d.monto        ? formatPeso(d.monto)         : '$ ...........';
  const fmontoL  = d.monto        ? numeroALetras(d.monto)      : '............';
  const ftasa    = d.tasa         ? formatPeso(d.tasa)          : '$ ...........';
  const fstasa   = d.sobretasa    ? formatPeso(d.sobretasa)     : '$ ...........';
  const fhono    = d.honorarios   ? formatPeso(d.honorarios)    : '$ ...........';
  const fhonoF   = d.honofisco    ? formatPeso(d.honofisco)     : '$ ...........';
  const fhonoA   = d.honoapod     ? formatPeso(d.honoapod)      : '$ ...........';
  const fhonoL   = d.honorarios   ? numeroALetras(d.honorarios) : '............';
  const faporte  = d.aportes      ? formatPeso(d.aportes)       : '$ ...........';
  const fgastos  = d.gastos       ? formatPeso(d.gastos)        : '$ ...........';
  const fserv    = (d.servMonto && d.servMonto > 0) ? formatPeso(d.servMonto) : '$ ...........';
  const fjuicio  = d.juicio       || document.getElementById('liq-juicio')?.value || '—';

  const html = `
<div style="text-align: center; margin-bottom: 20px;">
  <h2 style="font-size: 18px; letter-spacing: 1px; margin: 0;">ACUERDO DE PAGO</h2>
</div>

<p>En la ciudad de San Isidro a los <strong>${dia}</strong> días del mes de <strong>${mes}</strong> de <strong>${anio}</strong>, entre <strong>${contrib.toUpperCase()} (${tipoId} ${dni})</strong>, con domicilio en <strong>${dom}</strong>, en su carácter de <strong>${caracter}</strong>, por una parte, en adelante "LA DEMANDADA"; y por otra, el <strong>${apodNombre}</strong>, con domicilio constituido en calle Ituzaingo Nº 321 Depto. 52 de San Isidro, en su carácter de apoderado fiscal de la Provincia de Buenos Aires en los términos de los arts. 4 y 4 bis del decreto ley 7543/69 (t.o. y sus modificatorios), en adelante "EL APODERADO", tal como se acredita en los autos caratulados <strong>"FISCO DE LA PROVINCIA DE BUENOS AIRES c/ ${contrib.toUpperCase()} s/APREMIO"</strong>, expediente <strong>${expediente}</strong>, de trámite por ante el Juzgado en lo Contencioso Administrativo N° <strong>${juzgado}</strong>, del Departamento Judicial de San Isidro, con relación a la ejecución fiscal citada por la que se persigue el cobro de los períodos/adelantos individualizados en el/los título/s ejecutivo/s <strong>${titulosTexto}</strong>, juicio <strong>${fjuicio}</strong>, con el objeto de poner fin al apremio se hace constar lo siguiente:</p>

<p><strong>PRIMERA:</strong> LA DEMANDADA reconoce adeudar al Fisco la totalidad de la deuda reclamada en el apremio detallado, renunciando a toda reclamación impugnatoria administrativa o judicial de la deuda mencionada, también se notifica y consiente expresamente las medidas cautelares trabadas o a trabarse, ya sean judiciales -cuyo levantamiento queda a cargo de LA DEMANDADA cuando corresponda según el plan- o administrativas que se efectivicen sobre bienes muebles, inmuebles, financieros o de cualquier otra naturaleza. El monto que se utiliza, salvo error u omisión, para el presente acuerdo es el de <strong>${fmontoL} (${fmonto})</strong>, que surge de la página de ARBA con la salvedad contenida en los párrafos siguientes de esta primera cláusula. Se adjunta el presente convenio, y como parte integrante del mismo, una impresión conocida y consentida por LA DEMANDADA de los diferentes montos que surgen de la página Web de ARBA, que arrojan importes disímiles según opta LA DEMANDADA y que inciden directamente en el mayor o menor monto de costas que debe abonar. Es decir, conoce que su elección irrevocable es abonar la deuda en ARBA en la cantidad de CUOTAS que declara al apoderado, y que en caso de posteriormente a la suscripción del acuerdo cambiar de plan, dicho acto tiene consecuencias directas en el monto total de las costas a abonar por este convenio. Conste que se ha llegado a concluir este acuerdo por vía telefónica y/o electrónica quedando sujeto su perfeccionamiento a los pagos íntegros de todos los rubros detallados abajo y a la regularización del crédito de ARBA por parte del LA DEMANDADA. Se deja expresa constancia que en el juicio objeto del presente NO existe oposición de excepciones pendiente de tratamiento. De verificarse el ingreso a un plan con cuotas superior a las manifestadas, LA DEMANDADA deberá cancelar las diferencias que resulten de calcular nuevamente las costas. Para ello, se tomará el monto que corresponda consignado en el formulario de acogimiento expedido por ARBA. Si se le permitiere por cualquier motivo ingresar en un plan por un monto menor nada podrá reclamar LA DEMANDADA a LA ACTORA, ni al Apoderada/o Fiscal, ni a la Caja de Previsión, ni al Poder Judicial ya que lo ha hecho voluntariamente y prestando su consentimiento claramente informado. En el caso en que se disponga judicialmente, aún contra la voluntad expresada por LA DEMANDADA en este convenio, la devolución de Tasa de Justicia, Sobre Tasa, Gastos de Estudio, Aportes Previsionales u Honorarios, se exime a LA ACTORA y a su apoderada/o de cualquier gestión personal o profesional al respecto. Se deja constancia que los montos tenidos en cuenta surgen de la Web de ARBA en el día de hoy, que varían diariamente por acumulación de intereses, que si existe plan de facilidades puede haber vencido el horario para efectuar el acogimiento, que el eventual plan de facilidades tiene fecha de finalización que LA DEMANDADA manifiesta conocer, e incluso que para su tipo de deuda puede no existir plan de facilidades.-</p>

<p><strong>SEGUNDA:</strong> Las partes acuerdan fijar los honorarios del letrado apoderada/o del Fisco de la Provincia de Buenos Aires, en la suma total de <strong>${fhonoL} (${fhono})</strong> pactados de conformidad a las resoluciones dictadas al efecto, las cuales LA DEMANDADA declara conocer y consiente. Asimismo, las partes convienen que dicho monto será cancelado por LA DEMANDADA en UN PAGO.-</p>

<p><strong>TERCERA:</strong> En consecuencia, LA DEMANDADA asume e integra, las siguientes sumas:<br>
1) <strong>${ftasa}</strong> en concepto de Tasa de Justicia,<br>
2) <strong>${fstasa}</strong> en concepto de Sobre Tasa de Justicia,<br>
3) <strong>${fhono}</strong> en concepto de honorarios convenidos por la actuación profesional de la parte actora en los autos citados (de los cuales <strong>${fhonoF}</strong> en concepto de Honorarios para Fiscalía de Estado en virtud del convenio de cesión correspondiente al 40% de la totalidad de los honorarios convenidos, y <strong>${fhonoA}</strong> (60%) por la actuación del profesional de la parte actora) –pactados de conformidad a las resoluciones dictadas al efecto, las cuales LA DEMANDADA declara conocer y consiente–<br>
4) <strong>${faporte}</strong> en concepto de aportes previsionales a cargo de LA DEMANDADA;<br>
5) <strong>${fserv}</strong> en concepto de Servicios Registrales;<br>
6) <strong>${fgastos}</strong> en concepto de gastos generales.</p>

<p>Se deja constancia que el demandado asume la responsabilidad de cancelar cualquier saldo que se pueda adeudar por diferencias a su cargo. Asimismo, se aclara que los gastos causídicos pactados en el presente corresponden a la actividad procesal desarrollada por el Apoderada/o hasta la fecha del presente, pudiendo modificarse dicho concepto en caso de incumplimiento del acuerdo.-</p>

<p><strong>CUARTA:</strong> LA DEMANDADA regularizará el importe adeudado acogiéndose al plan de pago EN CUOTAS A LAS QUE PUEDA ACCEDER VÍA WEB NO PRESENCIAL EN EL SITIO DE ARBA dentro de los cinco (5) días corridos a partir de la fecha de celebración del acuerdo, por el importe que le liquide la ARBA –con la reserva realizada en la cláusula PRIMERA-. En caso de no poder regularizar la deuda en la Web de ARBA con la ayuda que se encuentra operativa de Chat, WhatsApp, Instagram, Facebook y Twitter, deberá hacerlo presencialmente dentro de los primeros treinta (30) días luego de reanudada la atención presencial, con un máximo de 90 días a partir de la fecha. Asume igualmente la obligación de acreditar ante el apoderada/o el efectivo acogimiento dentro de los dos (2) días hábiles posteriores a su ingreso en ARBA con lo que quedará perfeccionado el presente acuerdo. Se deja constancia que ante el incumplimiento del plan de pago que se acuerde con ARBA, el Fisco queda facultado para: a) ejecutar la totalidad del crédito consignado en el título de apremio, imputándose los pagos parciales del crédito de ARBA realizados de acuerdo con lo prescripto en el art. 142 del C.F.; b) iniciar un nuevo juicio fundado en el título ejecutivo que emita la autoridad de aplicación.-</p>

<p>En prueba de conformidad, se firman tres ejemplares de un mismo tenor y a un solo efecto y recibiendo cada parte el suyo y el restante para acompañar al expediente judicial por parte del apoderada/o, pudiendo cualquiera de las partes solicitar su homologación judicial si fuere menester.-</p>

<div style="display: flex; justify-content: space-between; margin-top: 50px; padding: 0 20px;">
  <div style="text-align: center; width: 40%;">
    <p>_____________________________________</p>
    <p style="font-size: 13px; margin-top: 5px;">Firma del Apoderada/o Fiscal</p>
  </div>
  <div style="text-align: center; width: 40%;">
    <p>_____________________________________</p>
    <p style="font-size: 13px; margin-top: 5px;">Firma del demandado</p>
  </div>
</div>
`;

  document.getElementById('acuerdo-texto').innerHTML = html;
  document.getElementById('form-acuerdo').style.display = 'none';
  document.getElementById('resultado-acuerdo').classList.add('visible');
}
window.generarAcuerdo = generarAcuerdo;

function limpiarAcuerdo() {
  ['ac-apoderado','ac-juzgado','ac-expediente','ac-dni','ac-contribuyente','ac-domicilio']
    .forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (el.tagName === 'SELECT') el.selectedIndex = 0;
      else el.value = '';
    });
  document.getElementById('titulos-inputs').innerHTML = '';
  document.querySelectorAll('#rg-ident .radio-opt, #rg-titulos .radio-opt, #rg-caracter .radio-opt')
    .forEach(o => o.classList.remove('selected'));
  document.querySelectorAll('input[name="identificacion"], input[name="cantTitulos"], input[name="caracter"]')
    .forEach(i => i.checked = false);
}
window.limpiarAcuerdo = limpiarAcuerdo;

function ocultarAcuerdo() {
  document.getElementById('resultado-acuerdo').classList.remove('visible');
  document.getElementById('form-acuerdo').style.display = 'block';
}
window.ocultarAcuerdo = ocultarAcuerdo;

// ── Init ──────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('rg-minimo')?.addEventListener('click', onMinimoChange);
  document.getElementById('rg-servicios')?.addEventListener('click', onServiciosChange);
  onMinimoChange();
  onServiciosChange();
});
