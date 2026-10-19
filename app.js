// ============================================================
// ESTADO GLOBAL
// ============================================================
let guardias = [];
const DIAS_SEMANA = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
const NOMBRES_MESES = [
  'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
  'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'
];

// ============================================================
// UTILIDADES
// ============================================================
function formatearMoneda(valor) {
  return new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS' }).format(valor);
}

function obtenerValoresConfig() {
  return {
    semana: parseFloat(document.getElementById('valorSemana').value) || 0,
    finde: parseFloat(document.getElementById('valorFinde').value) || 0,
    viatico: parseFloat(document.getElementById('valorViatico').value) || 0
  };
}

// ============================================================
// IMPORTAR EXCEL (SheetJS)
// ============================================================
function importarExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    
    const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    
    let headerIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const rowStr = rows[i].join(' ').toUpperCase();
      if (rowStr.includes('APELLIDO') || rowStr.includes('NOMBRE') || rowStr.includes('PROFESIONAL')) {
        headerIndex = i;
        break;
      }
    }

    if (headerIndex === -1) {
      alert('No se pudo encontrar la fila de encabezados en el Excel. Asegurate de que tenga una columna "Apellido y Nombre".');
      return;
    }

    const headers = rows[headerIndex].map(h => String(h).toUpperCase().trim());
    const colDia = headers.findIndex(h => h.includes('DIA') || h.includes('DÍA'));
    const colNumero = headers.findIndex(h => h.includes('Nº') || h.includes('N°') || h === 'NUMERO' || h === 'NÚMERO');
    const colNombre = headers.findIndex(h => h.includes('APELLIDO') || h.includes('NOMBRE') || h.includes('PROFESIONAL'));
    const colViatico = headers.findIndex(h => h.includes('VIATICO') || h.includes('VIÁTICO'));

    if (colNombre === -1) {
      alert('No se encontró la columna de Profesional/Apellido.');
      return;
    }

    const nuevasGuardias = [];
    let mesDetectado = '';

    for (let i = headerIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const diaTexto = colDia !== -1 ? String(row[colDia]).trim() : '';
      const numero = colNumero !== -1 ? parseInt(row[colNumero]) : null;
      const nombre = colNombre !== -1 ? String(row[colNombre]).trim() : '';
      
      if (!mesDetectado && diaTexto) {
        const matchMes = diaTexto.match(/(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)/i);
        if (matchMes) mesDetectado = matchMes[0];
      }

      if (!nombre) continue;

      let diaSemana = '';
      if (diaTexto) {
        const diaUpper = diaTexto.toUpperCase();
        if (diaUpper.includes('LUNES')) diaSemana = 'Lunes';
        else if (diaUpper.includes('MARTES')) diaSemana = 'Martes';
        else if (diaUpper.includes('MIERCOLES') || diaUpper.includes('MIÉRCOLES')) diaSemana = 'Miércoles';
        else if (diaUpper.includes('JUEVES')) diaSemana = 'Jueves';
        else if (diaUpper.includes('VIERNES')) diaSemana = 'Viernes';
        else if (diaUpper.includes('SABADO') || diaUpper.includes('SÁBADO')) diaSemana = 'Sábado';
        else if (diaUpper.includes('DOMINGO')) diaSemana = 'Domingo';
      }

      const esFinde = diaSemana === 'Sábado' || diaSemana === 'Domingo';
      
      let horas = 24;
      const matchHoras = nombre.match(/\((\d+)\s*HS?\)/i);
      if (matchHoras) {
        horas = parseInt(matchHoras[1]);
      }
      
      const nombreLimpio = nombre.replace(/\(\d+\s*HS?\)/i, '').trim();
      const viaticoExcel = colViatico !== -1 ? parseFloat(row[colViatico]) || 0 : 0;

      nuevasGuardias.push({
        id: Date.now() + Math.random(),
        dia: diaSemana || 'Desconocido',
        numero: numero,
        profesional: nombreLimpio,
        horas: horas,
        viatico: viaticoExcel > 0,
        esFinde: esFinde,
        esFeriado: false,
        mes: mesDetectado || 'Actual'
      });
    }

    if (nuevasGuardias.length === 0) {
      alert('No se encontraron datos de guardias en el Excel.');
      return;
    }

    guardias = nuevasGuardias;
    document.getElementById('mesTitulo').textContent = mesDetectado ? `${mesDetectado} 2026` : 'Datos Importados';
    
    recalcularTodo();
    guardarEnLocalStorage();
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// CÁLCULOS
// ============================================================
function calcularMontoBase(guardia, config) {
  const valorDia = (guardia.esFinde || guardia.esFeriado) ? config.finde : config.semana;
  const proporcion = guardia.horas / 24;
  return Math.round(valorDia * proporcion);
}

function calcularViatico(guardia, config) {
  return guardia.viatico ? config.viatico : 0;
}

function recalcularTodo() {
  const config = obtenerValoresConfig();
  
  guardias.forEach(g => {
    g.montoBase = calcularMontoBase(g, config);
    g.montoViatico = calcularViatico(g, config);
    g.total = g.montoBase + g.montoViatico;
  });

  renderizarTabla();
  renderizarResumen();
  guardarEnLocalStorage();
}

// ============================================================
// RENDERIZAR TABLA PRINCIPAL
// ============================================================
function renderizarTabla() {
  const tbody = document.getElementById('cuerpoTabla');
  if (!tbody) return;

  if (guardias.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:24px;color:#9ca3af;">No hay guardias cargadas. Importá un Excel o agregá un día manualmente.</td></tr>';
    return;
  }

  guardias.sort((a, b) => (a.numero || 0) - (b.numero || 0));

  tbody.innerHTML = guardias.map((g) => {
    let rowClass = '';
    if (g.esFeriado) rowClass = 'feriado';
    else if (g.esFinde) rowClass = 'finde';

    return `
      <tr class="${rowClass}" data-id="${g.id}">
        <td>
          <select onchange="actualizarGuardia('${g.id}', 'dia', this.value); recalcularTodo();">
            ${DIAS_SEMANA.map(d => `<option value="${d}" ${g.dia === d ? 'selected' : ''}>${d}</option>`).join('')}
          </select>
        </td>
        <td>
          <input type="number" value="${g.numero || ''}" style="width:60px;" onchange="actualizarGuardia('${g.id}', 'numero', parseInt(this.value)); recalcularTodo();">
        </td>
        <td>
          <input type="text" value="${g.profesional}" onchange="actualizarGuardia('${g.id}', 'profesional', this.value); recalcularTodo();">
        </td>
        <td>
          <input type="number" value="${g.horas}" style="width:70px;" onchange="actualizarGuardia('${g.id}', 'horas', parseInt(this.value)); recalcularTodo();">
        </td>
        <td style="text-align:center;">
          <input type="checkbox" ${g.viatico ? 'checked' : ''} onchange="actualizarGuardia('${g.id}', 'viatico', this.checked); recalcularTodo();">
        </td>
        <td style="text-align:center;">
          <input type="checkbox" ${g.esFeriado ? 'checked' : ''} onchange="actualizarGuardia('${g.id}', 'esFeriado', this.checked); recalcularTodo();">
        </td>
        <td class="monto-calc">${formatearMoneda(g.montoBase || 0)}</td>
        <td class="monto-calc">${formatearMoneda(g.montoViatico || 0)}</td>
        <td class="total-calc">${formatearMoneda(g.total || 0)}</td>
        <td style="text-align:center;">
          <button class="danger" style="padding:4px 8px;font-size:11px;" onclick="eliminarGuardia('${g.id}')">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="icon"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function actualizarGuardia(id, campo, valor) {
  const g = guardias.find(x => x.id == id);
  if (!g) return;
  
  if (campo === 'dia') {
    g.dia = valor;
    g.esFinde = valor === 'Sábado' || valor === 'Domingo';
  } else if (campo === 'viatico') {
    g.viatico = valor;
  } else if (campo === 'esFeriado') {
    g.esFeriado = valor;
  } else if (campo === 'horas') {
    g.horas = valor;
  } else {
    g[campo] = valor;
  }
}

function eliminarGuardia(id) {
  if (!confirm('¿Eliminar esta guardia?')) return;
  guardias = guardias.filter(g => g.id != id);
  recalcularTodo();
}

// ============================================================
// MODAL NUEVO DÍA
// ============================================================
function abrirModalNuevoDia() {
  // Poblar el select de días si está vacío
  const select = document.getElementById('nuevoDiaSemana');
  if (select.options.length === 0) {
    select.innerHTML = DIAS_SEMANA.map(d => `<option value="${d}">${d}</option>`).join('');
  }
  
  // Limpiar campos
  document.getElementById('nuevoDiaSemana').value = 'Lunes';
  document.getElementById('nuevoDiaNumero').value = '';
  document.getElementById('nuevoDiaProfesional').value = '';
  document.getElementById('nuevoDiaHoras').value = 24;
  document.getElementById('nuevoDiaViatico').checked = false;
  document.getElementById('nuevoDiaFeriado').checked = false;
  
  document.getElementById('modalNuevoDia').classList.remove('hidden');
}

function cerrarModalNuevoDia() {
  document.getElementById('modalNuevoDia').classList.add('hidden');
}

function guardarNuevoDia() {
  const dia = document.getElementById('nuevoDiaSemana').value;
  const numero = parseInt(document.getElementById('nuevoDiaNumero').value) || null;
  const profesional = document.getElementById('nuevoDiaProfesional').value.trim();
  const horas = parseInt(document.getElementById('nuevoDiaHoras').value) || 24;
  const viatico = document.getElementById('nuevoDiaViatico').checked;
  const esFeriado = document.getElementById('nuevoDiaFeriado').checked;

  if (!profesional && numero === null) {
    alert('Completá al menos el número de día o el nombre del profesional.');
    return;
  }

  guardias.push({
    id: Date.now(),
    dia: dia,
    numero: numero,
    profesional: profesional,
    horas: horas,
    viatico: viatico,
    esFinde: dia === 'Sábado' || dia === 'Domingo',
    esFeriado: esFeriado,
    mes: 'Manual'
  });

  recalcularTodo();
  cerrarModalNuevoDia();
}

function vaciarTodo() {
  if (!confirm('¿Vaciar todas las guardias? Esta acción no se puede deshacer.')) return;
  guardias = [];
  recalcularTodo();
  document.getElementById('mesTitulo').textContent = 'Sin datos';
}

// ============================================================
// RENDERIZAR RESUMEN POR QUINCENA
// ============================================================
function renderizarResumen() {
  const tbody = document.getElementById('cuerpoResumen');
  const tfoot = document.getElementById('pieResumen');
  if (!tbody) return;

  if (guardias.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:24px;color:#9ca3af;">Sin datos para resumir.</td></tr>';
    tfoot.innerHTML = '';
    return;
  }

  const profesionales = {};
  
  guardias.forEach(g => {
    const nombre = g.profesional || 'Sin nombre';
    if (!profesionales[nombre]) {
      profesionales[nombre] = { q1: 0, q2: 0 };
    }
    
    const esPrimeraQuincena = (g.numero || 0) <= 15;
    if (esPrimeraQuincena) {
      profesionales[nombre].q1 += g.total || 0;
    } else {
      profesionales[nombre].q2 += g.total || 0;
    }
  });

  const nombresOrdenados = Object.keys(profesionales).sort();

  let totalQ1 = 0;
  let totalQ2 = 0;

  tbody.innerHTML = nombresOrdenados.map(nombre => {
    const p = profesionales[nombre];
    totalQ1 += p.q1;
    totalQ2 += p.q2;
    return `
      <tr>
        <td><strong>${nombre}</strong></td>
        <td>${formatearMoneda(p.q1)}</td>
        <td>${formatearMoneda(p.q2)}</td>
        <td><strong>${formatearMoneda(p.q1 + p.q2)}</strong></td>
        <td style="text-align:center;">
          <button class="btn-recibo" onclick="descargarRecibo('${nombre.replace(/'/g, "\\'")}')" title="Ver recibo individual">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="icon"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            Recibo
          </button>
        </td>
      </tr>
    `;
  }).join('');

  tfoot.innerHTML = `
    <tr>
      <td>TOTAL GENERAL</td>
      <td>${formatearMoneda(totalQ1)}</td>
      <td>${formatearMoneda(totalQ2)}</td>
      <td>${formatearMoneda(totalQ1 + totalQ2)}</td>
      <td></td>
    </tr>
  `;
}

// ============================================================
// RECIBO INDIVIDUAL EN NUEVA PESTAÑA
// ============================================================
function descargarRecibo(nombreProfesional) {
  const fechaHoy = new Date().toLocaleDateString('es-AR', { day: 'numeric', month: 'long', year: 'numeric' });
  const mesTexto = document.getElementById('mesTitulo').textContent;

  const misGuardias = guardias.filter(g => g.profesional === nombreProfesional);
  if (misGuardias.length === 0) {
    alert('No hay guardias para este profesional.');
    return;
  }

  let q1 = 0, q2 = 0;
  misGuardias.forEach(g => {
    if ((g.numero || 0) <= 15) q1 += g.total || 0;
    else q2 += g.total || 0;
  });
  const total = q1 + q2;

  const baseUrl = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);
  const urlSantaFe = baseUrl + 'santa fe.webp';
  const urlHospital = baseUrl + 'hospital.png';

  const htmlRecibo = `
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8">
      <title>Recibo - ${nombreProfesional}</title>
      <style>
        @page { size: A4 portrait; margin: 15mm; }
        * { box-sizing: border-box; }
        body { font-family: Arial, Helvetica, sans-serif; color: #000; margin: 0; padding: 20px; background: #fff; font-size: 11pt; }
        .header { position: relative; padding-bottom: 12px; margin-bottom: 20px; border-bottom: 2px solid #000; min-height: 65px; }
        .header .logo-santafe { position: absolute; top: 0; right: 0; height: 40px; width: auto; object-fit: contain; }
        .header .logo-hospital { display: block; margin: 0 auto 10px auto; height: 60px; width: auto; object-fit: contain; }
        .header h1 { text-align: center; font-size: 15pt; font-weight: 700; margin: 0 0 4px 0; letter-spacing: 0.04em; text-transform: uppercase; }
        .header .sub { text-align: center; font-size: 10pt; color: #333; }
        .datos { display: flex; justify-content: space-between; font-size: 11pt; margin-bottom: 12px; }
        .datos p { margin: 3px 0; }
        table { width: 100%; border-collapse: collapse; font-size: 10pt; margin-bottom: 12px; }
        th, td { border: 1px solid #000; padding: 5px 7px; text-align: left; }
        th { background: #e5e7eb; font-weight: 700; text-transform: uppercase; font-size: 9pt; }
        .totales { display: flex; justify-content: flex-end; gap: 25px; font-size: 11pt; font-weight: 700; margin-top: 10px; margin-bottom: 20px; }
        .firmas { margin-top: 60px; display: flex; justify-content: space-between; }
        .firmas div { border-top: 1px solid #000; width: 220px; text-align: center; padding-top: 6px; font-size: 9pt; }
        .no-print { text-align: center; margin-bottom: 20px; display: flex; gap: 10px; justify-content: center; }
        .no-print button { padding: 8px 20px; font-size: 13px; background: #111827; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-family: inherit; }
        .no-print button:hover { background: #000; }
        .no-print button.secundario { background: #fff; color: #111827; border: 1px solid #d1d5db; }
        .no-print button.secundario:hover { background: #f9fafb; }
        @media print { .no-print { display: none !important; } body { padding: 0; margin: 0; } }
      </style>
    </head>
    <body>
      <div class="no-print">
        <button onclick="window.print()">🖨️ Imprimir / Guardar como PDF</button>
        <button class="secundario" onclick="window.close()">Cerrar</button>
      </div>

      <div class="header">
        <img src="${urlSantaFe}" alt="Santa Fe" class="logo-santafe">
        <img src="${urlHospital}" alt="Hospital" class="logo-hospital">
        <h1>Hospital Julio César Villanueva</h1>
        <div class="sub">San Cristóbal — Santa Fe</div>
        <div class="sub"><strong>Recibo de Haberes — Guardias Médicas</strong></div>
        <div class="sub">${mesTexto}</div>
      </div>

      <div class="datos">
        <p><strong>Profesional:</strong> ${nombreProfesional}</p>
        <p><strong>Fecha de emisión:</strong> ${fechaHoy}</p>
      </div>

      <table>
        <thead>
          <tr>
            <th>Día</th>
            <th>Nº</th>
            <th>Horas</th>
            <th>Viático</th>
            <th>Monto Base</th>
            <th>Viático $</th>
            <th>Total</th>
          </tr>
        </thead>
        <tbody>
          ${misGuardias.map(g => `
            <tr>
              <td>${g.dia}</td>
              <td>${g.numero || ''}</td>
              <td>${g.horas}</td>
              <td>${g.viatico ? 'Sí' : 'No'}</td>
              <td>${formatearMoneda(g.montoBase || 0)}</td>
              <td>${formatearMoneda(g.montoViatico || 0)}</td>
              <td><strong>${formatearMoneda(g.total || 0)}</strong></td>
            </tr>
          `).join('')}
        </tbody>
      </table>

      <div class="totales">
        <span>1° Quincena: ${formatearMoneda(q1)}</span>
        <span>2° Quincena: ${formatearMoneda(q2)}</span>
        <span>TOTAL A LIQUIDAR: ${formatearMoneda(total)}</span>
      </div>

      <div class="firmas">
        <div>Firma del Profesional</div>
        <div>Firma y Sello — Administración</div>
      </div>
    </body>
    </html>
  `;

  const ventana = window.open('', '_blank');
  if (!ventana) {
    alert('El navegador bloqueó la apertura de la nueva pestaña. Permití las ventanas emergentes para este sitio.');
    return;
  }
  ventana.document.write(htmlRecibo);
  ventana.document.close();
}

// ============================================================
// IMPRIMIR SOLO HABERES
// ============================================================
function imprimirHaberes() {
  if (guardias.length === 0) {
    alert('No hay datos para imprimir.');
    return;
  }

  const mesTexto = document.getElementById('mesTitulo').textContent;
  const fechaHoy = new Date().toLocaleDateString('es-AR', { day: 'numeric', month: 'long', year: 'numeric' });

  const profesionales = {};
  guardias.forEach(g => {
    const nombre = g.profesional || 'Sin nombre';
    if (!profesionales[nombre]) {
      profesionales[nombre] = { q1: 0, q2: 0, guardias: [] };
    }
    const esPrimeraQuincena = (g.numero || 0) <= 15;
    if (esPrimeraQuincena) {
      profesionales[nombre].q1 += g.total || 0;
    } else {
      profesionales[nombre].q2 += g.total || 0;
    }
    profesionales[nombre].guardias.push(g);
  });

  const nombresOrdenados = Object.keys(profesionales).sort();

  const baseUrl = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);
  const urlSantaFe = baseUrl + 'santa fe.webp';
  const urlHospital = baseUrl + 'hospital.png';

  let htmlHaberes = `
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8">
      <title>Haberes Profesionales - ${mesTexto}</title>
      <style>
        @page { size: A4 portrait; margin: 15mm; }
        * { box-sizing: border-box; }
        body { font-family: Arial, Helvetica, sans-serif; color: #000; margin: 0; padding: 20px; background: #fff; font-size: 11pt; }
        .header { position: relative; padding-bottom: 12px; margin-bottom: 20px; border-bottom: 2px solid #000; min-height: 65px; }
        .header .logo-santafe { position: absolute; top: 0; right: 0; height: 40px; width: auto; object-fit: contain; }
        .header .logo-hospital { display: block; margin: 0 auto 10px auto; height: 60px; width: auto; object-fit: contain; }
        .header h1 { text-align: center; font-size: 15pt; font-weight: 700; margin: 0 0 4px 0; letter-spacing: 0.04em; text-transform: uppercase; }
        .header .sub { text-align: center; font-size: 10pt; color: #333; }
        .recibo { page-break-inside: avoid; border: 1px solid #000; padding: 16px 18px; margin-bottom: 24px; }
        .recibo-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #000; padding-bottom: 6px; margin-bottom: 12px; }
        .recibo-header h3 { margin: 0; font-size: 12pt; text-transform: uppercase; }
        .recibo-header span { font-size: 10pt; color: #333; }
        .recibo-datos { display: flex; justify-content: space-between; font-size: 10pt; margin-bottom: 10px; }
        .recibo-datos p { margin: 3px 0; }
        table { width: 100%; border-collapse: collapse; font-size: 10pt; margin-bottom: 12px; }
        th, td { border: 1px solid #000; padding: 5px 7px; text-align: left; }
        th { background: #e5e7eb; font-weight: 700; text-transform: uppercase; font-size: 9pt; }
        .recibo-totales { display: flex; justify-content: flex-end; gap: 25px; font-size: 11pt; font-weight: 700; margin-top: 8px; }
        .recibo-firma { margin-top: 45px; display: flex; justify-content: space-between; }
        .recibo-firma div { border-top: 1px solid #000; width: 220px; text-align: center; padding-top: 6px; font-size: 9pt; }
        .no-print { text-align: center; margin-bottom: 20px; display: flex; gap: 10px; justify-content: center; }
        .no-print button { padding: 8px 20px; font-size: 13px; background: #111827; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-family: inherit; }
        .no-print button:hover { background: #000; }
        .no-print button.secundario { background: #fff; color: #111827; border: 1px solid #d1d5db; }
        .no-print button.secundario:hover { background: #f9fafb; }
        .page-break { page-break-after: always; }
        @media print { .no-print { display: none !important; } body { padding: 0; margin: 0; } }
      </style>
    </head>
    <body>
      <div class="no-print">
        <button onclick="window.print()">🖨️ Imprimir / Guardar como PDF</button>
        <button class="secundario" onclick="window.close()">Cerrar</button>
      </div>

      <div class="header">
        <img src="${urlSantaFe}" alt="Santa Fe" class="logo-santafe">
        <img src="${urlHospital}" alt="Hospital" class="logo-hospital">
        <h1>Hospital Julio César Villanueva</h1>
        <div class="sub">San Cristóbal — Santa Fe</div>
        <div class="sub"><strong>Recibos de Haberes — Guardias Médicas</strong></div>
        <div class="sub">${mesTexto} | Generado el ${fechaHoy}</div>
      </div>

      <h2 class="seccion" style="font-size:12pt; margin:20px 0 8px 0; text-transform:uppercase; border-bottom:1px solid #666; padding-bottom:3px;">Recibos Individuales por Profesional</h2>
      
      ${nombresOrdenados.map((nombre, index) => {
        const p = profesionales[nombre];
        const total = p.q1 + p.q2;
        return `
          <div class="recibo">
            <div class="recibo-header">
              <h3>Recibo de Haberes — Guardias Médicas</h3>
              <span>${mesTexto}</span>
            </div>
            <div class="recibo-datos">
              <p><strong>Profesional:</strong> ${nombre}</p>
              <p><strong>Fecha de emisión:</strong> ${fechaHoy}</p>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Día</th>
                  <th>Nº</th>
                  <th>Horas</th>
                  <th>Viático</th>
                  <th>Monto Base</th>
                  <th>Viático $</th>
                  <th>Total</th>
                </tr>
              </thead>
              <tbody>
                ${p.guardias.map(g => `
                  <tr>
                    <td>${g.dia}</td>
                    <td>${g.numero || ''}</td>
                    <td>${g.horas}</td>
                    <td>${g.viatico ? 'Sí' : 'No'}</td>
                    <td>${formatearMoneda(g.montoBase || 0)}</td>
                    <td>${formatearMoneda(g.montoViatico || 0)}</td>
                    <td>${formatearMoneda(g.total || 0)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
            <div class="recibo-totales">
              <span>1° Quincena: ${formatearMoneda(p.q1)}</span>
              <span>2° Quincena: ${formatearMoneda(p.q2)}</span>
              <span>TOTAL A LIQUIDAR: ${formatearMoneda(total)}</span>
            </div>
            <div class="recibo-firma">
              <div>Firma del Profesional</div>
              <div>Firma y Sello — Administración</div>
            </div>
          </div>
          ${index < nombresOrdenados.length - 1 ? '<div class="page-break"></div>' : ''}
        `;
      }).join('')}
    </body>
    </html>
  `;

  const ventana = window.open('', '_blank');
  if (!ventana) {
    alert('El navegador bloqueó la apertura de la nueva pestaña. Permití las ventanas emergentes para este sitio.');
    return;
  }
  ventana.document.write(htmlHaberes);
  ventana.document.close();
}

// ============================================================
// REPORTE COMPLETO EN NUEVA PESTAÑA
// ============================================================
function imprimirReporte() {
  if (guardias.length === 0) {
    alert('No hay datos para imprimir.');
    return;
  }

  const mesTexto = document.getElementById('mesTitulo').textContent;
  const fechaHoy = new Date().toLocaleDateString('es-AR', { day: 'numeric', month: 'long', year: 'numeric' });

  const profesionales = {};
  guardias.forEach(g => {
    const nombre = g.profesional || 'Sin nombre';
    if (!profesionales[nombre]) {
      profesionales[nombre] = { q1: 0, q2: 0, guardias: [] };
    }
    const esPrimeraQuincena = (g.numero || 0) <= 15;
    if (esPrimeraQuincena) {
      profesionales[nombre].q1 += g.total || 0;
    } else {
      profesionales[nombre].q2 += g.total || 0;
    }
    profesionales[nombre].guardias.push(g);
  });

  const nombresOrdenados = Object.keys(profesionales).sort();
  const totalGeneral = guardias.reduce((acc, g) => acc + (g.total || 0), 0);

  const baseUrl = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);
  const urlSantaFe = baseUrl + 'santa fe.webp';
  const urlHospital = baseUrl + 'hospital.png';

  const htmlReporte = `
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8">
      <title>Reporte de Guardias - ${mesTexto}</title>
      <style>
        @page { size: A4 portrait; margin: 15mm; }
        * { box-sizing: border-box; }
        body { font-family: Arial, Helvetica, sans-serif; color: #000; margin: 0; padding: 20px; background: #fff; font-size: 11pt; }
        .header { position: relative; padding-bottom: 12px; margin-bottom: 20px; border-bottom: 2px solid #000; min-height: 65px; }
        .header .logo-santafe { position: absolute; top: 0; right: 0; height: 40px; width: auto; object-fit: contain; }
        .header .logo-hospital { display: block; margin: 0 auto 10px auto; height: 60px; width: auto; object-fit: contain; }
        .header h1 { text-align: center; font-size: 15pt; font-weight: 700; margin: 0 0 4px 0; letter-spacing: 0.04em; text-transform: uppercase; }
        .header .sub { text-align: center; font-size: 10pt; color: #333; }
        h2.seccion { font-size: 12pt; margin: 20px 0 8px 0; text-transform: uppercase; border-bottom: 1px solid #666; padding-bottom: 3px; }
        table { width: 100%; border-collapse: collapse; font-size: 10pt; margin-bottom: 12px; }
        th, td { border: 1px solid #000; padding: 5px 7px; text-align: left; }
        th { background: #e5e7eb; font-weight: 700; text-transform: uppercase; font-size: 9pt; }
        tfoot td { font-weight: 700; background: #f3f4f6; }
        .recibo { page-break-inside: avoid; border: 1px solid #000; padding: 16px 18px; margin-bottom: 24px; }
        .recibo-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #000; padding-bottom: 6px; margin-bottom: 12px; }
        .recibo-header h3 { margin: 0; font-size: 12pt; text-transform: uppercase; }
        .recibo-header span { font-size: 10pt; color: #333; }
        .recibo-datos { display: flex; justify-content: space-between; font-size: 10pt; margin-bottom: 10px; }
        .recibo-datos p { margin: 3px 0; }
        .recibo-totales { display: flex; justify-content: flex-end; gap: 25px; font-size: 11pt; font-weight: 700; margin-top: 8px; }
        .recibo-firma { margin-top: 45px; display: flex; justify-content: space-between; }
        .recibo-firma div { border-top: 1px solid #000; width: 220px; text-align: center; padding-top: 6px; font-size: 9pt; }
        .no-print { text-align: center; margin-bottom: 20px; display: flex; gap: 10px; justify-content: center; }
        .no-print button { padding: 8px 20px; font-size: 13px; background: #111827; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-family: inherit; }
        .no-print button:hover { background: #000; }
        .no-print button.secundario { background: #fff; color: #111827; border: 1px solid #d1d5db; }
        .no-print button.secundario:hover { background: #f9fafb; }
        .page-break { page-break-after: always; }
        @media print { .no-print { display: none !important; } body { padding: 0; margin: 0; } .recibo { border: 1px solid #000; } }
      </style>
    </head>
    <body>
      <div class="no-print">
        <button onclick="window.print()">🖨️ Imprimir / Guardar como PDF</button>
        <button class="secundario" onclick="window.close()">Cerrar</button>
      </div>

      <div class="header">
        <img src="${urlSantaFe}" alt="Santa Fe" class="logo-santafe">
        <img src="${urlHospital}" alt="Hospital" class="logo-hospital">
        <h1>Hospital Julio César Villanueva</h1>
        <div class="sub">San Cristóbal — Santa Fe</div>
        <div class="sub"><strong>Reporte de Guardias Médicas — ${mesTexto}</strong></div>
        <div class="sub">Generado el ${fechaHoy}</div>
      </div>

      <h2 class="seccion">Detalle General de Guardias</h2>
      <table>
        <thead>
          <tr>
            <th>Día</th>
            <th>Nº</th>
            <th>Profesional</th>
            <th>Horas</th>
            <th>Viático</th>
            <th>Feriado</th>
            <th>Monto Base</th>
            <th>Viático $</th>
            <th>Total</th>
          </tr>
        </thead>
        <tbody>
          ${guardias.map(g => `
            <tr>
              <td>${g.dia}</td>
              <td>${g.numero || ''}</td>
              <td>${g.profesional}</td>
              <td>${g.horas}</td>
              <td>${g.viatico ? 'Sí' : 'No'}</td>
              <td>${g.esFeriado ? 'Sí' : 'No'}</td>
              <td>${formatearMoneda(g.montoBase || 0)}</td>
              <td>${formatearMoneda(g.montoViatico || 0)}</td>
              <td><strong>${formatearMoneda(g.total || 0)}</strong></td>
            </tr>
          `).join('')}
        </tbody>
        <tfoot>
          <tr>
            <td colspan="8" style="text-align:right;">TOTAL GENERAL</td>
            <td>${formatearMoneda(totalGeneral)}</td>
          </tr>
        </tfoot>
      </table>

      <div class="page-break"></div>

      <h2 class="seccion">Recibos Individuales por Profesional</h2>
      
      ${nombresOrdenados.map((nombre, index) => {
        const p = profesionales[nombre];
        const total = p.q1 + p.q2;
        return `
          <div class="recibo">
            <div class="recibo-header">
              <h3>Recibo de Haberes — Guardias Médicas</h3>
              <span>${mesTexto}</span>
            </div>
            <div class="recibo-datos">
              <p><strong>Profesional:</strong> ${nombre}</p>
              <p><strong>Fecha de emisión:</strong> ${fechaHoy}</p>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Día</th>
                  <th>Nº</th>
                  <th>Horas</th>
                  <th>Viático</th>
                  <th>Feriado</th>
                  <th>Monto Base</th>
                  <th>Viático $</th>
                  <th>Total</th>
                </tr>
              </thead>
              <tbody>
                ${p.guardias.map(g => `
                  <tr>
                    <td>${g.dia}</td>
                    <td>${g.numero || ''}</td>
                    <td>${g.horas}</td>
                    <td>${g.viatico ? 'Sí' : 'No'}</td>
                    <td>${g.esFeriado ? 'Sí' : 'No'}</td>
                    <td>${formatearMoneda(g.montoBase || 0)}</td>
                    <td>${formatearMoneda(g.montoViatico || 0)}</td>
                    <td>${formatearMoneda(g.total || 0)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
            <div class="recibo-totales">
              <span>1° Quincena: ${formatearMoneda(p.q1)}</span>
              <span>2° Quincena: ${formatearMoneda(p.q2)}</span>
              <span>TOTAL A LIQUIDAR: ${formatearMoneda(total)}</span>
            </div>
            <div class="recibo-firma">
              <div>Firma del Profesional</div>
              <div>Firma y Sello — Administración</div>
            </div>
          </div>
          ${index < nombresOrdenados.length - 1 ? '<div class="page-break"></div>' : ''}
        `;
      }).join('')}
    </body>
    </html>
  `;

  const ventana = window.open('', '_blank');
  if (!ventana) {
    alert('El navegador bloqueó la apertura de la nueva pestaña. Permití las ventanas emergentes para este sitio.');
    return;
  }
  ventana.document.write(htmlReporte);
  ventana.document.close();
}

// ============================================================
// EXPORTAR A EXCEL
// ============================================================
function exportarAExcel() {
  if (guardias.length === 0) {
    alert('No hay datos para exportar.');
    return;
  }

  const datosExcel = guardias.map(g => ({
    'Día': g.dia,
    'Nº': g.numero,
    'Profesional': g.profesional,
    'Horas': g.horas,
    'Viático': g.viatico ? 'Sí' : 'No',
    'Feriado': g.esFeriado ? 'Sí' : 'No',
    'Monto Base': g.montoBase,
    'Monto Viático': g.montoViatico,
    'Total': g.total
  }));

  const ws = XLSX.utils.json_to_sheet(datosExcel);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Guardias');
  XLSX.writeFile(wb, `Guardias_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ============================================================
// EXPORTAR / IMPORTAR JSON
// ============================================================
function exportarJSON() {
  const data = {
    guardias: guardias,
    config: obtenerValoresConfig(),
    mes: document.getElementById('mesTitulo').textContent,
    fechaExportacion: new Date().toISOString()
  };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `guardias_backup_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function importarJSON(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      if (!data.guardias || !Array.isArray(data.guardias)) {
        alert('El archivo JSON no tiene el formato esperado.');
        return;
      }
      guardias = data.guardias;
      if (data.config) {
        document.getElementById('valorSemana').value = data.config.semana || 340000;
        document.getElementById('valorFinde').value = data.config.finde || 380000;
        document.getElementById('valorViatico').value = data.config.viatico || 60000;
      }
      if (data.mes) {
        document.getElementById('mesTitulo').textContent = data.mes;
      } else {
        document.getElementById('mesTitulo').textContent = 'Datos Importados';
      }
      recalcularTodo();
      alert('Datos importados correctamente.');
    } catch (err) {
      alert('Error al leer el archivo JSON: ' + err.message);
    }
    event.target.value = '';
  };
  reader.readAsText(file);
}

// ============================================================
// MODAL PLANTILLA EXCEL
// ============================================================
function abrirModalPlantilla() {
  const hoy = new Date();
  document.getElementById('plantillaMes').value = hoy.getMonth();
  document.getElementById('plantillaAnio').value = hoy.getFullYear();
  document.getElementById('modalPlantilla').classList.remove('hidden');
}

function cerrarModalPlantilla() {
  document.getElementById('modalPlantilla').classList.add('hidden');
}

// ============================================================
// DESCARGAR PLANTILLA EXCEL EN BLANCO
// ============================================================
function descargarPlantillaExcel() {
  const mes = parseInt(document.getElementById('plantillaMes').value);
  const anio = parseInt(document.getElementById('plantillaAnio').value);
  
  if (isNaN(mes) || isNaN(anio)) {
    alert('Seleccioná un mes y un año válidos.');
    return;
  }

  const diasSemana = ['DOMINGO', 'LUNES', 'MARTES', 'MIÉRCOLES', 'JUEVES', 'VIERNES', 'SÁBADO'];
  const diasEnMes = new Date(anio, mes + 1, 0).getDate();
  const nombreMes = NOMBRES_MESES[mes];

  const filas = [];
  filas.push([`GUARDIAS MÉDICAS ${nombreMes} ${anio}`, '', '', '', '']);
  filas.push(['DÍA', 'Nº', 'APELLIDO Y NOMBRE', 'MONTO', 'VIÁTICO']);
  
  for (let dia = 1; dia <= diasEnMes; dia++) {
    const fecha = new Date(anio, mes, dia);
    const nombreDia = diasSemana[fecha.getDay()];
    filas.push([nombreDia, dia, '', '', '']);
  }

  const ws = XLSX.utils.aoa_to_sheet(filas);

  ws['!cols'] = [
    { wch: 14 },
    { wch: 5 },
    { wch: 28 },
    { wch: 15 },
    { wch: 12 }
  ];

  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `${nombreMes} ${anio}`);
  XLSX.writeFile(wb, `Plantilla_Guardias_${nombreMes}_${anio}.xlsx`);

  cerrarModalPlantilla();
}

// ============================================================
// PERSISTENCIA LOCAL
// ============================================================
function guardarEnLocalStorage() {
  try {
    localStorage.setItem('guardias_data', JSON.stringify(guardias));
    const config = obtenerValoresConfig();
    localStorage.setItem('guardias_config', JSON.stringify(config));
  } catch(e) { console.warn('No se pudo guardar en localStorage', e); }
}

function cargarDeLocalStorage() {
  try {
    const data = localStorage.getItem('guardias_data');
    if (data) {
      guardias = JSON.parse(data);
    }
    const config = localStorage.getItem('guardias_config');
    if (config) {
      const c = JSON.parse(config);
      document.getElementById('valorSemana').value = c.semana || 340000;
      document.getElementById('valorFinde').value = c.finde || 380000;
      document.getElementById('valorViatico').value = c.viatico || 60000;
    }
    if (guardias.length > 0) {
      document.getElementById('mesTitulo').textContent = 'Datos Restaurados';
      recalcularTodo();
    }
  } catch(e) { console.warn('Error cargando de localStorage', e); }
}

// ============================================================
// INICIALIZACIÓN
// ============================================================
window.addEventListener('DOMContentLoaded', () => {
  cargarDeLocalStorage();
  if (guardias.length === 0) {
    renderizarTabla();
    renderizarResumen();
  }
});