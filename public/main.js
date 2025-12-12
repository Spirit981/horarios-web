document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('form-upload');
  const inputExcel = document.getElementById('input-excel');
  const chkAusencias = document.getElementById('chk-ausencias');
  const chkBolsa = document.getElementById('chk-bolsa');
  const chkDisfruteHH = document.getElementById('chk-disfrutehh');
  const chk6Dias = document.getElementById('chk-6dias');
  const divResultado = document.getElementById('resultado');

  // Filtros por semana
  const inputSemanaDesde = document.getElementById('semana-desde');
  const inputSemanaHasta = document.getElementById('semana-hasta');

  // Subfiltros de ausencias
  const ausPanel = document.getElementById('filtros-ausencias');
  const ausAP = document.getElementById('aus-AP');
  const ausBM = document.getElementById('aus-BM');
  const ausHH = document.getElementById('aus-HH');
  const ausVAC = document.getElementById('aus-VAC');
  const ausOTRAS = document.getElementById('aus-OTRAS');

  const btnConfig = document.getElementById('btn-config');
  const panelConfig = document.getElementById('panel-config');
  const btnGuardarMinimos = document.getElementById('btn-guardar-minimos');
  const msgMinimos = document.getElementById('msg-minimos');

  const btnEstadisticas = document.getElementById('btn-estadisticas');

  let registrosGlobal = [];
  let estadisticasDetalladas = null;

  // Configuración de mínimos en memoria
  let minimosConfig = {
    MANANA: {
      Lunes: 0,
      Martes: 0,
      'Miércoles': 0,
      Jueves: 0,
      Viernes: 0,
      'Sábado': 0
    },
    TARDE: {
      Lunes: 0,
      Martes: 0,
      'Miércoles': 0,
      Jueves: 0,
      Viernes: 0
    },
    NOCHE: {
      Domingo: 0,
      Lunes: 0,
      Martes: 0,
      'Miércoles': 0,
      Jueves: 0,
      Viernes: 0
    }
  };

  function cargarMinimosDesdeStorage() {
    try {
      const raw = localStorage.getItem('minimosConfig');
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        minimosConfig = parsed;
      }
    } catch (e) {
      console.error('Error leyendo minimosConfig de localStorage', e);
    }
  }

  function guardarMinimosEnStorage() {
    try {
      localStorage.setItem('minimosConfig', JSON.stringify(minimosConfig));
    } catch (e) {
      console.error('Error guardando minimosConfig en localStorage', e);
    }
  }

  function rellenarInputsMinimos() {
    // Mañana
    document.getElementById('min-manana-lunes').value = minimosConfig.MANANA.Lunes || 0;
    document.getElementById('min-manana-martes').value = minimosConfig.MANANA.Martes || 0;
    document.getElementById('min-manana-miercoles').value = minimosConfig.MANANA['Miércoles'] || 0;
    document.getElementById('min-manana-jueves').value = minimosConfig.MANANA.Jueves || 0;
    document.getElementById('min-manana-viernes').value = minimosConfig.MANANA.Viernes || 0;
    document.getElementById('min-manana-sabado').value = minimosConfig.MANANA['Sábado'] || 0;

    // Tarde
    document.getElementById('min-tarde-lunes').value = minimosConfig.TARDE.Lunes || 0;
    document.getElementById('min-tarde-martes').value = minimosConfig.TARDE.Martes || 0;
    document.getElementById('min-tarde-miercoles').value = minimosConfig.TARDE['Miércoles'] || 0;
    document.getElementById('min-tarde-jueves').value = minimosConfig.TARDE.Jueves || 0;
    document.getElementById('min-tarde-viernes').value = minimosConfig.TARDE.Viernes || 0;

    // Noche
    document.getElementById('min-noche-domingo').value = minimosConfig.NOCHE.Domingo || 0;
    document.getElementById('min-noche-lunes').value = minimosConfig.NOCHE.Lunes || 0;
    document.getElementById('min-noche-martes').value = minimosConfig.NOCHE.Martes || 0;
    document.getElementById('min-noche-miercoles').value = minimosConfig.NOCHE['Miércoles'] || 0;
    document.getElementById('min-noche-jueves').value = minimosConfig.NOCHE.Jueves || 0;
    document.getElementById('min-noche-viernes').value = minimosConfig.NOCHE.Viernes || 0;
  }

  function leerInputsMinimos() {
    minimosConfig.MANANA.Lunes = Number(document.getElementById('min-manana-lunes').value) || 0;
    minimosConfig.MANANA.Martes = Number(document.getElementById('min-manana-martes').value) || 0;
    minimosConfig.MANANA['Miércoles'] = Number(document.getElementById('min-manana-miercoles').value) || 0;
    minimosConfig.MANANA.Jueves = Number(document.getElementById('min-manana-jueves').value) || 0;
    minimosConfig.MANANA.Viernes = Number(document.getElementById('min-manana-viernes').value) || 0;
    minimosConfig.MANANA['Sábado'] = Number(document.getElementById('min-manana-sabado').value) || 0;

    minimosConfig.TARDE.Lunes = Number(document.getElementById('min-tarde-lunes').value) || 0;
    minimosConfig.TARDE.Martes = Number(document.getElementById('min-tarde-martes').value) || 0;
    minimosConfig.TARDE['Miércoles'] = Number(document.getElementById('min-tarde-miercoles').value) || 0;
    minimosConfig.TARDE.Jueves = Number(document.getElementById('min-tarde-jueves').value) || 0;
    minimosConfig.TARDE.Viernes = Number(document.getElementById('min-tarde-viernes').value) || 0;

    minimosConfig.NOCHE.Domingo = Number(document.getElementById('min-noche-domingo').value) || 0;
    minimosConfig.NOCHE.Lunes = Number(document.getElementById('min-noche-lunes').value) || 0;
    minimosConfig.NOCHE.Martes = Number(document.getElementById('min-noche-martes').value) || 0;
    minimosConfig.NOCHE['Miércoles'] = Number(document.getElementById('min-noche-miercoles').value) || 0;
    minimosConfig.NOCHE.Jueves = Number(document.getElementById('min-noche-jueves').value) || 0;
    minimosConfig.NOCHE.Viernes = Number(document.getElementById('min-noche-viernes').value) || 0;
  }

  // Cargar mínimos al arrancar
  cargarMinimosDesdeStorage();
  rellenarInputsMinimos();

  // Mostrar/ocultar panel de configuración
  if (btnConfig && panelConfig) {
    btnConfig.addEventListener('click', () => {
      const visible = panelConfig.style.display === 'block';
      panelConfig.style.display = visible ? 'none' : 'block';
    });
  }

  // Guardar mínimos
  if (btnGuardarMinimos) {
    btnGuardarMinimos.addEventListener('click', () => {
      leerInputsMinimos();
      guardarMinimosEnStorage();
      msgMinimos.textContent = 'Guardado';
      setTimeout(() => { msgMinimos.textContent = ''; }, 2000);
      renderInforme();
    });
  }

  console.log('DOM elements:', {
    form: !!form,
    inputExcel: !!inputExcel,
    chkAusencias: !!chkAusencias,
    chkBolsa: !!chkBolsa,
    chkDisfruteHH: !!chkDisfruteHH,
    chk6Dias: !!chk6Dias
  });

  if (!form || !inputExcel || !chkAusencias || !chkBolsa || !chkDisfruteHH || !chk6Dias || !divResultado) {
    console.error('Faltan elementos en el DOM');
    return;
  }

  // Eventos filtros
  chkAusencias.addEventListener('change', () => {
    ausPanel.style.display = chkAusencias.checked ? 'block' : 'none';
    renderInforme();
  });
  chkBolsa.addEventListener('change', renderInforme);
  chkDisfruteHH.addEventListener('change', renderInforme);
  chk6Dias.addEventListener('change', renderInforme);

  // Subfiltros ausencias
  [ausAP, ausBM, ausHH, ausVAC, ausOTRAS].forEach(el => {
    el.addEventListener('change', renderInforme);
  });

  // Filtros semana desde/hasta
  if (inputSemanaDesde) inputSemanaDesde.addEventListener('input', renderInforme);
  if (inputSemanaHasta) inputSemanaHasta.addEventListener('input', renderInforme);

  // Botón estadísticas
  if (btnEstadisticas) {
    btnEstadisticas.addEventListener('click', abrirVentanaEstadisticas);
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    console.log('🔄 FORM SUBMIT - procesando...');

    const files = inputExcel.files;
    if (!files || files.length === 0) {
      alert('Selecciona al menos un fichero Excel');
      return;
    }

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
      formData.append('excel', files[i]);
    }

    divResultado.innerHTML = '<p>Procesando ficheros...</p>';

    try {
      const resp = await fetch('/upload', {
        method: 'POST',
        body: formData
      });

      const data = await resp.json();
      console.log('📥 RESPUESTA:', data);

      if (!data.ok) {
        divResultado.innerHTML = `<p style="color: red;">Error: ${data.message}</p>`;
        return;
      }

      registrosGlobal = data.registros || [];
      console.log('✅ REGISTROS:', registrosGlobal.length);

      autocompletarRangoSemanas();
      renderInforme();
    } catch (err) {
      console.error('❌ ERROR:', err);
      divResultado.innerHTML = '<p style="color: red;">Error al enviar los ficheros.</p>';
    }
  });

  function autocompletarRangoSemanas() {
    if (!inputSemanaDesde || !inputSemanaHasta || registrosGlobal.length === 0) return;

    const semanas = Array.from(new Set(registrosGlobal.map(r => r.semana)))
      .filter(s => s != null)
      .sort((a, b) => a - b);

    if (semanas.length === 0) return;

    inputSemanaDesde.value = semanas[0];
    inputSemanaHasta.value = semanas[semanas.length - 1];
  }

  function pasaSubfiltroAusencia(reg) {
    if (reg.tipo !== 'Ausencia') return true;

    const texto = (reg.persona || '').toUpperCase();

    if (texto.includes('(AP)') && !ausAP.checked) return false;
    if (texto.includes('(BM)') && !ausBM.checked) return false;
    if ((texto.includes('HHEE') || texto.includes('HHDD')) && !ausHH.checked) return false;
    if ((texto.includes('(VC)') || texto.includes('(VP)') || texto.includes('(VX)') || texto.includes('(V)')) && !ausVAC.checked) return false;
    if ((texto.includes('(PAT)') || texto.includes('(LIC)') || texto.includes(' HF')) && !ausOTRAS.checked) return false;

    return true;
  }

  function renderInforme() {
    const mostrarAusencias = chkAusencias.checked;
    const mostrarBolsa = chkBolsa.checked;
    const mostrarDisfruteHH = chkDisfruteHH.checked;
    const mostrar6Dias = chk6Dias.checked;

    const semDesde = inputSemanaDesde && inputSemanaDesde.value ? Number(inputSemanaDesde.value) : null;
    const semHasta = inputSemanaHasta && inputSemanaHasta.value ? Number(inputSemanaHasta.value) : null;

    console.log('Filtros:', {
      mostrarAusencias,
      mostrarBolsa,
      mostrarDisfruteHH,
      mostrar6Dias,
      semDesde,
      semHasta
    });

    const registrosFiltrados = registrosGlobal.filter(reg => {
      // Filtro por semana desde/hasta
      if (semDesde !== null && reg.semana < semDesde) return false;
      if (semHasta !== null && reg.semana > semHasta) return false;

      if (reg.tipo === 'Ausencia') {
        if (!mostrarAusencias) return false;
        if (!pasaSubfiltroAusencia(reg)) return false;
      }
      if (reg.tipo === 'Bolsa' && !mostrarBolsa) return false;
      if (reg.tipo === 'DisfruteHH' && !mostrarDisfruteHH) return false;
      if (reg.tipo === '6Dias' && !mostrar6Dias) return false;
      return true;
    });

    // Estadísticas por tipo (de TODOS)
    const stats = {};
    registrosGlobal.forEach(r => {
      stats[r.tipo] = (stats[r.tipo] || 0) + 1;
    });

    // Agrupar por semana, día y turno
    const conteos = {};
    registrosFiltrados.forEach(reg => {
      const clave = `${reg.semana}|${reg.dia}|${reg.turno}`;
      if (!conteos[clave]) {
        conteos[clave] = {
          semana: reg.semana,
          dia: reg.dia,
          turno: reg.turno,
          total: 0,
          tieneBolsa: false
        };
      }
      conteos[clave].total++;
      if (reg.tipo === 'Bolsa' && mostrarBolsa) {
        conteos[clave].tieneBolsa = true;
      }
    });

    const ordenDias = {
      'Domingo': 0,
      'Lunes': 1,
      'Martes': 2,
      'Miércoles': 3,
      'Jueves': 4,
      'Viernes': 5,
      'Sábado': 6
    };
    const ordenTurnos = { 'MAÑANA': 0, 'TARDE': 1, 'NOCHE': 2 };

    const conteosArray = Object.values(conteos).sort((a, b) => {
      if (a.semana !== b.semana) return a.semana - b.semana;
      const diaA = a.dia.split(' ')[0];
      const diaB = b.dia.split(' ')[0];
      if (ordenDias[diaA] !== ordenDias[diaB]) {
        return ordenDias[diaA] - ordenDias[diaB];
      }
      return ordenTurnos[a.turno] - ordenTurnos[b.turno];
    });

    const coloresTurno = {
      'MAÑANA': '#E3F2FD',
      'TARDE': '#FFF8E1',
      'NOCHE': '#F5F5F5'
    };

    // Estadísticas avanzadas
    const casosSinMinimos = [];
    const casosSinMinimosSinBolsa = [];
    const ausenciasPorSemanaDia = {};
    const ausenciasPorSemanaYTipo = {};
    let totalAusencias = 0;

    // Recorremos conteosArray para mínimos/bolsa
    conteosArray.forEach(conteo => {
      const diaNombre = conteo.dia.split(' ')[0];
      const turnoClave = conteo.turno.toUpperCase();

      let minRequerido = 0;
      if (turnoClave === 'MAÑANA') {
        minRequerido = minimosConfig.MANANA[diaNombre] || 0;
      } else if (turnoClave === 'TARDE') {
        minRequerido = minimosConfig.TARDE[diaNombre] || 0;
      } else if (turnoClave === 'NOCHE') {
        minRequerido = minimosConfig.NOCHE[diaNombre] || 0;
      }

      const respeta = conteo.total >= minRequerido;
      if (minRequerido > 0 && !respeta) {
        const item = {
          semana: conteo.semana,
          dia: conteo.dia,
          turno: conteo.turno,
          total: conteo.total,
          minimos: minRequerido,
          tieneBolsa: conteo.tieneBolsa
        };
        casosSinMinimos.push(item);
        if (!conteo.tieneBolsa) {
          casosSinMinimosSinBolsa.push(item);
        }
      }
    });

    // Recorremos registrosGlobal para ausencias (tabla + subtipos)
    registrosGlobal.forEach(reg => {
      if (reg.tipo !== 'Ausencia') return;
      totalAusencias++;

      // Tabla por semana/día
      const clave = `${reg.semana}|${reg.dia}`;
      if (!ausenciasPorSemanaDia[clave]) {
        ausenciasPorSemanaDia[clave] = {
          semana: reg.semana,
          dia: reg.dia,
          total: 0,
          detalles: []
        };
      }
      ausenciasPorSemanaDia[clave].total++;
      ausenciasPorSemanaDia[clave].detalles.push({
        turno: reg.turno,
        persona: reg.persona
      });

      // Subtipo para gráfico por semana
      const texto = (reg.persona || '').toUpperCase();
      let subtipo = 'OTRAS';
      if (texto.includes('(AP)')) subtipo = 'AP';
      else if (texto.includes('(BM)')) subtipo = 'BM';
      else if (texto.includes('HHEE') || texto.includes('HHDD')) subtipo = 'HH';
      else if (texto.includes('(VC)') || texto.includes('(VP)') || texto.includes('(VX)') || texto.includes('(V)')) subtipo = 'VAC';
      else if (texto.includes('(PAT)') || texto.includes('(LIC)') || texto.includes(' HF')) subtipo = 'OTRAS';

      const claveSemanaTipo = `${reg.semana}|${subtipo}`;
      if (!ausenciasPorSemanaYTipo[claveSemanaTipo]) {
        ausenciasPorSemanaYTipo[claveSemanaTipo] = { semana: reg.semana, subtipo, total: 0 };
      }
      ausenciasPorSemanaYTipo[claveSemanaTipo].total++;
    });

    // Rango real de semanas según conteosArray
    const semanasUsadas = Array.from(new Set(conteosArray.map(c => c.semana)))
      .filter(s => s != null)
      .sort((a, b) => a - b);

    const semanaMin = semanasUsadas.length ? semanasUsadas[0] : null;
    const semanaMax = semanasUsadas.length ? semanasUsadas[semanasUsadas.length - 1] : null;

    estadisticasDetalladas = {
      casosSinMinimos,
      casosSinMinimosSinBolsa,
      ausenciasPorSemanaDia,
      ausenciasPorSemanaYTipo,
      totalAusencias,
      statsTipos: stats,
      semanaMin,
      semanaMax
    };

    let html = `
      <h2>📊 Informe de Horarios</h2>
      <p><strong>Total mostrados:</strong> ${registrosFiltrados.length} | 
         <strong>Total registros:</strong> ${registrosGlobal.length}</p>
      <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 15px;">
        <tr style="background: linear-gradient(90deg, #4CAF50, #45a049); color: white;">
          <th style="padding: 15px;">Semana</th>
          <th style="padding: 15px;">Día</th>
          <th style="padding: 15px;">Turno</th>
          <th style="padding: 15px; text-align: center;">Nº Personas</th>
          <th style="padding: 15px; text-align: center;">Respeta mínimos</th>
          <th style="padding: 15px; text-align: center;">Bolsa Activada</th>
        </tr>`;

    conteosArray.forEach(conteo => {
      const colorFila = coloresTurno[conteo.turno] || '#FFFFFF';

      const diaNombre = conteo.dia.split(' ')[0];
      const turnoClave = conteo.turno.toUpperCase();

      let minRequerido = 0;
      if (turnoClave === 'MAÑANA') {
        minRequerido = minimosConfig.MANANA[diaNombre] || 0;
      } else if (turnoClave === 'TARDE') {
        minRequerido = minimosConfig.TARDE[diaNombre] || 0;
      } else if (turnoClave === 'NOCHE') {
        minRequerido = minimosConfig.NOCHE[diaNombre] || 0;
      }

      const respeta = conteo.total >= minRequerido;
      const textoMinimos = minRequerido > 0 ? (respeta ? '✅ SI' : '❌ NO') : '-';

      const bolsaTexto = conteo.tieneBolsa ? '✅ SI' : '❌ NO';

      const casoCritico = (minRequerido > 0 && !respeta && !conteo.tieneBolsa);

      let colorBolsa;
      let estiloBolsaExtra = '';

      if (casoCritico) {
        colorBolsa = '#FFCDD2'; // rojo claro
        estiloBolsaExtra = 'border: 2px solid #D32F2F; color:#D32F2F;';
      } else {
        colorBolsa = colorFila; // mismo fondo que la fila
      }

      html += `
        <tr style="background-color: ${colorFila};">
          <td style="font-weight: bold; padding: 12px;">${conteo.semana}</td>
          <td style="padding: 12px;">${conteo.dia}</td>
          <td style="font-weight: bold; padding: 12px; text-transform: uppercase;">${conteo.turno}</td>
          <td style="text-align: center; font-size: 24px; font-weight: bold; padding: 12px; color: #1976D2;">${conteo.total}</td>
          <td style="text-align: center; padding: 12px; font-weight: bold;">${textoMinimos}</td>
          <td style="text-align: center; padding: 12px; background-color: ${colorBolsa}; font-weight: bold; ${estiloBolsaExtra}">${bolsaTexto}</td>
        </tr>`;
    });

    html += `</table><hr style="margin: 25px 0;">`;

    html += `
      <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 5px solid #2196F3; margin-top: 15px;">
        <strong>📈 ESTADÍSTICAS:</strong>
        <div style="margin-top: 10px;">
          ${Object.entries(stats)
            .sort()
            .map(([tipo, count]) => 
              `<span style="display:inline-block;margin:6px 15px 0 0;padding:6px 12px;background:${getColorTipo(tipo)};color:#333;border-radius:6px;font-weight:bold;">
                 ${tipo}: ${count}
               </span>`
            )
            .join('')}
        </div>
      </div>`;

    divResultado.innerHTML = html;
  }

  function getColorTipo(tipo) {
    const colores = {
      Normal: '#E0E0E0',
      Ausencia: '#FF4444',
      Bolsa: '#FFC000',
      DisfruteHH: '#FFFF00',
      '6Dias': '#00B050'
    };
    return colores[tipo] || '#FFFFFF';
  }

  // -------- Estadísticas avanzadas en nueva ventana --------

  function construirHtmlTablasEstadisticas() {
    if (!estadisticasDetalladas) return '<p>No hay estadísticas detalladas disponibles.</p>';

    const {
      casosSinMinimos,
      casosSinMinimosSinBolsa,
      ausenciasPorSemanaDia,
      ausenciasPorSemanaYTipo,
      totalAusencias,
      statsTipos,
      semanaMin,
      semanaMax
    } = estadisticasDetalladas;

    const textoRango =
      semanaMin != null && semanaMax != null
        ? `Estadísticas desde la semana <strong>${semanaMin}</strong> hasta la semana <strong>${semanaMax}</strong>.`
        : 'Estadísticas sin rango de semanas definido.';

    const panelSeleccion = `
      <div style="margin:10px 0 20px 0; padding:8px 12px; background:#f1f5fb; border-radius:6px;">
        <strong>Mostrar secciones:</strong>
        <label style="margin-left:10px;"><input type="checkbox" id="sel-resumen-tipos" checked> Resumen por tipo</label>
        <label style="margin-left:10px;"><input type="checkbox" id="sel-sin-minimos" checked> Baja mínimos</label>
        <label style="margin-left:10px;"><input type="checkbox" id="sel-sin-minimos-sin-bolsa" checked> Baja mínimos y sin activar bolsa</label>
        <label style="margin-left:10px;"><input type="checkbox" id="sel-ausencias" checked> Ausencias</label>
        <label style="margin-left:10px;"><input type="checkbox" id="sel-graficas" checked> Gráficas</label>
      </div>
    `;

    const resumenTipos = `
      <div id="bloque-resumen-tipos">
        <h3>Resumen por tipo</h3>
        <div style="margin-top:8px;">
          ${Object.entries(statsTipos)
            .sort()
            .map(([tipo, count]) => 
              `<span style="display:inline-block;margin:4px 10px 0 0;
                            padding:4px 10px;background:${getColorTipo(tipo)};
                            color:#333;border-radius:6px;font-weight:bold;font-size:12px;">
                 ${tipo}: ${count}
               </span>`
            ).join('')}
        </div>
      </div>`;

    // Tabla 1: sin mínimos
    let tablaSinMinimos = `
      <div id="bloque-sin-minimos">
        <h3>Semanas y días BAJANDO mínimos</h3>
        <p>Total: <strong>${casosSinMinimos.length}</strong></p>
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 13px;">
          <tr style="background:#f0f0f0;">
            <th>Semana</th><th>Día</th><th>Turno</th><th>Nº Personas</th><th>Mínimos</th><th>Bolsa</th>
          </tr>`;
    casosSinMinimos.forEach(c => {
      const estiloCritico = !c.tieneBolsa
        ? 'background:#FFCDD2;border:2px solid #D32F2F;'
        : '';

      tablaSinMinimos += `
          <tr>
            <td>${c.semana}</td>
            <td>${c.dia}</td>
            <td>${c.turno}</td>
            <td style="text-align:center;">${c.total}</td>
            <td style="text-align:center;">${c.minimos}</td>
            <td style="text-align:center; font-weight:bold; ${estiloCritico}">
              ${c.tieneBolsa ? 'SI' : 'NO'}
            </td>
          </tr>`;
    });

    tablaSinMinimos += `</table></div>`;

    // Tabla 2: sin mínimos y sin bolsa
    let tablaSinMinimosSinBolsa = `
      <div id="bloque-sin-minimos-sin-bolsa">
        <h3>Casos BAJANDO mínimos y SIN ACTIVAR bolsa</h3>
        <p>Total: <strong>${casosSinMinimosSinBolsa.length}</strong></p>
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 13px;">
          <tr style="background:#f0f0f0;">
            <th>Semana</th><th>Día</th><th>Turno</th><th>Nº Personas</th><th>Mínimos</th>
          </tr>`;
    casosSinMinimosSinBolsa.forEach(c => {
      tablaSinMinimosSinBolsa += `
          <tr style="background:#FFEBEE;">
            <td>${c.semana}</td>
            <td>${c.dia}</td>
            <td>${c.turno}</td>
            <td style="text-align:center;">${c.total}</td>
            <td style="text-align:center; font-weight:bold; background:#FFCDD2; border:2px solid #D32F2F;">
              ${c.minimos}
            </td>
          </tr>`;
    });

    tablaSinMinimosSinBolsa += `</table></div>`;

    // Tabla 3: ausencias por semana/día
    const clavesAus = Object.keys(ausenciasPorSemanaDia).sort((a, b) => {
      const [semA] = a.split('|'); const [semB] = b.split('|');
      return Number(semA) - Number(semB);
    });

    let tablaAusencias = `
      <div id="bloque-ausencias">
        <h3>Ausencias por semana y día</h3>
        <p>Total de ausencias: <strong>${totalAusencias}</strong></p>
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 13px;">
          <tr style="background:#f0f0f0;">
            <th>Semana</th><th>Día</th><th>Total ausencias</th><th>Detalle</th>
          </tr>`;
    clavesAus.forEach(clave => {
      const info = ausenciasPorSemanaDia[clave];
      const detalle = info.detalles
        .map(d => `${d.turno}: ${d.persona}`)
        .join('<br>');
      tablaAusencias += `
          <tr>
            <td>${info.semana}</td>
            <td>${info.dia}</td>
            <td style="text-align:center;">${info.total}</td>
            <td>${detalle}</td>
          </tr>`;
    });
    tablaAusencias += `</table></div>`;

    return `
      <p style="margin-bottom:8px;">${textoRango}</p>
      ${panelSeleccion}
      ${resumenTipos}
      <hr style="margin:20px 0;">
      ${tablaSinMinimos}
      <hr style="margin:20px 0;">
      ${tablaSinMinimosSinBolsa}
      <hr style="margin:20px 0%;">
      ${tablaAusencias}
    `;
  }

  function abrirVentanaEstadisticas() {
    if (!registrosGlobal || registrosGlobal.length === 0) {
      alert('Primero procesa algún horario para ver estadísticas.');
      return;
    }
    if (!estadisticasDetalladas) {
      alert('No hay estadísticas detalladas calculadas todavía.');
      return;
    }

    const win = window.open('', '_blank', 'width=1100,height=800');
    if (!win) {
      alert('El navegador ha bloqueado la ventana emergente. Permite pop-ups para esta página.');
      return;
    }

    const htmlTablas = construirHtmlTablasEstadisticas();
    const {
      casosSinMinimos,
      casosSinMinimosSinBolsa,
      ausenciasPorSemanaDia,
      ausenciasPorSemanaYTipo,
      statsTipos
    } = estadisticasDetalladas;

    // Datos para gráficos
    const conteoPorSemanaSinMin = {};
    casosSinMinimos.forEach(c => {
      conteoPorSemanaSinMin[c.semana] = (conteoPorSemanaSinMin[c.semana] || 0) + 1;
    });
    const conteoPorSemanaSinMinSinBolsa = {};
    casosSinMinimosSinBolsa.forEach(c => {
      conteoPorSemanaSinMinSinBolsa[c.semana] = (conteoPorSemanaSinMinSinBolsa[c.semana] || 0) + 1;
    });
    const semanasGraf1 = Array.from(new Set(Object.keys(conteoPorSemanaSinMin).concat(Object.keys(conteoPorSemanaSinMinSinBolsa))))
      .sort((a, b) => Number(a) - Number(b));
    const datosSinMin = semanasGraf1.map(s => conteoPorSemanaSinMin[s] || 0);
    const datosSinMinSinBolsa = semanasGraf1.map(s => conteoPorSemanaSinMinSinBolsa[s] || 0);

    const conteoAusPorSemana = {};
    Object.values(ausenciasPorSemanaDia).forEach(info => {
      conteoAusPorSemana[info.semana] = (conteoAusPorSemana[info.semana] || 0) + info.total;
    });
    const semanasGraf2 = Object.keys(conteoAusPorSemana).sort((a, b) => Number(a) - Number(b));
    const datosAusPorSemana = semanasGraf2.map(s => conteoAusPorSemana[s] || 0);

    const labelsTipos = Object.keys(statsTipos).sort();
    const datosTipos = labelsTipos.map(t => statsTipos[t]);

    const subtiposAus = ['AP', 'BM', 'HH', 'VAC', 'OTRAS'];
    const semanasAusTiposSet = new Set();
    Object.values(ausenciasPorSemanaYTipo).forEach(info => {
      semanasAusTiposSet.add(info.semana);
    });
    const semanasGrafAusTipos = Array.from(semanasAusTiposSet).sort((a, b) => Number(a) - Number(b));
    const datosAusPorSemanaYTipo = {};
    subtiposAus.forEach(st => {
      datosAusPorSemanaYTipo[st] = semanasGrafAusTipos.map(sem => {
        const clave = `${sem}|${st}`;
        const reg = ausenciasPorSemanaYTipo[clave];
        return reg ? reg.total : 0;
      });
    });

    win.document.write(`
      <!DOCTYPE html>
      <html lang="es">
      <head>
        <meta charset="UTF-8">
        <title>Estadísticas de horarios</title>
        <style>
          body { font-family: Arial, sans-serif; padding:20px; }
          h2, h3 { margin-top: 0; }
          .chart-container { width: 100%; max-width: 900px; margin: 25px auto; }
          .modo-toggle { margin: 0 0 10px 0; font-size: 13px; }
          .modo-toggle button { margin-left: 8px; padding: 3px 8px; font-size: 12px; }
        </style>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
           </head>
      <body>
        <div style="display:flex;justify-content:space-between;align-items:center;">
          <h2>📊 Estadísticas detalladas</h2>
          <button id="btn-imprimir" style="padding:6px 10px;font-size:13px;cursor:pointer;">
            🖨 Imprimir
          </button>
        </div>
        ${htmlTablas}

        <div id="bloque-graficas">
          <hr style="margin:30px 0;">

          <h2>📉 Gráficas</h2>

          <div class="chart-container">
            <h3>Casos bajando mínimos por semana</h3>
            <canvas id="chartMinimos"></canvas>
          </div>

          <div class="chart-container">
            <h3>Ausencias por semana</h3>
            <canvas id="chartAusencias"></canvas>
          </div>

          <div class="chart-container">
            <div class="modo-toggle">
              <strong>Tipos de ausencias por semana</strong>
              <button id="btn-modo-aus-tipos">Cambiar a barras agrupadas</button>
            </div>
            <div id="legend-aus-tipos" style="margin-bottom:8px;"></div>
            <canvas id="chartAusenciasTipos"></canvas>
          </div>

          <div class="chart-container">
            <h3>Distribución por tipo</h3>
            <canvas id="chartTipos"></canvas>
          </div>
        </div>

        <script>
          const semanasGraf1 = ${JSON.stringify(semanasGraf1)};
          const datosSinMin = ${JSON.stringify(datosSinMin)};
          const datosSinMinSinBolsa = ${JSON.stringify(datosSinMinSinBolsa)};

          const semanasGraf2 = ${JSON.stringify(semanasGraf2)};
          const datosAusPorSemana = ${JSON.stringify(datosAusPorSemana)};

          const labelsTipos = ${JSON.stringify(labelsTipos)};
          const datosTipos = ${JSON.stringify(datosTipos)};

          const subtiposAus = ${JSON.stringify(subtiposAus)};
          const semanasGrafAusTipos = ${JSON.stringify(semanasGrafAusTipos)};
          const datosAusPorSemanaYTipo = ${JSON.stringify(datosAusPorSemanaYTipo)};

          // Gráfico 1: sin mínimos
          const ctx1 = document.getElementById('chartMinimos').getContext('2d');
          new Chart(ctx1, {
            type: 'bar',
            data: {
              labels: semanasGraf1,
              datasets: [
                {
                  label: 'Baja mínimos',
                  data: datosSinMin,
                  backgroundColor: 'rgba(244, 67, 54, 0.6)'
                },
                {
                  label: 'Baja mínimos y sin activar bolsa',
                  data: datosSinMinSinBolsa,
                  backgroundColor: 'rgba(255, 152, 0, 0.6)'
                }
              ]
            },
            options: {
              responsive: true,
              scales: {
                x: { title: { display: true, text: 'Semana' } },
                y: { title: { display: true, text: 'Nº de casos' }, beginAtZero: true }
              }
            }
          });

          // Gráfico 2: ausencias por semana
          const ctx2 = document.getElementById('chartAusencias').getContext('2d');
          new Chart(ctx2, {
            type: 'bar',
            data: {
              labels: semanasGraf2,
              datasets: [
                {
                  label: 'Ausencias',
                  data: datosAusPorSemana,
                  backgroundColor: 'rgba(33, 150, 243, 0.6)'
                }
              ]
            },
            options: {
              responsive: true,
              scales: {
                x: { title: { display: true, text: 'Semana' } },
                y: { title: { display: true, text: 'Nº de ausencias' }, beginAtZero: true }
              }
            }
          });

          // Gráfico 3: pastel por tipo (Normal verde, 6Dias gris)
          const ctx3 = document.getElementById('chartTipos').getContext('2d');
          new Chart(ctx3, {
            type: 'pie',
            data: {
              labels: labelsTipos,
              datasets: [{
                data: datosTipos,
                backgroundColor: [
                  '#00B050', // Normal
                  '#FF4444', // Ausencia
                  '#FFC000', // Bolsa
                  '#FFFF00', // DisfruteHH
                  '#E0E0E0', // 6Dias
                  '#90CAF9',
                  '#CE93D8'
                ]
              }]
            },
            options: {
              responsive: true
            }
          });

          // Gráfico 4: tipos de ausencias por semana (apilado / agrupado)
          const ctx4 = document.getElementById('chartAusenciasTipos').getContext('2d');

          const coloresSubtipos = {
            AP:   '#1976D2', // azul
            BM:   '#D32F2F', // rojo
            HH:   '#FFB300', // amarillo/naranja
            VAC:  '#388E3C', // verde
            OTRAS:'#7B1FA2'  // morado
          };

          const datasetsAusTipos = subtiposAus.map(st => ({
            label: st,
            data: datosAusPorSemanaYTipo[st],
            backgroundColor: coloresSubtipos[st],
            hidden: false
          }));

          let modoApilado = true; // true = barras apiladas, false = agrupadas

          const chartAusTipos = new Chart(ctx4, {
            type: 'bar',
            data: {
              labels: semanasGrafAusTipos,
              datasets: datasetsAusTipos
            },
            options: {
              responsive: true,
              plugins: {
                legend: { display: false }
              },
              scales: {
                x: { stacked: true, title: { display: true, text: 'Semana' } },
                y: { stacked: true, title: { display: true, text: 'Nº ausencias' }, beginAtZero: true }
              }
            }
          });

          // Botón de conmutación apilado / agrupado
          const btnModoAusTipos = document.getElementById('btn-modo-aus-tipos');
          btnModoAusTipos.addEventListener('click', () => {
            modoApilado = !modoApilado;
            chartAusTipos.options.scales.x.stacked = modoApilado;
            chartAusTipos.options.scales.y.stacked = modoApilado;
            btnModoAusTipos.textContent = modoApilado
              ? 'Cambiar a barras agrupadas'
              : 'Cambiar a barras apiladas';
            chartAusTipos.update();
          });

          // Leyenda HTML para mostrar/ocultar subtipos
          const legendContainer = document.getElementById('legend-aus-tipos');
          subtiposAus.forEach((st, idx) => {
            const span = document.createElement('span');
            span.textContent = st;
            span.style.display = 'inline-block';
            span.style.marginRight = '10px';
            span.style.padding = '4px 8px';
            span.style.borderRadius = '4px';
            span.style.cursor = 'pointer';
            span.style.background = coloresSubtipos[st];
            span.style.color = '#fff';
            span.dataset.index = idx;

            span.onclick = () => {
              const i = Number(span.dataset.index);
              const ds = chartAusTipos.data.datasets[i];
              ds.hidden = !ds.hidden;
              span.style.opacity = ds.hidden ? '0.3' : '1';
              chartAusTipos.update();
            };

            legendContainer.appendChild(span);
          });

          // Toggles de secciones
          function conectarToggles() {
            const m = (id) => document.getElementById(id);

            const mapa = {
              'sel-resumen-tipos': 'bloque-resumen-tipos',
              'sel-sin-minimos': 'bloque-sin-minimos',
              'sel-sin-minimos-sin-bolsa': 'bloque-sin-minimos-sin-bolsa',
              'sel-ausencias': 'bloque-ausencias',
              'sel-graficas': 'bloque-graficas'
            };

            Object.entries(mapa).forEach(([chkId, bloqueId]) => {
              const chk = m(chkId);
              const bloque = m(bloqueId);
              if (!chk || !bloque) return;
              const actualizar = () => {
                bloque.style.display = chk.checked ? 'block' : 'none';
              };
              chk.addEventListener('change', actualizar);
              actualizar();
            });
          }
          // Botón imprimir: usa el estado actual de los checkboxes
          const btnImprimir = document.getElementById('btn-imprimir');
          btnImprimir.addEventListener('click', () => {
            // Ya que las secciones se ocultan/ muestran con conectarToggles,
            // basta con llamar a print: solo se imprimirá lo que esté visible.
            window.print();
          });
          window.addEventListener('load', conectarToggles);
        </script>
      </body>
      </html>
    `);
    win.document.close();
  }
});
