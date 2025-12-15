document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('form-upload');
  const inputExcel = document.getElementById('input-excel');
  const inputCuadrante = document.getElementById('input-cuadrante');
  const chkAusencias = document.getElementById('chk-ausencias');
  const chkBolsa = document.getElementById('chk-bolsa');
  const chkDisfruteHH = document.getElementById('chk-disfrutehh');
  const chk6Dias = document.getElementById('chk-6dias');
  const chkMostrarTeorico = document.getElementById('chk-mostrar-teorico');
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
  let teoricoMap = {}; // { "semana|dia|turno": count }

  // Configuración de mínimos en memoria
  let minimosConfig = {
    MANANA: { Lunes: 0, Martes: 0, Miércoles: 0, Jueves: 0, Viernes: 0, Sábado: 0 },
    TARDE: { Lunes: 0, Martes: 0, Miércoles: 0, Jueves: 0, Viernes: 0 },
    NOCHE: { Domingo: 0, Lunes: 0, Martes: 0, Miércoles: 0, Jueves: 0, Viernes: 0 }
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
    document.getElementById('min-manana-lunes').value = minimosConfig.MANANA['Lunes'] || 0;
    document.getElementById('min-manana-martes').value = minimosConfig.MANANA['Martes'] || 0;
    document.getElementById('min-manana-miercoles').value = minimosConfig.MANANA['Miércoles'] || 0;
    document.getElementById('min-manana-jueves').value = minimosConfig.MANANA['Jueves'] || 0;
    document.getElementById('min-manana-viernes').value = minimosConfig.MANANA['Viernes'] || 0;
    document.getElementById('min-manana-sabado').value = minimosConfig.MANANA['Sábado'] || 0;
    // Tarde
    document.getElementById('min-tarde-lunes').value = minimosConfig.TARDE['Lunes'] || 0;
    document.getElementById('min-tarde-martes').value = minimosConfig.TARDE['Martes'] || 0;
    document.getElementById('min-tarde-miercoles').value = minimosConfig.TARDE['Miércoles'] || 0;
    document.getElementById('min-tarde-jueves').value = minimosConfig.TARDE['Jueves'] || 0;
    document.getElementById('min-tarde-viernes').value = minimosConfig.TARDE['Viernes'] || 0;
    // Noche
    document.getElementById('min-noche-domingo').value = minimosConfig.NOCHE['Domingo'] || 0;
    document.getElementById('min-noche-lunes').value = minimosConfig.NOCHE['Lunes'] || 0;
    document.getElementById('min-noche-martes').value = minimosConfig.NOCHE['Martes'] || 0;
    document.getElementById('min-noche-miercoles').value = minimosConfig.NOCHE['Miércoles'] || 0;
    document.getElementById('min-noche-jueves').value = minimosConfig.NOCHE['Jueves'] || 0;
    document.getElementById('min-noche-viernes').value = minimosConfig.NOCHE['Viernes'] || 0;
  }

  function leerInputsMinimos() {
    minimosConfig.MANANA['Lunes'] = Number(document.getElementById('min-manana-lunes').value) || 0;
    minimosConfig.MANANA['Martes'] = Number(document.getElementById('min-manana-martes').value) || 0;
    minimosConfig.MANANA['Miércoles'] = Number(document.getElementById('min-manana-miercoles').value) || 0;
    minimosConfig.MANANA['Jueves'] = Number(document.getElementById('min-manana-jueves').value) || 0;
    minimosConfig.MANANA['Viernes'] = Number(document.getElementById('min-manana-viernes').value) || 0;
    minimosConfig.MANANA['Sábado'] = Number(document.getElementById('min-manana-sabado').value) || 0;

    minimosConfig.TARDE['Lunes'] = Number(document.getElementById('min-tarde-lunes').value) || 0;
    minimosConfig.TARDE['Martes'] = Number(document.getElementById('min-tarde-martes').value) || 0;
    minimosConfig.TARDE['Miércoles'] = Number(document.getElementById('min-tarde-miercoles').value) || 0;
    minimosConfig.TARDE['Jueves'] = Number(document.getElementById('min-tarde-jueves').value) || 0;
    minimosConfig.TARDE['Viernes'] = Number(document.getElementById('min-tarde-viernes').value) || 0;

    minimosConfig.NOCHE['Domingo'] = Number(document.getElementById('min-noche-domingo').value) || 0;
    minimosConfig.NOCHE['Lunes'] = Number(document.getElementById('min-noche-lunes').value) || 0;
    minimosConfig.NOCHE['Martes'] = Number(document.getElementById('min-noche-martes').value) || 0;
    minimosConfig.NOCHE['Miércoles'] = Number(document.getElementById('min-noche-miercoles').value) || 0;
    minimosConfig.NOCHE['Jueves'] = Number(document.getElementById('min-noche-jueves').value) || 0;
    minimosConfig.NOCHE['Viernes'] = Number(document.getElementById('min-noche-viernes').value) || 0;
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
      msgMinimos.textContent = '✓ Guardado';
      setTimeout(() => { msgMinimos.textContent = ''; }, 2000);
      renderInforme();
    });
  }

  console.log('DOM elements:', {
    form: !!form,
    inputExcel: !!inputExcel,
    inputCuadrante: !!inputCuadrante,
    chkAusencias: !!chkAusencias,
    chkBolsa: !!chkBolsa,
    chkDisfruteHH: !!chkDisfruteHH,
    chk6Dias: !!chk6Dias,
    chkMostrarTeorico: !!chkMostrarTeorico
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
  if (chkMostrarTeorico) {
    chkMostrarTeorico.addEventListener('change', renderInforme);
  }

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

  // ========== LECTURA DEL CUADRANTE ANUAL (CLIENTE) ==========
  if (inputCuadrante) {
    inputCuadrante.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;

      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        teoricoMap = procesarCuadranteAnual(sheet);
        console.log('Cuadrante anual cargado:', Object.keys(teoricoMap).length, 'entradas');
        renderInforme();
      } catch (err) {
        console.error('Error leyendo cuadrante anual:', err);
        alert('Error al procesar el cuadrante anual.');
      }
    });
  }

  function procesarCuadranteAnual(sheet) {
    const map = {};

    // Leer año (B3)
    const year = sheet['B3'] ? sheet['B3'].v : null;

    // Leer nombres (C3:AQ3)
    const nombres = [];
    for (let col = 2; col <= 42; col++) { // C=2, AQ=42
      const cellAddr = XLSX.utils.encode_cell({ r: 2, c: col });
      const val = sheet[cellAddr] ? String(sheet[cellAddr].v).trim() : '';
      nombres.push(val);
    }

    // Leer filas A4:A55 (semanas) y turnos C4:AQ55
    for (let row = 3; row <= 54; row++) { // fila 4 (índice 3) a fila 55 (índice 54)
      const semanaCell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
      const semana = semanaCell ? Number(semanaCell.v) : null;
      if (semana === null || isNaN(semana)) continue;

      // Leer turnos para cada persona (C=2..AQ=42)
      for (let col = 2; col <= 42; col++) {
        const nombre = nombres[col - 2];
        if (!nombre) continue;

        const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
        const turnoVal = sheet[cellAddr] ? String(sheet[cellAddr].v).trim().toUpperCase() : '';
        if (!turnoVal || turnoVal === 'V') continue; // ignorar vacaciones

        // Interpretar patrón de turno
        const diasTurnos = interpretarPatronTurno(turnoVal);
        diasTurnos.forEach(({ dia, turno }) => {
          const clave = `${semana}|${dia}|${turno}`;
          map[clave] = (map[clave] || 0) + 1;
        });
      }
    }

    return map;
  }

  function interpretarPatronTurno(patron) {
    const result = [];
    const p = patron.toUpperCase().trim();

    // M-LU → Martes a Sábado, MAÑANA
    if (p === 'M-LU') {
      ['Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'].forEach(d => {
        result.push({ dia: d, turno: 'MAÑANA' });
      });
    }
    // M-MI → Lunes, Martes, Jueves, Viernes, Sábado, MAÑANA
    else if (p === 'M-MI') {
      ['Lunes', 'Martes', 'Jueves', 'Viernes', 'Sábado'].forEach(d => {
        result.push({ dia: d, turno: 'MAÑANA' });
      });
    }
    // M-SA → Lunes a Viernes, MAÑANA
    else if (p === 'M-SA') {
      ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'].forEach(d => {
        result.push({ dia: d, turno: 'MAÑANA' });
      });
    }
    // T-SA → Lunes a Viernes, TARDE
    else if (p === 'T-SA') {
      ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'].forEach(d => {
        result.push({ dia: d, turno: 'TARDE' });
      });
    }
    // N-VI o N-VI-A → Domingo a Jueves, NOCHE
    else if (p === 'N-VI' || p === 'N-VI-A') {
      ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves'].forEach(d => {
        result.push({ dia: d, turno: 'NOCHE' });
      });
    }
    // N-DO → Lunes a Viernes, NOCHE
    else if (p === 'N-DO') {
      ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'].forEach(d => {
        result.push({ dia: d, turno: 'NOCHE' });
      });
    }

    return result;
  }

  // ========== SUBMIT FORM (PROCESAMIENTO BACKEND) ==========
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    console.log('FORM SUBMIT - procesando...');

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
      const resp = await fetch('/upload', { method: 'POST', body: formData });
      const data = await resp.json();
      console.log('RESPUESTA:', data);

      if (!data.ok) {
        divResultado.innerHTML = `<p style="color:red;">Error: ${data.message}</p>`;
        return;
      }

      registrosGlobal = data.registros;
      console.log('REGISTROS:', registrosGlobal.length);
      autocompletarRangoSemanas();
      renderInforme();
    } catch (err) {
      console.error('ERROR:', err);
      divResultado.innerHTML = '<p style="color:red;">Error al enviar los ficheros.</p>';
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
    if (texto.includes('AP') && !ausAP.checked) return false;
    if (texto.includes('BM') && !ausBM.checked) return false;
    if ((texto.includes('HHEE') || texto.includes('HHDD')) && !ausHH.checked) return false;
    if ((texto.includes('VC') || texto.includes('VP') || texto.includes('VX') || texto.includes('V')) && !ausVAC.checked) return false;
    if ((texto.includes('PAT') || texto.includes('LIC') || texto.includes('HF') || texto.includes('HP')) && !ausOTRAS.checked) return false;
    return true;
  }

  function renderInforme() {
    if (!registrosGlobal || registrosGlobal.length === 0) {
      divResultado.innerHTML = '<p>No hay datos para mostrar.</p>';
      return;
    }

    const mostrarAusencias = chkAusencias.checked;
    const mostrarBolsa = chkBolsa.checked;
    const mostrarDisfruteHH = chkDisfruteHH.checked;
    const mostrar6Dias = chk6Dias.checked;
    const mostrarTeorico = chkMostrarTeorico ? chkMostrarTeorico.checked : false;

    const semDesde = inputSemanaDesde && inputSemanaDesde.value ? Number(inputSemanaDesde.value) : null;
    const semHasta = inputSemanaHasta && inputSemanaHasta.value ? Number(inputSemanaHasta.value) : null;

    console.log('Filtros:', { mostrarAusencias, mostrarBolsa, mostrarDisfruteHH, mostrar6Dias, mostrarTeorico, semDesde, semHasta });

    const registrosFiltrados = registrosGlobal.filter(reg => {
      // Filtro por semana desde/hasta
      if (semDesde != null && reg.semana < semDesde) return false;
      if (semHasta != null && reg.semana > semHasta) return false;

      if (reg.tipo === 'Ausencia') {
        if (!mostrarAusencias) return false;
        if (!pasaSubfiltroAusencia(reg)) return false;
      }
      if (reg.tipo === 'Bolsa' && !mostrarBolsa) return false;
      if (reg.tipo === 'DisfruteHH' && !mostrarDisfruteHH) return false;
      if (reg.tipo === '6Dias' && !mostrar6Dias) return false;

      return true;
    });

    // Estadísticas por tipo de TODOS
    const stats = {};
    registrosGlobal.forEach(r => {
      stats[r.tipo] = (stats[r.tipo] || 0) + 1;
    });

    // Agrupar por semana, día y turno
    const conteos = {};
    registrosFiltrados.forEach(reg => {
      const clave = `${reg.semana}|${reg.dia}|${reg.turno}`;
      if (!conteos[clave]) {
        conteos[clave] = { semana: reg.semana, dia: reg.dia, turno: reg.turno, total: 0, tieneBolsa: false };
      }
      conteos[clave].total++;
      if (reg.tipo === 'Bolsa' && mostrarBolsa) {
        conteos[clave].tieneBolsa = true;
      }
    });

    const ordenDias = { Domingo: 0, Lunes: 1, Martes: 2, Miércoles: 3, Jueves: 4, Viernes: 5, Sábado: 6 };
    const ordenTurnos = { MAÑANA: 0, TARDE: 1, NOCHE: 2 };

    const conteosArray = Object.values(conteos).sort((a, b) => {
      if (a.semana !== b.semana) return a.semana - b.semana;
      const diaA = a.dia.split(' ')[0];
      const diaB = b.dia.split(' ')[0];
      if (ordenDias[diaA] !== ordenDias[diaB]) return ordenDias[diaA] - ordenDias[diaB];
      return ordenTurnos[a.turno] - ordenTurnos[b.turno];
    });

    const coloresTurno = { MAÑANA: '#E3F2FD', TARDE: '#FFF8E1', NOCHE: '#F5F5F5' };

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
      if (turnoClave === 'MAÑANA') minRequerido = minimosConfig.MANANA[diaNombre] || 0;
      else if (turnoClave === 'TARDE') minRequerido = minimosConfig.TARDE[diaNombre] || 0;
      else if (turnoClave === 'NOCHE') minRequerido = minimosConfig.NOCHE[diaNombre] || 0;

      const respeta = conteo.total >= minRequerido;

      // NUEVO: calcular teórico para las tablas detalladas
      const claveTeorico = `${conteo.semana}|${diaNombre}|${turnoClave}`;
      const teorico = teoricoMap ? (teoricoMap[claveTeorico] || 0) : 0;

      if (minRequerido > 0 && !respeta) {
        const item = {
          semana: conteo.semana,
          dia: conteo.dia,
          turno: conteo.turno,
          teorico,          // NUEVO
          total: conteo.total,
          minimos: minRequerido,
          tieneBolsa: conteo.tieneBolsa
        };
        casosSinMinimos.push(item);
        if (!conteo.tieneBolsa) casosSinMinimosSinBolsa.push(item);
      }
    });

    // Recorremos registrosGlobal para ausencias (tabla + subtipos)
    registrosGlobal.forEach(reg => {
      if (reg.tipo !== 'Ausencia') return;
      totalAusencias++;

      // Tabla por semana+día
      const clave = `${reg.semana}|${reg.dia}`;
      if (!ausenciasPorSemanaDia[clave]) {
        ausenciasPorSemanaDia[clave] = { semana: reg.semana, dia: reg.dia, total: 0, detalles: [] };
      }
      ausenciasPorSemanaDia[clave].total++;
      ausenciasPorSemanaDia[clave].detalles.push({ turno: reg.turno, persona: reg.persona });

      // Subtipo para gráfico por semana
      const texto = (reg.persona || '').toUpperCase();
      let subtipo = 'OTRAS';
      if (texto.includes('AP')) subtipo = 'AP';
      else if (texto.includes('BM')) subtipo = 'BM';
      else if (texto.includes('HHEE') || texto.includes('HHDD')) subtipo = 'HH';
      else if (texto.includes('VC') || texto.includes('VP') || texto.includes('VX') || texto.includes('V')) subtipo = 'VAC';
      else if (texto.includes('PAT') || texto.includes('LIC') || texto.includes('HF') || texto.includes('HP')) subtipo = 'OTRAS';

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

    // NUEVO: guardar conteosArray para que la ventana de estadísticas respete filtros
    estadisticasDetalladas = {
      casosSinMinimos,
      casosSinMinimosSinBolsa,
      ausenciasPorSemanaDia,
      ausenciasPorSemanaYTipo,
      totalAusencias,
      statsTipos: stats,
      semanaMin,
      semanaMax,
      conteosArray // NUEVO
    };

    // ========== CONSTRUIR HTML DE LA TABLA ==========
    let html = `<h2>📊 Informe de Horarios</h2>`;
    html += `<p><strong>Total mostrados:</strong> ${registrosFiltrados.length} | <strong>Total registros:</strong> ${registrosGlobal.length}</p>`;
    html += `<table border="1" cellpadding="10" cellspacing="0" style="border-collapse:collapse; width:100%; font-size:15px;">`;

    // CABECERA
    html += `<tr style="background:linear-gradient(90deg, #4CAF50, #45a049); color:white;">`;
    html += `<th style="padding:15px;">Semana</th>`;
    html += `<th style="padding:15px;">Día</th>`;
    html += `<th style="padding:15px;">Turno</th>`;
    html += `<th style="padding:15px; text-align:center;">Nº Personas</th>`;
    if (mostrarTeorico) {
      html += `<th style="padding:15px; text-align:center;">Nº Teórico</th>`;
    }
    html += `<th style="padding:15px; text-align:center;">Respeta mínimos</th>`;
    html += `<th style="padding:15px; text-align:center;">Bolsa Activada</th>`;
    html += `</tr>`;

    // FILAS
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
      const textoMinimos = minRequerido > 0 ? (respeta ? 'SÍ' : 'NO') : '-';
      const bolsaTexto = conteo.tieneBolsa ? 'SÍ' : 'NO';

      const casoCritico = minRequerido > 0 && !respeta && !conteo.tieneBolsa;
      let colorBolsa, estiloBolsaExtra = '';
      if (casoCritico) {
        colorBolsa = '#FFCDD2'; // rojo claro
        estiloBolsaExtra = 'border:2px solid #D32F2F; color:#D32F2F;';
      } else {
        colorBolsa = colorFila;
      }

      // Consultar teórico
      let teorico = '-';
      if (mostrarTeorico) {
        const claveTeorico = `${conteo.semana}|${diaNombre}|${turnoClave}`;
        teorico = teoricoMap[claveTeorico] || 0;
      }

      html += `<tr style="background-color:${colorFila};">`;
      html += `<td style="font-weight:bold; padding:12px;">${conteo.semana}</td>`;
      html += `<td style="padding:12px;">${conteo.dia}</td>`;
      html += `<td style="font-weight:bold; padding:12px; text-transform:uppercase;">${conteo.turno}</td>`;
      html += `<td style="text-align:center; font-size:24px; font-weight:bold; padding:12px; color:#1976D2;">${conteo.total}</td>`;
      if (mostrarTeorico) {
        html += `<td style="text-align:center; font-size:20px; font-weight:bold; padding:12px; color:#558B2F;">${teorico}</td>`;
      }
      html += `<td style="text-align:center; padding:12px; font-weight:bold;">${textoMinimos}</td>`;
      html += `<td style="text-align:center; padding:12px; background-color:${colorBolsa}; font-weight:bold; ${estiloBolsaExtra}">${bolsaTexto}</td>`;
      html += `</tr>`;
    });

    html += `</table><hr style="margin:25px 0;">`;

    // Estadísticas resumen
    html += `<div style="background:#f8f9fa; padding:15px; border-radius:8px; border-left:5px solid #2196F3; margin-top:15px;">`;
    html += `<strong>📈 ESTADÍSTICAS</strong>`;
    html += `<div style="margin-top:10px;">`;
    html += Object.entries(stats)
      .sort()
      .map(([tipo, count]) => {
        return `<span style="display:inline-block;margin:6px 15px 0 0;padding:6px 12px;background:${getColorTipo(tipo)};color:#333;border-radius:6px;font-weight:bold;">${tipo}: ${count}</span>`;
      })
      .join('');
    html += `</div></div>`;

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

  // ========== Estadísticas avanzadas en nueva ventana ==========
  function construirHtmlTablasEstadisticas() {
    if (!estadisticasDetalladas) return '<p>No hay estadísticas detalladas disponibles.</p>';

    const { casosSinMinimos, casosSinMinimosSinBolsa, ausenciasPorSemanaDia, ausenciasPorSemanaYTipo, totalAusencias, statsTipos, semanaMin, semanaMax } = estadisticasDetalladas;

    const textoRango = semanaMin != null && semanaMax != null
      ? `Estadísticas desde la semana <strong>${semanaMin}</strong> hasta la semana <strong>${semanaMax}</strong>.`
      : 'Estadísticas sin rango de semanas definido.';

    const panelSeleccion = `
      <div style="margin:10px 0 20px 0; padding:8px 12px; background:#f1f5fb; border-radius:6px;">
        <strong>Mostrar secciones</strong>
        <label style="margin-left:10px"><input type="checkbox" id="sel-resumen-tipos" checked> Resumen por tipo</label>
        <label style="margin-left:10px"><input type="checkbox" id="sel-sin-minimos" checked> Baja mínimos</label>
        <label style="margin-left:10px"><input type="checkbox" id="sel-sin-minimos-sin-bolsa" checked> Baja mínimos y sin activar bolsa</label>
        <label style="margin-left:10px"><input type="checkbox" id="sel-ausencias" checked> Ausencias</label>
        <label style="margin-left:10px"><input type="checkbox" id="sel-graficas" checked> Gráficas</label>

        <!-- NUEVO -->
        <label style="margin-left:10px"><input type="checkbox" id="sel-mostrar-minimos" checked> Mostrar columna mínimos</label>
      </div>
    `;

    const resumenTipos = `
      <div id="bloque-resumen-tipos">
        <h3>Resumen por tipo</h3>
        <div style="margin-top:8px;">
          ${Object.entries(statsTipos)
            .sort()
            .map(([tipo, count]) => {
              return `<span style="display:inline-block;margin:4px 10px 0 0; padding:4px 10px;background:${getColorTipo(tipo)};color:#333;border-radius:6px;font-weight:bold;font-size:12px;">${tipo}: ${count}</span>`;
            })
            .join('')}
        </div>
      </div>
    `;

    // Tabla 1: sin mínimos
    let tablaSinMinimos = `
      <div id="bloque-sin-minimos">
        <h3>Semanas y días BAJANDO mínimos</h3>
        <p>Total: <strong>${casosSinMinimos.length}</strong></p>

        <table id="tbl-sin-minimos" border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; width:100%; font-size:13px;">
          <tr style="background:#f0f0f0">
            <th>Semana</th><th>Día</th><th>Turno</th><th>Teórico</th><th>Nº Personas</th>
            <th class="col-minimos">Mínimos</th>
            <th>Bolsa</th>
          </tr>
          ${casosSinMinimos.map(c => {
            const estiloCritico = !c.tieneBolsa ? 'background:#FFCDD2; border:2px solid #D32F2F;' : '';
            return `
              <tr>
                <td>${c.semana}</td>
                <td>${c.dia}</td>
                <td>${c.turno}</td>
                <td style="text-align:center">${c.teorico ?? 0}</td>
                <td style="text-align:center">${c.total}</td>
                <td class="col-minimos" style="text-align:center">${c.minimos}</td>
                <td style="text-align:center; font-weight:bold; ${estiloCritico}">${c.tieneBolsa ? 'SÍ' : 'NO'}</td>
              </tr>
            `;
          }).join('')}
        </table>
      </div>
    `;

    // Tabla 2: sin mínimos y sin bolsa
    let tablaSinMinimosSinBolsa = `
      <div id="bloque-sin-minimos-sin-bolsa">
        <h3>Casos BAJANDO mínimos y SIN ACTIVAR bolsa</h3>
        <p>Total: <strong>${casosSinMinimosSinBolsa.length}</strong></p>

        <table id="tbl-sin-minimos-sin-bolsa" border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; width:100%; font-size:13px;">
          <tr style="background:#f0f0f0">
            <th>Semana</th><th>Día</th><th>Turno</th><th>Teórico</th><th>Nº Personas</th>
            <th class="col-minimos">Mínimos</th>
          </tr>
          ${casosSinMinimosSinBolsa.map(c => `
            <tr style="background:#FFEBEE">
              <td>${c.semana}</td>
              <td>${c.dia}</td>
              <td>${c.turno}</td>
              <td style="text-align:center">${c.teorico ?? 0}</td>
              <td style="text-align:center">${c.total}</td>
              <td class="col-minimos" style="text-align:center; font-weight:bold; background:#FFCDD2; border:2px solid #D32F2F">${c.minimos}</td>
            </tr>
          `).join('')}
        </table>
      </div>
    `;

    // Tabla 3: ausencias por semana+día
    const clavesAus = Object.keys(ausenciasPorSemanaDia).sort((a, b) => {
      const semA = a.split('|');
      const semB = b.split('|');
      return Number(semA) - Number(semB);
    });

    let tablaAusencias = `
      <div id="bloque-ausencias">
        <h3>Ausencias por semana y día</h3>
        <p>Total de ausencias: <strong>${totalAusencias}</strong></p>
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; width:100%; font-size:13px;">
          <tr style="background:#f0f0f0;">
            <th>Semana</th><th>Día</th><th>Total ausencias</th><th>Detalle</th>
          </tr>
    `;
    clavesAus.forEach(clave => {
      const info = ausenciasPorSemanaDia[clave];
      const detalle = info.detalles.map(d => `${d.turno}: ${d.persona}`).join('<br>');
      tablaAusencias += `
        <tr>
          <td>${info.semana}</td>
          <td>${info.dia}</td>
          <td style="text-align:center;">${info.total}</td>
          <td>${detalle}</td>
        </tr>
      `;
    });
    tablaAusencias += `</table></div>`;

    return `<p style="margin-bottom:8px;">${textoRango}</p>${panelSeleccion} ${resumenTipos}<hr style="margin:20px 0;"> ${tablaSinMinimos}<hr style="margin:20px 0;"> ${tablaSinMinimosSinBolsa}<hr style="margin:20px 0;"> ${tablaAusencias}`;
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

  const win = window.open('', '_blank', 'width=1200,height=850');
  if (!win) {
    alert('El navegador ha bloqueado la ventana emergente. Permite pop-ups para esta página.');
    return;
  }

  const htmlTablas = construirHtmlTablasEstadisticas();

  // IMPORTANTE: para Real vs Teórico necesitamos conteosArray y teoricoMap
  const { casosSinMinimos, casosSinMinimosSinBolsa, ausenciasPorSemanaDia, ausenciasPorSemanaYTipo, statsTipos, conteosArray } = estadisticasDetalladas;

  // ====== Helpers ======
  const normDia = (diaTexto) => (diaTexto || '').split(' ')[0];
  const keyTeo = (semana, dia, turno) => `${semana}|${dia}|${turno}`;

  // ====== Datos existentes (tus gráficas actuales) ======
  const conteoPorSemanaSinMin = {};
  casosSinMinimos.forEach(c => { conteoPorSemanaSinMin[c.semana] = (conteoPorSemanaSinMin[c.semana] || 0) + 1; });

  const conteoPorSemanaSinMinSinBolsa = {};
  casosSinMinimosSinBolsa.forEach(c => { conteoPorSemanaSinMinSinBolsa[c.semana] = (conteoPorSemanaSinMinSinBolsa[c.semana] || 0) + 1; });

  const semanasGraf1 = Array.from(new Set([
    ...Object.keys(conteoPorSemanaSinMin),
    ...Object.keys(conteoPorSemanaSinMinSinBolsa)
  ])).sort((a, b) => Number(a) - Number(b));
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
  Object.values(ausenciasPorSemanaYTipo).forEach(info => semanasAusTiposSet.add(info.semana));
  const semanasGrafAusTipos = Array.from(semanasAusTiposSet).sort((a, b) => Number(a) - Number(b));

  const datosAusPorSemanaYTipo = {};
  subtiposAus.forEach(st => {
    datosAusPorSemanaYTipo[st] = semanasGrafAusTipos.map(sem => {
      const clave = `${sem}|${st}`;
      const reg = ausenciasPorSemanaYTipo[clave];
      return reg ? reg.total : 0;
    });
  });

  // ====== NUEVO: Real vs Teórico (3 modos) ======
  const turnosOrden = ['MAÑANA', 'TARDE', 'NOCHE'];
  const ordenDias = { Domingo: 0, Lunes: 1, Martes: 2, Miércoles: 3, Jueves: 4, Viernes: 5, Sábado: 6 };

  // semanas disponibles según conteosArray (respeta filtros de la pantalla principal)
  const semanasRT = Array.from(new Set((conteosArray || []).map(c => c.semana)))
    .filter(s => s != null)
    .sort((a, b) => a - b);

  // Precalcular agregados
  const realPorSemana = {};
  const teoPorSemana = {};
  const realPorSemanaTurno = {};     // key: "semana|TURNO"
  const teoPorSemanaTurno = {};
  const realPorSemanaDia = {};       // key: "semana|DIA"
  const teoPorSemanaDia = {};
  const realPorSemanaDiaTurno = {};  // key: "semana|DIA|TURNO"
  const teoPorSemanaDiaTurno = {};

  (conteosArray || []).forEach(c => {
    const sem = c.semana;
    const dia = normDia(c.dia);
    const turno = (c.turno || '').toUpperCase();
    const real = c.total || 0;
    const teo = teoricoMap ? (teoricoMap[keyTeo(sem, dia, turno)] || 0) : 0;

    realPorSemana[sem] = (realPorSemana[sem] || 0) + real;
    teoPorSemana[sem] = (teoPorSemana[sem] || 0) + teo;

    const kST = `${sem}|${turno}`;
    realPorSemanaTurno[kST] = (realPorSemanaTurno[kST] || 0) + real;
    teoPorSemanaTurno[kST] = (teoPorSemanaTurno[kST] || 0) + teo;

    const kSD = `${sem}|${dia}`;
    realPorSemanaDia[kSD] = (realPorSemanaDia[kSD] || 0) + real;
    teoPorSemanaDia[kSD] = (teoPorSemanaDia[kSD] || 0) + teo;

    const kSDT = `${sem}|${dia}|${turno}`;
    realPorSemanaDiaTurno[kSDT] = (realPorSemanaDiaTurno[kSDT] || 0) + real;
    teoPorSemanaDiaTurno[kSDT] = (teoPorSemanaDiaTurno[kSDT] || 0) + teo;
  });

  // ====== HTML (sin backticks) ======
  const head =
    '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<title>Estadísticas de horarios</title>' +
    '<style>' +
      'body{font-family:Arial,sans-serif;padding:20px}' +
      'h2,h3{margin-top:0}' +
      '.chart-container{width:100%;max-width:980px;margin:25px auto}' +
      '.modo-toggle{margin:0 0 10px 0;font-size:13px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}' +
      '.modo-toggle button{padding:3px 8px;font-size:12px;cursor:pointer}' +
      '.btn-ampliar{background:#eee;border:1px solid #ccc;color:#333;border-radius:6px}' +
      '.filtro-global{margin:10px 0 15px 0;padding:10px 12px;background:#f7f9fc;border:1px solid #e3e8f5;border-radius:8px}' +
      '.filtro-global input{width:80px}' +
    '</style></head><body>';

  const body =
    '<div style="display:flex;justify-content:space-between;align-items:center;">' +
      '<h2>📊 Estadísticas detalladas</h2>' +
      '<button id="btn-imprimir" style="padding:6px 10px;font-size:13px;cursor:pointer;">🖨 Imprimir</button>' +
    '</div>' +

    // Filtro global por semana (afecta a TODAS las gráficas de la ventana)
    '<div class="filtro-global">' +
      '<strong>Filtro global de semanas (estadísticas)</strong> ' +
      '<span style="margin-left:10px">Desde: <input id="filtro-sem-desde" type="number"></span>' +
      '<span style="margin-left:10px">Hasta: <input id="filtro-sem-hasta" type="number"></span>' +
      '<button id="btn-aplicar-filtro" style="margin-left:10px">Aplicar</button>' +
      '<button id="btn-reset-filtro" style="margin-left:6px">Reset</button>' +
    '</div>' +

    htmlTablas +

    '<div id="bloque-graficas">' +
      '<hr style="margin:30px 0;">' +
      '<h2>📈 Gráficas</h2>' +

      '<div class="chart-container">' +
        '<div class="modo-toggle"><strong>Casos bajando mínimos por semana</strong>' +
          '<button class="btn-ampliar" data-chart="chartMinimos">Ampliar</button>' +
        '</div>' +
        '<canvas id="chartMinimos"></canvas>' +
      '</div>' +

      '<div class="chart-container">' +
        '<div class="modo-toggle"><strong>Ausencias por semana</strong>' +
          '<button class="btn-ampliar" data-chart="chartAusencias">Ampliar</button>' +
        '</div>' +
        '<canvas id="chartAusencias"></canvas>' +
      '</div>' +

      '<div class="chart-container">' +
        '<div class="modo-toggle">' +
          '<strong>Tipos de ausencias por semana</strong>' +
          '<button id="btn-modo-aus-tipos">Cambiar a barras agrupadas</button>' +
          '<button class="btn-ampliar" data-chart="chartAusenciasTipos">Ampliar</button>' +
        '</div>' +
        '<div id="legend-aus-tipos" style="margin-bottom:8px;"></div>' +
        '<canvas id="chartAusenciasTipos"></canvas>' +
      '</div>' +

      // NUEVO bloque: Real vs Teórico con selector de modo
      '<div class="chart-container">' +
        '<div class="modo-toggle">' +
          '<strong>Teórico vs Real</strong>' +
          '<label style="margin-left:10px">Modo: ' +
            '<select id="sel-modo-rt">' +
              '<option value="SEMANA">Por semana</option>' +
              '<option value="SEMANA_TURNO">Por semana y turno</option>' +
              '<option value="SEMANA_DIA">Por semana y día</option>' +
			    '<option value="SEMANA_DIA_TURNO">Por semana, día y turno</option>' +

            '</select>' +
          '</label>' +
          '<button class="btn-ampliar" data-chart="chartTeoReal">Ampliar</button>' +
        '</div>' +
        '<canvas id="chartTeoReal"></canvas>' +
      '</div>' +

      '<div class="chart-container">' +
        '<div class="modo-toggle"><strong>Distribución por tipo</strong>' +
          '<button class="btn-ampliar" data-chart="chartTipos">Ampliar</button>' +
        '</div>' +
        '<canvas id="chartTipos"></canvas>' +
      '</div>' +
    '</div>';

  const dataObj = {
    // charts existentes
    semanasGraf1, datosSinMin, datosSinMinSinBolsa,
    semanasGraf2, datosAusPorSemana,
    labelsTipos, datosTipos,
    subtiposAus, semanasGrafAusTipos, datosAusPorSemanaYTipo,

    // Real vs Teórico
    semanasRT,
    turnosOrden,
    ordenDias,
    realPorSemana,
    teoPorSemana,
    realPorSemanaTurno,
    teoPorSemanaTurno,
    realPorSemanaDia,
    teoPorSemanaDia,
    realPorSemanaDiaTurno,
    teoPorSemanaDiaTurno
  };

  const script =
    '<script>window.__DATA__=' + JSON.stringify(dataObj).replace(/</g, '\\u003c') + ';</script>' +
    '<script src="https://cdn.jsdelivr.net/npm/chart.js"><\/script>' +
    '<script>' +
      'const D=window.__DATA__;' +
      'document.getElementById("btn-imprimir").addEventListener("click",()=>window.print());' +

      // ====== Toggles (checks superiores de secciones + columna mínimos) ======
      'function aplicarColMinimos(){const chk=document.getElementById("sel-mostrar-minimos");const show=!chk||chk.checked;document.querySelectorAll(".col-minimos").forEach(el=>{el.style.display=show?"":"none";});}' +
      'function conectarToggles(){' +
        'const m=(id)=>document.getElementById(id);' +
        'const mapa={' +
          '"sel-resumen-tipos":"bloque-resumen-tipos",' +
          '"sel-sin-minimos":"bloque-sin-minimos",' +
          '"sel-sin-minimos-sin-bolsa":"bloque-sin-minimos-sin-bolsa",' +
          '"sel-ausencias":"bloque-ausencias",' +
          '"sel-graficas":"bloque-graficas"' +
        '};' +
        'Object.entries(mapa).forEach(([chkId,bloqueId])=>{' +
          'const chk=m(chkId);const bloque=m(bloqueId);if(!chk||!bloque)return;' +
          'const actualizar=()=>{bloque.style.display=chk.checked?"block":"none";};' +
          'chk.addEventListener("change",actualizar);actualizar();' +
        '});' +
        'const chkMin=m("sel-mostrar-minimos"); if(chkMin) chkMin.addEventListener("change", aplicarColMinimos); aplicarColMinimos();' +
      '}' +

      // ====== Utilidades filtro global ======
      'let filtroDesde=null; let filtroHasta=null;' +
      'function enRangoSemana(s){' +
        'const n=Number(s); if(isNaN(n)) return false;' +
        'if(filtroDesde!=null && n<Number(filtroDesde)) return false;' +
        'if(filtroHasta!=null && n>Number(filtroHasta)) return false;' +
        'return true;' +
      '}' +

      // ====== Charts: crear/actualizar con un registro global ======
      'const __charts={};' +
      'function setChart(id, chart){__charts[id]=chart;}' +

      // ====== Ampliar: abrir una ventana con el config del chart ======
      'function openChartInNewWindow(chartId){' +
        'const ch=__charts[chartId]; if(!ch) return;' +
        'const w=window.open("","_blank","width=1200,height=800"); if(!w) return;' +
        'const cfg=JSON.parse(JSON.stringify(ch.config));' +
        'const cfgStr=JSON.stringify(cfg).replace(/</g,"\\\\u003c");' +
        'const html=' +
          '"<!doctype html><html lang=\\"es\\"><head><meta charset=\\"UTF-8\\">"+' +
          '"<title>"+chartId+"</title>"+' +
          '"<style>body{font-family:Arial,sans-serif;padding:20px}</style>"+' +
          '"<script src=\\"https://cdn.jsdelivr.net/npm/chart.js\\"><\\\\/script>"+' +
          '"</head><body>"+' +
          '"<h3 style=\\"margin-top:0\\">"+chartId+"</h3>"+' +
          '"<canvas id=\\"c\\"></canvas>"+' +
          '"<script>const cfg="+cfgStr+"; const ctx=document.getElementById(\\"c\\").getContext(\\"2d\\"); new Chart(ctx,cfg);<\\\\/script>"+' +
          '"</body></html>";' +
        'w.document.open(); w.document.write(html); w.document.close();' +
      '}' +
      'document.querySelectorAll(".btn-ampliar").forEach(btn=>{btn.addEventListener("click",()=>openChartInNewWindow(btn.dataset.chart));});' +

      // ====== Chart 1: mínimos ======
      'function buildChartMinimos(){' +
        'const labels=[]; const a=[]; const b=[];' +
        'D.semanasGraf1.forEach((s,i)=>{ if(enRangoSemana(s)){labels.push(s); a.push(D.datosSinMin[i]); b.push(D.datosSinMinSinBolsa[i]);} });' +
        'return {labels, a, b};' +
      '}' +
      'function renderChartMinimos(){' +
        'const x=buildChartMinimos();' +
        'if(!__charts.chartMinimos){' +
          'const ch=new Chart(document.getElementById("chartMinimos").getContext("2d"),{' +
            'type:"bar",data:{labels:x.labels,datasets:[' +
              '{label:"Baja mínimos",data:x.a,backgroundColor:"rgba(244, 67, 54, 0.6)"},' +
              '{label:"Baja mínimos y sin activar bolsa",data:x.b,backgroundColor:"rgba(255, 152, 0, 0.6)"}' +
            ']},options:{responsive:true,scales:{x:{title:{display:true,text:"Semana"}},y:{title:{display:true,text:"Nº de casos"},beginAtZero:true}}}' +
          '}); setChart("chartMinimos",ch);' +
        '} else {' +
          '__charts.chartMinimos.data.labels=x.labels;' +
          '__charts.chartMinimos.data.datasets[0].data=x.a;' +
          '__charts.chartMinimos.data.datasets[1].data=x.b;' +
          '__charts.chartMinimos.update();' +
        '}' +
      '}' +

      // ====== Chart 2: ausencias ======
      'function buildChartAusencias(){' +
        'const labels=[]; const a=[];' +
        'D.semanasGraf2.forEach((s,i)=>{ if(enRangoSemana(s)){labels.push(s); a.push(D.datosAusPorSemana[i]);} });' +
        'return {labels,a};' +
      '}' +
      'function renderChartAusencias(){' +
        'const x=buildChartAusencias();' +
        'if(!__charts.chartAusencias){' +
          'const ch=new Chart(document.getElementById("chartAusencias").getContext("2d"),{' +
            'type:"bar",data:{labels:x.labels,datasets:[' +
              '{label:"Ausencias",data:x.a,backgroundColor:"rgba(33, 150, 243, 0.6)"}' +
            ']},options:{responsive:true,scales:{x:{title:{display:true,text:"Semana"}},y:{title:{display:true,text:"Nº de ausencias"},beginAtZero:true}}}' +
          '}); setChart("chartAusencias",ch);' +
        '} else {' +
          '__charts.chartAusencias.data.labels=x.labels;' +
          '__charts.chartAusencias.data.datasets[0].data=x.a;' +
          '__charts.chartAusencias.update();' +
        '}' +
      '}' +

      // ====== Chart 3: tipos por semana ======
      'let modoApilado=true;' +
      'const coloresSubtipos={AP:"#1976D2",BM:"#D32F2F",HH:"#FFB300",VAC:"#388E3C",OTRAS:"#7B1FA2"};' +
      'function buildChartAusTipos(){' +
        'const labels=[]; D.semanasGrafAusTipos.forEach(s=>{ if(enRangoSemana(s)) labels.push(s); });' +
        'const idxMap={}; labels.forEach((s,i)=>{idxMap[s]=i;});' +
        'const datasets=D.subtiposAus.map(st=>({label:st,data:labels.map(()=>0),backgroundColor:coloresSubtipos[st],hidden:false}));' +
        'D.semanasGrafAusTipos.forEach((s,origI)=>{' +
          'if(!enRangoSemana(s)) return;' +
          'const i=idxMap[s];' +
          'D.subtiposAus.forEach((st,dsI)=>{ datasets[dsI].data[i]=D.datosAusPorSemanaYTipo[st][origI]||0; });' +
        '});' +
        'return {labels,datasets};' +
      '}' +
      'function renderChartAusTipos(){' +
        'const x=buildChartAusTipos();' +
        'if(!__charts.chartAusenciasTipos){' +
          'const ch=new Chart(document.getElementById("chartAusenciasTipos").getContext("2d"),{' +
            'type:"bar",data:{labels:x.labels,datasets:x.datasets},' +
            'options:{responsive:true,plugins:{legend:{display:false}},scales:{x:{stacked:true,title:{display:true,text:"Semana"}},y:{stacked:true,title:{display:true,text:"Nº ausencias"},beginAtZero:true}}}' +
          '}); setChart("chartAusenciasTipos",ch);' +
        '} else {' +
          '__charts.chartAusenciasTipos.data.labels=x.labels;' +
          '__charts.chartAusenciasTipos.data.datasets=x.datasets;' +
          '__charts.chartAusenciasTipos.options.scales.x.stacked=modoApilado;' +
          '__charts.chartAusenciasTipos.options.scales.y.stacked=modoApilado;' +
          '__charts.chartAusenciasTipos.update();' +
        '}' +
        // leyenda
        'const legend=document.getElementById("legend-aus-tipos"); legend.innerHTML="";' +
        'D.subtiposAus.forEach((st,idx)=>{' +
          'const span=document.createElement("span");' +
          'span.textContent=st;' +
          'span.style.cssText="display:inline-block;margin-right:10px;padding:4px 8px;border-radius:4px;cursor:pointer;color:#fff;background:"+coloresSubtipos[st]+";";' +
          'span.dataset.index=String(idx);' +
          'span.onclick=()=>{' +
            'const ds=__charts.chartAusenciasTipos.data.datasets[idx]; ds.hidden=!ds.hidden; span.style.opacity=ds.hidden?"0.3":"1"; __charts.chartAusenciasTipos.update();' +
          '};' +
          'legend.appendChild(span);' +
        '});' +
      '}' +
      'document.getElementById("btn-modo-aus-tipos").addEventListener("click",()=>{' +
        'modoApilado=!modoApilado;' +
        'document.getElementById("btn-modo-aus-tipos").textContent=modoApilado?"Cambiar a barras agrupadas":"Cambiar a barras apiladas";' +
        'renderChartAusTipos();' +
      '});' +

      // ====== Chart 4: pastel tipos ======
      'function renderChartTipos(){' +
        'if(!__charts.chartTipos){' +
          'const ch=new Chart(document.getElementById("chartTipos").getContext("2d"),{' +
            'type:"pie",data:{labels:D.labelsTipos,datasets:[{data:D.datosTipos,backgroundColor:["#00B050","#FF4444","#FFC000","#FFFF00","#E0E0E0","#90CAF9","#CE93D8"]}]},options:{responsive:true}' +
          '}); setChart("chartTipos",ch);' +
        '}' +
      '}' +

      // ====== NUEVO: Chart Teórico vs Real (3 modos) ======
      'function buildRT_Semana(){' +
        'const labels=[]; const real=[]; const teo=[];' +
        'D.semanasRT.forEach(s=>{ if(!enRangoSemana(s)) return; labels.push(String(s)); real.push(D.realPorSemana[s]||0); teo.push(D.teoPorSemana[s]||0); });' +
        'return {labels,real,teo,xTitle:"Semana"};' +
      '}' +
      'function buildRT_SemanaTurno(){' +
        'const labels=[]; const real=[]; const teo=[];' +
        'D.semanasRT.forEach(s=>{' +
          'if(!enRangoSemana(s)) return;' +
          'D.turnosOrden.forEach(t=>{' +
            'const k=String(s)+"|"+String(t);' +
            'const r=D.realPorSemanaTurno[k]||0; const te=D.teoPorSemanaTurno[k]||0;' +
            'if(r===0 && te===0) return;' +
            'labels.push("Sem "+s+" - "+t); real.push(r); teo.push(te);' +
          '});' +
        '});' +
        'return {labels,real,teo,xTitle:"Semana y turno"};' +
      '}' +
      'function buildRT_SemanaDia(){' +
        'const labels=[]; const real=[]; const teo=[];' +
        'D.semanasRT.forEach(s=>{' +
          'if(!enRangoSemana(s)) return;' +
          // días presentes esa semana (según realPorSemanaDiaTurno)
          'const diasSet=new Set();' +
          'D.turnosOrden.forEach(t=>{Object.keys(D.realPorSemanaDiaTurno).forEach(k=>{ if(k.indexOf(String(s)+"|")===0){ const parts=k.split("|"); diasSet.add(parts[1]); } });});' +
          'const dias=Array.from(diasSet).sort((a,b)=>(D.ordenDias[a]??99)-(D.ordenDias[b]??99));' +
          'dias.forEach(d=>{' +
            'const k=String(s)+"|"+String(d);' +
            'const r=D.realPorSemanaDia[k]||0; const te=D.teoPorSemanaDia[k]||0;' +
            'labels.push("Sem "+s+" - "+d); real.push(r); teo.push(te);' +
          '});' +
        '});' +
        'return {labels,real,teo,xTitle:"Semana y día"};' +
      '}' +
	  'function buildRT_SemanaDiaTurno(){' +
  'const labels=[]; const real=[]; const teo=[];' +
  'D.semanasRT.forEach(s=>{' +
    'if(!enRangoSemana(s)) return;' +
    // sacar días presentes en esa semana mirando claves semana|dia|turno
    'const diasSet=new Set();' +
    'Object.keys(D.realPorSemanaDiaTurno).forEach(k=>{' +
      'if(k.indexOf(String(s)+"|")===0){ const parts=k.split("|"); diasSet.add(parts[1]); }' +
    '});' +
    'const dias=Array.from(diasSet).sort((a,b)=>(D.ordenDias[a]??99)-(D.ordenDias[b]??99));' +
    'dias.forEach(d=>{' +
      'D.turnosOrden.forEach(t=>{' +
        'const k=String(s)+"|"+String(d)+"|"+String(t);' +
        'const r=D.realPorSemanaDiaTurno[k]||0; const te=D.teoPorSemanaDiaTurno[k]||0;' +
        'if(r===0 && te===0) return;' +
        'labels.push("Sem "+s+" - "+d+" - "+t); real.push(r); teo.push(te);' +
      '});' +
    '});' +
  '});' +
  'return {labels,real,teo,xTitle:"Semana, día y turno"};' +
'}' +

      'function renderChartRT(){' +
'const modo=document.getElementById("sel-modo-rt").value;' +
'let x=null;' +
'if(modo==="SEMANA") x=buildRT_Semana();' +
'else if(modo==="SEMANA_TURNO") x=buildRT_SemanaTurno();' +
'else if(modo==="SEMANA_DIA") x=buildRT_SemanaDia();' +
'else x=buildRT_SemanaDiaTurno();' +

        'if(!__charts.chartTeoReal){' +
          'const ch=new Chart(document.getElementById("chartTeoReal").getContext("2d"),{' +
            'type:"bar",' +
            'data:{labels:x.labels,datasets:[' +
              '{label:"Real",data:x.real,backgroundColor:"rgba(25, 118, 210, 0.6)"},' +
              '{label:"Teórico",data:x.teo,backgroundColor:"rgba(85, 139, 47, 0.6)"}' +
            ']},' +
            'options:{responsive:true,scales:{x:{title:{display:true,text:x.xTitle}},y:{beginAtZero:true,title:{display:true,text:"Personas"}}}}' +
          '}); setChart("chartTeoReal",ch);' +
        '} else {' +
          '__charts.chartTeoReal.data.labels=x.labels;' +
          '__charts.chartTeoReal.data.datasets[0].data=x.real;' +
          '__charts.chartTeoReal.data.datasets[1].data=x.teo;' +
          '__charts.chartTeoReal.options.scales.x.title.text=x.xTitle;' +
          '__charts.chartTeoReal.update();' +
        '}' +
      '}' +
      'document.getElementById("sel-modo-rt").addEventListener("change", renderChartRT);' +

      // ====== Render inicial + filtro global ======
      'function renderAll(){renderChartMinimos(); renderChartAusencias(); renderChartAusTipos(); renderChartRT(); renderChartTipos();}' +
      'document.getElementById("btn-aplicar-filtro").addEventListener("click",()=>{' +
        'const d=document.getElementById("filtro-sem-desde").value; const h=document.getElementById("filtro-sem-hasta").value;' +
        'filtroDesde = (d!=="" ? Number(d) : null); filtroHasta = (h!=="" ? Number(h) : null);' +
        'renderAll();' +
      '});' +
      'document.getElementById("btn-reset-filtro").addEventListener("click",()=>{' +
        'filtroDesde=null; filtroHasta=null;' +
        'document.getElementById("filtro-sem-desde").value="";' +
        'document.getElementById("filtro-sem-hasta").value="";' +
        'renderAll();' +
      '});' +

      // ====== Init ======
      'conectarToggles();' +
      // autocompletar filtro con min/max de semanas RT
      'if(D.semanasRT && D.semanasRT.length){' +
        'document.getElementById("filtro-sem-desde").value=D.semanasRT[0];' +
        'document.getElementById("filtro-sem-hasta").value=D.semanasRT[D.semanasRT.length-1];' +
        'filtroDesde=D.semanasRT[0]; filtroHasta=D.semanasRT[D.semanasRT.length-1];' +
      '}' +
      'renderAll();' +
    '</script>';

  const footer = '</body></html>';

  win.document.open();
  win.document.write(head + body + script + footer);
  win.document.close();
}


});
