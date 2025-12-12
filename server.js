const express = require('express');
const path = require('path');
const multer = require('multer');
const xlsx = require('xlsx');     // Datos
const ExcelJS = require('exceljs'); // Colores
const cors = require('cors');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;


// Middleware
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Subida de ficheros
const upload = multer({ dest: path.join(__dirname, 'uploads') });

// Ruta simple de test
app.get('/api/ping', (req, res) => {
  res.json({ ok: true, message: 'Servidor funcionando' });
});

// Colores ARGB
const COLOR_6_DIAS_ARGB = 'FF00AF50';  // Verde 6 días
const COLOR_HH_ARGB     = 'FFFFFF00';  // Amarillo HH
const COLOR_BOLSA_ARGB  = 'FFFFBF00';  // Naranja Bolsa
const COLOR_BOLSA_ARGB2  = 'FFFFC000';  // Naranja Bolsa2

/**
 * Escanea colores con ExcelJS y devuelve un mapa:
 * { "A1": "FF00AF50", "C2": "FFFFFF00", ... }
 */
async function escanearColores(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.getWorksheet(1); // Primera hoja
  const colorMap = {};

  if (!worksheet) {
    console.warn('No se encontró hoja en ExcelJS');
    return colorMap;
  }

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const cellAddress = xlsx.utils.encode_cell({ r: rowNumber - 1, c: colNumber - 1 });
      let color = null;

      if (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
        color = cell.fill.fgColor ? cell.fill.fgColor.argb : null;
      }

      colorMap[cellAddress] = color;
    });
  });

  return colorMap;
}

/**
 * Busca la fila de días (Domingo, Lunes, ...) y devuelve:
 * - filaDias: índice de fila
 * - dayCols: array de objetos { col, label } en orden Domingo..Sábado
 */
function detectarFilaYColumnasDias(rows) {
  const nombresDias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  let filaDias = -1;
  let dayCols = [];

  for (let r = 0; r < rows.length; r++) {
    const fila = rows[r];
    const encontrados = [];

    for (let c = 0; c < fila.length; c++) {
      const valor = (fila[c]?.valor || '').toString().trim();
      if (!valor) continue;

      const nombre = valor.split(' ')[0]; // "Lunes 3" -> "Lunes"
      if (nombresDias.includes(nombre)) {
        encontrados.push({ col: c, label: nombre });
      }
    }

    if (encontrados.length >= 2) {
      dayCols = encontrados.sort(
        (a, b) => nombresDias.indexOf(a.label) - nombresDias.indexOf(b.label)
      );
      filaDias = r;
      break;
    }
  }

  return { filaDias, dayCols };
}

/**
 * Procesa un workbook de Excel con el formato de tu hoja de horarios.
 * Acepta un mapa de colores como argumento.
 * Devuelve array de registros { semana, dia, turno, persona, tipo }.
 */
function procesarHorario(workbook, colorMap) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const range = sheet['!ref'];

  if (!range) return [];

  const rows = [];
  const ref = xlsx.utils.decode_range(range);

  // Leer todas las celdas con valor + tipo + color
  for (let R = ref.s.r; R <= ref.e.r; R++) {
    const row = [];
    for (let C = ref.s.c; C <= ref.e.c; C++) {
      const cellAddress = xlsx.utils.encode_cell({ r: R, c: C });
      const cell = sheet[cellAddress];
      const valor = cell ? String(cell.v || '').trim() : '';
      let tipo = 'Normal';

      const colorCelda = colorMap[cellAddress] || null;

      if (valor) {
        const t = valor.toUpperCase();

        // 1) PRIORIDAD COLOR 6 DÍAS
        if (colorCelda === COLOR_6_DIAS_ARGB) {
          tipo = '6Dias';
        }
        // 2) PRIORIDAD COLOR BOLSA (dejamos de usar nombres)
        else if ((colorCelda === COLOR_BOLSA_ARGB) || (colorCelda === COLOR_BOLSA_ARGB2)) {
          tipo = 'Bolsa';
        }
        // 3) PRIORIDAD COLOR HH
        else if (colorCelda === COLOR_HH_ARGB) {
          tipo = 'DisfruteHH';
        }
        // 4) AUSENCIAS POR CÓDIGO (agrupadas)
        else if (
          // AP
          t.includes('(AP)') ||
          // BM
          t.includes('(BM)') ||
          // HHEE/HHDD
          t.includes('HHEE') || t.includes('HHDD') ||
          // Vacaciones VC/VP/Vx/V
          t.includes('(VC)') || t.includes('(VP)') || t.includes('(VX)') || t.includes('(V)') ||
          // Resto licencias: PAT, LIC, HF, etc.
          t.includes('(PAT)') || t.includes('(LIC)') || t.includes(' HF')
        ) {
          tipo = 'Ausencia';
        }
      }

      row.push({ valor, tipo });
    }
    rows.push(row);
  }

  // Detectar fila y columnas de días
  const { filaDias, dayCols } = detectarFilaYColumnasDias(rows);
  if (filaDias === -1 || dayCols.length === 0) {
    console.warn('No se pudo detectar la fila de días');
    return [];
  }

  // Buscar nº de semana
  let semanaNumero = null;
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const texto = rows[r][c].valor;
      if (texto && texto.toLowerCase().includes('semana') && texto.match(/(\d+)/)) {
        semanaNumero = parseInt(texto.match(/(\d+)/)[1]);
        break;
      }
    }
    if (semanaNumero !== null) break;
  }
  if (semanaNumero === null) semanaNumero = 0;

  const dias = rows[filaDias] || [];

  // Buscar filas donde aparece MAÑANA / TARDE / NOCHE en la primera columna
  let filaManana = -1;
  let filaTarde = -1;
  let filaNoche = -1;
  for (let r = 0; r < rows.length; r++) {
    const primeraCol = (rows[r][0]?.valor || '').toUpperCase();
    if (primeraCol.includes('MAÑANA')) filaManana = r;
    if (primeraCol.includes('TARDE')) filaTarde = r;
    if (primeraCol.includes('NOCHE')) filaNoche = r;
  }

  const registros = [];

  // Helper para recorrer un bloque de filas y un turno concreto
  function procesarBloqueTurno(filaInicio, filaFin, turnoLabel) {
    if (filaInicio === -1 || filaFin === -1) return;

    for (let r = filaInicio; r < filaFin; r++) {
      const fila = rows[r];
      for (const dayInfo of dayCols) {
        const c = dayInfo.col; // columnas de Domingo..Sábado
        const diaTexto = dias[c] ? dias[c].valor : '';
        const personaData = fila[c];

        if (diaTexto && personaData && personaData.valor) {
          registros.push({
            semana: semanaNumero,
            dia: diaTexto,
            turno: turnoLabel,
            persona: personaData.valor,
            tipo: personaData.tipo
          });
        }
      }
    }
  }

  // MAÑANA
  if (filaManana !== -1 && filaTarde !== -1) {
    procesarBloqueTurno(filaManana, filaTarde, 'MAÑANA');
  }

  // TARDE
  if (filaTarde !== -1 && filaNoche !== -1) {
    procesarBloqueTurno(filaTarde, filaNoche, 'TARDE');
  }

  // NOCHE
  if (filaNoche !== -1) {
    let filaFinNoche = rows.length;
    for (let r = filaNoche + 1; r < rows.length; r++) {
      const primeraCol = (rows[r][0]?.valor || '').toUpperCase();
      const filaVacia = !rows[r].some(celda => celda.valor && celda.valor.trim());
      if (primeraCol.includes('VACACIONES') || primeraCol.includes('ADMIN') || filaVacia) {
        filaFinNoche = r;
        break;
      }
    }
    procesarBloqueTurno(filaNoche, filaFinNoche, 'NOCHE');
  }

  // Estadísticas en consola
  const stats = {};
  registros.forEach(r => {
    stats[r.tipo] = (stats[r.tipo] || 0) + 1;
  });
  console.log(`✅ Semana ${semanaNumero}: ${registros.length} registros`, stats);

  return registros;
}

// Endpoint para subir y procesar Excels
app.post('/upload', upload.array('excel'), async (req, res) => {
  try {
    const files = req.files || [];
    if (files.length === 0) {
      return res.status(400).json({ ok: false, message: 'No se recibieron ficheros.' });
    }

    let registrosTotal = [];
    for (const file of files) {
      // 1. Escanear colores con ExcelJS
      const colorMap = await escanearColores(file.path);

      // 2. Leer datos con XLSX
      const workbook = xlsx.readFile(file.path);

      // 3. Procesar usando mapa de colores
      const regs = procesarHorario(workbook, colorMap);

      registrosTotal = registrosTotal.concat(regs);
      fs.unlinkSync(file.path);
    }

    console.log(`🎉 TOTAL: ${registrosTotal.length} registros de ${files.length} fichero(s).`);
    res.json({ ok: true, registros: registrosTotal });
  } catch (err) {
    console.error('❌ Error procesando Excel:', err);
    res.status(500).json({ ok: false, message: 'Error procesando los Excels.' });
  }
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en http://localhost:${PORT}`);
});
