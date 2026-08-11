const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const BASE = __dirname;

const RESUMEN_ELECTIVOS = path.join(BASE, 'resumen electivos.xls');
const RESUMEN_INGLES = path.join(BASE, 'resumen ingles.xls');
const HORARIOS_ELECTIVOS = path.join(BASE, 'horarios_maestro_2026-2.json');
const HORARIOS_INGLES = path.join(BASE, 'horarios_ingles_2026-2.json');
const DESCRIPCIONES_ELECTIVOS = path.join(BASE, 'descripciones_electivos_2026-2.json');
const ENCABEZADO_UFRO = path.join(BASE, 'encabezado-ufro.png');
const ENCABEZADO_DATA_URI = `data:image/png;base64,${fs.readFileSync(ENCABEZADO_UFRO).toString('base64')}`;

const SITE_DIR = path.join(BASE, 'site');
const SALIDA_HTML = path.join(SITE_DIR, 'index.html');
const SALIDA_CSV = path.join(SITE_DIR, 'cupos-disponibles-2026-2.csv');

const CODIGOS_MINOR_INGLES = new Set([
  'DFI183','DFI185','ELL591','ELL822','ELL605','ELL768',
  'ELL599','ELL753','DFI181','DFI638','COD302'
]);

const CUPOS_INGLES_POR_DIA = {
  COD590: { lunes: 5, martes: 12, miercoles: 16, jueves: 20, viernes: 22 },
  COD594: { lunes: 6, martes: 12, miercoles: 16, jueves: 20, viernes: 22 },
  COD598: { lunes: 5, martes: 11, miercoles: 15, jueves: 20, viernes: 22 },
  COD602: { lunes: 5, martes: 10, miercoles: 15, jueves: 20, viernes: 22 }
};

function norm(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v.text !== undefined) return String(v.text).trim();
  if (typeof v === 'object' && v.result !== undefined) return String(v.result).trim();
  return String(v).trim();
}

function escHtml(v) {
  return norm(v)
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;');
}

function csvCell(v) {
  return `"${norm(v).replaceAll('"','""')}"`;
}

function chileNow() {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Santiago',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    hourCycle: 'h23',
    weekday: 'long'
  }).formatToParts(new Date());

  const obj = {};
  for (const p of parts) obj[p.type] = p.value;

  const weekdayMap = {
    Monday: 'lunes',
    Tuesday: 'martes',
    Wednesday: 'miercoles',
    Thursday: 'jueves',
    Friday: 'viernes',
    Saturday: 'sabado',
    Sunday: 'domingo'
  };

  return {
    year: Number(obj.year),
    month: Number(obj.month),
    day: Number(obj.day),
    hour: Number(obj.hour),
    minute: Number(obj.minute),
    weekday: weekdayMap[obj.weekday],
    isoDate: `${obj.year}-${obj.month}-${obj.day}`
  };
}

function fechaBonitaChile() {
  return new Intl.DateTimeFormat('es-CL', {
    timeZone: 'America/Santiago',
    dateStyle: 'long',
    timeStyle: 'short'
  }).format(new Date());
}

async function readDisguisedXlsx(filePath) {
  if (!fs.existsSync(filePath)) throw new Error(`No encontré: ${filePath}`);
  const tmp = filePath.replace(/\.xls$/i, '') + '_temporal.xlsx';
  fs.copyFileSync(filePath, tmp);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(tmp);
  try { fs.unlinkSync(tmp); } catch {}
  return wb.worksheets[0];
}

function findHeaderRow(ws, requiredHeaders, maxRows = 10) {
  for (let r = 1; r <= Math.min(maxRows, ws.rowCount); r++) {
    const found = new Set();
    ws.getRow(r).eachCell((cell) => found.add(norm(cell.value)));
    if (requiredHeaders.every(h => found.has(h))) return r;
  }
  throw new Error(`No encontré encabezados: ${requiredHeaders.join(', ')}`);
}

function headerMapForRow(ws, rowNum) {
  const map = {};
  ws.getRow(rowNum).eachCell((cell, col) => {
    map[norm(cell.value)] = col;
  });
  return map;
}

async function procesarElectivos() {
  const ws = await readDisguisedXlsx(RESUMEN_ELECTIVOS);
  const headerRow = findHeaderRow(ws, ['Código','Módulo','Vacantes']);
  const headers = headerMapForRow(ws, headerRow);

  const horarios = JSON.parse(fs.readFileSync(HORARIOS_ELECTIVOS, 'utf8'));
  const descripciones = fs.existsSync(DESCRIPCIONES_ELECTIVOS)
    ? JSON.parse(fs.readFileSync(DESCRIPCIONES_ELECTIVOS, 'utf8'))
    : {};
  const hMap = new Map();

  for (const h of horarios) {
    const codigo = norm(h['Código']).replace(/\s*\(Pilares\)\s*/gi,'').trim();
    const modulo = norm(h['Módulo']);
    hMap.set(`${codigo}|${modulo}`, {...h, 'Código': codigo});
  }

  const resultado = [];
  const sinHorario = [];

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const codigo = norm(row.getCell(headers['Código']).value);
    const modulo = norm(row.getCell(headers['Módulo']).value);
    const vacantes = Number(row.getCell(headers['Vacantes']).value ?? 0);

    if (!codigo || !modulo || !(vacantes > 0)) continue;
    if (CODIGOS_MINOR_INGLES.has(codigo)) continue;
    if (['COD590','COD594','COD598','COD602'].includes(codigo)) continue;

    const h = hMap.get(`${codigo}|${modulo}`);
    if (!h) {
      sinHorario.push({codigo, modulo, vacantes});
      continue;
    }

    resultado.push({
      tipo: 'Electivo',
      codigo,
      asignatura: norm(h['Asignatura']),
      formato: norm(h['Formato']),
      modulo,
      dia: norm(h['Día']),
      horario: norm(h['Horario']),
      sala: norm(h['Sala']),
      docente: norm(h['Docente']),
      correo: norm(h['Correo electrónico']),
      descripcion: norm(descripciones[codigo] || ''),
      vacantes
    });
  }

  return {resultado, sinHorario};
}

async function procesarIngles() {
  const now = chileNow();

  const inicio = '2026-08-10';
  const fin = '2026-08-14';

  if (now.isoDate < inicio) {
    return {resultado: [], sinHorario: [], nota: 'La liberación de cupos de Inglés aún no comienza.'};
  }

  let diaParaCupos = now.weekday;

  // Después del viernes 14 se conserva el máximo del viernes.
  if (now.isoDate > fin) diaParaCupos = 'viernes';

  // Si se ejecuta durante fin de semana dentro de otro contexto, usar viernes como último tope disponible.
  if (!['lunes','martes','miercoles','jueves','viernes'].includes(diaParaCupos)) {
    diaParaCupos = 'viernes';
  }

  const ws = await readDisguisedXlsx(RESUMEN_INGLES);
  const headerRow = findHeaderRow(ws, ['Código','Módulo','N° Inscritos']);
  const headers = headerMapForRow(ws, headerRow);

  const horarios = JSON.parse(fs.readFileSync(HORARIOS_INGLES, 'utf8'));
  const hMap = new Map();
  for (const h of horarios) {
    hMap.set(`${norm(h['Código'])}|${norm(h['Módulo'])}`, h);
  }

  const resultado = [];
  const sinHorario = [];

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const codigo = norm(row.getCell(headers['Código']).value);
    const modulo = norm(row.getCell(headers['Módulo']).value);

    if (!['COD590','COD594','COD598','COD602'].includes(codigo) || !modulo) continue;

    const inscritos = Number(row.getCell(headers['N° Inscritos']).value ?? 0);
    const tope = CUPOS_INGLES_POR_DIA[codigo][diaParaCupos];
    const vacantes = Math.max(0, tope - inscritos);

    if (!(vacantes > 0)) continue;

    const h = hMap.get(`${codigo}|${modulo}`);
    if (!h) {
      sinHorario.push({codigo, modulo, inscritos, tope, vacantes});
      continue;
    }

    resultado.push({
      tipo: 'Inglés',
      codigo,
      asignatura: norm(h['Nivel']),
      formato: 'Presencial',
      modulo,
      dia: '',
      horario: norm(h['Horario']),
      sala: norm(h['Sala']),
      docente: '',
      correo: '',
      inscritos,
      tope,
      vacantes
    });
  }

  return {
    resultado,
    sinHorario,
    nota: `Cupos de Inglés calculados con el tope correspondiente a ${diaParaCupos}.`
  };
}

function tablaHtml(items, incluirDocente = true) {
  if (!items.length) {
    return `<div class="empty">No hay cupos disponibles en este momento.</div>`;
  }

  const rows = items.map(x => {
    const descripcion = x.descripcion
      ? `<details class="desc"><summary>Ver descripción</summary><div>${escHtml(x.descripcion)}</div></details>`
      : '';

    return `
    <tr>
      <td>${escHtml(x.codigo)}</td>
      <td class="asig">
        ${escHtml(x.asignatura)}
        ${descripcion}
      </td>
      <td>${escHtml(x.modulo)}</td>
      <td>${escHtml(x.dia || x.horario.split(' ')[0])}</td>
      <td>${escHtml(x.horario)}</td>
      <td>${escHtml(x.sala)}</td>
      ${incluirDocente ? `<td>${escHtml(x.docente)}</td>` : ''}
      <td class="vac">${x.vacantes}</td>
    </tr>`;
  }).join('');

  return `
  <div class="tablewrap">
    <table>
      <thead><tr>
        <th>Código</th><th>Asignatura / Nivel</th><th>Módulo</th><th>Día</th>
        <th>Horario</th><th>Sala</th>${incluirDocente ? '<th>Docente</th>' : ''}<th>Cupos</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

function cardsHtml(items) {
  if (!items.length) return '';
  return `<div class="cards">${items.map(x => `
    <article class="card">
      <div class="top">
        <span class="code">${escHtml(x.codigo)} · Módulo ${escHtml(x.modulo)}</span>
        <strong>${x.vacantes} cupo${x.vacantes === 1 ? '' : 's'}</strong>
      </div>
      <h3>${escHtml(x.asignatura)}</h3>
      <p><b>${escHtml(x.horario)}</b></p>
      <p>${escHtml(x.sala)}</p>
      ${x.docente ? `<p>${escHtml(x.docente)}</p>` : ''}
      ${x.descripcion ? `<details class="desc"><summary>Ver descripción</summary><div>${escHtml(x.descripcion)}</div></details>` : ''}
    </article>`).join('')}</div>`;
}

function inglesPorNivelHtml(items) {
  const niveles = [
    {codigo:'COD590', titulo:'Inglés Principiante'},
    {codigo:'COD594', titulo:'Inglés Básico'},
    {codigo:'COD598', titulo:'Inglés Pre-Intermedio'},
    {codigo:'COD602', titulo:'Inglés Intermedio'}
  ];

  const botones = niveles.map((n, i) =>
    `<button class="levelbtn ${i === 0 ? 'active' : ''}" data-level="${n.codigo}">${n.titulo}</button>`
  ).join('');

  const paneles = niveles.map((n, i) => {
    const lista = items.filter(x => x.codigo === n.codigo);
    const contenido = lista.length
      ? `${tablaHtml(lista, false)}${cardsHtml(lista)}`
      : `<div class="no-cupos">Lamentablemente este nivel no tiene cupos en este momento. Vuelve a consultar más tarde.</div>`;

    return `
      <div id="nivel-${n.codigo}" class="levelpanel ${i === 0 ? 'active' : ''}">
        <h3 class="leveltitle">${n.titulo}</h3>
        ${contenido}
      </div>`;
  }).join('');

  return `
    <div class="leveltabs">${botones}</div>
    ${paneles}
  `;
}

(async () => {
  fs.mkdirSync(SITE_DIR, { recursive: true });
  for (const f of [HORARIOS_ELECTIVOS, HORARIOS_INGLES, DESCRIPCIONES_ELECTIVOS]) {
    if (!fs.existsSync(f)) throw new Error(`No encontré: ${f}`);
  }

  const electivos = await procesarElectivos();
  const ingles = await procesarIngles();

  electivos.resultado.sort((a,b) =>
    a.asignatura.localeCompare(b.asignatura,'es') || Number(a.modulo)-Number(b.modulo)
  );

  ingles.resultado.sort((a,b) =>
    a.codigo.localeCompare(b.codigo) || Number(a.modulo)-Number(b.modulo)
  );

  const advertencias = [
    ...electivos.sinHorario.map(x => `Electivo sin horario: ${x.codigo}-${x.modulo}`),
    ...ingles.sinHorario.map(x => `Inglés sin horario: ${x.codigo}-${x.modulo}`)
  ];

  const html = `<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cupos disponibles 2026-2</title>
<style>
*{box-sizing:border-box}
body{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f4f6f9;color:#172033}
.wrap{max-width:1500px;margin:auto;padding:28px}
.brand-header{background:#000;border-radius:16px;overflow:hidden;margin-bottom:18px;box-shadow:0 2px 10px #00000012}
.brand-header img{display:block;width:100%;height:auto}
header,.section{background:#fff;border-radius:16px;padding:24px;margin-bottom:18px;box-shadow:0 2px 10px #00000012}
h1{margin:0 0 8px;font-size:30px}
h2{margin:0 0 6px;font-size:24px}
.meta,.note{color:#616b7d}
.note{font-size:14px}
.tablewrap{overflow:auto;margin-top:16px}
table{border-collapse:collapse;width:100%;min-width:1000px}
th,td{padding:11px 10px;border-bottom:1px solid #e7eaf0;text-align:left;vertical-align:top;font-size:14px}
th{background:#172033;color:white;position:sticky;top:0}
td.asig{font-weight:700;min-width:260px}
td.vac{font-size:18px;font-weight:800;text-align:center}
.warning{background:#fff3cd;border:1px solid #ffe69c;padding:12px;border-radius:10px;margin-bottom:18px}
.cards{display:none}
.empty{padding:18px;background:#f8f9fb;border-radius:10px;margin-top:14px}
.tabs{display:flex;gap:10px;margin:0 0 18px;position:sticky;top:0;z-index:20;background:#f4f6f9;padding:10px 0}
.tabbtn{border:0;border-radius:999px;padding:12px 18px;font-weight:700;cursor:pointer;background:#e3e7ee;color:#172033}
.tabbtn.active{background:#172033;color:#fff}
.tabpanel{display:none}
.tabpanel.active{display:block}
.desc{margin-top:8px;font-weight:400}
.desc summary{cursor:pointer;font-weight:700;color:#314a72;display:inline-block}
.desc div{margin-top:8px;padding:10px 12px;background:#f3f5f8;border-radius:9px;line-height:1.45;font-weight:400;min-width:260px}
.leveltabs{display:flex;flex-wrap:wrap;gap:9px;margin:18px 0}
.levelbtn{border:1px solid #ccd3df;background:#fff;color:#172033;border-radius:10px;padding:10px 14px;font-weight:700;cursor:pointer}
.levelbtn.active{background:#172033;color:#fff;border-color:#172033}
.levelpanel{display:none}
.levelpanel.active{display:block}
.leveltitle{margin:6px 0 12px;font-size:20px}
.no-cupos{padding:22px;background:#f8f9fb;border:1px solid #e4e8ef;border-radius:12px;line-height:1.5;color:#4f5a6c}

@media(max-width:800px){
  .wrap{padding:14px}
  h1{font-size:24px}
  h2{font-size:21px}
  .tablewrap{display:none}
  .cards{display:grid;gap:12px;margin-top:14px}
  .card{background:#f8f9fb;border-radius:14px;padding:16px;border:1px solid #e7eaf0}
  .card h3{font-size:18px;margin:10px 0}
  .card p{margin:5px 0}
  .top{display:flex;justify-content:space-between;gap:10px;align-items:center}
  .top strong{white-space:nowrap}
  .code{font-size:13px;color:#5f6878}
}
</style>
</head>
<body>
<div class="wrap">
<div class="brand-header">
  <img src="${ENCABEZADO_DATA_URI}" alt="Universidad de La Frontera - Dirección de Trayectoria Formativa - Coordinación de Formación General e Idiomas">
</div>
<header>
  <h1>Cupos disponibles · Segundo semestre 2026</h1>
  <p class="meta">Última actualización: ${escHtml(fechaBonitaChile())}</p>
  <p class="note">Los cupos pueden variar durante el proceso de inscripción.</p>
</header>

${advertencias.length ? `<div class="warning"><b>Advertencia interna:</b><br>${advertencias.map(escHtml).join('<br>')}</div>` : ''}

<nav class="tabs" aria-label="Tipo de oferta">
  <button class="tabbtn active" data-tab="electivos">Electivos con cupo</button>
  <button class="tabbtn" data-tab="ingles">Módulos de Inglés con Cupo</button>
</nav>

<div id="electivos" class="tabpanel active">
<section class="section">
  <h2>Electivos de Formación General con cupos</h2>
  <p class="note">Solo se muestran módulos con vacantes. Se excluyen las asignaturas del Minor en Inglés.</p>
  ${tablaHtml(electivos.resultado, true)}
  ${cardsHtml(electivos.resultado)}
</section>
</div>

<div id="ingles" class="tabpanel">
<section class="section">
  <h2>Módulos de Inglés con cupos</h2>
  <p class="note">${escHtml(ingles.nota || '')}</p>
  ${inglesPorNivelHtml(ingles.resultado)}
</section>
</div>

</div>
<script>
document.querySelectorAll('.tabbtn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tabbtn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tabpanel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
    window.scrollTo({top: 0, behavior: 'smooth'});
  });
});

document.querySelectorAll('.levelbtn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.levelbtn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.levelpanel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('nivel-' + btn.dataset.level).classList.add('active');
  });
});
</script>
</body>
</html>`;

  fs.writeFileSync(SALIDA_HTML, html, 'utf8');

  const todos = [...electivos.resultado, ...ingles.resultado];
  const csvHeader = ['Tipo','Código','Asignatura/Nivel','Módulo','Horario','Sala','Docente','Vacantes'];
  const csvRows = todos.map(x => [
    x.tipo,x.codigo,x.asignatura,x.modulo,x.horario,x.sala,x.docente,x.vacantes
  ].map(csvCell).join(';'));
  fs.writeFileSync(
    SALIDA_CSV,
    '\ufeff' + csvHeader.map(csvCell).join(';') + '\n' + csvRows.join('\n'),
    'utf8'
  );

  console.log(`OK Electivos: ${electivos.resultado.length} módulos con cupo.`);
  console.log(`OK Inglés: ${ingles.resultado.length} módulos con cupo.`);
  console.log(`HTML: ${SALIDA_HTML}`);
  console.log(`CSV: ${SALIDA_CSV}`);

  if (electivos.sinHorario.length) {
    console.log('Electivos con cupo sin coincidencia en horario:');
    console.table(electivos.sinHorario);
  }
  if (ingles.sinHorario.length) {
    console.log('Módulos de Inglés con cupo sin coincidencia en horario:');
    console.table(ingles.sinHorario);
  }
})().catch(err => {
  console.error('ERROR:', err.message);
  process.exit(1);
});
