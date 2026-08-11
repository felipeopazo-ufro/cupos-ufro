const { chromium } = require('playwright');
const path = require('path');

const USER = process.env.UFRO_USER;
const PASS = process.env.UFRO_PASSWORD;

if (!USER || !PASS) {
  console.error('Faltan UFRO_USER o UFRO_PASSWORD.');
  process.exit(1);
}

async function conectarIntranet(page) {
  console.log('Intentando conectar con Intranet UFRO...');

  let conectado = false;

  for (let intento = 1; intento <= 3; intento++) {
    try {
      console.log(`Intento ${intento} de 3...`);

      await page.goto('https://intranet.ufro.cl/', {
        waitUntil: 'commit',
        timeout: 90000
      });

      await page.locator('#POPUSERNAME').waitFor({
        state: 'visible',
        timeout: 30000
      });

      conectado = true;
      console.log('Conexión con Intranet UFRO exitosa.');
      break;

    } catch (error) {
      console.log(`Intento ${intento} fallido: ${error.message}`);

      if (intento < 3) {
        console.log('Esperando 10 segundos antes de reintentar...');
        await page.waitForTimeout(10000);
      }
    }
  }

  if (!conectado) {
    throw new Error(
      'No fue posible acceder a Intranet UFRO después de 3 intentos.'
    );
  }
}

async function iniciarSesion(page) {
  console.log('Ingresando credenciales...');

  await page.locator('#POPUSERNAME').fill(USER);
  await page.locator('#XYZ').fill(PASS);

  await page.getByRole('link', {
    name: 'INGRESO INTRANET'
  }).click();

  console.log('Inicio de sesión enviado.');
}

async function abrirGestionCarreras(page) {
  console.log('Abriendo Dirección de Carrera...');

  await page.getByRole('link', {
    name: 'Dirección de Carrera'
  }).click();

  const popupPromise = page.waitForEvent('popup');

  await page.getByRole('link', {
    name: 'Gestión de Carreras'
  }).click();

  const gestion = await popupPromise;

  await gestion.waitForLoadState('domcontentloaded');

  return gestion;
}

async function descargarReporte(
  gestion,
  unidad,
  carpetaConsulta,
  nombreReporte,
  nombreArchivo
) {
  console.log(`Seleccionando unidad: ${unidad}`);

  await gestion.getByRole('link', {
    name: unidad
  }).click();

  await gestion.getByRole('link', {
    name: 'Consultas '
  }).click();

  await gestion.getByRole('link', {
    name: 'Consultas Personalizadas'
  }).click();

  await gestion.getByRole('link', {
    name: ' Consultas Globales Pregrado'
  }).click();

  await gestion.getByRole('link', {
    name: carpetaConsulta
  }).click();

  await gestion.getByRole('link', {
    name: nombreReporte
  }).click();

  console.log(`Ejecutando reporte: ${nombreReporte}`);

  await gestion.getByRole('button', {
    name: 'Buscar '
  }).click();

  gestion.once('dialog', async dialog => {
    console.log(`Cuadro de descarga: ${dialog.message()}`);

    await dialog.accept(
      nombreArchivo.replace(/\.xls$/i, '')
    );
  });

  const downloadPromise = gestion.waitForEvent('download', {
    timeout: 60000
  });

  await gestion.getByRole('button', {
    name: ' XLS'
  }).click();

  const download = await downloadPromise;

  const destino = path.join(
    process.cwd(),
    nombreArchivo
  );

  await download.saveAs(destino);

  console.log(`Archivo descargado: ${nombreArchivo}`);
}

(async () => {
  console.log('Iniciando Playwright...');

  const browser = await chromium.launch({
    headless: true
  });

  const context = await browser.newContext({
    acceptDownloads: true
  });

  const page = await context.newPage();

  page.setDefaultTimeout(60000);

  try {
    /*
     * 1. CONECTAR A INTRANET
     */
    await conectarIntranet(page);

    /*
     * 2. INICIAR SESIÓN
     */
    await iniciarSesion(page);

    /*
     * 3. DESCARGAR ELECTIVOS
     */
    console.log('--- Descargando Electivos ---');

    let gestion = await abrirGestionCarreras(page);

    await descargarReporte(
      gestion,
      'CIP - Centro Innovación',
      ' Electivos Formación General',
      'Resumen Electivos F.Gral.',
      'resumen electivos.xls'
    );

    await gestion.close();

    /*
     * 4. DESCARGAR INGLÉS
     */
    console.log('--- Descargando Inglés ---');

    gestion = await abrirGestionCarreras(page);

    await descargarReporte(
      gestion,
      'codi - Coordinación Idiomas',
      ' Consultas CODI',
      'Resumen Inscritos CODI',
      'resumen ingles.xls'
    );

    await gestion.close();

    console.log('--------------------------------');
    console.log('Descargas completadas correctamente.');
    console.log('--------------------------------');

  } catch (error) {
    console.error('ERROR EN LA AUTOMATIZACIÓN:');
    console.error(error);

    process.exitCode = 1;

  } finally {
    await browser.close();
  }
})();
