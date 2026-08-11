const { chromium } = require('playwright');
const path = require('path');

const USER = process.env.UFRO_USER;
const PASS = process.env.UFRO_PASSWORD;

if (!USER || !PASS) {
  console.error('Faltan UFRO_USER o UFRO_PASSWORD.');
  process.exit(1);
}

async function descargar(page, unidadLink, rutaConsultas, reporte, nombreArchivo) {
  await page.getByRole('link', { name: unidadLink }).click();
  await page.getByRole('link', { name: 'Consultas ' }).click();
  await page.getByRole('link', { name: 'Consultas Personalizadas' }).click();
  await page.getByRole('link', { name: ' Consultas Globales Pregrado' }).click();

  for (const item of rutaConsultas) {
    await page.getByRole('link', { name: item }).click();
  }

  await page.getByRole('link', { name: reporte }).click();
  await page.getByRole('button', { name: 'Buscar ' }).click();

  page.once('dialog', async dialog => {
    await dialog.accept(nombreArchivo.replace(/\.xls$/i, ''));
  });

  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: ' XLS' }).click();
  const download = await downloadPromise;
  await download.saveAs(path.join(process.cwd(), nombreArchivo));
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  try {
    await page.goto('https://intranet.ufro.cl/', { waitUntil: 'domcontentloaded' });
    await page.locator('#POPUSERNAME').fill(USER);
    await page.locator('#XYZ').fill(PASS);
    await page.getByRole('link', { name: 'INGRESO INTRANET' }).click();

    await page.getByRole('link', { name: 'Dirección de Carrera' }).click();

    // Primera descarga: Electivos
    let popupPromise = page.waitForEvent('popup');
    await page.getByRole('link', { name: 'Gestión de Carreras' }).click();
    let gestion = await popupPromise;

    await descargar(
      gestion,
      'CIP - Centro Innovación',
      [' Electivos Formación General'],
      'Resumen Electivos F.Gral.',
      'resumen electivos.xls'
    );
    await gestion.close();

    // Segunda descarga: Inglés
    popupPromise = page.waitForEvent('popup');
    await page.getByRole('link', { name: 'Gestión de Carreras' }).click();
    gestion = await popupPromise;

    await descargar(
      gestion,
      'codi - Coordinación Idiomas',
      [' Consultas CODI'],
      'Resumen Inscritos CODI',
      'resumen ingles.xls'
    );

    console.log('Descargas completadas correctamente.');
  } finally {
    await browser.close();
  }
})().catch(err => {
  console.error(err);
  process.exit(1);
});
