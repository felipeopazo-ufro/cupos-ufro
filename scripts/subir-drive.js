const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

const credsRaw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const folderId = process.env.DRIVE_FOLDER_ID;

if (!credsRaw || !folderId) {
  console.log('Drive no configurado: se omite la copia a Google Drive.');
  process.exit(0);
}

const credentials = JSON.parse(credsRaw);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/drive']
});
const drive = google.drive({ version: 'v3', auth });

async function upsert(localPath, name, mimeType) {
  const q = `'${folderId}' in parents and name='${name.replaceAll("'", "\\'")}' and trashed=false`;
  const found = await drive.files.list({
    q,
    fields: 'files(id,name)',
    spaces: 'drive'
  });

  const media = {
    mimeType,
    body: fs.createReadStream(localPath)
  };

  if (found.data.files.length) {
    const fileId = found.data.files[0].id;
    await drive.files.update({
      fileId,
      media,
      fields: 'id,name,modifiedTime'
    });
    console.log(`Drive actualizado: ${name}`);
  } else {
    await drive.files.create({
      requestBody: {
        name,
        parents: [folderId]
      },
      media,
      fields: 'id,name'
    });
    console.log(`Drive creado: ${name}`);
  }
}

(async () => {
  await upsert(
    path.join(process.cwd(), 'site', 'index.html'),
    'cupos-disponibles-2026-2.html',
    'text/html'
  );
  await upsert(
    path.join(process.cwd(), 'site', 'cupos-disponibles-2026-2.csv'),
    'cupos-disponibles-2026-2.csv',
    'text/csv'
  );
})().catch(err => {
  console.error('Error al subir a Drive:', err.message);
  process.exit(1);
});
