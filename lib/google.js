'use strict';

// One Google client for both Sheets and Drive.
//
// Two ways to authenticate, checked in this order:
//
//  1. OAUTH  – GOOGLE_OAUTH_CLIENT_ID / _SECRET / _REFRESH_TOKEN
//     Acts as the practice's own Google account. Files it creates are owned by
//     that account and use its storage. This is the one to use for Drive
//     uploads on a consumer (gmail.com) account.
//
//  2. SERVICE ACCOUNT – GOOGLE_CREDENTIALS_JSON
//     Fine for Sheets. On a consumer account it can FAIL for Drive uploads:
//     a service account has no storage quota of its own, so creating a file
//     can come back as "storage quota exceeded" even with Editor on the folder.
//
// Whichever is configured, the Drive folders and both sheets must be shared
// with that identity as Editor.

const { google } = require('googleapis');
const { Readable } = require('stream');

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive'
];

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const GDOC_MIME = 'application/vnd.google-apps.document';
const PDF_MIME = 'application/pdf';

let authClient = null;
let sheetsApi = null;
let driveApi = null;

function authMode() {
  if (process.env.GOOGLE_OAUTH_REFRESH_TOKEN && process.env.GOOGLE_OAUTH_CLIENT_ID) return 'oauth';
  if (process.env.GOOGLE_CREDENTIALS_JSON) return 'service-account';
  return 'none';
}

function getAuth() {
  if (authClient) return authClient;
  const mode = authMode();
  if (mode === 'oauth') {
    const o = new google.auth.OAuth2(
      process.env.GOOGLE_OAUTH_CLIENT_ID.trim(),
      (process.env.GOOGLE_OAUTH_CLIENT_SECRET || '').trim()
    );
    o.setCredentials({ refresh_token: process.env.GOOGLE_OAUTH_REFRESH_TOKEN.trim() });
    authClient = o;
  } else if (mode === 'service-account') {
    const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    authClient = new google.auth.GoogleAuth({
      credentials: {
        client_email: creds.client_email,
        private_key: String(creds.private_key).replace(/\\n/g, '\n')
      },
      scopes: SCOPES
    });
  } else {
    throw new Error('No Google credentials set (GOOGLE_OAUTH_REFRESH_TOKEN or GOOGLE_CREDENTIALS_JSON).');
  }
  return authClient;
}

const getSheets = () => (sheetsApi ||= google.sheets({ version: 'v4', auth: getAuth() }));
const getDrive = () => (driveApi ||= google.drive({ version: 'v3', auth: getAuth() }));

const bufferStream = buf => Readable.from(buf);
const viewUrl = id => `https://drive.google.com/file/d/${id}/view`;

/* ---------- Drive ---------- */

// Uploads a file as-is (no Google conversion) and returns { id, url }.
async function uploadFile(buffer, name, folderId, mimeType = DOCX_MIME) {
  const res = await getDrive().files.create({
    requestBody: { name, parents: folderId ? [folderId] : undefined },
    media: { mimeType, body: bufferStream(buffer) },
    fields: 'id,name,webViewLink',
    supportsAllDrives: true
  });
  return { id: res.data.id, url: res.data.webViewLink || viewUrl(res.data.id), name: res.data.name };
}

// Replaces the contents of a file already in Drive, keeping its id and link.
async function replaceFile(fileId, buffer, mimeType = DOCX_MIME) {
  const res = await getDrive().files.update({
    fileId,
    media: { mimeType, body: bufferStream(buffer) },
    fields: 'id,webViewLink',
    supportsAllDrives: true
  });
  return { id: res.data.id, url: res.data.webViewLink || viewUrl(res.data.id) };
}

// Makes the PDF the practice files: Word -> Google Doc (converted on upload)
// -> exported as PDF -> saved beside it. The temporary Google Doc is removed.
// Never throws: a PDF that cannot be made must not lose the Word file.
async function docxToPdf(buffer, baseName, folderId) {
  const drive = getDrive();
  let tempId = null;
  try {
    const temp = await drive.files.create({
      requestBody: { name: `~converting ${baseName}`, mimeType: GDOC_MIME, parents: folderId ? [folderId] : undefined },
      media: { mimeType: DOCX_MIME, body: bufferStream(buffer) },
      fields: 'id',
      supportsAllDrives: true
    });
    tempId = temp.data.id;
    const pdf = await drive.files.export(
      { fileId: tempId, mimeType: PDF_MIME },
      { responseType: 'arraybuffer' }
    );
    const out = await uploadFile(Buffer.from(pdf.data), `${baseName}.pdf`, folderId, PDF_MIME);
    return { ok: true, ...out };
  } catch (err) {
    console.error(`PDF conversion failed for ${baseName}:`, err.message);
    return { ok: false, error: err.message };
  } finally {
    if (tempId) {
      try { await drive.files.delete({ fileId: tempId, supportsAllDrives: true }); }
      catch (e) { console.error('Could not remove the temporary Google Doc:', e.message); }
    }
  }
}

// Reads a Drive file back as a Buffer, whatever Drive is holding.
// A .docx uploaded without conversion comes back as-is; one that Drive turned
// into a Google Doc has to be exported back to .docx instead.
async function downloadDoc(fileId) {
  const meta = await getDrive().files.get({
    fileId, fields: 'id,name,mimeType', supportsAllDrives: true
  });
  if (meta.data.mimeType === GDOC_MIME) {
    const res = await getDrive().files.export(
      { fileId, mimeType: DOCX_MIME },
      { responseType: 'arraybuffer' }
    );
    return Buffer.from(res.data);
  }
  return downloadFile(fileId);
}

// Reads a Drive file back as a Buffer (used for the invoice template).
async function downloadFile(fileId) {
  const res = await getDrive().files.get(
    { fileId, alt: 'media', supportsAllDrives: true },
    { responseType: 'arraybuffer' }
  );
  return Buffer.from(res.data);
}

module.exports = {
  SCOPES, DOCX_MIME, PDF_MIME, GDOC_MIME,
  authMode, getAuth, getSheets, getDrive,
  uploadFile, replaceFile, docxToPdf, downloadFile, downloadDoc, viewUrl
};
