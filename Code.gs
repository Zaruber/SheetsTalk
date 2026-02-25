const MESSAGES_SHEET_NAME = 'Messages';
const MESSAGES_HEADER = [
  'id',
  'room',
  'createdAt',
  'sender',
  'kind',
  'mime',
  'ivB64',
  'storage',
  'pointer',
  'byteLength',
];

const INLINE_PAYLOAD_LIMIT = 45000;
const MAX_PAYLOAD_LENGTH = 2400000;
const MAX_FETCH_LIMIT = 120;
const PAYLOAD_FOLDER_NAME = 'SheetsTalkEncryptedPayloads';

function doGet(e) {
  if (isApiRequest_(e)) {
    return handleApiGet_(e);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SheetsTalk')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  if (!isApiRequest_(e)) {
    return respondJson_(
      {
        ok: false,
        error: 'API endpoint expects ?api=1',
      },
      null
    );
  }

  return handleApiPost_(e);
}

function isApiRequest_(e) {
  return Boolean(e && e.parameter && String(e.parameter.api || '') === '1');
}

function handleApiGet_(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').toLowerCase();
  const callback = String((e && e.parameter && e.parameter.callback) || '');

  try {
    if (action === 'config') {
      return respondJson_(
        {
          ok: true,
          data: getConfig(),
        },
        callback
      );
    }

    if (action === 'fetch') {
      const room = String((e.parameter && e.parameter.room) || '');
      const afterCursor = Number((e.parameter && e.parameter.afterCursor) || 1);
      const limit = Number((e.parameter && e.parameter.limit) || 50);

      return respondJson_(
        {
          ok: true,
          data: fetchMessages(room, afterCursor, limit),
        },
        callback
      );
    }

    if (action === 'send') {
      const payloadB64 = String((e.parameter && e.parameter.payloadB64) || '');
      if (!payloadB64) {
        throw new Error('payloadB64 is required for GET send action.');
      }

      const payload = decodeWebSafeJson_(payloadB64);
      return respondJson_(
        {
          ok: true,
          data: sendMessage(payload),
        },
        callback
      );
    }

    return respondJson_(
      {
        ok: false,
        error: 'Unknown action.',
      },
      callback
    );
  } catch (error) {
    return respondJson_(
      {
        ok: false,
        error: String(error && error.message ? error.message : error),
      },
      callback
    );
  }
}

function handleApiPost_(e) {
  try {
    const raw = String((e && e.postData && e.postData.contents) || '');
    if (!raw) {
      throw new Error('Missing POST body.');
    }

    const parsed = JSON.parse(raw);
    const action = String(parsed.action || '').toLowerCase();

    if (action !== 'send') {
      throw new Error('POST action must be send.');
    }

    const payload = parsed.payload;
    const result = sendMessage(payload);

    return respondJson_({
      ok: true,
      data: result,
    });
  } catch (error) {
    return respondJson_({
      ok: false,
      error: String(error && error.message ? error.message : error),
    });
  }
}

function decodeWebSafeJson_(payloadB64) {
  const bytes = Utilities.base64DecodeWebSafe(String(payloadB64 || ''));
  const text = Utilities.newBlob(bytes).getDataAsString('UTF-8');
  return JSON.parse(text);
}

function respondJson_(obj, callbackName) {
  const payload = JSON.stringify(obj);
  const callback = String(callbackName || '').trim();

  if (isSafeCallbackName_(callback)) {
    return ContentService.createTextOutput(callback + '(' + payload + ');').setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
  }

  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

function isSafeCallbackName_(value) {
  return /^[A-Za-z_$][0-9A-Za-z_$\.]{0,120}$/.test(String(value || ''));
}

function getConfig() {
  return {
    pollIntervalMs: 450,
    maxFetch: MAX_FETCH_LIMIT,
    inlinePayloadLimit: INLINE_PAYLOAD_LIMIT,
    webAppUrl: resolveWebAppUrl_(),
  };
}

function resolveWebAppUrl_() {
  try {
    const url = ScriptApp.getService().getUrl();
    return url ? String(url) : '';
  } catch (error) {
    return '';
  }
}

function sendMessage(input) {
  if (!input || typeof input !== 'object') {
    throw new Error('Invalid message payload.');
  }

  const room = normalizeRoom_(input.room);
  const sender = normalizeSender_(input.sender);
  const kind = normalizeKind_(input.kind);
  const mime = normalizeMime_(input.mime, kind);
  const ivB64 = normalizeBase64_(input.ivB64, 'iv');
  const payloadB64 = normalizeBase64_(input.payloadB64, 'payload');

  if (payloadB64.length > MAX_PAYLOAD_LENGTH) {
    throw new Error('Payload is too large.');
  }

  const byteLength = normalizeByteLength_(input.byteLength, payloadB64);
  const createdAt = new Date().toISOString();
  const id = Utilities.getUuid();

  let storage = 'INLINE';
  let pointer = payloadB64;

  if (payloadB64.length > INLINE_PAYLOAD_LIMIT) {
    storage = 'DRIVE';
    pointer = storePayload_(payloadB64, id);
  }

  const sheet = ensureMessagesSheet_();
  sheet.appendRow([
    id,
    room,
    createdAt,
    sender,
    kind,
    mime,
    ivB64,
    storage,
    pointer,
    byteLength,
  ]);

  return {
    id: id,
    cursor: sheet.getLastRow(),
    createdAt: createdAt,
    storage: storage,
  };
}

function fetchMessages(room, afterCursor, limit) {
  const safeRoom = normalizeRoom_(room);
  const safeAfterCursor = Math.max(1, Number(afterCursor) || 1);
  const safeLimit = Math.min(Math.max(1, Number(limit) || 50), MAX_FETCH_LIMIT);

  const sheet = ensureMessagesSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1 || safeAfterCursor >= lastRow) {
    return {
      messages: [],
      nextCursor: lastRow,
      lastRow: lastRow,
    };
  }

  const startRow = Math.max(2, safeAfterCursor + 1);
  const rowCount = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, rowCount, MESSAGES_HEADER.length).getValues();

  const messages = [];
  let nextCursor = safeAfterCursor;

  for (let i = 0; i < values.length; i += 1) {
    const cursor = startRow + i;
    const row = values[i];
    nextCursor = cursor;

    if (String(row[1]) !== safeRoom) {
      continue;
    }

    const message = {
      cursor: cursor,
      id: String(row[0]),
      room: String(row[1]),
      createdAt: String(row[2]),
      sender: String(row[3]),
      kind: String(row[4]),
      mime: String(row[5]),
      ivB64: String(row[6]),
      storage: String(row[7]),
      byteLength: Number(row[9]) || 0,
      payloadB64: '',
    };

    const pointer = String(row[8]);
    if (message.storage === 'DRIVE') {
      try {
        message.payloadB64 = readPayload_(pointer);
      } catch (error) {
        message.payloadB64 = '';
      }
    } else {
      message.payloadB64 = pointer;
    }

    messages.push(message);

    if (messages.length >= safeLimit) {
      break;
    }
  }

  return {
    messages: messages,
    nextCursor: nextCursor,
    lastRow: lastRow,
  };
}

function ensureMessagesSheet_() {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(MESSAGES_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(MESSAGES_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MESSAGES_HEADER);
  }

  const header = sheet.getRange(1, 1, 1, MESSAGES_HEADER.length).getValues()[0];
  const headerMismatch = MESSAGES_HEADER.some(function (name, idx) {
    return String(header[idx]) !== name;
  });

  if (headerMismatch) {
    throw new Error('Messages sheet header does not match expected format.');
  }

  return sheet;
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const configuredId = props.getProperty('SHEET_ID');
  if (configuredId) {
    return SpreadsheetApp.openById(configuredId);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  throw new Error('Spreadsheet is not configured. Set Script Property SHEET_ID or bind this script to a Sheet.');
}

function storePayload_(payloadB64, messageId) {
  const folder = getPayloadFolder_();
  const blob = Utilities.newBlob(payloadB64, 'text/plain', messageId + '.payload.txt');
  const file = folder.createFile(blob);
  return file.getId();
}

function readPayload_(fileId) {
  if (!fileId) {
    return '';
  }

  const file = DriveApp.getFileById(fileId);
  return file.getBlob().getDataAsString();
}

function getPayloadFolder_() {
  const existing = DriveApp.getFoldersByName(PAYLOAD_FOLDER_NAME);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(PAYLOAD_FOLDER_NAME);
}

function normalizeRoom_(value) {
  const room = String(value || '').trim().toLowerCase();
  if (!/^[a-z0-9_-]{3,64}$/.test(room)) {
    throw new Error('Room must be 3-64 characters: a-z, 0-9, _ or -');
  }
  return room;
}

function normalizeSender_(value) {
  const sender = String(value || '').trim();
  if (sender.length < 1 || sender.length > 40) {
    throw new Error('Sender must be 1-40 characters.');
  }
  return sender;
}

function normalizeKind_(value) {
  const kind = String(value || '').toLowerCase();
  if (kind !== 'text' && kind !== 'audio' && kind !== 'event' && kind !== 'receipt') {
    throw new Error('Kind must be text, audio, event or receipt.');
  }
  return kind;
}

function normalizeMime_(value, kind) {
  const fallback = kind === 'audio' ? 'audio/webm' : 'text/plain;charset=utf-8';
  const mime = String(value || fallback).trim();
  if (mime.length < 3 || mime.length > 120) {
    throw new Error('Invalid mime value.');
  }
  return mime;
}

function normalizeBase64_(value, fieldName) {
  const text = String(value || '').replace(/\s+/g, '');
  if (!text) {
    throw new Error('Missing base64 value for ' + fieldName + '.');
  }
  if (!/^[A-Za-z0-9+/=]+$/.test(text)) {
    throw new Error('Invalid base64 value for ' + fieldName + '.');
  }
  return text;
}

function normalizeByteLength_(value, payloadB64) {
  const candidate = Number(value);
  if (Number.isFinite(candidate) && candidate > 0 && candidate <= MAX_PAYLOAD_LENGTH) {
    return Math.floor(candidate);
  }

  const estimated = Math.max(1, Math.floor((payloadB64.length * 3) / 4));
  return estimated;
}
