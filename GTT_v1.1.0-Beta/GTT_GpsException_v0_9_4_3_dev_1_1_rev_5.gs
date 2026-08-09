/**
 * GRAND TIME TRACKER
 * File: GTT_GpsException_v0_9_4_3_dev_1_1.gs
 * Build: GTT v0.9.4.3-dev.1.1 (rev-5)
 *
 * Migration-ready boundary:
 * - Frontend calls API functions only.
 * - Sheet access is isolated in repository helpers in this file.
 * - Business validation is isolated from persistence.
 */

const GTT_DEV11 = Object.freeze({
  TZ: 'Asia/Makassar',
  MASTER_SA: 'MASTER_SA',
  MASTER_SETTING: 'MASTER_SETTING',
  ABSENSI: 'ABSENSI_HARIAN',
  LOG: 'LOG_GPS_EXCEPTION',
  MAX_GPS_ACCURACY_M: 150,
  MIN_NOTE_LENGTH: 5,
  MAX_NOTE_LENGTH: 200,
  LOCK_TIMEOUT_MS: 15000,
  PROFILE_FOLDER_NAME: 'GTT_PROFILE_PHOTOS',
  PROFILE_MAX_BYTES: 900000
});

function apiPingServerGttDev11() {
  return {
    success: true,
    data: {
      timestamp: Date.now(),
      zonaWaktu: GTT_DEV11.TZ,
      version: 'GTT v0.9.4.3-dev.1.1 (rev-5)'
    }
  };
}

function apiGetGpsExceptionApproversGttDev11(saPin) {
  try {
    const pin = normalizePinDev11_(saPin);

    if (!pin) {
      return failDev11_('PIN SA tidak valid.');
    }

    const users = readMasterSaDev11_();

    const sa = users.find(function (user) {
      return user.pin === pin && user.active;
    });

    if (!sa) {
      return failDev11_('Data SA aktif tidak ditemukan.');
    }

    const allowed = users.filter(function (user) {
      const sameOutlet = user.outlet === sa.outlet;
      const differentUser = user.pin !== sa.pin;

      return user.active &&
        sameOutlet &&
        differentUser &&
        isApproverDev11_(user);
    });

    const approvers = allowed.map(function (user) {
      return {
        approverId: user.saId,
        nama: user.name,
        jabatan: user.position,
        outlet: user.outlet
      };
    });

    approvers.sort(function (first, second) {
      return first.nama.localeCompare(second.nama, 'id');
    });

    return {
      success: true,
      data: approvers
    };
  } catch (error) {
    console.error('apiGetGpsExceptionApproversGttDev11:', error);
    return failDev11_('Daftar SL/SPV gagal dibaca.');
  }
}

function apiProcessGpsExceptionAttendanceGttDev11(payload) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(GTT_DEV11.LOCK_TIMEOUT_MS);

    const request = validateGpsExceptionRequestDev11_(payload);

    if (!request.success) {
      return request;
    }

    const users = readMasterSaDev11_();

    const sa = users.find(function (user) {
      return user.pin === request.data.saPin && user.active;
    });

    const approver = users.find(function (user) {
      return (
        user.saId === request.data.approverId &&
        user.pin === request.data.approverPin &&
        user.active
      );
    });

    const authorization = validateApproverDev11_(
      sa,
      approver,
      request.data
    );

    if (!authorization.success) {
      writeGpsExceptionLogDev11_(
        request.data,
        sa,
        approver,
        'DITOLAK',
        authorization.message,
        ''
      );

      return authorization;
    }

    let attendanceResult;

    if (request.data.actionCode === 'ABSEN_PULANG') {
      attendanceResult = processGpsExceptionCheckoutDev11_(
        sa,
        approver,
        request.data
      );
    } else {
      attendanceResult = processGpsExceptionCheckinDev11_(
        sa,
        approver,
        request.data
      );
    }

    let attendanceId = '';

    if (
      attendanceResult &&
      attendanceResult.data &&
      attendanceResult.data.attendanceId
    ) {
      attendanceId = attendanceResult.data.attendanceId;
    }

    writeGpsExceptionLogDev11_(
      request.data,
      sa,
      approver,
      attendanceResult.success ? 'DISETUJUI' : 'GAGAL',
      attendanceResult.message,
      attendanceId
    );

    if (attendanceResult.success) {
      attendanceResult.data = attendanceResult.data || {};
      attendanceResult.data.approverName = approver.name;
      attendanceResult.data.approverPosition = approver.position;
    }

    return attendanceResult;
  } catch (error) {
    console.error(
      'apiProcessGpsExceptionAttendanceGttDev11:',
      error
    );

    if (error && error.message === 'Timeout') {
      return failDev11_('Sistem sedang sibuk. Silakan coba lagi.');
    }

    return failDev11_('Otorisasi GPS gagal diproses.');
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      console.warn(
        'apiProcessGpsExceptionAttendanceGttDev11: lock release gagal',
        releaseError
      );
    }
  }
}

function validateGpsExceptionRequestDev11_(payload) {
  if (!payload || typeof payload !== 'object') {
    return failDev11_('Data permintaan tidak valid.');
  }

  const saPin = normalizePinDev11_(payload.saPin);
  const approverId = String(payload.approverId || '').trim();
  const approverPin = normalizePinDev11_(payload.approverPin);
  const actionCode = String(payload.actionCode || '').trim().toUpperCase();
  const note = String(payload.note || '').trim();
  const gps = payload.gps || {};
  const latitude = Number(gps.latitude);
  const longitude = Number(gps.longitude);
  const accuracy = Number(gps.accuracy);
  const attempts = Math.floor(Number(payload.attempts));

  if (!saPin) {
    return failDev11_('PIN SA tidak valid.');
  }
  if (!approverId) {
    return failDev11_('SL/SPV pemberi otorisasi belum dipilih.');
  }
  if (!approverPin) {
    return failDev11_('PIN SL/SPV tidak valid.');
  }
  if (saPin === approverPin) {
    return failDev11_('SA tidak boleh mengotorisasi dirinya sendiri.');
  }
  if (actionCode !== 'ABSEN_MASUK' && actionCode !== 'ABSEN_PULANG') {
    return failDev11_('Jenis absensi tidak didukung.');
  }
  if (note.length < GTT_DEV11.MIN_NOTE_LENGTH) {
    return failDev11_('Catatan minimal 5 karakter.');
  }
  if (note.length > GTT_DEV11.MAX_NOTE_LENGTH) {
    return failDev11_('Catatan maksimal 200 karakter.');
  }

  const latitudeInvalid = !Number.isFinite(latitude) || latitude < -90 || latitude > 90;
  if (latitudeInvalid) {
    return failDev11_('Latitude GPS tidak valid.');
  }

  const longitudeInvalid = !Number.isFinite(longitude) || longitude < -180 || longitude > 180;
  if (longitudeInvalid) {
    return failDev11_('Longitude GPS tidak valid.');
  }

  const isAccuracyInvalid = !Number.isFinite(accuracy);
  const isGpsStillAcceptable = accuracy <= GTT_DEV11.MAX_GPS_ACCURACY_M;

  if (isAccuracyInvalid || isGpsStillAcceptable) {
    return failDev11_(
      'Otorisasi hanya tersedia jika akurasi GPS masih di atas 150 meter.'
    );
  }

  if (!Number.isFinite(attempts) || attempts < 5) {
    return failDev11_(
      'Otorisasi hanya tersedia setelah 5 percobaan GPS.'
    );
  }

  return {
    success: true,
    data: {
      saPin: saPin,
      approverId: approverId,
      approverPin: approverPin,
      actionCode: actionCode,
      note: note,
      gps: {
        latitude: latitude,
        longitude: longitude,
        accuracy: accuracy,
        timestamp: Number(gps.timestamp) || Date.now()
      },
      attempts: attempts,
      clientVersion: String(payload.clientVersion || '').trim()
    }
  };
}

function validateApproverDev11_(sa, approver, request) {
  if (!sa) return failDev11_('Data SA aktif tidak ditemukan.');
  if (!approver) return failDev11_('PIN SL/SPV tidak valid atau akun tidak aktif.');
  if (sa.outlet !== approver.outlet) return failDev11_('SL/SPV harus berasal dari outlet yang sama.');
  if (!isApproverDev11_(approver)) return failDev11_('Pengguna tersebut tidak memiliki otoritas SL/SPV.');
  return { success: true };
}

function isApproverDev11_(user) {
  const positionValue = (user && user.position) ? user.position : '';
  const roleValue = (user && user.role) ? user.role : '';
  const position = String(positionValue).trim().toUpperCase();
  const role = String(roleValue).trim().toUpperCase();
  const positionAllowed = position === 'SL' || position === 'SPV';
  const roleAllowed = role === 'SL' || role === 'SPV' || role === 'ADMIN';

  return positionAllowed || roleAllowed;
}

function processGpsExceptionCheckinDev11_(sa, approver, request) {
  const sheet = getRequiredSheetDev11_(GTT_DEV11.ABSENSI);
  const table = readTableDev11_(sheet);
  const now = new Date();
  const dateKey = Utilities.formatDate(now, GTT_DEV11.TZ, 'dd/MM/yyyy');
  const existing = findAttendanceRowDev11_(table, sa.pin, dateKey);

  if (existing) {
    const existingCheckin = valueByAliasesDev11_(
      existing.row,
      table.map,
      ['JAM MASUK', 'CHECK IN']
    );

    if (existingCheckin) {
      return failDev11_('SA sudah melakukan absen masuk hari ini.');
    }
  }

  const decision = calculateCheckinDecisionDev11_(now);
  let attendanceId = buildIdDev11_('ABS', sa.pin);

  if (existing) {
    const existingId = valueByAliasesDev11_(
      existing.row,
      table.map,
      ['ID']
    );
    attendanceId = String(existingId || '');
  }

  const values = {
    'ID': attendanceId,
    'TANGGAL': dateKey,
    'NAMA SA': sa.name,
    'PIN': sa.pin,
    'OUTLET': sa.outlet,
    'JAM MASUK': Utilities.formatDate(now, GTT_DEV11.TZ, 'HH:mm:ss'),
    'STATUS JAM MASUK': decision.statusJamMasuk,
    'TERLAMBAT MENIT': decision.terlambatMenit,
    'STATUS KEHADIRAN': decision.statusKehadiran,
    'KETERANGAN':
      'ABSEN PENGECUALIAN GPS — DIOTORISASI ' +
      approver.name +
      ' (' +
      approver.position +
      '). ' +
      request.note
  };

  if (existing) {
    updateRowByHeaderDev11_(sheet, table, existing.rowNumber, values);
  } else {
    appendRowByHeaderDev11_(sheet, table, values);
  }

  return {
    success: true,
    message: 'Absen masuk pengecualian GPS berhasil dicatat.',
    data: {
      attendanceId: attendanceId,
      exceptionType: 'GPS',
      actionCode: request.actionCode
    }
  };
}

function processGpsExceptionCheckoutDev11_(sa, approver, request) {
  const sheet = getRequiredSheetDev11_(GTT_DEV11.ABSENSI);
  const table = readTableDev11_(sheet);
  const now = new Date();
  const dateKey = Utilities.formatDate(now, GTT_DEV11.TZ, 'dd/MM/yyyy');
  const existing = findAttendanceRowDev11_(table, sa.pin, dateKey);

  if (!existing) {
    return failDev11_('Absen masuk hari ini tidak ditemukan.');
  }

  const checkoutValue = valueByAliasesDev11_(
    existing.row,
    table.map,
    ['JAM PULANG', 'CHECK OUT']
  );

  if (checkoutValue) {
    return failDev11_('SA sudah melakukan absen pulang hari ini.');
  }

  const decision = calculateCheckoutDecisionDev11_(now);
  const currentNoteValue = valueByAliasesDev11_(
    existing.row,
    table.map,
    ['KETERANGAN']
  );
  const currentNote = String(currentNoteValue || '').trim();
  const currentStatus = valueByAliasesDev11_(
    existing.row,
    table.map,
    ['STATUS KEHADIRAN']
  );
  const finalStatus = decision.statusKehadiran ? decision.statusKehadiran : currentStatus;

  const values = {
    'JAM PULANG': Utilities.formatDate(now, GTT_DEV11.TZ, 'HH:mm:ss'),
    'STATUS JAM PULANG': decision.statusJamPulang,
    'STATUS KEHADIRAN': finalStatus || '',
    'KETERANGAN': [
      currentNote,
      'ABSEN PULANG PENGECUALIAN GPS — DIOTORISASI ' +
        approver.name +
        ' (' +
        approver.position +
        '). ' +
        request.note
    ].filter(Boolean).join(' | ')
  };

  updateRowByHeaderDev11_(sheet, table, existing.rowNumber, values);

  const attendanceIdValue = valueByAliasesDev11_(
    existing.row,
    table.map,
    ['ID']
  );

  return {
    success: true,
    message: 'Absen pulang pengecualian GPS berhasil dicatat.',
    data: {
      attendanceId: String(attendanceIdValue || ''),
      exceptionType: 'GPS',
      actionCode: request.actionCode
    }
  };
}

function calculateCheckinDecisionDev11_(now) {
  const normal = settingTimeMinutesDev11_('ABSENSI - JAM MASUK NORMAL', '08:15');
  const tolerance = settingNumberDev11_('ABSENSI - TOLERANSI TERLAMBAT (MENIT)', 5);
  const halfDay = settingTimeMinutesDev11_('ABSENSI - JAM MASUK 1/2 HARI', '09:01');
  const current = timeMinutesDev11_(now);
  const lateStart = normal + tolerance;

  if (current >= halfDay) {
    return { statusJamMasuk: '1/2 HARI', terlambatMenit: 0, statusKehadiran: 'HADIR 1/2 HARI' };
  }
  if (current > lateStart) {
    return {
      statusJamMasuk: 'TERLAMBAT',
      terlambatMenit: Math.max(1, current - normal),
      statusKehadiran: 'HADIR'
    };
  }
  return { statusJamMasuk: 'NORMAL', terlambatMenit: 0, statusKehadiran: 'HADIR' };
}

function calculateCheckoutDecisionDev11_(now) {
  const halfDayMinimum = settingTimeMinutesDev11_(
    'ABSENSI - BATAS PULANG MINIMAL 1/2 HARI', '14:01'
  );
  const normalCheckout = settingTimeMinutesDev11_('ABSENSI - JAM PULANG NORMAL', '20:30');
  const current = timeMinutesDev11_(now);

  if (current < halfDayMinimum) {
    return { statusJamPulang: 'PULANG TERLALU AWAL', statusKehadiran: 'ALPA' };
  }
  if (current < normalCheckout) {
    return { statusJamPulang: 'PULANG 1/2 HARI', statusKehadiran: 'HADIR 1/2 HARI' };
  }
  return { statusJamPulang: 'NORMAL', statusKehadiran: '' };
}

function writeGpsExceptionLogDev11_(
  request,
  sa,
  approver,
  status,
  message,
  attendanceId
) {
  try {
    const sheet = ensureGpsExceptionLogSheetDev11_();
    const saName = (sa && sa.name) ? sa.name : '';
    const saOutlet = (sa && sa.outlet) ? sa.outlet : '';
    const approverName = (approver && approver.name) ? approver.name : '';
    const approverPosition = (approver && approver.position) ? approver.position : '';

    sheet.appendRow([
      buildIdDev11_('GPSX', request.saPin),
      new Date(),
      Utilities.formatDate(new Date(), GTT_DEV11.TZ, 'dd/MM/yyyy'),
      request.actionCode,
      request.saPin,
      saName,
      saOutlet,
      request.gps.latitude,
      request.gps.longitude,
      request.gps.accuracy,
      request.attempts,
      request.approverPin,
      approverName,
      approverPosition,
      request.note,
      status,
      message,
      attendanceId || '',
      request.clientVersion || ''
    ]);
  } catch (error) {
    console.error('writeGpsExceptionLogDev11_:', error);
  }
}

function ensureGpsExceptionLogSheetDev11_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(GTT_DEV11.LOG);

  if (!logSheet) {
    logSheet = ss.insertSheet(GTT_DEV11.LOG);

    logSheet.getRange(1, 1, 1, 19).setValues([[
      'ID EXCEPTION',
      'WAKTU SERVER',
      'TANGGAL',
      'AKSI',
      'PIN SA',
      'NAMA SA',
      'OUTLET',
      'LATITUDE',
      'LONGITUDE',
      'AKURASI GPS (M)',
      'JUMLAH PERCOBAAN',
      'PIN APPROVER',
      'NAMA APPROVER',
      'JABATAN APPROVER',
      'CATATAN',
      'STATUS',
      'PESAN',
      'ID ABSENSI',
      'VERSI CLIENT'
    ]]);

    logSheet.setFrozenRows(1);
  }

  return logSheet;
}

function readMasterSaDev11_() {
  const sheet = getRequiredSheetDev11_(GTT_DEV11.MASTER_SA);
  const table = readTableDev11_(sheet);
  return table.rows.map(row => ({
    pin: normalizePinDev11_(valueByAliasesDev11_(row, table.map, ['PIN'])),
    name: String(valueByAliasesDev11_(row, table.map, ['NAMA SA', 'NAMA']) || '').trim(),
    outlet: String(valueByAliasesDev11_(row, table.map, ['OUTLET']) || '').trim().toUpperCase(),
    status: String(valueByAliasesDev11_(row, table.map, ['STATUS']) || '').trim().toUpperCase(),
    saId: String(valueByAliasesDev11_(row, table.map, ['SA_ID', 'SA ID']) || '').trim(),
    position: String(valueByAliasesDev11_(row, table.map, ['JABATAN']) || '').trim().toUpperCase(),
    role: String(valueByAliasesDev11_(row, table.map, ['ROLE']) || '').trim().toUpperCase(),
    photoFileId: String(valueByAliasesDev11_(row, table.map, ['PHOTO_FILE_ID']) || '').trim(),
    active: String(valueByAliasesDev11_(row, table.map, ['STATUS']) || '').trim().toUpperCase() === 'AKTIF'
  })).filter(user => user.pin);
}

function getRequiredSheetDev11_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`Sheet ${name} tidak ditemukan.`);
  return sheet;
}

function readTableDev11_(sheet) {
  if (!sheet) {
    throw new Error('Sheet tidak tersedia.');
  }

  const lastRow = Math.max(1, sheet.getLastRow());
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const dataRange = sheet.getRange(1, 1, lastRow, lastColumn);
  const values = dataRange.getValues();
  const firstRow = Array.isArray(values[0]) ? values[0] : [];

  const headers = firstRow.map(function (value) {
    return normalizeHeaderDev11_(value);
  });

  const map = {};

  headers.forEach(function (header, headerIndex) {
    const isNewHeader = header &&
      Object.prototype.hasOwnProperty.call(map, header) === false;

    if (isNewHeader) {
      map[header] = headerIndex;
    }
  });

  return {
    headers: headers,
    map: map,
    rows: values.slice(1),
    rawValues: values
  };
}

function findAttendanceRowDev11_(table, pin, dateKey) {
  let rowIndex = table.rows.length - 1;

  while (rowIndex >= 0) {
    const row = table.rows[rowIndex];
    const rowPin = normalizePinDev11_(
      valueByAliasesDev11_(row, table.map, ['PIN'])
    );
    const rowDate = normalizeDateKeyDev11_(
      valueByAliasesDev11_(row, table.map, ['TANGGAL'])
    );

    if (rowPin === pin && rowDate === dateKey) {
      return {
        row: row,
        rowNumber: rowIndex + 2
      };
    }

    rowIndex = rowIndex - 1;
  }

  return null;
}

function normalizeDateKeyDev11_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, GTT_DEV11.TZ, 'dd/MM/yyyy');
  }
  const text = String(value || '').trim();
  const dmy = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (dmy) return `${dmy[1].padStart(2, '0')}/${dmy[2].padStart(2, '0')}/${dmy[3]}`;
  const ymd = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (ymd) return `${ymd[3].padStart(2, '0')}/${ymd[2].padStart(2, '0')}/${ymd[1]}`;
  return text;
}

function appendRowByHeaderDev11_(sheet, table, valuesByHeader) {
  const row = Array(table.headers.length).fill('');
  Object.keys(valuesByHeader).forEach(header => {
    const index = findHeaderIndexDev11_(table.map, [header]);
    if (index !== null) row[index] = valuesByHeader[header];
  });
  sheet.appendRow(row);
}

function updateRowByHeaderDev11_(sheet, table, rowNumber, valuesByHeader) {
  Object.keys(valuesByHeader).forEach(header => {
    const index = findHeaderIndexDev11_(table.map, [header]);
    if (index !== null) sheet.getRange(rowNumber, index + 1).setValue(valuesByHeader[header]);
  });
}

function valueByAliasesDev11_(row, map, aliases) {
  const index = findHeaderIndexDev11_(map, aliases);
  return index === null ? '' : row[index];
}

function findHeaderIndexDev11_(map, aliases) {
  if (!map || typeof map !== 'object') {
    return null;
  }

  if (!Array.isArray(aliases)) {
    return null;
  }

  let aliasIndex = 0;

  while (aliasIndex < aliases.length) {
    const key = normalizeHeaderDev11_(
      aliases[aliasIndex]
    );

    if (
      Object.prototype.hasOwnProperty.call(map, key)
    ) {
      return map[key];
    }

    aliasIndex = aliasIndex + 1;
  }

  return null;
}

function normalizeHeaderDev11_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function normalizePinDev11_(value) {
  const text = String(value == null ? '' : value).trim();
  return /^\d{4}$/.test(text) ? text : '';
}

function buildIdDev11_(prefix, pin) {
  return [
    prefix,
    Utilities.formatDate(new Date(), GTT_DEV11.TZ, 'yyyyMMddHHmmss'),
    pin,
    Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase()
  ].join('-');
}

function settingNumberDev11_(parameter, fallback) {
  const raw = settingValueDev11_(parameter, fallback);
  const number = Number(raw);
  return Number.isFinite(number) ? number : fallback;
}

function settingTimeMinutesDev11_(parameter, fallback) {
  const raw = String(settingValueDev11_(parameter, fallback) || fallback).trim();
  const match = raw.match(/^(\d{1,2})[:.](\d{2})$/);
  if (!match) return timeStringToMinutesDev11_(fallback);
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    return timeStringToMinutesDev11_(fallback);
  }
  return hour * 60 + minute;
}

function settingValueDev11_(parameter, fallback) {
  try {
    const sheet = getRequiredSheetDev11_(GTT_DEV11.MASTER_SETTING);
    const table = readTableDev11_(sheet);
    for (let i = 0; i < table.rows.length; i += 1) {
      const key = String(valueByAliasesDev11_(table.rows[i], table.map, ['PARAMETER']) || '').trim();
      if (key.toUpperCase() === String(parameter).toUpperCase()) {
        const value = valueByAliasesDev11_(table.rows[i], table.map, ['NILAI']);
        return value === '' || value == null ? fallback : value;
      }
    }
  } catch (error) {
    console.warn('settingValueDev11_:', error);
  }
  return fallback;
}

function timeMinutesDev11_(date) {
  const hour = Number(Utilities.formatDate(date, GTT_DEV11.TZ, 'H'));
  const minute = Number(Utilities.formatDate(date, GTT_DEV11.TZ, 'm'));
  return hour * 60 + minute;
}

function timeStringToMinutesDev11_(text) {
  const match = String(text || '').match(/^(\d{1,2})[:.](\d{2})$/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : 0;
}

function failDev11_(message) {
  return { success: false, message: String(message || 'Permintaan gagal.') };
}


/* =========================================================
   Profile Photo API — GTT v0.9.4.3-dev.1.1 (rev-5)
   File bytes live in Drive. MASTER_SA stores only metadata.
========================================================== */

function apiGetProfilePhotoGttDev12(saPin) {
  try {
    const pin = normalizePinDev11_(saPin);
    if (!pin) return failDev11_('PIN SA tidak valid.');
    const record = findMasterSaRecordDev12_(pin);
    if (!record) return failDev11_('Data SA tidak ditemukan.');
    const fileId = String(record.photoFileId || '').trim();
    if (!fileId) return { success: true, data: { dataUrl: '', fileId: '' } };

    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > GTT_DEV11.PROFILE_MAX_BYTES) {
      return failDev11_('Foto profil terlalu besar untuk ditampilkan.');
    }
    return {
      success: true,
      data: {
        fileId,
        dataUrl: `data:${blob.getContentType()};base64,${Utilities.base64Encode(bytes)}`
      }
    };
  } catch (error) {
    console.error('apiGetProfilePhotoGttDev12:', error);
    return { success: true, data: { dataUrl: '', fileId: '' } };
  }
}

function apiSaveProfilePhotoGttDev12(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(GTT_DEV11.LOCK_TIMEOUT_MS);
    if (!payload || typeof payload !== 'object') return failDev11_('Data foto tidak valid.');

    const pin = normalizePinDev11_(payload.saPin);
    const mimeType = String(payload.mimeType || '').trim().toLowerCase();
    const base64 = String(payload.base64 || '').replace(/\s/g, '');
    if (!pin) return failDev11_('PIN SA tidak valid.');
    if (!['image/jpeg', 'image/png', 'image/webp'].includes(mimeType)) {
      return failDev11_('Format foto harus JPG, PNG, atau WEBP.');
    }
    if (!base64) return failDev11_('Data foto kosong.');

    const bytes = decodeProfilePhotoBase64Dev12_(base64);

if (!bytes) {
  return failDev11_('Data foto rusak atau tidak valid.');
}
    if (!bytes.length || bytes.length > GTT_DEV11.PROFILE_MAX_BYTES) {
      return failDev11_('Ukuran foto setelah kompresi maksimal 900 KB.');
    }

    const record = findMasterSaRecordDev12_(pin);
    if (!record) return failDev11_('Data SA tidak ditemukan.');

    const folder = getOrCreateProfileFolderDev12_();
    const extension = mimeType === 'image/png' ? 'png' : mimeType === 'image/webp' ? 'webp' : 'jpg';
    const safeSaId = String(record.saId || pin).replace(/[^A-Za-z0-9_]/g, '_');
    const fileName = `${safeSaId}_${Utilities.formatDate(new Date(), GTT_DEV11.TZ, 'yyyyMMdd_HHmmss')}.${extension}`;
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const file = folder.createFile(blob);

    const previousFileId = String(record.photoFileId || '').trim();
    updateMasterSaPhotoMetadataDev12_(record.rowNumber, file.getId(), new Date());

    if (previousFileId && previousFileId !== file.getId()) {
      try {
        DriveApp.getFileById(previousFileId).setTrashed(true);
      } catch (trashError) {
        console.warn('Foto lama tidak dapat dipindahkan ke sampah:', trashError);
      }
    }

    return {
      success: true,
      message: 'Foto profil berhasil disimpan.',
      data: {
        fileId: file.getId(),
        dataUrl: `data:${mimeType};base64,${Utilities.base64Encode(bytes)}`
      }
    };
  } catch (error) {
    console.error('apiSaveProfilePhotoGttDev12:', error);
    return failDev11_('Foto profil gagal disimpan.');
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      console.warn('apiSaveProfilePhotoGttDev12: lock release gagal', releaseError);
    }
  }
}

function apiDeleteProfilePhotoGttDev12(saPin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(GTT_DEV11.LOCK_TIMEOUT_MS);
    const pin = normalizePinDev11_(saPin);
    if (!pin) return failDev11_('PIN SA tidak valid.');
    const record = findMasterSaRecordDev12_(pin);
    if (!record) return failDev11_('Data SA tidak ditemukan.');

    const fileId = String(record.photoFileId || '').trim();
    updateMasterSaPhotoMetadataDev12_(record.rowNumber, '', new Date());
    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (trashError) {
        console.warn('Foto tidak dapat dipindahkan ke sampah:', trashError);
      }
    }
    return { success: true, message: 'Foto profil berhasil dihapus.' };
  } catch (error) {
    console.error('apiDeleteProfilePhotoGttDev12:', error);
    return failDev11_('Foto profil gagal dihapus.');
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      console.warn('apiDeleteProfilePhotoGttDev12: lock release gagal', releaseError);
    }
  }
}

function findMasterSaRecordDev12_(pin) {
  const sheet = getRequiredSheetDev11_(GTT_DEV11.MASTER_SA);

  ensureMasterSaPhotoColumnsDev12_(sheet);

  const table = readTableDev11_(sheet);
  let rowIndex = 0;

  while (rowIndex < table.rows.length) {
    const row = table.rows[rowIndex];
    const rowPin = normalizePinDev11_(
      valueByAliasesDev11_(
        row,
        table.map,
        ['PIN']
      )
    );

    if (rowPin === pin) {
      return {
        sheet: sheet,
        table: table,
        row: row,
        rowNumber: rowIndex + 2,
        saId: String(
          valueByAliasesDev11_(
            row,
            table.map,
            ['SA_ID', 'SA ID']
          ) || ''
        ).trim(),
        photoFileId: String(
          valueByAliasesDev11_(
            row,
            table.map,
            ['PHOTO_FILE_ID']
          ) || ''
        ).trim()
      };
    }

    rowIndex = rowIndex + 1;
  }

  return null;
}

function ensureMasterSaPhotoColumnsDev12_(sheet) {
  if (!sheet) {
    throw new Error('MASTER_SA tidak tersedia.');
  }

  const requiredHeaders = [
    'PHOTO_FILE_ID',
    'PHOTO_UPDATED_AT'
  ];

  const lastColumn = Math.max(1, sheet.getLastColumn());
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  const headerValues = headerRange.getValues();
  const firstRow = Array.isArray(headerValues[0]) ? headerValues[0] : [];

  const headers = firstRow.map(function (value) {
    return normalizeHeaderDev11_(value);
  });

  requiredHeaders.forEach(function (header) {
    if (headers.indexOf(header) === -1) {
      const newColumn = sheet.getLastColumn() + 1;
      const headerCell = sheet.getRange(1, newColumn);

      headerCell.setValue(header);
      headers.push(header);
    }
  });
}

function updateMasterSaPhotoMetadataDev12_(rowNumber, fileId, updatedAt) {
  const sheet = getRequiredSheetDev11_(GTT_DEV11.MASTER_SA);
  ensureMasterSaPhotoColumnsDev12_(sheet);
  const table = readTableDev11_(sheet);
  updateRowByHeaderDev11_(sheet, table, rowNumber, {
    'PHOTO_FILE_ID': fileId,
    'PHOTO_UPDATED_AT': updatedAt
  });
}

function getOrCreateProfileFolderDev12_() {
  const folderName = GTT_DEV11.PROFILE_FOLDER_NAME;
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}

// tempel helper di bawah sini
function decodeProfilePhotoBase64Dev12_(base64) {
  try {
    return Utilities.base64Decode(base64);
  } catch (decodeError) {
    console.warn(
      'decodeProfilePhotoBase64Dev12_: Base64 tidak valid',
      decodeError
    );

    return null;
  }
}

