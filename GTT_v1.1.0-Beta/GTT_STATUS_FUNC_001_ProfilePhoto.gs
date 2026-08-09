/**
 * ============================================================
 * GTT V1.1.0-ALPHA (STATUS_FUNC_001)
 * Modul : Foto Profil SA
 * Scope : Ambil Foto -> Drive -> Metadata Sheet -> Hero Beranda
 *
 * PROTECTED / TIDAK DIUBAH:
 * - Logika absensi
 * - Break
 * - GPS
 * - Navigasi
 * - Layout/CSS Hero Beranda
 * - Alignment nama pengguna
 *
 * Folder foto Drive:
 * GTT_SA_PROFILE_PHOTOS
 * ID: 1yTDBLl1rP5XasLgY_6HfVMDvR-Uub9hj
 * ============================================================
 */

const GTT_STATUS_FUNC_001 = Object.freeze({
  PROFILE_FOLDER_ID: '1yTDBLl1rP5XasLgY_6HfVMDvR-Uub9hj',
  MASTER_SA: 'MASTER_SA',
  MASTER_FOTO: 'MASTER_FOTO_PROFIL',
  TIMEZONE: 'Asia/Makassar',
  MAX_BYTES: 4 * 1024 * 1024
});

/**
 * API upload foto profil.
 *
 * payload:
 * {
 *   saId: 'GP001',
 *   pin: '9616',
 *   base64Image: 'data:image/jpeg;base64,...',
 *   timestamp: 172...
 * }
 */
function uploadProfilePhotoGttStatusFunc001(payload) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const request = payload || {};
    const saId = String(request.saId || '').trim();
    const pin = String(request.pin || '').replace(/\D/g, '').trim();
    const imageData = String(request.base64Image || '').trim();

    if (!saId) {
      return gttPhotoFail_('PHOTO_SA_ID_EMPTY', 'SA_ID tidak tersedia. Silakan login ulang.');
    }

    if (!/^\d{4}$/.test(pin)) {
      return gttPhotoFail_('PHOTO_PIN_INVALID', 'PIN sesi tidak valid. Silakan login ulang.');
    }

    if (!imageData) {
      return gttPhotoFail_('PHOTO_IMAGE_EMPTY', 'Foto belum dipilih.');
    }

    // Validasi user berdasarkan PIN + SA_ID. User tidak boleh mengunggah foto untuk akun lain.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSa = ss.getSheetByName(GTT_STATUS_FUNC_001.MASTER_SA);

    if (!masterSa) {
      return gttPhotoFail_('PHOTO_MASTER_SA_NOT_FOUND', 'Sheet MASTER_SA tidak ditemukan.');
    }

    const saTable = gttPhotoReadTable_(masterSa);
    const saIdCol = gttPhotoHeaderIndex_(saTable.headers, 'SA_ID');
    const pinCol = gttPhotoHeaderIndex_(saTable.headers, 'PIN');
    const nameCol = gttPhotoHeaderIndex_(saTable.headers, 'NAMA SA');
    const photoIdCol = gttPhotoHeaderIndex_(saTable.headers, 'PHOTO_FILE_ID');
    const photoUpdatedCol = gttPhotoHeaderIndex_(saTable.headers, 'PHOTO_UPDATED_AT');

    if ([saIdCol, pinCol, nameCol, photoIdCol, photoUpdatedCol].some(index => index < 0)) {
      return gttPhotoFail_(
        'PHOTO_MASTER_SA_COLUMNS_INVALID',
        'Kolom foto pada MASTER_SA belum lengkap.'
      );
    }

    let saRowIndex = -1;
    let saRow = null;

    for (let i = 0; i < saTable.rows.length; i += 1) {
      const rowSaId = String(saTable.rows[i][saIdCol] || '').trim();
      const rowPin = String(saTable.rows[i][pinCol] || '').replace(/\D/g, '').trim();

      if (rowSaId === saId && rowPin === pin) {
        saRowIndex = i + 2; // row 1 = header
        saRow = saTable.rows[i];
        break;
      }
    }

    if (!saRow) {
      return gttPhotoFail_(
        'PHOTO_USER_MISMATCH',
        'Data pengguna tidak cocok dengan sesi login.'
      );
    }

    const saName = String(saRow[nameCol] || '').trim();
    const previousFileId = String(saRow[photoIdCol] || '').trim();

    const parsed = gttPhotoParseDataUrl_(imageData);

    if (!parsed.success) {
      return gttPhotoFail_(parsed.code, parsed.message);
    }

    if (parsed.bytes.length > GTT_STATUS_FUNC_001.MAX_BYTES) {
      return gttPhotoFail_(
        'PHOTO_TOO_LARGE',
        'Ukuran foto terlalu besar. Ambil foto ulang.'
      );
    }

    const folder = DriveApp.getFolderById(GTT_STATUS_FUNC_001.PROFILE_FOLDER_ID);
    const now = new Date();
    const stamp = Utilities.formatDate(
      now,
      GTT_STATUS_FUNC_001.TIMEZONE,
      'yyyyMMdd_HHmmss'
    );

    const extension = parsed.mimeType === 'image/png' ? 'png' : 'jpg';
    const safeSaId = saId.replace(/[^A-Za-z0-9_-]/g, '_');
    const fileName = `${safeSaId}_${stamp}.${extension}`;

    const blob = Utilities.newBlob(parsed.bytes, parsed.mimeType, fileName);
    const file = folder.createFile(blob);
    const newFileId = file.getId();

    // 1) Update MASTER_SA: foto aktif user.
    masterSa.getRange(saRowIndex, photoIdCol + 1).setValue(newFileId);
    masterSa.getRange(saRowIndex, photoUpdatedCol + 1).setValue(now);

    // 2) Update MASTER_FOTO_PROFIL: metadata yang dibaca/tersedia untuk AppSheet.
    const masterFoto = ss.getSheetByName(GTT_STATUS_FUNC_001.MASTER_FOTO);

    if (!masterFoto) {
      // Rollback metadata MASTER_SA jika tabel metadata tidak ada.
      masterSa.getRange(saRowIndex, photoIdCol + 1).setValue(previousFileId);
      file.setTrashed(true);
      return gttPhotoFail_(
        'PHOTO_MASTER_FOTO_NOT_FOUND',
        'Sheet MASTER_FOTO_PROFIL tidak ditemukan.'
      );
    }

    const fotoTable = gttPhotoReadTable_(masterFoto);
    const requiredFotoHeaders = [
      'SA_ID',
      'PIN',
      'NAMA SA',
      'PHOTO_FILE_ID',
      'PHOTO_FILE_NAME',
      'PHOTO_UPDATED_AT',
      'STATUS'
    ];

    const fotoIndexes = {};
    for (let i = 0; i < requiredFotoHeaders.length; i += 1) {
      const header = requiredFotoHeaders[i];
      fotoIndexes[header] = gttPhotoHeaderIndex_(fotoTable.headers, header);
      if (fotoIndexes[header] < 0) {
        masterSa.getRange(saRowIndex, photoIdCol + 1).setValue(previousFileId);
        file.setTrashed(true);
        return gttPhotoFail_(
          'PHOTO_MASTER_FOTO_COLUMNS_INVALID',
          `Kolom ${header} tidak ditemukan pada MASTER_FOTO_PROFIL.`
        );
      }
    }

    let fotoRowNumber = -1;

    for (let i = 0; i < fotoTable.rows.length; i += 1) {
      if (
        String(fotoTable.rows[i][fotoIndexes['SA_ID']] || '').trim() === saId
      ) {
        fotoRowNumber = i + 2;
        break;
      }
    }

    const metadata = {
      'SA_ID': saId,
      'PIN': pin,
      'NAMA SA': saName,
      'PHOTO_FILE_ID': newFileId,
      'PHOTO_FILE_NAME': fileName,
      'PHOTO_UPDATED_AT': now,
      'STATUS': 'AKTIF'
    };

    if (fotoRowNumber > 0) {
      Object.keys(metadata).forEach(header => {
        masterFoto
          .getRange(fotoRowNumber, fotoIndexes[header] + 1)
          .setValue(metadata[header]);
      });
    } else {
      const newRow = new Array(Math.max(masterFoto.getLastColumn(), 7)).fill('');
      Object.keys(metadata).forEach(header => {
        newRow[fotoIndexes[header]] = metadata[header];
      });
      masterFoto.appendRow(newRow);
    }

    // Cache foto Beranda harus dibuang agar login/refresh berikutnya membaca foto terbaru.
    try {
      CacheService.getScriptCache().remove('GTT_HOME_PHOTO_' + saId);
    } catch (cacheError) {
      console.warn('STATUS_FUNC_001 cache remove:', cacheError);
    }

    // Bersihkan foto aktif lama hanya setelah file baru + metadata berhasil.
    // Sumber gambar baru tetap file asli; tidak ada penggantian background.
    if (previousFileId && previousFileId !== newFileId) {
      try {
        const oldFile = DriveApp.getFileById(previousFileId);
        const parents = oldFile.getParents();
        let belongsToProfileFolder = false;

        while (parents.hasNext()) {
          if (parents.next().getId() === GTT_STATUS_FUNC_001.PROFILE_FOLDER_ID) {
            belongsToProfileFolder = true;
            break;
          }
        }

        if (belongsToProfileFolder) {
          oldFile.setTrashed(true);
        }
      } catch (oldFileError) {
        console.warn('STATUS_FUNC_001 old photo cleanup:', oldFileError);
      }
    }

    return {
      success: true,
      code: 'STATUS_FUNC_001_PHOTO_SAVED',
      message: 'Foto profil berhasil disimpan.',
      data: {
        saId: saId,
        fileId: newFileId,
        fileName: fileName,
        updatedAt: Utilities.formatDate(
          now,
          GTT_STATUS_FUNC_001.TIMEZONE,
          'dd/MM/yyyy HH:mm:ss'
        ),
        // Data URL hanya dikembalikan untuk update Hero secara instan.
        // Foto sumber tetap tersimpan di Drive.
        dataUrl: `data:${parsed.mimeType};base64,${Utilities.base64Encode(parsed.bytes)}`
      }
    };

  } catch (error) {
    console.error('uploadProfilePhotoGttStatusFunc001:', error);

    return gttPhotoFail_(
      'STATUS_FUNC_001_SERVER_ERROR',
      error && error.message
        ? 'Foto gagal disimpan: ' + error.message
        : 'Foto gagal disimpan.'
    );

  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      // Lock mungkin belum sempat diperoleh.
    }
  }
}


function gttPhotoParseDataUrl_(dataUrl) {
  const match = String(dataUrl || '').match(
    /^data:(image\/(?:jpeg|png));base64,([A-Za-z0-9+/=\s]+)$/i
  );

  if (!match) {
    return {
      success: false,
      code: 'PHOTO_FORMAT_INVALID',
      message: 'Format foto tidak didukung. Gunakan JPG atau PNG.'
    };
  }

  const mimeType = String(match[1]).toLowerCase();
  let bytes;

  try {
    bytes = Utilities.base64Decode(String(match[2]).replace(/\s/g, ''));
  } catch (error) {
    return {
      success: false,
      code: 'PHOTO_BASE64_INVALID',
      message: 'Data foto rusak. Ambil foto ulang.'
    };
  }

  return {
    success: true,
    mimeType: mimeType,
    bytes: bytes
  };
}


function gttPhotoReadTable_(sheet) {
  const lastRow = Math.max(1, sheet.getLastRow());
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  return {
    headers: (values[0] || []).map(value =>
      String(value || '').trim().toUpperCase()
    ),
    rows: values.slice(1)
  };
}


function gttPhotoHeaderIndex_(headers, name) {
  const target = String(name || '').trim().toUpperCase();
  return (headers || []).findIndex(header => header === target);
}


function gttPhotoFail_(code, message) {
  return {
    success: false,
    code: code,
    message: message
  };
}
