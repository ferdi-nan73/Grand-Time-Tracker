/**
 * GTT v1.0.0-Alpha — HOME PATCH 001
 * Backend READ-ONLY foto profil untuk Beranda SA.
 * Tidak menulis data dan tidak menyentuh logika absensi/break/GPS.
 */
function getHomeProfilePhotoGttAlpha(saId) {
  var cleanSaId = String(saId || '').trim();
  if (!cleanSaId) {
    return { success: false, code: 'HOME_PHOTO_SA_ID_EMPTY' };
  }

  var cache = CacheService.getScriptCache();
  var cacheKey = 'GTT_HOME_PHOTO_' + cleanSaId;
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      console.warn('GTT HOME PHOTO - Cache JSON parse gagal:', err);
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSa = ss.getSheetByName('MASTER_SA');
  if (!masterSa) {
    return { success: false, code: 'HOME_PHOTO_MASTER_SA_NOT_FOUND' };
  }

  var lastRow = masterSa.getLastRow();
  var lastCol = masterSa.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { success: false, code: 'HOME_PHOTO_MASTER_SA_EMPTY' };
  }

  var values = masterSa.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function(v) { return String(v || '').trim().toUpperCase(); });
  var saIdCol = headers.indexOf('SA_ID');
  var photoCol = headers.indexOf('PHOTO_FILE_ID');

  if (saIdCol < 0 || photoCol < 0) {
    return { success: false, code: 'HOME_PHOTO_COLUMNS_NOT_FOUND' };
  }

  var photoFileId = '';
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][saIdCol] || '').trim() === cleanSaId) {
      photoFileId = String(values[r][photoCol] || '').trim();
      break;
    }
  }

  if (!photoFileId) {
    return { success: false, code: 'HOME_PHOTO_NOT_SET' };
  }

  try {
    var blob = DriveApp.getFileById(photoFileId).getBlob();
    var mime = String(blob.getContentType() || 'image/jpeg');
    if (mime.indexOf('image/') !== 0) {
      return { success: false, code: 'HOME_PHOTO_INVALID_MIME' };
    }

    var dataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
    var result = {
      success: true,
      code: 'HOME_PHOTO_OK',
      fileId: photoFileId,
      dataUrl: dataUrl
    };

    // Cache singkat: mengurangi pembacaan Drive saat user bolak-balik Beranda.
    cache.put(cacheKey, JSON.stringify(result), 300);
    return result;
  } catch (err) {
    return {
      success: false,
      code: 'HOME_PHOTO_READ_FAILED',
      message: err && err.message ? err.message : String(err)
    };
  }
}
