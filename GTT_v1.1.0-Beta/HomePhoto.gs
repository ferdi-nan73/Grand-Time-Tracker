/**
 * GTT V1.1.0-ALPHA — HOME PHOTO READER FIX 001
 * Backend READ-ONLY foto profil untuk Beranda SA.
 *
 * FIX:
 * - Tidak menyimpan Base64 foto ke CacheService.
 * - Tidak mengubah absensi/break/GPS.
 * - Membaca PHOTO_FILE_ID dari MASTER_SA.
 */

function getHomeProfilePhotoGttAlpha(saId) {
  var cleanSaId = String(saId || '').trim();

  if (!cleanSaId) {
    return {
      success: false,
      code: 'HOME_PHOTO_SA_ID_EMPTY'
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSa = ss.getSheetByName('MASTER_SA');

  if (!masterSa) {
    return {
      success: false,
      code: 'HOME_PHOTO_MASTER_SA_NOT_FOUND'
    };
  }

  var lastRow = masterSa.getLastRow();
  var lastCol = masterSa.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    return {
      success: false,
      code: 'HOME_PHOTO_MASTER_SA_EMPTY'
    };
  }

  var values = masterSa
    .getRange(1, 1, lastRow, lastCol)
    .getDisplayValues();

  var headers = values[0].map(function(value) {
    return String(value || '')
      .trim()
      .toUpperCase();
  });

  var saIdCol = headers.indexOf('SA_ID');
  var photoCol = headers.indexOf('PHOTO_FILE_ID');

  if (saIdCol < 0 || photoCol < 0) {
    return {
      success: false,
      code: 'HOME_PHOTO_COLUMNS_NOT_FOUND'
    };
  }

  var photoFileId = '';

  for (var row = 1; row < values.length; row++) {
    var rowSaId = String(
      values[row][saIdCol] || ''
    ).trim();

    if (rowSaId === cleanSaId) {
      photoFileId = String(
        values[row][photoCol] || ''
      ).trim();

      break;
    }
  }

  if (!photoFileId) {
    return {
      success: false,
      code: 'HOME_PHOTO_NOT_SET'
    };
  }

  try {
    var file = DriveApp.getFileById(photoFileId);
    var blob = file.getBlob();

    var mime = String(
      blob.getContentType() || 'image/jpeg'
    );

    if (mime.indexOf('image/') !== 0) {
      return {
        success: false,
        code: 'HOME_PHOTO_INVALID_MIME'
      };
    }

    var dataUrl =
      'data:' +
      mime +
      ';base64,' +
      Utilities.base64Encode(
        blob.getBytes()
      );

    return {
      success: true,
      code: 'HOME_PHOTO_OK',
      fileId: photoFileId,
      dataUrl: dataUrl
    };

  } catch (error) {
    console.error(
      'getHomeProfilePhotoGttAlpha:',
      error
    );

    return {
      success: false,
      code: 'HOME_PHOTO_READ_FAILED',
      message:
        error && error.message
          ? error.message
          : String(error)
    };
  }
}
