/**
 * =====================================================
 * GRAND TIME TRACKER — GTT
 * Module : Mode Uji Waktu
 * Version: 1.1.0
 * Sprint : GTT-002
 * =====================================================
 *
 * Parameter MASTER_SETTING:
 * - MODE UJI WAKTU
 * - PIN UJI WAKTU
 * - TAMBAH MENIT UJI
 */

var GTT_SETTING_MODE_UJI_WAKTU_ = 'MODE UJI WAKTU';
var GTT_SETTING_PIN_UJI_WAKTU_ = 'PIN UJI WAKTU';
var GTT_SETTING_TAMBAH_MENIT_UJI_ = 'TAMBAH MENIT UJI';
var GTT_CACHE_MODE_UJI_WAKTU_ = 'GTT_MODE_UJI_WAKTU_V1_1';
var GTT_CACHE_MODE_UJI_DETIK_ = 10;
var GTT_BATAS_MENIT_UJI_ = 1440;

/**
 * Waktu transaksi GTT. Tambahan waktu hanya berlaku untuk PIN uji.
 *
 * @param {string|number} pinInput
 * @return {Date}
 */
function gttSekarang_(pinInput) {
  return gttInfoWaktu_(pinInput).sekarang;
}

/**
 * Informasi waktu lengkap untuk backend/API/UI.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function gttInfoWaktu_(pinInput) {
  var waktuServerAsli = new Date();
  var konfigurasi = gttAmbilKonfigurasiWaktu_();
  var pin = String(pinInput || '').trim();
  var modeUjiAktif = Boolean(
    konfigurasi.aktif === true &&
    pin !== '' &&
    pin === konfigurasi.pinUji
  );

  var sekarang = gttHitungWaktuUji_(
    waktuServerAsli,
    pin,
    konfigurasi
  );

  return {
    sekarang: sekarang,
    waktuServerAsli: waktuServerAsli,
    modeUjiAktif: modeUjiAktif,
    pinUji: konfigurasi.pinUji,
    tambahMenitDiterapkan: modeUjiAktif
      ? konfigurasi.tambahMenit
      : 0,
    timestampServerAsli: waktuServerAsli.getTime(),
    timestampDigunakan: sekarang.getTime()
  };
}

/**
 * Membaca konfigurasi Mode Uji Waktu dari MASTER_SETTING.
 *
 * @return {Object}
 */
function gttAmbilKonfigurasiWaktu_() {
  var cache = CacheService.getScriptCache();
  var tersimpan = cache.get(GTT_CACHE_MODE_UJI_WAKTU_);

  if (tersimpan) {
    try {
      return JSON.parse(tersimpan);
    } catch (errorCache) {
      cache.remove(GTT_CACHE_MODE_UJI_WAKTU_);
    }
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('Spreadsheet aktif GTT tidak ditemukan.');
  }

  var sheet = spreadsheet.getSheetByName('MASTER_SETTING');

  if (!sheet) {
    throw new Error('Sheet MASTER_SETTING tidak ditemukan.');
  }

  var barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    throw new Error('MASTER_SETTING masih kosong.');
  }

  var jumlahKolom = Math.max(2, Math.min(3, sheet.getLastColumn()));
  var data = sheet
    .getRange(2, 1, barisTerakhir - 1, jumlahKolom)
    .getDisplayValues();

  var settingMap = {};

  data.forEach(function (baris) {
    var parameter = gttNormalisasiSettingWaktu_(baris[0]);

    if (parameter) {
      settingMap[parameter] = String(baris[1] || '').trim();
    }
  });

  var mode = gttWajibSettingWaktu_(
    settingMap,
    GTT_SETTING_MODE_UJI_WAKTU_
  ).toUpperCase();

  var pinUji = gttWajibSettingWaktu_(
    settingMap,
    GTT_SETTING_PIN_UJI_WAKTU_
  );

  var teksTambahMenit = gttWajibSettingWaktu_(
    settingMap,
    GTT_SETTING_TAMBAH_MENIT_UJI_
  );

  if (mode !== 'YA' && mode !== 'TIDAK') {
    throw new Error('MODE UJI WAKTU harus bernilai YA atau TIDAK.');
  }

  if (!/^\d{4}$/.test(pinUji)) {
    throw new Error('PIN UJI WAKTU harus terdiri dari 4 angka.');
  }

  var tambahMenit = Number(teksTambahMenit);

  if (
    !Number.isInteger(tambahMenit) ||
    Math.abs(tambahMenit) > GTT_BATAS_MENIT_UJI_
  ) {
    throw new Error(
      'TAMBAH MENIT UJI harus berupa bilangan bulat antara -1440 sampai 1440.'
    );
  }

  var konfigurasi = {
    aktif: mode === 'YA',
    pinUji: pinUji,
    tambahMenit: tambahMenit
  };

  cache.put(
    GTT_CACHE_MODE_UJI_WAKTU_,
    JSON.stringify(konfigurasi),
    GTT_CACHE_MODE_UJI_DETIK_
  );

  return konfigurasi;
}

/**
 * Fungsi murni penghitung waktu simulasi.
 *
 * @param {Date} waktuServerAsli
 * @param {string|number} pinInput
 * @param {Object} konfigurasi
 * @return {Date}
 */
function gttHitungWaktuUji_(waktuServerAsli, pinInput, konfigurasi) {
  if (!(waktuServerAsli instanceof Date)) {
    throw new Error('Waktu server asli harus berupa Date.');
  }

  var waktuAsliMs = waktuServerAsli.getTime();

  if (!Number.isFinite(waktuAsliMs)) {
    throw new Error('Waktu server asli tidak valid.');
  }

  var config = konfigurasi || {};
  var pin = String(pinInput || '').trim();
  var pinUji = String(config.pinUji || '').trim();
  var tambahMenit = Number(config.tambahMenit || 0);

  if (!Number.isFinite(tambahMenit)) {
    tambahMenit = 0;
  }

  var bolehDisesuaikan = Boolean(
    config.aktif === true &&
    pin !== '' &&
    pin === pinUji
  );

  return new Date(
    waktuAsliMs +
    (bolehDisesuaikan ? tambahMenit : 0) * 60000
  );
}

/** Menghapus cache setelah MASTER_SETTING diubah. */
function gttResetCacheModeUjiWaktu_() {
  CacheService
    .getScriptCache()
    .remove(GTT_CACHE_MODE_UJI_WAKTU_);
}

function gttNormalisasiSettingWaktu_(nilai) {
  return String(nilai || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function gttWajibSettingWaktu_(settingMap, nama) {
  var kunci = gttNormalisasiSettingWaktu_(nama);

  if (!Object.prototype.hasOwnProperty.call(settingMap, kunci)) {
    throw new Error('Parameter "' + nama + '" tidak ditemukan.');
  }

  return String(settingMap[kunci] || '').trim();
}

/**
 * Tes logika murni dan konfigurasi MASTER_SETTING.
 * Tidak menulis data ke Sheet.
 *
 * @return {Object}
 */
function ujiModeUjiWaktuGTT() {
  var dasar = new Date('2026-08-02T00:00:00.000Z');
  var configAktif = {
    aktif: true,
    pinUji: '3710',
    tambahMenit: 180
  };

  var hasilNonaktif = gttHitungWaktuUji_(
    dasar,
    '3710',
    { aktif: false, pinUji: '3710', tambahMenit: 180 }
  );

  var hasilPinLain = gttHitungWaktuUji_(
    dasar,
    '9999',
    configAktif
  );

  var hasilPinUji = gttHitungWaktuUji_(
    dasar,
    '3710',
    configAktif
  );

  var konfigurasiAktif = gttAmbilKonfigurasiWaktu_();

  var hasil = {
    modeNonaktifTidakBerubah:
      hasilNonaktif.getTime() === dasar.getTime(),
    pinLainTidakBerubah:
      hasilPinLain.getTime() === dasar.getTime(),
    pinUjiBertambah180Menit:
      hasilPinUji.getTime() === dasar.getTime() + 180 * 60000,
    masterSettingTerbaca:
      typeof konfigurasiAktif.aktif === 'boolean' &&
      /^\d{4}$/.test(konfigurasiAktif.pinUji) &&
      Number.isInteger(konfigurasiAktif.tambahMenit)
  };

  var seluruhPass = Object.keys(hasil).every(function (kunci) {
    return hasil[kunci] === true;
  });

  var output = {
    success: seluruhPass,
    code: seluruhPass
      ? 'UJI_MODE_WAKTU_PASS'
      : 'UJI_MODE_WAKTU_GAGAL',
    hasil: hasil,
    konfigurasiAktif: konfigurasiAktif
  };

  console.log(JSON.stringify(output, null, 2));
  return output;
}
