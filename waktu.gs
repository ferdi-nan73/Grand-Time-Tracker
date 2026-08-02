/**
 * =====================================================
 * GRAND TIME TRACKER — GTT
 * Module : GTT-06 Mode Uji Waktu
 * Version: 1.0.0
 * Status : Development — MENUNGGU UJI WEB APP
 * =====================================================
 *
 * Fungsi publik internal:
 * - gttSekarang_(pinInput)
 * - gttInfoWaktu_(pinInput)
 *
 * Parameter MASTER_SETTING:
 * - MODE UJI WAKTU
 * - PIN UJI WAKTU
 * - TAMBAH MENIT UJI
 */


const GTT_SETTING_MODE_UJI_WAKTU_ =
  'MODE UJI WAKTU';

const GTT_SETTING_PIN_UJI_WAKTU_ =
  'PIN UJI WAKTU';

const GTT_SETTING_TAMBAH_MENIT_UJI_ =
  'TAMBAH MENIT UJI';

const GTT_CACHE_MODE_UJI_WAKTU_ =
  'GTT_MODE_UJI_WAKTU_V1';

const GTT_CACHE_MODE_UJI_DETIK_ = 10;

const GTT_BATAS_MENIT_UJI_ = 1440;


/**
 * Menghasilkan waktu yang dipakai transaksi GTT.
 * Waktu uji hanya diterapkan kepada PIN uji.
 *
 * @param {string|number} pinInput
 * @return {Date}
 */
function gttSekarang_(pinInput) {
  return gttInfoWaktu_(pinInput).sekarang;
}


/**
 * Menghasilkan informasi waktu lengkap untuk API/UI.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function gttInfoWaktu_(pinInput) {
  const waktuServerAsli = new Date();
  const konfigurasi = gttAmbilKonfigurasiWaktu_();
  const pin = String(pinInput || '').trim();

  const modeUjiAktif =
    konfigurasi.aktif === true &&
    pin === konfigurasi.pinUji;

  const sekarang = gttHitungWaktuUji_(
    waktuServerAsli,
    pin,
    konfigurasi
  );

  return {
    sekarang: sekarang,
    waktuServerAsli: waktuServerAsli,
    modeUjiAktif: modeUjiAktif,
    pinUji: konfigurasi.pinUji,
    tambahMenitDiterapkan:
      modeUjiAktif
        ? konfigurasi.tambahMenit
        : 0
  };
}


/**
 * Membaca dan memvalidasi konfigurasi Mode Uji Waktu.
 * Cache singkat mencegah pembacaan Sheet pada setiap detik refresh.
 *
 * @return {Object}
 */
function gttAmbilKonfigurasiWaktu_() {
  const cache = CacheService.getScriptCache();
  const tersimpan = cache.get(
    GTT_CACHE_MODE_UJI_WAKTU_
  );

  if (tersimpan) {
    try {
      return JSON.parse(tersimpan);
    } catch (error) {
      cache.remove(GTT_CACHE_MODE_UJI_WAKTU_);
    }
  }

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName('MASTER_SETTING');

  if (!sheet) {
    throw new Error(
      'Sheet MASTER_SETTING tidak ditemukan.'
    );
  }

  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    throw new Error(
      'MASTER_SETTING masih kosong.'
    );
  }

  const data = sheet
    .getRange(
      2,
      1,
      barisTerakhir - 1,
      Math.min(2, sheet.getLastColumn())
    )
    .getDisplayValues();

  const settingMap = {};

  data.forEach(function (baris) {
    const parameter = gttNormalisasiSettingWaktu_(
      baris[0]
    );

    if (parameter) {
      settingMap[parameter] =
        String(baris[1] || '').trim();
    }
  });

  const mode = gttWajibSettingWaktu_(
    settingMap,
    GTT_SETTING_MODE_UJI_WAKTU_
  ).toUpperCase();

  const pinUji = gttWajibSettingWaktu_(
    settingMap,
    GTT_SETTING_PIN_UJI_WAKTU_
  );

  const teksTambahMenit = gttWajibSettingWaktu_(
    settingMap,
    GTT_SETTING_TAMBAH_MENIT_UJI_
  );

  if (!['YA', 'TIDAK'].includes(mode)) {
    throw new Error(
      'MODE UJI WAKTU harus bernilai YA atau TIDAK.'
    );
  }

  if (!/^\d{4}$/.test(pinUji)) {
    throw new Error(
      'PIN UJI WAKTU harus terdiri dari 4 angka.'
    );
  }

  const tambahMenit = Number(teksTambahMenit);

  if (
    !Number.isInteger(tambahMenit) ||
    Math.abs(tambahMenit) > GTT_BATAS_MENIT_UJI_
  ) {
    throw new Error(
      'TAMBAH MENIT UJI harus berupa bilangan bulat ' +
      'antara -1440 sampai 1440.'
    );
  }

  const konfigurasi = {
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
 * Fungsi murni untuk menghitung waktu hasil simulasi.
 * Dipisahkan supaya dapat diuji tanpa akses Spreadsheet.
 *
 * @param {Date} waktuServerAsli
 * @param {string|number} pinInput
 * @param {Object} konfigurasi
 * @return {Date}
 */
function gttHitungWaktuUji_(
  waktuServerAsli,
  pinInput,
  konfigurasi
) {
  if (!(waktuServerAsli instanceof Date)) {
    throw new Error(
      'Waktu server asli harus berupa Date.'
    );
  }

  const waktuAsliMs = waktuServerAsli.getTime();

  if (!Number.isFinite(waktuAsliMs)) {
    throw new Error('Waktu server asli tidak valid.');
  }

  const config = konfigurasi || {};
  const pin = String(pinInput || '').trim();
  const pinUji = String(config.pinUji || '').trim();
  const tambahMenit = Number(config.tambahMenit || 0);

  const bolehDisesuaikan =
    config.aktif === true &&
    pin !== '' &&
    pin === pinUji;

  const menitDiterapkan =
    bolehDisesuaikan
      ? tambahMenit
      : 0;

  return new Date(
    waktuAsliMs + menitDiterapkan * 60000
  );
}


/**
 * Menghapus cache setelah parameter diubah saat pengujian.
 */
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
  const kunci = gttNormalisasiSettingWaktu_(nama);

  if (!Object.prototype.hasOwnProperty.call(
    settingMap,
    kunci
  )) {
    throw new Error(
      'Parameter "' + nama + '" tidak ditemukan.'
    );
  }

  return String(settingMap[kunci] || '').trim();
}


/**
 * Tes logika yang aman: tidak menulis data ke Sheet.
 * Jalankan dari editor Apps Script setelah pemasangan.
 *
 * @return {Object}
 */
function ujiModeUjiWaktuGTT() {
  const dasar = new Date('2026-08-02T00:00:00.000Z');
  const configAktif = {
    aktif: true,
    pinUji: '3710',
    tambahMenit: 180
  };

  const hasilNonaktif = gttHitungWaktuUji_(
    dasar,
    '3710',
    {
      aktif: false,
      pinUji: '3710',
      tambahMenit: 180
    }
  );

  const hasilPinLain = gttHitungWaktuUji_(
    dasar,
    '9999',
    configAktif
  );

  const hasilPinUji = gttHitungWaktuUji_(
    dasar,
    '3710',
    configAktif
  );

  const hasil = {
    modeNonaktifTidakBerubah:
      hasilNonaktif.getTime() === dasar.getTime(),
    pinLainTidakBerubah:
      hasilPinLain.getTime() === dasar.getTime(),
    pinUjiBertambah180Menit:
      hasilPinUji.getTime() ===
      dasar.getTime() + 180 * 60000
  };

  const seluruhPass = Object.keys(hasil).every(
    function (kunci) {
      return hasil[kunci] === true;
    }
  );

  const output = {
    success: seluruhPass,
    code: seluruhPass
      ? 'UJI_MODE_WAKTU_PASS'
      : 'UJI_MODE_WAKTU_GAGAL',
    hasil: hasil,
    konfigurasiAktif: gttAmbilKonfigurasiWaktu_()
  };

  console.log(JSON.stringify(output, null, 2));
  return output;
}
