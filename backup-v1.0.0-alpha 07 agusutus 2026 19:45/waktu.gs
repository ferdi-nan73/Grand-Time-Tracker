/**
 * =====================================================
 * GRAND TIME TRACKER — GTT
 * Module : Engine Waktu & Mode Uji
 * Version: 1.2.0
 * Sprint : GTT-09 — Absolute Test Clock
 * =====================================================
 *
 * SUMBER PARAMETER: MASTER_SETTING
 *
 * Mode lama (kompatibilitas):
 * - MODE UJI WAKTU
 * - PIN UJI WAKTU
 * - TAMBAH MENIT UJI
 *
 * Mode baru (prioritas utama):
 * - MODE UJI SKENARIO
 * - PIN UJI SKENARIO
 * - WAKTU UJI SEKARANG
 *
 * ATURAN ARSITEKTUR:
 * 1. waktuServerAsli selalu dipakai untuk timestamp permanen.
 * 2. sekarang dipakai hanya untuk validasi logika bisnis.
 * 3. Mode Uji Skenario mengalahkan Mode Uji Waktu.
 * 4. Mode uji hanya berlaku untuk PIN uji yang sesuai.
 * 5. PIN lain selalu memakai waktu server asli.
 */

var GTT_WAKTU_SETTING_MODE_OFFSET_ =
  'MODE UJI WAKTU';

var GTT_WAKTU_SETTING_PIN_OFFSET_ =
  'PIN UJI WAKTU';

var GTT_WAKTU_SETTING_TAMBAH_MENIT_ =
  'TAMBAH MENIT UJI';

var GTT_WAKTU_SETTING_MODE_SKENARIO_ =
  'MODE UJI SKENARIO';

var GTT_WAKTU_SETTING_PIN_SKENARIO_ =
  'PIN UJI SKENARIO';

var GTT_WAKTU_SETTING_WAKTU_SKENARIO_ =
  'WAKTU UJI SEKARANG';

var GTT_WAKTU_CACHE_KEY_ =
  'GTT_ENGINE_WAKTU_V1_2';

var GTT_WAKTU_CACHE_DETIK_ = 5;

var GTT_WAKTU_BATAS_OFFSET_MENIT_ =
  1440;


/**
 * Mengembalikan waktu validasi bisnis GTT.
 *
 * CATATAN:
 * Jangan gunakan fungsi ini sebagai timestamp permanen.
 * Untuk simpan transaksi, gunakan:
 * gttInfoWaktu_(pin).waktuServerAsli
 *
 * @param {string|number} pinInput
 * @return {Date}
 */
function gttSekarang_(pinInput) {
  return gttInfoWaktu_(pinInput).sekarang;
}


/**
 * Mengembalikan waktu server asli untuk timestamp permanen.
 *
 * @return {Date}
 */
function gttWaktuServerAsli_() {
  return new Date();
}


/**
 * Informasi waktu lengkap untuk backend, API, dan UI.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function gttInfoWaktu_(pinInput) {
  var waktuServerAsli =
    gttWaktuServerAsli_();

  var konfigurasi =
    gttAmbilKonfigurasiWaktu_();

  var pin =
    String(pinInput || '').trim();

  var hasil =
    gttTentukanWaktuValidasi_(
      waktuServerAsli,
      pin,
      konfigurasi
    );

  return {
    // Waktu yang dipakai untuk validasi jeda/durasi.
    sekarang:
      hasil.sekarang,

    // Waktu yang wajib dipakai untuk menyimpan transaksi.
    waktuServerAsli:
      waktuServerAsli,

    modeUjiAktif:
      hasil.modeUjiAktif,

    modeUjiJenis:
      hasil.modeUjiJenis,

    pinUji:
      hasil.pinUji,

    // Tetap disediakan agar Api.gs lama kompatibel.
    tambahMenitDiterapkan:
      hasil.tambahMenitDiterapkan,

    waktuSkenario:
      hasil.waktuSkenario
        ? new Date(
            hasil.waktuSkenario.getTime()
          )
        : null,

    timestampServerAsli:
      waktuServerAsli.getTime(),

    timestampDigunakan:
      hasil.sekarang.getTime()
  };
}


/**
 * Menentukan waktu validasi berdasarkan prioritas:
 *
 * 1. MODE UJI SKENARIO
 * 2. MODE UJI WAKTU (offset lama)
 * 3. Waktu server asli
 *
 * @param {Date} waktuServerAsli
 * @param {string} pin
 * @param {Object} konfigurasi
 * @return {Object}
 */
function gttTentukanWaktuValidasi_(
  waktuServerAsli,
  pin,
  konfigurasi
) {
  if (
    !(waktuServerAsli instanceof Date) ||
    isNaN(waktuServerAsli.getTime())
  ) {
    throw new Error(
      'Waktu server asli tidak valid.'
    );
  }

  var config =
    konfigurasi || {};

  var pinBersih =
    String(pin || '').trim();

  var skenarioAktif =
    Boolean(
      config.skenario &&
      config.skenario.aktif === true &&
      pinBersih !== '' &&
      pinBersih ===
        String(
          config.skenario.pinUji || ''
        ).trim()
    );

  if (skenarioAktif) {
    if (
      !(
        config.skenario.waktu instanceof Date
      ) ||
      isNaN(
        config.skenario.waktu.getTime()
      )
    ) {
      throw new Error(
        'WAKTU UJI SEKARANG tidak valid.'
      );
    }

    return {
      sekarang:
        new Date(
          config.skenario.waktu.getTime()
        ),

      modeUjiAktif:
        true,

      modeUjiJenis:
        'SKENARIO',

      pinUji:
        String(
          config.skenario.pinUji || ''
        ).trim(),

      tambahMenitDiterapkan:
        0,

      waktuSkenario:
        new Date(
          config.skenario.waktu.getTime()
        )
    };
  }

  var offsetAktif =
    Boolean(
      config.offset &&
      config.offset.aktif === true &&
      pinBersih !== '' &&
      pinBersih ===
        String(
          config.offset.pinUji || ''
        ).trim()
    );

  var tambahMenit =
    offsetAktif
      ? Number(
          config.offset.tambahMenit || 0
        )
      : 0;

  if (!Number.isFinite(tambahMenit)) {
    tambahMenit = 0;
  }

  return {
    sekarang:
      new Date(
        waktuServerAsli.getTime() +
        tambahMenit * 60000
      ),

    modeUjiAktif:
      offsetAktif,

    modeUjiJenis:
      offsetAktif
        ? 'OFFSET'
        : 'NORMAL',

    pinUji:
      offsetAktif
        ? String(
            config.offset.pinUji || ''
          ).trim()
        : '',

    tambahMenitDiterapkan:
      tambahMenit,

    waktuSkenario:
      null
  };
}


/**
 * Membaca seluruh konfigurasi waktu dari MASTER_SETTING.
 *
 * @return {Object}
 */
function gttAmbilKonfigurasiWaktu_() {
  var cache =
    CacheService.getScriptCache();

  var tersimpan =
    cache.get(
      GTT_WAKTU_CACHE_KEY_
    );

  if (tersimpan) {
    try {
      return gttPulihkanKonfigurasiCache_(
        JSON.parse(tersimpan)
      );
    } catch (errorCache) {
      cache.remove(
        GTT_WAKTU_CACHE_KEY_
      );
    }
  }

  var spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'Spreadsheet aktif GTT tidak ditemukan.'
    );
  }

  var sheet =
    spreadsheet.getSheetByName(
      'MASTER_SETTING'
    );

  if (!sheet) {
    throw new Error(
      'Sheet MASTER_SETTING tidak ditemukan.'
    );
  }

  var barisTerakhir =
    sheet.getLastRow();

  if (barisTerakhir < 2) {
    throw new Error(
      'MASTER_SETTING masih kosong.'
    );
  }

  var jumlahKolom =
    Math.max(
      2,
      Math.min(
        3,
        sheet.getLastColumn()
      )
    );

  var data =
    sheet
      .getRange(
        2,
        1,
        barisTerakhir - 1,
        jumlahKolom
      )
      .getDisplayValues();

  var settingMap = {};

  data.forEach(function (baris) {
    var parameter =
      gttNormalisasiSettingWaktu_(
        baris[0]
      );

    if (parameter) {
      settingMap[parameter] =
        String(
          baris[1] || ''
        ).trim();
    }
  });

  var zonaWaktu =
    spreadsheet
      .getSpreadsheetTimeZone();

  var konfigurasi = {
    zonaWaktu:
      zonaWaktu,

    offset:
      gttBacaModeOffset_(
        settingMap
      ),

    skenario:
      gttBacaModeSkenario_(
        settingMap,
        zonaWaktu
      )
  };

  gttValidasiKonflikModeUji_(
    konfigurasi
  );

  cache.put(
    GTT_WAKTU_CACHE_KEY_,
    JSON.stringify(
      gttKonfigurasiUntukCache_(
        konfigurasi
      )
    ),
    GTT_WAKTU_CACHE_DETIK_
  );

  return konfigurasi;
}


/**
 * Membaca konfigurasi mode offset lama.
 *
 * @param {Object} settingMap
 * @return {Object}
 */
function gttBacaModeOffset_(
  settingMap
) {
  var mode =
    gttAmbilSettingOpsional_(
      settingMap,
      GTT_WAKTU_SETTING_MODE_OFFSET_,
      'TIDAK'
    ).toUpperCase();

  var pinUji =
    gttAmbilSettingOpsional_(
      settingMap,
      GTT_WAKTU_SETTING_PIN_OFFSET_,
      ''
    );

  var teksTambahMenit =
    gttAmbilSettingOpsional_(
      settingMap,
      GTT_WAKTU_SETTING_TAMBAH_MENIT_,
      '0'
    );

  gttValidasiYaTidak_(
    mode,
    GTT_WAKTU_SETTING_MODE_OFFSET_
  );

  if (
    mode === 'YA' &&
    !/^\d{4}$/.test(pinUji)
  ) {
    throw new Error(
      'PIN UJI WAKTU harus terdiri dari 4 angka.'
    );
  }

  var tambahMenit =
    Number(teksTambahMenit);

  if (
    !Number.isInteger(tambahMenit) ||
    Math.abs(tambahMenit) >
      GTT_WAKTU_BATAS_OFFSET_MENIT_
  ) {
    throw new Error(
      'TAMBAH MENIT UJI harus berupa bilangan bulat antara -1440 sampai 1440.'
    );
  }

  return {
    aktif:
      mode === 'YA',

    pinUji:
      pinUji,

    tambahMenit:
      tambahMenit
  };
}


/**
 * Membaca konfigurasi Mode Uji Skenario.
 *
 * @param {Object} settingMap
 * @param {string} zonaWaktu
 * @return {Object}
 */
function gttBacaModeSkenario_(
  settingMap,
  zonaWaktu
) {
  var mode =
    gttAmbilSettingOpsional_(
      settingMap,
      GTT_WAKTU_SETTING_MODE_SKENARIO_,
      'TIDAK'
    ).toUpperCase();

  var pinUji =
    gttAmbilSettingOpsional_(
      settingMap,
      GTT_WAKTU_SETTING_PIN_SKENARIO_,
      ''
    );

  var teksWaktu =
    gttAmbilSettingOpsional_(
      settingMap,
      GTT_WAKTU_SETTING_WAKTU_SKENARIO_,
      ''
    );

  gttValidasiYaTidak_(
    mode,
    GTT_WAKTU_SETTING_MODE_SKENARIO_
  );

  if (
    mode === 'YA' &&
    !/^\d{4}$/.test(pinUji)
  ) {
    throw new Error(
      'PIN UJI SKENARIO harus terdiri dari 4 angka.'
    );
  }

  var waktu = null;

  if (mode === 'YA') {
    waktu =
      gttParseWaktuSkenario_(
        teksWaktu,
        zonaWaktu
      );
  }

  return {
    aktif:
      mode === 'YA',

    pinUji:
      pinUji,

    waktu:
      waktu,

    teksWaktu:
      teksWaktu
  };
}


/**
 * Parse format wajib:
 * yyyy-MM-dd HH:mm:ss
 *
 * @param {string} teks
 * @param {string} zonaWaktu
 * @return {Date}
 */
function gttParseWaktuSkenario_(
  teks,
  zonaWaktu
) {
  var nilai =
    String(teks || '').trim();

  if (
    !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(
      nilai
    )
  ) {
    throw new Error(
      'WAKTU UJI SEKARANG wajib memakai format yyyy-MM-dd HH:mm:ss. Contoh: 2026-08-02 11:01:00.'
    );
  }

  var waktu;

  try {
    waktu =
      Utilities.parseDate(
        nilai,
        zonaWaktu,
        'yyyy-MM-dd HH:mm:ss'
      );
  } catch (errorParse) {
    throw new Error(
      'WAKTU UJI SEKARANG tidak dapat dibaca: ' +
      nilai
    );
  }

  if (
    !(waktu instanceof Date) ||
    isNaN(waktu.getTime())
  ) {
    throw new Error(
      'WAKTU UJI SEKARANG tidak valid.'
    );
  }

  var hasilFormatUlang =
    Utilities.formatDate(
      waktu,
      zonaWaktu,
      'yyyy-MM-dd HH:mm:ss'
    );

  if (hasilFormatUlang !== nilai) {
    throw new Error(
      'WAKTU UJI SEKARANG berisi tanggal atau jam yang tidak valid.'
    );
  }

  return waktu;
}


/**
 * Mencegah dua mode uji aktif bersamaan.
 * Mode skenario tetap menjadi prioritas, tetapi konflik wajib diberi
 * error agar pengujian tidak ambigu.
 *
 * @param {Object} konfigurasi
 */
function gttValidasiKonflikModeUji_(
  konfigurasi
) {
  if (
    konfigurasi.skenario.aktif &&
    konfigurasi.offset.aktif
  ) {
    throw new Error(
      'MODE UJI SKENARIO dan MODE UJI WAKTU tidak boleh sama-sama YA. Nonaktifkan salah satunya.'
    );
  }
}


/**
 * Tanggal operasional transaksi.
 *
 * Selalu memakai tanggal server asli, bukan tanggal waktu uji.
 *
 * @param {string|number} pinInput
 * @param {string=} zonaWaktuInput
 * @return {string} yyyy-MM-dd
 */
function gttTanggalOperasional_(
  pinInput,
  zonaWaktuInput
) {
  var info =
    gttInfoWaktu_(pinInput);

  var zonaWaktu =
    String(
      zonaWaktuInput || ''
    ).trim();

  if (!zonaWaktu) {
    zonaWaktu =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSpreadsheetTimeZone();
  }

  return Utilities.formatDate(
    info.waktuServerAsli,
    zonaWaktu,
    'yyyy-MM-dd'
  );
}


/**
 * Waktu validasi bisnis.
 *
 * @param {string|number} pinInput
 * @return {Date}
 */
function gttWaktuValidasi_(
  pinInput
) {
  return gttInfoWaktu_(
    pinInput
  ).sekarang;
}


/**
 * Menghapus cache konfigurasi waktu.
 *
 * Jalankan setelah mengubah MASTER_SETTING saat pengujian.
 */
function gttResetCacheModeUjiWaktu_() {
  CacheService
    .getScriptCache()
    .remove(
      GTT_WAKTU_CACHE_KEY_
    );
}


/**
 * Alias baru yang lebih jelas.
 */
function gttResetCacheWaktu_() {
  gttResetCacheModeUjiWaktu_();
}


function gttNormalisasiSettingWaktu_(
  nilai
) {
  return String(
    nilai || ''
  )
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}


function gttAmbilSettingOpsional_(
  settingMap,
  nama,
  nilaiDefault
) {
  var kunci =
    gttNormalisasiSettingWaktu_(
      nama
    );

  if (
    !Object.prototype.hasOwnProperty.call(
      settingMap,
      kunci
    )
  ) {
    return String(
      nilaiDefault === undefined
        ? ''
        : nilaiDefault
    ).trim();
  }

  return String(
    settingMap[kunci] || ''
  ).trim();
}


function gttValidasiYaTidak_(
  nilai,
  namaParameter
) {
  if (
    nilai !== 'YA' &&
    nilai !== 'TIDAK'
  ) {
    throw new Error(
      namaParameter +
      ' harus bernilai YA atau TIDAK.'
    );
  }
}


/**
 * Mengubah konfigurasi menjadi bentuk JSON-safe untuk cache.
 *
 * @param {Object} konfigurasi
 * @return {Object}
 */
function gttKonfigurasiUntukCache_(
  konfigurasi
) {
  return {
    zonaWaktu:
      konfigurasi.zonaWaktu,

    offset:
      {
        aktif:
          konfigurasi.offset.aktif,

        pinUji:
          konfigurasi.offset.pinUji,

        tambahMenit:
          konfigurasi.offset.tambahMenit
      },

    skenario:
      {
        aktif:
          konfigurasi.skenario.aktif,

        pinUji:
          konfigurasi.skenario.pinUji,

        waktuIso:
          konfigurasi.skenario.waktu
            ? konfigurasi
                .skenario
                .waktu
                .toISOString()
            : '',

        teksWaktu:
          konfigurasi
            .skenario
            .teksWaktu
      }
  };
}


/**
 * Memulihkan objek Date setelah membaca cache.
 *
 * @param {Object} cacheData
 * @return {Object}
 */
function gttPulihkanKonfigurasiCache_(
  cacheData
) {
  var data =
    cacheData || {};

  var skenario =
    data.skenario || {};

  return {
    zonaWaktu:
      String(
        data.zonaWaktu || ''
      ),

    offset:
      {
        aktif:
          Boolean(
            data.offset &&
            data.offset.aktif
          ),

        pinUji:
          String(
            data.offset &&
            data.offset.pinUji ||
            ''
          ),

        tambahMenit:
          Number(
            data.offset &&
            data.offset.tambahMenit ||
            0
          )
      },

    skenario:
      {
        aktif:
          Boolean(
            skenario.aktif
          ),

        pinUji:
          String(
            skenario.pinUji || ''
          ),

        waktu:
          skenario.waktuIso
            ? new Date(
                skenario.waktuIso
              )
            : null,

        teksWaktu:
          String(
            skenario.teksWaktu || ''
          )
      }
  };
}


/**
 * Self-test Mode Uji Skenario.
 *
 * Tidak menulis data ke Sheet.
 *
 * Syarat sebelum Run:
 * - MODE UJI SKENARIO = TIDAK atau YA
 * - Parameter baru tersedia di MASTER_SETTING
 *
 * @return {Object}
 */
function ujiModeUjiSkenarioGTT() {
  var zonaWaktu =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSpreadsheetTimeZone();

  var dasar =
    Utilities.parseDate(
      '2026-08-02 08:00:00',
      zonaWaktu,
      'yyyy-MM-dd HH:mm:ss'
    );

  var waktuSkenario =
    Utilities.parseDate(
      '2026-08-02 13:01:00',
      zonaWaktu,
      'yyyy-MM-dd HH:mm:ss'
    );

  var konfigurasiTes = {
    offset:
      {
        aktif: false,
        pinUji: '3710',
        tambahMenit: 240
      },

    skenario:
      {
        aktif: true,
        pinUji: '3710',
        waktu: waktuSkenario
      }
  };

  var hasilPinUji =
    gttTentukanWaktuValidasi_(
      dasar,
      '3710',
      konfigurasiTes
    );

  var hasilPinLain =
    gttTentukanWaktuValidasi_(
      dasar,
      '4224',
      konfigurasiTes
    );

  var konfigurasiAktif =
    gttAmbilKonfigurasiWaktu_();

  var hasil = {
    pinUjiMemakaiWaktuAbsolut:
      hasilPinUji.sekarang.getTime() ===
      waktuSkenario.getTime(),

    jenisModeSkenario:
      hasilPinUji.modeUjiJenis ===
      'SKENARIO',

    pinLainMemakaiWaktuServer:
      hasilPinLain.sekarang.getTime() ===
      dasar.getTime(),

    timestampServerTidakBerubah:
      dasar.getTime() ===
      Utilities.parseDate(
        '2026-08-02 08:00:00',
        zonaWaktu,
        'yyyy-MM-dd HH:mm:ss'
      ).getTime(),

    masterSettingTerbaca:
      Boolean(
        konfigurasiAktif &&
        konfigurasiAktif.skenario &&
        typeof konfigurasiAktif
          .skenario
          .aktif === 'boolean'
      )
  };

  var seluruhPass = true;

Object.keys(hasil).forEach(function(kunci) {
  if (hasil[kunci] !== true) {
    seluruhPass = false;
  }
});

  var output = {
    success:
      seluruhPass,

    code:
      seluruhPass
        ? 'UJI_MODE_SKENARIO_PASS'
        : 'UJI_MODE_SKENARIO_FAIL',

    hasil:
      hasil,

    konfigurasiAktif:
      {
        modeOffset:
          konfigurasiAktif
            .offset
            .aktif,

        modeSkenario:
          konfigurasiAktif
            .skenario
            .aktif,

        pinSkenario:
          konfigurasiAktif
            .skenario
            .pinUji,

        waktuSkenario:
          konfigurasiAktif
            .skenario
            .teksWaktu
      }
  };

  console.log(
    JSON.stringify(
      output,
      null,
      2
    )
  );

  return output;
}
/**
 * Fungsi publik untuk menghapus cache konfigurasi waktu.
 * Fungsi ini sengaja tanpa akhiran "_" agar muncul di dropdown Run.
 *
 * @return {Object}
 */
function gttResetCacheWaktu() {
  gttResetCacheWaktu_();

  var hasil = {
    success: true,
    code: 'CACHE_WAKTU_DIRESET',
    message: 'Cache konfigurasi waktu berhasil dihapus.'
  };

  console.log(
    JSON.stringify(
      hasil,
      null,
      2
    )
  );

  return hasil;
}
