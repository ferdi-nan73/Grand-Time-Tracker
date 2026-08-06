/**
 * GRAND TIME TRACKER
 * File: BreakTime.gs
 * Build: GTT v0.9.3-dev(2) Rev (3.4)
 * Build Date: 2026-08-03
 * Module: Final Activity, Footer, Version Stamp & Favicon
 */

/**
 * Memulai Break 1.
 *
 * @param {string|number} pinInput PIN SA.
 * @return {Object}
 */
function mulaiBreak1(pinInput) {
  return mulaiBreak_(pinInput, 1);
}

/**
 * Menyelesaikan Break 1.
 *
 * @param {string|number} pinInput PIN SA.
 * @return {Object}
 */
function selesaiBreak1(pinInput) {
  return selesaiBreak_(pinInput, 1);
}

/**
 * Memulai Break 2.
 *
 * @param {string|number} pinInput PIN SA.
 * @return {Object}
 */
function mulaiBreak2(pinInput) {
  return mulaiBreak_(pinInput, 2);
}

/**
 * Menyelesaikan Break 2.
 *
 * @param {string|number} pinInput PIN SA.
 * @return {Object}
 */
function selesaiBreak2(pinInput) {
  return selesaiBreak_(pinInput, 2);
}

/**
 * Memulai break.
 *
 * @param {string|number} pinInput
 * @param {number} nomorBreak
 * @return {Object}
 */
function mulaiBreak_(pinInput, nomorBreak) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (![1, 2].includes(nomorBreak)) {
      throw new Error('Nomor break tidak valid.');
    }

    var konteks = ambilKonteksBreak_(pinInput);

    if (!konteks.success) {
      return konteks;
    }

    var {
      pengguna,
      spreadsheet,
      sheetLog,
      headerMap,
      zonaWaktu,
      sekarang,
      tanggalHariIni,
      pengaturan
    } = konteks.data;

    var hasilAbsensi = cariAbsensiAktifHariIni_(
      spreadsheet,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    if (!hasilAbsensi.ditemukan) {
      return {
        success: false,
        code: 'BELUM_ABSEN',
        message:
          'Anda belum melakukan absensi masuk hari ini. ' +
          'Silakan melakukan absensi terlebih dahulu.'
      };
    }

    var dataLog = cariLogIstirahatHariIni_(
      sheetLog,
      headerMap,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    if (nomorBreak === 1) {
      var validasiBreak1 = validasiMulaiBreak1_(
        dataLog,
        headerMap,
        hasilAbsensi,
        sekarang,
        pengaturan,
        zonaWaktu
      );

      if (!validasiBreak1.success) {
        return validasiBreak1;
      }
    }

    if (nomorBreak === 2) {
      var validasiBreak2 = validasiMulaiBreak2_(
        dataLog,
        headerMap,
        sekarang,
        pengaturan,
        zonaWaktu
      );

      if (!validasiBreak2.success) {
        return validasiBreak2;
      }
    }

    var jumlahSedangBreak = hitungSedangBreakOutlet_(
      sheetLog,
      headerMap,
      pengguna.outlet,
      tanggalHariIni,
      zonaWaktu
    );

    var maxIstirahat = pengaturan.maxBreakBersamaan;

    if (jumlahSedangBreak >= maxIstirahat) {
      var templateKapasitas =
        pengaturan.warningKapasitas ||
        (
          'MAKSIMAL {{JUMLAH}} SA DAPAT BREAK ' +
          'BERSAMAAN. SILAKAN TUNGGU SA LAIN ' +
          'SELESAI BREAK.'
        );

      var pesanKapasitas =
        gantiPlaceholderBreak_(
          templateKapasitas,
          {
            JUMLAH:
              maxIstirahat,
            MAKSIMAL:
              maxIstirahat,
            SEDANG:
              jumlahSedangBreak
          }
        );

      return {
        success: false,
        code: 'KAPASITAS_BREAK_PENUH',
        message:
          pesanKapasitas,
        data: {
          outlet: pengguna.outlet,
          sedangBreak: jumlahSedangBreak,
          maxIstirahat: maxIstirahat
        }
      };
    }

    var nomorBaris;

    if (!dataLog.ditemukan) {
      nomorBaris = buatBarisLogIstirahatBaru_(
        sheetLog,
        headerMap,
        pengguna,
        sekarang,
        zonaWaktu,
        tanggalHariIni
      );
    } else {
      nomorBaris = dataLog.nomorBaris;
    }

    var namaKolomMulai =
      nomorBreak === 1 ? 'S1 MULAI' : 'S2 MULAI';

    sheetLog
      .getRange(
        nomorBaris,
        headerMap[namaKolomMulai] + 1
      )
      .setValue(sekarang)
      .setNumberFormat('HH:mm:ss');

    SpreadsheetApp.flush();

    var durasiBreak =
      nomorBreak === 1
        ? pengaturan.durasiBreak1
        : pengaturan.durasiBreak2;

    var estimasiSelesai = new Date(
      sekarang.getTime() + durasiBreak * 60000
    );

    return {
      success: true,
      code:
        nomorBreak === 1
          ? 'BREAK_1_DIMULAI'
          : 'BREAK_2_DIMULAI',
      message:
        'Break ' +
        nomorBreak +
        ' dimulai pada pukul ' +
        Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'HH:mm:ss'
        ) +
        '.',
      data: {
        pin: pengguna.pin,
        namaSA: pengguna.namaSA,
        outlet: pengguna.outlet,
        break: nomorBreak,
        waktuMulai: Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'HH:mm:ss'
        ),
        estimasiSelesai: Utilities.formatDate(
          estimasiSelesai,
          zonaWaktu,
          'HH:mm:ss'
        ),
        durasiMenit: durasiBreak,
        batasAkhirBreak2:
          nomorBreak === 2
            ? pengaturan.batasAkhirBreak2
            : null,
        reminderMenit:
          pengaturan.reminderSebelumBerakhir,
        statusOperasional:
          nomorBreak === 1
            ? 'BREAK 1'
            : 'BREAK 2'
      }
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      code: 'SYSTEM_ERROR',
      message:
        'Terjadi kesalahan sistem: ' +
        error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * Menyelesaikan break.
 *
 * @param {string|number} pinInput
 * @param {number} nomorBreak
 * @return {Object}
 */
function selesaiBreak_(pinInput, nomorBreak) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (![1, 2].includes(nomorBreak)) {
      throw new Error('Nomor break tidak valid.');
    }

    var konteks = ambilKonteksBreak_(pinInput);

    if (!konteks.success) {
      return konteks;
    }

    var {
      pengguna,
      sheetLog,
      headerMap,
      zonaWaktu,
      sekarang,
      tanggalHariIni,
      pengaturan
    } = konteks.data;

    var dataLog = cariLogIstirahatHariIni_(
      sheetLog,
      headerMap,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    if (!dataLog.ditemukan) {
      return {
        success: false,
        code: 'BREAK_BELUM_DIMULAI',
        message:
          'Data istirahat hari ini belum tersedia.'
      };
    }

    var namaKolomMulai =
      nomorBreak === 1 ? 'S1 MULAI' : 'S2 MULAI';

    var namaKolomSelesai =
      nomorBreak === 1 ? 'S1 SELESAI' : 'S2 SELESAI';

    var namaKolomDurasi =
      nomorBreak === 1 ? 'S1 DURASI' : 'S2 DURASI';

    var namaKolomOvertime =
      nomorBreak === 1 ? 'S1 OVERTIME' : 'S2 OVERTIME';

    var namaKolomStatus =
      nomorBreak === 1 ? 'S1 STATUS' : 'S2 STATUS';

    var namaKolomSanksi =
      nomorBreak === 1 ? 'S1 SANKSI' : 'S2 SANKSI';

    var rangeBaris = sheetLog.getRange(
      dataLog.nomorBaris,
      1,
      1,
      sheetLog.getLastColumn()
    );

    var nilaiBaris = rangeBaris.getValues()[0];

    var waktuMulai =
      nilaiBaris[headerMap[namaKolomMulai]];

    var waktuSelesaiLama =
      nilaiBaris[headerMap[namaKolomSelesai]];

    if (!(waktuMulai instanceof Date)) {
      return {
        success: false,
        code: 'BREAK_BELUM_DIMULAI',
        message:
          'Break ' +
          nomorBreak +
          ' belum dimulai.'
      };
    }

    if (waktuSelesaiLama instanceof Date) {
      return {
        success: false,
        code: 'BREAK_SUDAH_SELESAI',
        message:
          'Break ' +
          nomorBreak +
          ' sudah diselesaikan pada pukul ' +
          Utilities.formatDate(
            waktuSelesaiLama,
            zonaWaktu,
            'HH:mm:ss'
          ) +
          '. Data pertama tetap digunakan.'
      };
    }

    var durasiMenit = hitungDurasiMenit_(
      waktuMulai,
      sekarang
    );

    var batasDurasi =
      nomorBreak === 1
        ? pengaturan.durasiBreak1
        : pengaturan.durasiBreak2;

    var lewatDurasiMenit = Math.max(
      0,
      durasiMenit - batasDurasi
    );

    var overbreakMenit = Math.max(
      0,
      lewatDurasiMenit - pengaturan.toleransiOverbreak
    );

    var evaluasi = evaluasiOverbreak_(
      lewatDurasiMenit,
      overbreakMenit,
      pengaturan.toleransiOverbreak,
      pengaturan.sanksiOverbreak
    );

    sheetLog
      .getRange(
        dataLog.nomorBaris,
        headerMap[namaKolomSelesai] + 1
      )
      .setValue(sekarang)
      .setNumberFormat('HH:mm:ss');

    sheetLog
      .getRange(
        dataLog.nomorBaris,
        headerMap[namaKolomDurasi] + 1
      )
      .setValue(durasiMenit)
      .setNumberFormat('0');

    sheetLog
      .getRange(
        dataLog.nomorBaris,
        headerMap[namaKolomOvertime] + 1
      )
      .setValue(overbreakMenit)
      .setNumberFormat('0');

    sheetLog
      .getRange(
        dataLog.nomorBaris,
        headerMap[namaKolomStatus] + 1
      )
      .setValue(evaluasi.status);

    sheetLog
      .getRange(
        dataLog.nomorBaris,
        headerMap[namaKolomSanksi] + 1
      )
      .setValue(evaluasi.sanksi)
      .setNumberFormat('#,##0');

    perbaruiTotalIstirahat_(
      sheetLog,
      dataLog.nomorBaris,
      headerMap
    );

    SpreadsheetApp.flush();

    return {
      success: true,
      code:
        nomorBreak === 1
          ? 'BREAK_1_SELESAI'
          : 'BREAK_2_SELESAI',
      message:
        'Break ' +
        nomorBreak +
        ' selesai. Durasi ' +
        durasiMenit +
        ' menit.',
      data: {
        pin: pengguna.pin,
        namaSA: pengguna.namaSA,
        outlet: pengguna.outlet,
        break: nomorBreak,
        waktuMulai: Utilities.formatDate(
          waktuMulai,
          zonaWaktu,
          'HH:mm:ss'
        ),
        waktuSelesai: Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'HH:mm:ss'
        ),
        durasiMenit: durasiMenit,
        batasDurasiMenit: batasDurasi,
        lewatDurasiMenit: lewatDurasiMenit,
        overbreakMenit: overbreakMenit,
        overtimeMenit: overbreakMenit,
        toleransiOverbreak:
          pengaturan.toleransiOverbreak,
        status: evaluasi.status,
        sanksi: evaluasi.sanksi,
        potensiSanksi: evaluasi.sanksi,
        warning: buatPesanHasilBreak_(
          evaluasi,
          overbreakMenit,
          pengaturan
        ),
        statusOperasional: 'AKTIF BEKERJA'
      }
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      code: 'SYSTEM_ERROR',
      message:
        'Terjadi kesalahan sistem: ' +
        error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * Mengambil konteks umum Break Time.
 */
function ambilKonteksBreak_(pinInput) {
  var hasilLogin = validasiLoginPin(pinInput);

  if (!hasilLogin.success) {
    return hasilLogin;
  }

  var pengguna = Object.assign({}, hasilLogin.data, {
    pin: String(
      (hasilLogin.data && hasilLogin.data.pin) || pinInput || ''
    ).trim()
  });

  var spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  var sheetLog =
    spreadsheet.getSheetByName('LOG_ISTIRAHAT');

  if (!sheetLog) {
    return {
      success: false,
      code: 'SHEET_LOG_TIDAK_DITEMUKAN',
      message:
        'Sheet LOG_ISTIRAHAT tidak ditemukan.'
    };
  }

  var headerMap =
    ambilHeaderMapLogIstirahat_(sheetLog);

  var zonaWaktu =
    spreadsheet.getSpreadsheetTimeZone();

  var infoWaktu = typeof gttInfoWaktu_ === 'function'
    ? gttInfoWaktu_(pengguna.pin)
    : {
        sekarang: new Date(),
        waktuServerAsli: new Date(),
        modeUjiAktif: false,
        tambahMenitDiterapkan: 0
      };

  // Waktu virtual dipakai hanya untuk validasi jeda/durasi.
  var sekarang = infoWaktu.sekarang;

  // Tanggal transaksi tetap mengikuti tanggal server asli agar
  // Mode Uji tidak memindahkan data ke hari berikutnya.
  var tanggalHariIni = Utilities.formatDate(
    infoWaktu.waktuServerAsli,
    zonaWaktu,
    'yyyy-MM-dd'
  );

  var pengaturan =
    ambilPengaturanBreak_();

  return {
    success: true,
    code: 'KONTEKS_BREAK_VALID',
    data: {
      pengguna: pengguna,
      spreadsheet: spreadsheet,
      sheetLog: sheetLog,
      headerMap: headerMap,
      zonaWaktu: zonaWaktu,
      sekarang: sekarang,
      tanggalHariIni: tanggalHariIni,
      pengaturan: pengaturan
    }
  };
}

/**
 * Memvalidasi header LOG_ISTIRAHAT.
 */
function ambilHeaderMapLogIstirahat_(sheet) {
  var jumlahKolom = sheet.getLastColumn();

  if (jumlahKolom === 0) {
    throw new Error(
      'Header LOG_ISTIRAHAT belum tersedia.'
    );
  }

  var daftarHeader = sheet
    .getRange(1, 1, 1, jumlahKolom)
    .getDisplayValues()[0]
    .map(normalisasiHeader_);

  var kolomWajib = [
    'ID',
    'TANGGAL',
    'PIN',
    'NAMA SA',
    'OUTLET',
    'S1 MULAI',
    'S1 SELESAI',
    'S1 DURASI',
    'S1 OVERTIME',
    'S1 STATUS',
    'S1 SANKSI',
    'S2 MULAI',
    'S2 SELESAI',
    'S2 DURASI',
    'S2 OVERTIME',
    'S2 STATUS',
    'S2 SANKSI',
    'TOTAL OVERTIME',
    'TOTAL SANKSI'
  ];

  var headerMap = {};

  daftarHeader.forEach((namaHeader, index) => {
    if (namaHeader) {
      headerMap[namaHeader] = index;
    }
  });

  var kolomTidakAda = kolomWajib.filter(
    namaKolom =>
      headerMap[namaKolom] === undefined
  );

  if (kolomTidakAda.length > 0) {
    throw new Error(
      'Kolom berikut tidak ditemukan di ' +
      'LOG_ISTIRAHAT: ' +
      kolomTidakAda.join(', ')
    );
  }

  return headerMap;
}

/**
 * Membaca setting Break Time.
 */
function ambilPengaturanBreak_() {
  var spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  var sheet =
    spreadsheet.getSheetByName('MASTER_SETTING');

  if (!sheet) {
    throw new Error(
      'Sheet MASTER_SETTING tidak ditemukan.'
    );
  }

  var barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    throw new Error(
      'MASTER_SETTING masih kosong.'
    );
  }

  var jumlahKolom = Math.min(
    3,
    sheet.getLastColumn()
  );

  var data = sheet
    .getRange(
      2,
      1,
      barisTerakhir - 1,
      jumlahKolom
    )
    .getDisplayValues();

  var settingMap = {};

  data.forEach(baris => {
    var parameter =
      normalisasiHeader_(baris[0]);

    if (!parameter) return;

    settingMap[parameter] = {
      nilai: String(baris[1] || '').trim(),
      warning: String(baris[2] || '').trim()
    };
  });

  var durasiBreak1 =
    ambilItemSettingBreak_(
      settingMap,
      'BREAK - DURASI BREAK 1 (MENIT)'
    );

  var durasiBreak2 =
    ambilItemSettingBreak_(
      settingMap,
      'BREAK - DURASI BREAK 2 (MENIT)'
    );

  var batasAkhirBreak2Item =
    ambilItemSettingBreak_(
      settingMap,
      'BREAK - BATAS AKHIR BREAK 2'
    );

  var toleransiOverbreak =
    ambilItemSettingBreak_(
      settingMap,
      'BREAK - TOLERANSI OVERBREAK (MENIT)'
    );

  var reminder =
    ambilItemSettingBreak_(
      settingMap,
      'BREAK - REMINDER BREAK (MENIT)'
    );

  var sanksi =
    ambilItemSettingBreak_(
      settingMap,
      'BREAK - SANKSI OVERBREAK / BREAK (RUPIAH)'
    );

  var maxBreakBersamaan =
    ambilItemSettingBreak_(
      settingMap,
      'BREAK - MAKSIMAL SA BREAK BERSAMAAN'
    );

  var jedaMasukKeBreak1 =
    ambilNilaiSettingBreak_(
      settingMap,
      [
        'JEDA ABSEN MASUK KE BREAK 1 (MENIT)',
        'BREAK - JEDA DARI JAM MASUK KE BREAK (MENIT)'
      ],
      true
    );

  var jedaBreak1KeBreak2 =
    ambilNilaiSettingBreak_(
      settingMap,
      [
        'JEDA BREAK 1 KE BREAK 2 (MENIT)',
        'BREAK - JEDA ANTARA BREAK 1 KE BREAK 2 (MENIT)'
      ],
      true
    );

  var minimalSisaKerjaBreak2 =
    ambilNilaiSettingBreak_(
      settingMap,
      [
        'BREAK - MINIMAL SISA KERJA UNTUK BREAK 2 (MENIT)',
        'MINIMAL SISA KERJA UNTUK BREAK 2 (MENIT)'
      ],
      true
    );

  var jamPulangNormal =
    ambilItemSettingBreak_(
      settingMap,
      'ABSENSI - JAM PULANG NORMAL'
    );

  return {
    durasiBreak1:
      validasiAngkaSetting_(
        durasiBreak1.nilai,
        'Durasi Break 1'
      ),

    durasiBreak2:
      validasiAngkaSetting_(
        durasiBreak2.nilai,
        'Durasi Break 2'
      ),

    batasAkhirBreak2:
      validasiFormatWaktuSetting_(
        batasAkhirBreak2Item.nilai,
        'Batas Akhir Break 2'
      ),

    toleransiOverbreak:
      validasiAngkaSetting_(
        toleransiOverbreak.nilai,
        'Toleransi Overbreak'
      ),

    reminderSebelumBerakhir:
      validasiAngkaSetting_(
        reminder.nilai,
        'Reminder Break'
      ),

    sanksiOverbreak:
      validasiAngkaSetting_(
        sanksi.nilai,
        'Sanksi Overbreak / Break'
      ),

    maxBreakBersamaan:
      validasiAngkaPositifSetting_(
        maxBreakBersamaan.nilai,
        'Maksimal SA Break Bersamaan'
      ),

    jedaMasukKeBreak1:
      validasiAngkaPositifSetting_(
        jedaMasukKeBreak1.nilai,
        'Jeda Absen Masuk ke Break 1'
      ),

    jedaBreak1KeBreak2:
      validasiAngkaPositifSetting_(
        jedaBreak1KeBreak2.nilai,
        'Jeda Break 1 ke Break 2'
      ),

    minimalSisaKerjaBreak2:
      validasiAngkaPositifSetting_(
        minimalSisaKerjaBreak2.nilai,
        'Minimal Sisa Kerja untuk Break 2'
      ),

    jamPulangNormal:
      validasiFormatWaktuSetting_(
        jamPulangNormal.nilai,
        'Jam Pulang Normal'
      ),

    warningDurasiBreak1:
      durasiBreak1.warning,

    warningDurasiBreak2:
      durasiBreak2.warning,

    warningBatasAkhirBreak2:
      batasAkhirBreak2Item.warning,

    warningOverbreak:
      toleransiOverbreak.warning,

    warningReminder:
      reminder.warning,

    warningSanksi:
      sanksi.warning,

    warningKapasitas:
      maxBreakBersamaan.warning,

    warningJedaMasukKeBreak1:
      jedaMasukKeBreak1.warning,

    warningJedaBreak1KeBreak2:
      jedaBreak1KeBreak2.warning,

    warningMinimalSisaKerjaBreak2:
      minimalSisaKerjaBreak2.warning
  };
}

/**
 * Mengambil satu parameter MASTER_SETTING.
 */
function ambilItemSettingBreak_(settingMap, namaParameter) {
  var nama = normalisasiHeader_(namaParameter);

  if (
    !Object.prototype.hasOwnProperty.call(
      settingMap,
      nama
    )
  ) {
    throw new Error(
      'Parameter "' +
      namaParameter +
      '" tidak ditemukan.'
    );
  }

  return settingMap[nama];
}

/**
 * Mengambil nilai berdasarkan beberapa variasi nama.
 */
function ambilNilaiSettingBreak_(
  settingMap,
  daftarNama,
  wajib
) {
  var harusAda = wajib !== false;

  for (var index = 0;
       index < daftarNama.length;
       index++) {

    var nama =
      normalisasiHeader_(daftarNama[index]);

    if (
      Object.prototype.hasOwnProperty.call(
        settingMap,
        nama
      )
    ) {
      return settingMap[nama];
    }
  }

  if (harusAda) {
    throw new Error(
      'Parameter "' +
      daftarNama[0] +
      '" tidak ditemukan.'
    );
  }

  return null;
}

/**
 * Validasi nilai angka pada MASTER_SETTING.
 */
function validasiAngkaSetting_(
  nilai,
  namaParameter
) {
  var angka = Number(
    String(nilai)
      .replace(/\./g, '')
      .replace(/,/g, '')
      .trim()
  );

  if (
    !Number.isFinite(angka) ||
    angka < 0
  ) {
    throw new Error(
      'Nilai "' +
      namaParameter +
      '" harus berupa angka.'
    );
  }

  return angka;
}

function validasiAngkaPositifSetting_(nilai, namaParameter) {
  var angka = validasiAngkaSetting_(nilai, namaParameter);

  if (!Number.isInteger(angka) || angka < 1) {
    throw new Error(
      'Nilai "' + namaParameter + '" harus berupa bilangan bulat minimal 1.'
    );
  }

  return angka;
}

/**
 * Validasi waktu HH:mm.
 */
function validasiFormatWaktuSetting_(
  nilai,
  namaParameter
) {
  var teks =
    String(nilai || '').trim();

  var hasil =
    teks.match(/^(\d{1,2}):(\d{2})$/);

  if (!hasil) {
    throw new Error(
      'Nilai "' +
      namaParameter +
      '" harus menggunakan format HH:mm.'
    );
  }

  var jam = Number(hasil[1]);
  var menit = Number(hasil[2]);

  if (
    jam < 0 ||
    jam > 23 ||
    menit < 0 ||
    menit > 59
  ) {
    throw new Error(
      'Nilai waktu "' +
      namaParameter +
      '" tidak valid.'
    );
  }

  return (
    String(jam).padStart(2, '0') +
    ':' +
    String(menit).padStart(2, '0')
  );
}

/**
 * Mencari absensi yang mengizinkan Break Time.
 */
function cariAbsensiAktifHariIni_(
  spreadsheet,
  pin,
  tanggalHariIni,
  zonaWaktu
) {
  var sheet =
    spreadsheet.getSheetByName(
      'ABSENSI_HARIAN'
    );

  if (!sheet) {
    throw new Error(
      'Sheet ABSENSI_HARIAN tidak ditemukan.'
    );
  }

  var headerMap =
    ambilHeaderMapAbsensi_(sheet);

  var barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    return {
      ditemukan: false
    };
  }

  var data = sheet
    .getRange(
      2,
      1,
      barisTerakhir - 1,
      sheet.getLastColumn()
    )
    .getValues();

  for (var index = 0;
       index < data.length;
       index++) {

    var baris = data[index];

    var pinData = String(
      baris[headerMap['PIN']] || ''
    ).trim();

    var tanggalData =
      normalisasiTanggalAbsensi_(
        baris[headerMap['TANGGAL']],
        zonaWaktu
      );

    var statusKehadiran =
      headerMap['STATUS KEHADIRAN'] !==
        undefined
        ? normalisasiHeader_(
            baris[
              headerMap['STATUS KEHADIRAN']
            ]
          )
        : '';

    var jamMasuk =
      baris[
        headerMap['JAM MASUK']
      ];

    var jamMasukTerisi =
      jamMasuk !== '' &&
      jamMasuk !== null &&
      jamMasuk !== undefined;

    // Validasi Break tidak bergantung pada KETERANGAN
    // maupun STATUS KEHADIRAN. Status hanya informasi.
    if (
      pinData === String(pin).trim() &&
      tanggalData === tanggalHariIni &&
      jamMasukTerisi
    ) {
      return {
        ditemukan: true,
        nomorBaris: index + 2,
        statusKehadiran:
          statusKehadiran,
        jamMasuk:
          jamMasuk
      };
    }
  }

  return {
    ditemukan: false
  };
}

/**
 * Mencari LOG_ISTIRAHAT berdasarkan PIN dan tanggal.
 */
function cariLogIstirahatHariIni_(
  sheet,
  headerMap,
  pin,
  tanggalHariIni,
  zonaWaktu
) {
  var barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    return {
      ditemukan: false
    };
  }

  var data = sheet
    .getRange(
      2,
      1,
      barisTerakhir - 1,
      sheet.getLastColumn()
    )
    .getValues();

  for (var index = 0;
       index < data.length;
       index++) {

    var baris = data[index];

    var pinData = String(
      baris[headerMap['PIN']] || ''
    ).trim();

    var tanggalData =
      normalisasiTanggalAbsensi_(
        baris[headerMap['TANGGAL']],
        zonaWaktu
      );

    if (
      pinData === String(pin).trim() &&
      tanggalData === tanggalHariIni
    ) {
      return {
        ditemukan: true,
        nomorBaris: index + 2,
        data: baris
      };
    }
  }

  return {
    ditemukan: false
  };
}

/**
 * Membuat satu baris LOG_ISTIRAHAT baru.
 */
function buatBarisLogIstirahatBaru_(
  sheet,
  headerMap,
  pengguna,
  sekarang,
  zonaWaktu,
  tanggalOperasional
) {
  var nomorBaris =
    sheet.getLastRow() + 1;

  var jumlahKolom =
    sheet.getLastColumn();

  var output =
    new Array(jumlahKolom).fill('');

  var idLog = buatIdLogIstirahat_(
    pengguna.pin,
    sekarang,
    zonaWaktu
  );

  var dataBaru = {
    'ID': idLog,
    // Identitas tanggal LOG selalu memakai tanggal operasional
    // (tanggal server asli), bukan tanggal waktu virtual Mode Uji.
    'TANGGAL':
      buatTanggalOperasionalBreak_(
        tanggalOperasional,
        sekarang
      ),
    'PIN': pengguna.pin,
    'NAMA SA': pengguna.namaSA,
    'OUTLET': pengguna.outlet,
    'TOTAL OVERTIME': 0,
    'TOTAL SANKSI': 0
  };

  Object.keys(dataBaru).forEach(
    namaKolom => {
      output[headerMap[namaKolom]] =
        dataBaru[namaKolom];
    }
  );

  sheet
    .getRange(
      nomorBaris,
      1,
      1,
      jumlahKolom
    )
    .setValues([output]);

  sheet
    .getRange(
      nomorBaris,
      headerMap['TANGGAL'] + 1
    )
    .setNumberFormat('dd/MM/yyyy');

  sheet
    .getRange(
      nomorBaris,
      headerMap['PIN'] + 1
    )
    .setNumberFormat('@');

  return nomorBaris;
}


/**
 * Membuat nilai tanggal LOG_ISTIRAHAT berdasarkan tanggal operasional.
 * Mode Uji boleh melewati tengah malam, tetapi baris LOG tetap terikat
 * pada tanggal server asli saat transaksi dimulai.
 *
 * @param {string} tanggalOperasional yyyy-MM-dd
 * @param {Date} fallbackDate
 * @return {Date}
 */
function buatTanggalOperasionalBreak_(
  tanggalOperasional,
  fallbackDate
) {
  var teks = String(
    tanggalOperasional || ''
  ).trim();

  var cocok = teks.match(
    /^(\d{4})-(\d{2})-(\d{2})$/
  );

  if (cocok) {
    return new Date(
      Number(cocok[1]),
      Number(cocok[2]) - 1,
      Number(cocok[3]),
      12,
      0,
      0,
      0
    );
  }

  return buatTanggalTanpaJam_(
    fallbackDate instanceof Date
      ? fallbackDate
      : new Date()
  );
}


/**
 * Validasi sebelum mulai Break 1.
 */
function validasiMulaiBreak1_(
  dataLog,
  headerMap,
  hasilAbsensi,
  sekarang,
  pengaturan,
  zonaWaktu
) {
  var baris = dataLog.ditemukan
    ? dataLog.data
    : null;

  var s1Mulai = baris
    ? baris[headerMap['S1 MULAI']]
    : null;

  var s1Selesai = baris
    ? baris[headerMap['S1 SELESAI']]
    : null;

  var s2Mulai = baris
    ? baris[headerMap['S2 MULAI']]
    : null;

  if (s1Mulai instanceof Date && !(s1Selesai instanceof Date)) {
    return {
      success: false,
      code: 'BREAK_1_SEDANG_BERJALAN',
      message: 'Break 1 sedang berlangsung.'
    };
  }

  if (s1Mulai instanceof Date) {
    return {
      success: false,
      code: 'BREAK_1_SUDAH_DIGUNAKAN',
      message: 'Break 1 sudah digunakan hari ini.'
    };
  }

  if (s2Mulai instanceof Date) {
    return {
      success: false,
      code: 'URUTAN_BREAK_TIDAK_VALID',
      message: 'Break 1 tidak dapat dimulai setelah Break 2.'
    };
  }

  return validasiJedaAktivitasBreak_(
    hasilAbsensi.jamMasuk,
    sekarang,
    pengaturan.jedaMasukKeBreak1,
    zonaWaktu,
    'BREAK_1_BELUM_TERSEDIA',
    pengaturan.warningJedaMasukKeBreak1 ||
      'ANDA BARU MASUK KERJA. LANJUTKAN BEKERJA TERLEBIH DAHULU. BREAK 1 DAPAT DIMULAI PUKUL {{JAM}} (TERSISA {{SISA}} MENIT).'
  );
}

/**
 * Validasi sebelum mulai Break 2.
 */
function validasiMulaiBreak2_(
  dataLog,
  headerMap,
  sekarang,
  pengaturan,
  zonaWaktu
) {
  if (!dataLog.ditemukan) {
    return {
      success: false,
      code: 'BREAK_1_BELUM_DIMULAI',
      message: 'Selesaikan Break 1 terlebih dahulu.'
    };
  }

  var baris = dataLog.data;
  var s1Mulai = baris[headerMap['S1 MULAI']];
  var s1Selesai = baris[headerMap['S1 SELESAI']];
  var s2Mulai = baris[headerMap['S2 MULAI']];
  var s2Selesai = baris[headerMap['S2 SELESAI']];

  if (!(s1Mulai instanceof Date)) {
    return {
      success: false,
      code: 'BREAK_1_BELUM_DIMULAI',
      message: 'Break 1 belum dimulai.'
    };
  }

  if (!(s1Selesai instanceof Date)) {
    return {
      success: false,
      code: 'BREAK_1_BELUM_SELESAI',
      message: 'Selesaikan Break 1 terlebih dahulu.'
    };
  }

  if (s2Mulai instanceof Date && !(s2Selesai instanceof Date)) {
    return {
      success: false,
      code: 'BREAK_2_SEDANG_BERJALAN',
      message: 'Break 2 sedang berlangsung.'
    };
  }

  if (s2Mulai instanceof Date) {
    return {
      success: false,
      code: 'BREAK_2_SUDAH_DIGUNAKAN',
      message: 'Break 2 sudah digunakan hari ini.'
    };
  }

  if (
    cekMencapaiAtauMelewatiBatasWaktu_(
      sekarang,
      pengaturan.batasAkhirBreak2,
      zonaWaktu
    )
  ) {
    return {
      success: false,
      code: 'BATAS_AKHIR_BREAK_2',
      message:
        pengaturan.warningBatasAkhirBreak2 ||
        'Break 2 tidak dapat dimulai karena telah mencapai batas akhir.'
    };
  }

  var validasiSisaKerja =
    validasiMinimalSisaKerjaBreak2_(
      sekarang,
      pengaturan,
      zonaWaktu
    );

  if (!validasiSisaKerja.success) {
    return validasiSisaKerja;
  }

  return validasiJedaAktivitasBreak_(
    s1Selesai,
    sekarang,
    pengaturan.jedaBreak1KeBreak2,
    zonaWaktu,
    'BREAK_2_BELUM_TERSEDIA',
    pengaturan.warningJedaBreak1KeBreak2 ||
      'ANDA BARU MENYELESAIKAN BREAK 1. LANJUTKAN BEKERJA TERLEBIH DAHULU. BREAK 2 DAPAT DIMULAI PUKUL {{JAM}} (TERSISA {{SISA}} MENIT).'
  );
}

/**
 * Memastikan jeda minimum sejak aktivitas sebelumnya telah terpenuhi.
 */
function validasiJedaAktivitasBreak_(
  waktuDasar,
  sekarang,
  jedaMenit,
  zonaWaktu,
  code,
  templatePesan
) {
  var waktuDasarNormal =
    normalisasiWaktuAktivitasBreak_(
      waktuDasar,
      sekarang,
      zonaWaktu
    );

  if (!waktuDasarNormal) {
    return {
      success: false,
      code: 'WAKTU_DASAR_BREAK_TIDAK_VALID',
      message:
        'Waktu aktivitas sebelumnya tidak valid. ' +
        'Silakan hubungi Administrator.'
    };
  }

  var tersediaPada = new Date(
    waktuDasarNormal.getTime() +
    jedaMenit * 60000
  );

  var sisaMilidetik =
    tersediaPada.getTime() - sekarang.getTime();

  if (sisaMilidetik <= 0) {
    return {
      success: true,
      data: {
        tersediaPada:
          Utilities.formatDate(
            tersediaPada,
            zonaWaktu,
            'HH:mm:ss'
          ),
        sisaTungguMenit: 0
      }
    };
  }

  var sisaTungguMenit = Math.ceil(
    sisaMilidetik / 60000
  );

  var pesan = gantiPlaceholderBreak_(
    templatePesan,
    {
      JEDA: jedaMenit,
      JAM: Utilities.formatDate(
        tersediaPada,
        zonaWaktu,
        'HH:mm'
      ),
      SISA: sisaTungguMenit
    }
  );

  return {
    success: false,
    code: code,
    message:
      pesan +
      ' Tersedia pukul ' +
      Utilities.formatDate(
        tersediaPada,
        zonaWaktu,
        'HH:mm'
      ) +
      ' (tersisa ' +
      sisaTungguMenit +
      ' menit).',
    data: {
      tersediaPada:
        Utilities.formatDate(
          tersediaPada,
          zonaWaktu,
          'HH:mm:ss'
        ),
      sisaTungguMenit: sisaTungguMenit
    }
  };
}

/**
 * Menormalisasi nilai waktu Sheet ke tanggal kerja hari ini.
 */
function normalisasiWaktuAktivitasBreak_(
  nilai,
  sekarang,
  zonaWaktu
) {
  var teksWaktu = '';

  if (
    nilai instanceof Date &&
    Number.isFinite(nilai.getTime())
  ) {
    var tanggalNilai = Utilities.formatDate(
      nilai,
      zonaWaktu,
      'yyyy-MM-dd'
    );

    var tanggalSekarang = Utilities.formatDate(
      sekarang,
      zonaWaktu,
      'yyyy-MM-dd'
    );

    if (tanggalNilai === tanggalSekarang) {
      return new Date(nilai.getTime());
    }

    teksWaktu = Utilities.formatDate(
      nilai,
      zonaWaktu,
      'HH:mm:ss'
    );
  } else {
    teksWaktu = String(nilai || '').trim();
  }

  var hasil = teksWaktu.match(
    /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
  );

  if (!hasil) return null;

  var jam = Number(hasil[1]);
  var menit = Number(hasil[2]);
  var detik = Number(hasil[3] || 0);

  if (
    jam < 0 || jam > 23 ||
    menit < 0 || menit > 59 ||
    detik < 0 || detik > 59
  ) {
    return null;
  }

  var bagianSekarang = Utilities
    .formatDate(
      sekarang,
      zonaWaktu,
      'HH:mm:ss'
    )
    .split(':')
    .map(Number);

  var detikSekarang =
    bagianSekarang[0] * 3600 +
    bagianSekarang[1] * 60 +
    bagianSekarang[2];

  var detikTarget =
    jam * 3600 + menit * 60 + detik;

  var selisihDetik =
    detikTarget - detikSekarang;

  // Mode Uji dapat melewati tengah malam. Jika jam aktivitas
  // terlihat lebih besar daripada jam virtual sekarang, aktivitas
  // tersebut adalah kejadian pada hari sebelumnya, bukan masa depan.
  if (selisihDetik > 0) {
    selisihDetik -= 24 * 60 * 60;
  }

  return new Date(
    sekarang.getTime() +
    selisihDetik * 1000
  );
}

/**
 * Memblokir Break 2 bila sisa jam kerja di bawah parameter minimum.
 */
function validasiMinimalSisaKerjaBreak2_(
  sekarang,
  pengaturan,
  zonaWaktu
) {
  var menitSekarang = ubahWaktuKeMenit_(
    Utilities.formatDate(
      sekarang,
      zonaWaktu,
      'HH:mm'
    )
  );

  var menitPulang = ubahWaktuKeMenit_(
    pengaturan.jamPulangNormal
  );

  var sisaKerjaMenit =
    menitPulang - menitSekarang;

  if (
    sisaKerjaMenit >=
    pengaturan.minimalSisaKerjaBreak2
  ) {
    return {
      success: true,
      data: {
        sisaKerjaMenit: sisaKerjaMenit
      }
    };
  }

  var template =
    pengaturan.warningMinimalSisaKerjaBreak2 ||
    'Break 2 tidak dapat dimulai karena sisa jam kerja kurang dari {{MINIMAL}} menit.';

  return {
    success: false,
    code: 'SISA_KERJA_BREAK_2_TIDAK_CUKUP',
    message: gantiPlaceholderBreak_(
      template,
      {
        MINIMAL:
          pengaturan.minimalSisaKerjaBreak2,
        SISA: Math.max(0, sisaKerjaMenit)
      }
    ),
    data: {
      sisaKerjaMenit:
        Math.max(0, sisaKerjaMenit)
    }
  };
}

/**
 * Mengecek status kehadiran setengah hari.
 */
function apakahSetengahHari_(statusKehadiran) {
  var status = normalisasiHeader_(statusKehadiran);
  return ['1/2 HARI', 'HADIR 1/2 HARI'].includes(status);
}

/**
 * Membaca kapasitas break outlet.
 * Mendukung checkbox serta nilai TRUE, AKTIF, YA, YES, dan 1.
 */
function ambilKapasitasOutlet_(spreadsheet, outletSA) {
  var sheet = spreadsheet.getSheetByName('MASTER_OUTLET');

  if (!sheet) {
    throw new Error('Sheet MASTER_OUTLET tidak ditemukan.');
  }

  var jumlahKolom = sheet.getLastColumn();

  if (jumlahKolom === 0) {
    throw new Error('Header MASTER_OUTLET belum tersedia.');
  }

  var header = sheet
    .getRange(1, 1, 1, jumlahKolom)
    .getDisplayValues()[0]
    .map(normalisasiHeader_);

  var indexKode = header.indexOf('KODE');
  var indexNama = header.indexOf('NAMA OUTLET');
  var indexMax = header.indexOf('MAX ISTIRAHAT');
  var indexAktif = header.indexOf('AKTIF');

  var kolomTidakAda = [];

  if (indexKode === -1) kolomTidakAda.push('KODE');
  if (indexNama === -1) kolomTidakAda.push('NAMA OUTLET');
  if (indexMax === -1) kolomTidakAda.push('MAX ISTIRAHAT');
  if (indexAktif === -1) kolomTidakAda.push('AKTIF');

  if (kolomTidakAda.length > 0) {
    throw new Error(
      'Kolom berikut tidak ditemukan di MASTER_OUTLET: ' +
      kolomTidakAda.join(', ')
    );
  }

  var barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    throw new Error('MASTER_OUTLET masih kosong.');
  }

  var data = sheet
    .getRange(2, 1, barisTerakhir - 1, jumlahKolom)
    .getValues();

  var outletDicari = normalisasiHeader_(outletSA);

  for (var index = 0; index < data.length; index++) {
    var baris = data[index];
    var kode = normalisasiHeader_(baris[indexKode]);
    var namaOutlet = normalisasiHeader_(baris[indexNama]);

    if (kode !== outletDicari && namaOutlet !== outletDicari) {
      continue;
    }

    if (!statusAktifGtt_(baris[indexAktif])) {
      throw new Error(
        'Outlet ' +
        String(baris[indexKode] || outletSA).trim() +
        ' berstatus tidak aktif.'
      );
    }

    var maxIstirahat = Number(baris[indexMax]);

    if (!Number.isInteger(maxIstirahat) || maxIstirahat < 1) {
      throw new Error(
        'MAX ISTIRAHAT Outlet ' +
        String(baris[indexKode] || outletSA).trim() +
        ' tidak valid.'
      );
    }

    return {
      kode: String(baris[indexKode] || '').trim(),
      namaOutlet: String(baris[indexNama] || '').trim(),
      maxIstirahat: maxIstirahat,
      aktif: true
    };
  }

  throw new Error(
    'Outlet "' + outletSA + '" tidak ditemukan di MASTER_OUTLET.'
  );
}

/**
 * Mengecek apakah nilai status berarti aktif.
 */
function statusAktifGtt_(nilai) {
  if (nilai === true) return true;

  if (nilai === false || nilai === null || nilai === undefined) {
    return false;
  }

  var teks = String(nilai).trim().toUpperCase();

  return ['TRUE', 'AKTIF', 'YA', 'YES', '1'].includes(teks);
}

/**
 * Menghitung jumlah SA yang sedang break pada outlet dan tanggal yang sama.
 */
function hitungSedangBreakOutlet_(
  sheet,
  headerMap,
  outlet,
  tanggalHariIni,
  zonaWaktu
) {
  var barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) return 0;

  var data = sheet
    .getRange(2, 1, barisTerakhir - 1, sheet.getLastColumn())
    .getValues();

  var outletDicari = normalisasiHeader_(outlet);
  var jumlah = 0;

  data.forEach(baris => {
    var tanggalData = normalisasiTanggalAbsensi_(
      baris[headerMap['TANGGAL']],
      zonaWaktu
    );

    var outletData = normalisasiHeader_(
      baris[headerMap['OUTLET']]
    );

    if (tanggalData !== tanggalHariIni || outletData !== outletDicari) {
      return;
    }

    var s1Aktif =
      baris[headerMap['S1 MULAI']] instanceof Date &&
      !(baris[headerMap['S1 SELESAI']] instanceof Date);

    var s2Aktif =
      baris[headerMap['S2 MULAI']] instanceof Date &&
      !(baris[headerMap['S2 SELESAI']] instanceof Date);

    if (s1Aktif || s2Aktif) jumlah++;
  });

  return jumlah;
}

/**
 * Menghitung durasi break dalam menit.
 */
function hitungDurasiMenit_(waktuMulai, waktuSelesai) {
  var selisih = waktuSelesai.getTime() - waktuMulai.getTime();

  if (selisih < 0) {
    throw new Error('Waktu selesai lebih kecil dari waktu mulai.');
  }

  return Math.ceil(selisih / 60000);
}

/**
 * Menentukan status overtime dan potensi sanksi.
 */
function evaluasiOverbreak_(
  lewatDurasiMenit,
  overbreakMenit,
  toleransiMenit,
  sanksiPerBreak
) {
  if (lewatDurasiMenit <= 0) {
    return {
      status: 'NORMAL',
      sanksi: 0
    };
  }

  if (lewatDurasiMenit <= toleransiMenit) {
    return {
      status: 'TOLERANSI',
      sanksi: 0
    };
  }

  return {
    status: 'OVERBREAK',
    sanksi: sanksiPerBreak,
    overbreakMenit: overbreakMenit
  };
}

/**
 * Membuat pesan hasil break berdasarkan MASTER_SETTING.
 */
function buatPesanHasilBreak_(
  evaluasi,
  overbreakMenit,
  pengaturan
) {
  if (evaluasi.status !== 'OVERBREAK') {
    return '';
  }

  var pesanOverbreak = gantiPlaceholderBreak_(
    pengaturan.warningOverbreak ||
      'Anda sudah OVERBREAK selama {{MENIT}} menit. Segera akhiri break dan kembali bekerja.',
    {
      MENIT: overbreakMenit
    }
  );

  var pesanSanksi = gantiPlaceholderBreak_(
    pengaturan.warningSanksi ||
      'Anda dikenakan sanksi overbreak sebesar Rp{{RUPIAH}} untuk break ini.',
    {
      RUPIAH: formatRupiahBreak_(
        pengaturan.sanksiOverbreak
      )
    }
  );

  return pesanOverbreak + ' ' + pesanSanksi;
}

/**
 * Mengganti placeholder {{NAMA}} pada warning.
 */
function gantiPlaceholderBreak_(template, data) {
  var hasil = String(template || '');

  Object.keys(data).forEach(kunci => {
    var pola = new RegExp(
      '\\{\\{' + kunci + '\\}\\}',
      'g'
    );

    hasil = hasil.replace(
      pola,
      String(data[kunci])
    );
  });

  return hasil;
}

/**
 * Format angka rupiah tanpa simbol.
 */
function formatRupiahBreak_(nilai) {
  return Number(nilai || 0)
    .toLocaleString('id-ID');
}

/**
 * Memperbarui total overtime dan total sanksi.
 */
function perbaruiTotalIstirahat_(sheet, nomorBaris, headerMap) {
  var baris = sheet
    .getRange(nomorBaris, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  var s1Overtime = Number(baris[headerMap['S1 OVERTIME']]) || 0;
  var s2Overtime = Number(baris[headerMap['S2 OVERTIME']]) || 0;
  var s1Sanksi = Number(baris[headerMap['S1 SANKSI']]) || 0;
  var s2Sanksi = Number(baris[headerMap['S2 SANKSI']]) || 0;

  sheet
    .getRange(nomorBaris, headerMap['TOTAL OVERTIME'] + 1)
    .setValue(s1Overtime + s2Overtime)
    .setNumberFormat('0');

  sheet
    .getRange(nomorBaris, headerMap['TOTAL SANKSI'] + 1)
    .setValue(s1Sanksi + s2Sanksi)
    .setNumberFormat('#,##0');
}

/**
 * Mengecek apakah jam mulai melewati batas waktu break.
 */
function cekMencapaiAtauMelewatiBatasWaktu_(
  waktuSekarang,
  batasWaktu,
  zonaWaktu
) {
  var waktuTeks = Utilities.formatDate(
    waktuSekarang,
    zonaWaktu,
    'HH:mm'
  );

  return (
    ubahWaktuKeMenit_(waktuTeks) >=
    ubahWaktuKeMenit_(batasWaktu)
  );
}

/**
 * Mengubah waktu HH:mm menjadi total menit.
 */
function ubahWaktuKeMenit_(nilaiWaktu) {
  var teks = String(nilaiWaktu || '').trim();
  var hasil = teks.match(/^(\d{1,2}):(\d{2})$/);

  if (!hasil) {
    throw new Error('Format waktu tidak valid: ' + teks + '. Gunakan HH:mm.');
  }

  var jam = Number(hasil[1]);
  var menit = Number(hasil[2]);

  if (jam < 0 || jam > 23 || menit < 0 || menit > 59) {
    throw new Error('Nilai waktu tidak valid: ' + teks + '.');
  }

  return jam * 60 + menit;
}

/**
 * Membuat ID LOG_ISTIRAHAT unik.
 */
function buatIdLogIstirahat_(pin, waktu, zonaWaktu) {
  var timestamp = Utilities.formatDate(
    waktu,
    zonaWaktu,
    'yyyyMMddHHmmss'
  );

  var acak = Utilities
    .getUuid()
    .replace(/-/g, '')
    .substring(0, 6)
    .toUpperCase();

  return ['BRK', timestamp, String(pin).trim(), acak].join('-');
}

/**
 * Mengembalikan status operasional Break Time hari ini.
 */
function ambilStatusBreakHariIni(pinInput) {
  try {
    var konteks = ambilKonteksBreak_(pinInput);

    if (!konteks.success) return konteks;

    var {
      pengguna,
      spreadsheet,
      sheetLog,
      headerMap,
      zonaWaktu,
      tanggalHariIni,
      sekarang,
      pengaturan
    } = konteks.data;

    var hasilAbsensi = cariAbsensiAktifHariIni_(
      spreadsheet,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    if (!hasilAbsensi.ditemukan) {
      return {
        success: true,
        code: 'STATUS_BREAK_DITEMUKAN',
        data: {
          statusOperasional: 'BELUM ABSEN',
          tombolBerikutnya: 'ABSEN MASUK'
        }
      };
    }

    var dataLog = cariLogIstirahatHariIni_(
      sheetLog,
      headerMap,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    if (!dataLog.ditemukan) {
      var validasiBreak1 =
        validasiMulaiBreak1_(
          dataLog,
          headerMap,
          hasilAbsensi,
          sekarang,
          pengaturan,
          zonaWaktu
        );

      return buatStatusMulaiBreak_(
        'MULAI BREAK 1',
        validasiBreak1,
        sekarang,
        zonaWaktu,
        {
          break1Dilewati: false
        }
      );
    }

    var baris = dataLog.data;
    var s1Mulai = baris[headerMap['S1 MULAI']];
    var s1Selesai = baris[headerMap['S1 SELESAI']];
    var s2Mulai = baris[headerMap['S2 MULAI']];
    var s2Selesai = baris[headerMap['S2 SELESAI']];

    if (s1Mulai instanceof Date && !(s1Selesai instanceof Date)) {
      var statusWaktu = buatStatusWaktuBreak_(
        s1Mulai,
        1,
        konteks.data.sekarang,
        konteks.data.pengaturan
      );

      return {
        success: true,
        code: 'STATUS_BREAK_DITEMUKAN',
        data: Object.assign({
          statusOperasional: 'BREAK 1',
          tombolBerikutnya: 'SELESAI BREAK 1',
          serverTime: konteks.data.sekarang.toISOString()
        }, statusWaktu)
      };
    }

    if (s1Selesai instanceof Date && !(s2Mulai instanceof Date)) {
      var validasiBreak2 =
        validasiMulaiBreak2_(
          dataLog,
          headerMap,
          sekarang,
          pengaturan,
          zonaWaktu
        );

      return buatStatusMulaiBreak_(
        'MULAI BREAK 2',
        validasiBreak2,
        sekarang,
        zonaWaktu
      );
    }

    if (s2Mulai instanceof Date && !(s2Selesai instanceof Date)) {
      var statusWaktu = buatStatusWaktuBreak_(
        s2Mulai,
        2,
        konteks.data.sekarang,
        konteks.data.pengaturan
      );

      return {
        success: true,
        code: 'STATUS_BREAK_DITEMUKAN',
        data: Object.assign({
          statusOperasional: 'BREAK 2',
          tombolBerikutnya: 'SELESAI BREAK 2',
          serverTime: konteks.data.sekarang.toISOString()
        }, statusWaktu)
      };
    }

    if (s2Selesai instanceof Date) {
      return {
        success: true,
        code: 'STATUS_BREAK_DITEMUKAN',
        data: {
          statusOperasional: 'AKTIF BEKERJA',
          tombolBerikutnya: 'ABSEN PULANG',
          seluruhBreakSelesai: true
        }
      };
    }

    var validasiBreak1 =
      validasiMulaiBreak1_(
        dataLog,
        headerMap,
        hasilAbsensi,
        sekarang,
        pengaturan,
        zonaWaktu
      );

    return buatStatusMulaiBreak_(
      'MULAI BREAK 1',
      validasiBreak1,
      sekarang,
      zonaWaktu,
      {
        break1Dilewati: false
      }
    );
  } catch (error) {
    console.error(error);

    return {
      success: false,
      code: 'SYSTEM_ERROR',
      message: 'Terjadi kesalahan sistem: ' + error.message
    };
  }
}

/**
 * Membentuk status tombol Mulai Break beserta alasan bila masih terkunci.
 */
function buatStatusMulaiBreak_(
  tombolBerikutnya,
  validasi,
  sekarang,
  zonaWaktu,
  dataTambahan
) {
  var hasilValidasi = validasi || {
    success: false,
    message: 'Status Break belum dapat ditentukan.'
  };

  var dataValidasi =
    hasilValidasi.data || {};

  return {
    success: true,
    code: 'STATUS_BREAK_DITEMUKAN',
    data: Object.assign(
      {
        statusOperasional: 'AKTIF BEKERJA',
        tombolBerikutnya: tombolBerikutnya,
        aksiDiizinkan:
          hasilValidasi.success === true,
        alasanTerkunci:
          hasilValidasi.success === true
            ? ''
            : String(
                hasilValidasi.message ||
                'Break belum tersedia.'
              ).trim(),
        kodeTerkunci:
          hasilValidasi.success === true
            ? ''
            : String(
                hasilValidasi.code || ''
              ).trim(),
        serverTime: sekarang.toISOString()
      },
      dataValidasi,
      dataTambahan || {}
    )
  };
}

/**
 * Menghasilkan countdown, reminder, toleransi, dan warning overbreak.
 */
function buatStatusWaktuBreak_(
  waktuMulai,
  nomorBreak,
  sekarang,
  pengaturan
) {
  var batasDurasi =
    nomorBreak === 1
      ? pengaturan.durasiBreak1
      : pengaturan.durasiBreak2;

  var durasiBerjalan = hitungDurasiMenit_(
    waktuMulai,
    sekarang
  );

  var sisaNormal = batasDurasi - durasiBerjalan;
  var lewatDurasi = Math.max(
    0,
    durasiBerjalan - batasDurasi
  );
  var overbreakMenit = Math.max(
    0,
    lewatDurasi - pengaturan.toleransiOverbreak
  );

  var fase = 'NORMAL';
  var warning = '';

  if (
    sisaNormal > 0 &&
    sisaNormal <= pengaturan.reminderSebelumBerakhir
  ) {
    fase = 'REMINDER';
    warning = gantiPlaceholderBreak_(
      pengaturan.warningReminder ||
        'Waktu break Anda akan berakhir dalam {{MENIT}} menit. Mohon segera bersiap untuk kembali bekerja.',
      {
        MENIT: sisaNormal
      }
    );
  } else if (
    lewatDurasi > 0 &&
    overbreakMenit === 0
  ) {
    fase = 'TOLERANSI';
    warning =
      nomorBreak === 1
        ? pengaturan.warningDurasiBreak1
        : pengaturan.warningDurasiBreak2;
  } else if (overbreakMenit > 0) {
    fase = 'OVERBREAK';
    warning = buatPesanHasilBreak_(
      {
        status: 'OVERBREAK',
        sanksi: pengaturan.sanksiOverbreak
      },
      overbreakMenit,
      pengaturan
    );
  }

  return {
    faseBreak: fase,
    durasiBerjalanMenit: durasiBerjalan,
    sisaNormalMenit: Math.max(0, sisaNormal),
    overbreakMenit: overbreakMenit,
    warning: warning
  };
}

/**
 * =====================================================
 * FUNGSI PENGUJIAN
 * =====================================================
 */

function ujiStatusBreakHariIni() {
  var PIN_UJI = '8415';
  var hasil = ambilStatusBreakHariIni(PIN_UJI);
  console.log(JSON.stringify(hasil, null, 2));
}

function ujiMulaiBreak1() {
  var PIN_UJI = '8415';
  var hasil = mulaiBreak1(PIN_UJI);
  console.log(JSON.stringify(hasil, null, 2));
}

function ujiSelesaiBreak1() {
  var PIN_UJI = '8415';
  var hasil = selesaiBreak1(PIN_UJI);
  console.log(JSON.stringify(hasil, null, 2));
}

function ujiMulaiBreak2() {
  var PIN_UJI = '9616';
  var hasil = mulaiBreak2(PIN_UJI);
  console.log(JSON.stringify(hasil, null, 2));
}

function ujiSelesaiBreak2() {
  var PIN_UJI = '8415';
  var hasil = selesaiBreak2(PIN_UJI);
  console.log(JSON.stringify(hasil, null, 2));
}

/**
 * Tidak mengubah data; hanya menguji MASTER_OUTLET.
 */
function ujiKapasitasOutlet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hasil = ambilKapasitasOutlet_(spreadsheet, 'GP');
  console.log(JSON.stringify(hasil, null, 2));
}


/**
 * =====================================================
 * SELF-TEST BREAK ENGINE 2.2.0
 * Tidak menulis transaksi. Aman dijalankan dari editor.
 * =====================================================
 */
function ujiBreakTimeGsLengkap() {
  var hasil = {
    success: false,
    code: 'UJI_BREAKTIME_FAIL',
    tests: {},
    pengaturan: null
  };

  try {
    var fungsiWajib = [
      'mulaiBreak1',
      'selesaiBreak1',
      'mulaiBreak2',
      'selesaiBreak2',
      'ambilStatusBreakHariIni',
      'ambilKonteksBreak_',
      'ambilPengaturanBreak_',
      'validasiMulaiBreak1_',
      'validasiMulaiBreak2_',
      'validasiJedaAktivitasBreak_',
      'validasiMinimalSisaKerjaBreak2_'
    ];

    var fungsiTersedia = true;
    fungsiWajib.forEach(function (namaFungsi) {
      var tersedia = typeof this[namaFungsi] === 'function';
      hasil.tests['fungsi_' + namaFungsi] = tersedia;
      if (!tersedia) fungsiTersedia = false;
    }, this);

    var pengaturan = ambilPengaturanBreak_();
    hasil.pengaturan = pengaturan;

    hasil.tests.jedaMasukKeBreak1Valid =
      Number(pengaturan.jedaMasukKeBreak1) === 120;

    hasil.tests.jedaBreak1KeBreak2Valid =
      Number(pengaturan.jedaBreak1KeBreak2) === 120;

    hasil.tests.durasiBreak1Valid =
      Number(pengaturan.durasiBreak1) > 0;

    hasil.tests.durasiBreak2Valid =
      Number(pengaturan.durasiBreak2) > 0;

    hasil.tests.maxBreakBersamaanValid =
      Number(pengaturan.maxBreakBersamaan) > 0;

    hasil.tests.minimalSisaKerjaBreak2Valid =
      Number(pengaturan.minimalSisaKerjaBreak2) >= 0;

    var zonaWaktu = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSpreadsheetTimeZone();

    var dasar = new Date('2026-08-02T08:00:00+08:00');
    var sebelum = new Date('2026-08-02T09:59:00+08:00');
    var tepat = new Date('2026-08-02T10:00:00+08:00');

    var hasilSebelum = validasiJedaAktivitasBreak_(
      dasar,
      sebelum,
      120,
      zonaWaktu,
      'TEST_BELUM_TERSEDIA',
      'Tersedia pukul {{JAM}} tersisa {{SISA}} menit.'
    );

    var hasilTepat = validasiJedaAktivitasBreak_(
      dasar,
      tepat,
      120,
      zonaWaktu,
      'TEST_BELUM_TERSEDIA',
      'Tersedia pukul {{JAM}} tersisa {{SISA}} menit.'
    );

    hasil.tests.jedaSebelum120MenitDitolak =
      hasilSebelum.success === false &&
      hasilSebelum.code === 'TEST_BELUM_TERSEDIA';

    hasil.tests.jedaTepat120MenitDiterima =
      hasilTepat.success === true;

    var semuaLulus = fungsiTersedia;
    Object.keys(hasil.tests).forEach(function (namaTest) {
      if (hasil.tests[namaTest] !== true) semuaLulus = false;
    });

    hasil.success = semuaLulus;
    hasil.code = semuaLulus
      ? 'UJI_BREAKTIME_PASS'
      : 'UJI_BREAKTIME_FAIL';

    console.log(JSON.stringify(hasil, null, 2));
    return hasil;
  } catch (error) {
    hasil.error = error && error.message
      ? error.message
      : String(error);
    console.error(JSON.stringify(hasil, null, 2));
    return hasil;
  }
}
