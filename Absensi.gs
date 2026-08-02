/**
 * =====================================================
 * GRAND TIME TRACKER — GTT
 * Module : GTT-05 Attendance Engine
 * Version: 1.3.0
 * Status : Reviewed Candidate — ABSENSI CORE
 * =====================================================
 */


/**
 * Menyimpan absensi masuk berdasarkan PIN.
 *
 * @param {string|number} pinInput PIN SA.
 * @return {Object} Hasil proses absensi.
 */
function simpanAbsensiMasuk(pinInput) {
  const SHEET_ABSENSI = 'ABSENSI_HARIAN';

  const lock = LockService.getScriptLock();

  try {
    // Mencegah dua proses absensi berjalan bersamaan.
    lock.waitLock(10000);

    const hasilLogin = validasiLoginPin(pinInput);

    if (!hasilLogin.success) {
      return hasilLogin;
    }

    const pengguna = absensiNormalisasiPengguna_(
      hasilLogin.data,
      pinInput
    );

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAbsensi = spreadsheet.getSheetByName(SHEET_ABSENSI);

    if (!sheetAbsensi) {
      return {
        success: false,
        code: 'SHEET_ABSENSI_TIDAK_DITEMUKAN',
        message: 'Sheet ABSENSI_HARIAN tidak ditemukan.'
      };
    }

    const zonaWaktu = spreadsheet.getSpreadsheetTimeZone();
    const sekarang = gttSekarang_(pengguna.pin);

    const tanggalHariIni = Utilities.formatDate(
      sekarang,
      zonaWaktu,
      'yyyy-MM-dd'
    );

    const headerMap = ambilHeaderMapAbsensi_(sheetAbsensi);

    const hasilCekDouble = cariAbsensiHariIni_(
      sheetAbsensi,
      headerMap,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    // Data pertama tetap digunakan.
    if (hasilCekDouble.ditemukan) {
      return {
        success: false,
        code: 'SUDAH_ABSEN',
        message:
          'Anda sudah melakukan absensi hari ini pada pukul ' +
          hasilCekDouble.jamMasuk +
          '. Data absensi pertama tetap digunakan.',
        data: hasilCekDouble.data
      };
    }

    const pengaturan = absensiAmbilPengaturan_();

    const hasilStatus = hitungStatusAbsensi_(
      sekarang,
      pengaturan,
      zonaWaktu
    );

    const idAbsensi = buatIdAbsensi_(
      pengguna.pin,
      sekarang,
      zonaWaktu
    );

    absensiValidasiDataPengguna_(pengguna);

    const dataBaru = {
      'ID': idAbsensi,
      'TANGGAL': buatTanggalTanpaJam_(sekarang),
      'NAMA SA': pengguna.namaSA,
      'PIN': pengguna.pin,
      'OUTLET': pengguna.outlet,
      'JAM MASUK': sekarang,
      'STATUS JAM MASUK': hasilStatus.statusJamMasuk,
      'TERLAMBAT MENIT': hasilStatus.terlambatMenit,
      'JAM PULANG': '',
      'STATUS JAM PULANG': '',
      'STATUS KEHADIRAN': hasilStatus.statusKehadiran,
      'KETERANGAN': hasilStatus.keterangan || ''
    };

    const nomorBarisBaru = sheetAbsensi.getLastRow() + 1;
    const jumlahKolom = sheetAbsensi.getLastColumn();
    const barisOutput = new Array(jumlahKolom).fill('');

    Object.keys(dataBaru).forEach(namaKolom => {
      const indexKolom = headerMap[namaKolom];

      if (indexKolom !== undefined) {
        barisOutput[indexKolom] = dataBaru[namaKolom];
      }
    });

    sheetAbsensi
      .getRange(nomorBarisBaru, 1, 1, jumlahKolom)
      .setValues([barisOutput]);

    formatBarisAbsensi_(
      sheetAbsensi,
      nomorBarisBaru,
      headerMap
    );

    SpreadsheetApp.flush();

    return {
      success: true,
      code: 'ABSENSI_BERHASIL',
      message:
        'Absensi berhasil dicatat pada pukul ' +
        Utilities.formatDate(sekarang, zonaWaktu, 'HH:mm') +
        '.',
      data: {
        id: idAbsensi,
        tanggal: Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'dd/MM/yyyy'
        ),
        pin: pengguna.pin,
        namaSA: pengguna.namaSA,
        outlet: pengguna.outlet,
        jamMasuk: Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'HH:mm:ss'
        ),
        statusJamMasuk: hasilStatus.statusJamMasuk,
        terlambatMenit: hasilStatus.terlambatMenit,
        statusKehadiran: hasilStatus.statusKehadiran
      }
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      code: 'SYSTEM_ERROR',
      message: 'Terjadi kesalahan sistem: ' + error.message
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


/**
 * Menyimpan absensi pulang pada baris absensi hari ini.
 *
 * Aturan:
 * - SA harus sudah absen masuk;
 * - absen pulang hanya boleh satu kali;
 * - tidak boleh pulang ketika break masih aktif;
 * - JAM PULANG dan STATUS JAM PULANG diperbarui pada baris yang sama.
 *
 * @param {string|number} pinInput PIN SA.
 * @return {Object} Hasil proses absensi pulang.
 */
function simpanAbsensiPulang(pinInput) {
  const SHEET_ABSENSI = 'ABSENSI_HARIAN';
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const hasilLogin = validasiLoginPin(pinInput);
    if (!hasilLogin.success) return hasilLogin;

    const pengguna = absensiNormalisasiPengguna_(
      hasilLogin.data,
      pinInput
    );

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAbsensi = spreadsheet.getSheetByName(SHEET_ABSENSI);

    if (!sheetAbsensi) {
      return {
        success: false,
        code: 'SHEET_ABSENSI_TIDAK_DITEMUKAN',
        message: 'Sheet ABSENSI_HARIAN tidak ditemukan.'
      };
    }

    const zonaWaktu = spreadsheet.getSpreadsheetTimeZone();
    const sekarang = gttSekarang_(pengguna.pin);
    const tanggalHariIni = Utilities.formatDate(
      sekarang,
      zonaWaktu,
      'yyyy-MM-dd'
    );

    const headerMap = ambilHeaderMapAbsensi_(sheetAbsensi);
    const hasilAbsensi = cariAbsensiHariIni_(
      sheetAbsensi,
      headerMap,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    if (!hasilAbsensi.ditemukan) {
      return {
        success: false,
        code: 'BELUM_ABSEN_MASUK',
        message: 'Anda belum melakukan absen masuk hari ini.'
      };
    }

    const nomorBaris = hasilAbsensi.data.nomorBaris;
    const jamPulangTersimpan = sheetAbsensi
      .getRange(nomorBaris, headerMap['JAM PULANG'] + 1)
      .getValue();

    if (jamPulangTersimpan instanceof Date || String(jamPulangTersimpan || '').trim()) {
      const teksJamPulang = jamPulangTersimpan instanceof Date
        ? Utilities.formatDate(jamPulangTersimpan, zonaWaktu, 'HH:mm:ss')
        : String(jamPulangTersimpan).trim();

      return {
        success: false,
        code: 'SUDAH_ABSEN_PULANG',
        message: 'Anda sudah melakukan absen pulang pada pukul ' + teksJamPulang + '.',
        data: hasilAbsensi.data
      };
    }

    const hasilStatusBreak = absensiAmbilStatusBreakAman_(pinInput);
    if (!hasilStatusBreak.success) return hasilStatusBreak;

    const statusOperasional = String(
      hasilStatusBreak.data.statusOperasional || ''
    ).toUpperCase();

    if (statusOperasional === 'BREAK 1' || statusOperasional === 'BREAK 2') {
      return {
        success: false,
        code: 'BREAK_MASIH_AKTIF',
        message: 'Selesaikan break sebelum melakukan absen pulang.'
      };
    }

    const jamMasukValue = sheetAbsensi
      .getRange(nomorBaris, headerMap['JAM MASUK'] + 1)
      .getValue();

    if (!(jamMasukValue instanceof Date)) {
      return {
        success: false,
        code: 'JAM_MASUK_TIDAK_VALID',
        message: 'Data JAM MASUK tidak valid dan tidak dapat dihitung.'
      };
    }

    const pengaturan = absensiAmbilPengaturanPulang_();
    const hasilPulang = hitungStatusPulang_(
      sekarang,
      jamMasukValue,
      hasilAbsensi.data.statusKehadiran,
      pengaturan,
      zonaWaktu
    );
    sheetAbsensi
      .getRange(nomorBaris, headerMap['JAM PULANG'] + 1)
      .setValue(sekarang)
      .setNumberFormat('HH:mm:ss');

    sheetAbsensi
      .getRange(nomorBaris, headerMap['STATUS JAM PULANG'] + 1)
      .setValue(hasilPulang.statusJamPulang);

    sheetAbsensi
      .getRange(nomorBaris, headerMap['STATUS KEHADIRAN'] + 1)
      .setValue(hasilPulang.statusKehadiranAkhir);

    sheetAbsensi
      .getRange(nomorBaris, headerMap['KETERANGAN'] + 1)
      .setValue(hasilPulang.keterangan || '');

    SpreadsheetApp.flush();

    return {
      success: true,
      code: 'ABSENSI_PULANG_BERHASIL',
      message: 'Absen pulang berhasil dicatat pada pukul ' +
        Utilities.formatDate(sekarang, zonaWaktu, 'HH:mm') + '.',
      data: {
        id: hasilAbsensi.data.id,
        tanggal: Utilities.formatDate(sekarang, zonaWaktu, 'dd/MM/yyyy'),
        pin: pengguna.pin,
        namaSA: pengguna.namaSA,
        outlet: pengguna.outlet,
        jamMasuk: hasilAbsensi.data.jamMasuk,
        jamPulang: Utilities.formatDate(sekarang, zonaWaktu, 'HH:mm:ss'),
        statusJamPulang: hasilPulang.statusJamPulang,
        selisihMenit: hasilPulang.selisihMenit,
        statusKehadiran: hasilPulang.statusKehadiranAkhir,
        keterangan: hasilPulang.keterangan || ''
      }
    };
  } catch (error) {
    console.error(error);
    return {
      success: false,
      code: 'SYSTEM_ERROR',
      message: 'Terjadi kesalahan sistem: ' + error.message
    };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}


/**
 * Membaca parameter pulang dari MASTER_SETTING yang sudah LOCK.
 */
function absensiAmbilPengaturanPulang_() {
  return absensiAmbilPengaturan_();
}


/**
 * Menentukan status jam pulang dan status kehadiran akhir.
 */
function hitungStatusPulang_(
  waktuPulang,
  jamMasuk,
  statusKehadiranMasuk,
  pengaturan,
  zonaWaktu
) {
  const menitPulang = absensiUbahWaktuKeMenit_(
    Utilities.formatDate(waktuPulang, zonaWaktu, 'HH:mm')
  );

  const menitBatasSetengahHari =
    absensiUbahWaktuKeMenit_(pengaturan.batasPulangSetengahHari);

  const menitPulangNormal =
    absensiUbahWaktuKeMenit_(pengaturan.jamPulangNormal);

  const statusMasuk = normalisasiHeader_(
    statusKehadiranMasuk
  );

  const masukSudahSetengahHari = [
    '1/2 HARI',
    'HADIR 1/2 HARI'
  ].includes(statusMasuk);

  let statusJamPulang;
  let statusKehadiranAkhir;
  let keterangan = '';

  if (masukSudahSetengahHari) {
    statusKehadiranAkhir = 'HADIR 1/2 HARI';

    if (menitPulang < menitBatasSetengahHari) {
      statusJamPulang = 'PULANG TERLALU AWAL';
      keterangan =
        'STATUS TETAP HADIR 1/2 HARI BERDASARKAN JAM MASUK.';
    } else if (menitPulang < menitPulangNormal) {
      statusJamPulang = '1/2 HARI';
    } else {
      statusJamPulang = 'NORMAL';
    }
  } else if (menitPulang < menitBatasSetengahHari) {
    statusJamPulang = 'ALPA';
    statusKehadiranAkhir = 'ALPA';
    keterangan =
      pengaturan.warningPulangTerlaluAwal ||
      'PULANG SEBELUM BATAS MINIMAL 1/2 HARI.';
  } else if (menitPulang < menitPulangNormal) {
    statusJamPulang = '1/2 HARI';
    statusKehadiranAkhir = 'HADIR 1/2 HARI';
    keterangan =
      'PULANG MULAI PUKUL ' +
      pengaturan.batasPulangSetengahHari +
      ' DAN SEBELUM JAM PULANG NORMAL.';
  } else {
    statusJamPulang = 'NORMAL';
    statusKehadiranAkhir = 'HADIR PENUH';
  }

  return {
    statusJamPulang: statusJamPulang,
    statusKehadiranAkhir: statusKehadiranAkhir,
    selisihMenit:
      menitPulang - menitPulangNormal,
    keterangan: keterangan
  };
}




/**
 * Mengambil dan memvalidasi header ABSENSI_HARIAN.
 */
function ambilHeaderMapAbsensi_(sheet) {
  const jumlahKolom = sheet.getLastColumn();

  if (jumlahKolom === 0) {
    throw new Error('Header ABSENSI_HARIAN belum tersedia.');
  }

  const daftarHeader = sheet
    .getRange(1, 1, 1, jumlahKolom)
    .getDisplayValues()[0]
    .map(normalisasiHeader_);

  const kolomWajib = [
    'ID',
    'TANGGAL',
    'NAMA SA',
    'PIN',
    'OUTLET',
    'JAM MASUK',
    'STATUS JAM MASUK',
    'TERLAMBAT MENIT',
    'JAM PULANG',
    'STATUS JAM PULANG',
    'STATUS KEHADIRAN',
    'KETERANGAN'
  ];

  const headerMap = {};

  daftarHeader.forEach((namaHeader, index) => {
    if (namaHeader) {
      headerMap[namaHeader] = index;
    }
  });

  const kolomTidakAda = kolomWajib.filter(
    namaKolom => headerMap[namaKolom] === undefined
  );

  if (kolomTidakAda.length > 0) {
    throw new Error(
      'Kolom berikut tidak ditemukan di ABSENSI_HARIAN: ' +
      kolomTidakAda.join(', ')
    );
  }

  return headerMap;
}


/**
 * Mencari absensi terbaru berdasarkan PIN dan tanggal.
 */
function cariAbsensiHariIni_(
  sheet,
  headerMap,
  pin,
  tanggalHariIni,
  zonaWaktu
) {
  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    return {
      ditemukan: false
    };
  }

  const jumlahBaris = barisTerakhir - 1;
  const jumlahKolom = sheet.getLastColumn();

  const seluruhData = sheet
    .getRange(2, 1, jumlahBaris, jumlahKolom)
    .getValues();

  for (let index = seluruhData.length - 1; index >= 0; index--) {
    const baris = seluruhData[index];

    const pinData = String(
      baris[headerMap['PIN']] || ''
    ).trim();

    const tanggalData = normalisasiTanggalAbsensi_(
      baris[headerMap['TANGGAL']],
      zonaWaktu
    );

    if (
      pinData === String(pin).trim() &&
      tanggalData === tanggalHariIni
    ) {
      const jamMasukValue =
        baris[headerMap['JAM MASUK']];

      const jamMasuk = jamMasukValue instanceof Date
        ? Utilities.formatDate(
            jamMasukValue,
            zonaWaktu,
            'HH:mm:ss'
          )
        : String(jamMasukValue || '-');

      return {
        ditemukan: true,
        jamMasuk: jamMasuk,
        data: {
          nomorBaris: index + 2,
          id: baris[headerMap['ID']],
          tanggal: tanggalData,
          pin: pinData,
          namaSA: baris[headerMap['NAMA SA']],
          outlet: baris[headerMap['OUTLET']],
          jamMasuk: jamMasuk,
          statusJamMasuk:
            baris[headerMap['STATUS JAM MASUK']],
          terlambatMenit:
            baris[headerMap['TERLAMBAT MENIT']],
          statusKehadiran:
            baris[headerMap['STATUS KEHADIRAN']]
        }
      };
    }
  }

  return {
    ditemukan: false
  };
}


/**
 * Menormalisasi struktur pengguna dari hasil validasi login.
 * Mendukung perbedaan nama field antarversi PinSA.gs.
 */
function absensiNormalisasiPengguna_(dataLogin, pinInput) {
  const sumberAwal = dataLogin || {};
  const sumber = sumberAwal.pengguna || sumberAwal;

  const pin = String(
    sumber.pin !== undefined && sumber.pin !== null
      ? sumber.pin
      : (
          sumber.PIN !== undefined && sumber.PIN !== null
            ? sumber.PIN
            : pinInput
        )
  ).trim();

  const namaSA = String(
    sumber.namaSA !== undefined && sumber.namaSA !== null
      ? sumber.namaSA
      : (
          sumber.nama !== undefined && sumber.nama !== null
            ? sumber.nama
            : (
                sumber['NAMA SA'] !== undefined &&
                sumber['NAMA SA'] !== null
                  ? sumber['NAMA SA']
                  : ''
              )
        )
  ).trim();

  const outlet = String(
    sumber.outlet !== undefined && sumber.outlet !== null
      ? sumber.outlet
      : (
          sumber.OUTLET !== undefined && sumber.OUTLET !== null
            ? sumber.OUTLET
            : ''
        )
  ).trim();

  const status = String(
    sumber.status !== undefined && sumber.status !== null
      ? sumber.status
      : (
          sumber.STATUS !== undefined && sumber.STATUS !== null
            ? sumber.STATUS
            : ''
        )
  ).trim();

  return {
    pin: pin,
    namaSA: namaSA,
    outlet: outlet,
    status: status
  };
}


/**
 * Memastikan data wajib pengguna tersedia sebelum disimpan.
 */
function absensiValidasiDataPengguna_(pengguna) {
  const fieldKosong = [];

  if (!pengguna.pin) {
    fieldKosong.push('PIN');
  }

  if (!pengguna.namaSA) {
    fieldKosong.push('NAMA SA');
  }

  if (!pengguna.outlet) {
    fieldKosong.push('OUTLET');
  }

  if (fieldKosong.length > 0) {
    throw new Error(
      'Data pengguna tidak lengkap: ' +
      fieldKosong.join(', ') +
      '. Periksa hasil validasi login PinSA.gs.'
    );
  }
}


/**
 * Mengambil parameter absensi dari MASTER_SETTING.
 */
function absensiAmbilPengaturan_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('MASTER_SETTING');

  if (!sheet) {
    throw new Error('Sheet MASTER_SETTING tidak ditemukan.');
  }

  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    throw new Error('MASTER_SETTING masih kosong.');
  }

  const dataSetting = sheet
    .getRange(2, 1, barisTerakhir - 1, Math.min(3, sheet.getLastColumn()))
    .getDisplayValues();

  const settingMap = {};
  const warningMap = {};

  dataSetting.forEach(function (baris) {
    const parameter = normalisasiHeader_(baris[0]);
    if (!parameter) return;

    settingMap[parameter] = String(baris[1] || '').trim();
    warningMap[parameter] = String(baris[2] || '').trim();
  });

  const namaParameter = {
    jamMasukNormal: 'ABSENSI - JAM MASUK NORMAL',
    toleransiTerlambat: 'ABSENSI - TOLERANSI TERLAMBAT (MENIT)',
    batasAkhirTerlambat: 'ABSENSI - BATAS AKHIR TERLAMBAT',
    jamMasukSetengahHari: 'ABSENSI - JAM MASUK 1/2 HARI',
    batasPulangSetengahHari:
      'ABSENSI - BATAS PULANG MINIMAL 1/2 HARI',
    jamPulangNormal: 'ABSENSI - JAM PULANG NORMAL'
  };

  Object.keys(namaParameter).forEach(function (kunci) {
    const parameter = namaParameter[kunci];
    if (!settingMap[parameter]) {
      throw new Error(
        'Parameter "' + parameter + '" tidak ditemukan.'
      );
    }
  });

  const toleransiTerlambat = Number(
    settingMap[namaParameter.toleransiTerlambat]
  );

  if (
    !Number.isFinite(toleransiTerlambat) ||
    toleransiTerlambat < 0
  ) {
    throw new Error(
      'Nilai "' +
      namaParameter.toleransiTerlambat +
      '" harus berupa angka tidak negatif.'
    );
  }

  const pengaturan = {
    jamMasukNormal:
      settingMap[namaParameter.jamMasukNormal],
    toleransiTerlambat: toleransiTerlambat,
    batasAkhirTerlambat:
      settingMap[namaParameter.batasAkhirTerlambat],
    jamMasukSetengahHari:
      settingMap[namaParameter.jamMasukSetengahHari],
    batasPulangSetengahHari:
      settingMap[namaParameter.batasPulangSetengahHari],
    jamPulangNormal:
      settingMap[namaParameter.jamPulangNormal],
    warningPulangTerlaluAwal:
      warningMap[namaParameter.batasPulangSetengahHari] || ''
  };

  [
    pengaturan.jamMasukNormal,
    pengaturan.batasAkhirTerlambat,
    pengaturan.jamMasukSetengahHari,
    pengaturan.batasPulangSetengahHari,
    pengaturan.jamPulangNormal
  ].forEach(absensiUbahWaktuKeMenit_);

  const batasNormalDariToleransi =
    absensiUbahWaktuKeMenit_(pengaturan.jamMasukNormal) +
    pengaturan.toleransiTerlambat;

  const batasAkhirTerlambatMenit =
    absensiUbahWaktuKeMenit_(pengaturan.batasAkhirTerlambat);

  const jamMasukSetengahHariMenit =
    absensiUbahWaktuKeMenit_(pengaturan.jamMasukSetengahHari);

  if (batasAkhirTerlambatMenit < batasNormalDariToleransi) {
    throw new Error(
      'ABSENSI - BATAS AKHIR TERLAMBAT tidak boleh lebih awal ' +
      'dari batas masuk normal.'
    );
  }

  if (jamMasukSetengahHariMenit <= batasAkhirTerlambatMenit) {
    throw new Error(
      'ABSENSI - JAM MASUK 1/2 HARI harus setelah ' +
      'ABSENSI - BATAS AKHIR TERLAMBAT.'
    );
  }

  return pengaturan;
}


/**
 * Menentukan status kehadiran dan status jam masuk.
 */
function hitungStatusAbsensi_(
  waktuMasuk,
  pengaturan,
  zonaWaktu
) {
  const totalMenitMasuk = absensiUbahWaktuKeMenit_(
    Utilities.formatDate(waktuMasuk, zonaWaktu, 'HH:mm')
  );

  const menitMasukNormal =
    absensiUbahWaktuKeMenit_(pengaturan.jamMasukNormal);

  const batasMasukNormal =
    menitMasukNormal + pengaturan.toleransiTerlambat;

  const batasAkhirTerlambat =
    absensiUbahWaktuKeMenit_(pengaturan.batasAkhirTerlambat);

  const menitMasukSetengahHari =
    absensiUbahWaktuKeMenit_(pengaturan.jamMasukSetengahHari);

  if (totalMenitMasuk >= menitMasukSetengahHari) {
    return {
      success: true,
      statusKehadiran: 'HADIR 1/2 HARI',
      statusJamMasuk: '1/2 HARI',
      terlambatMenit: 0,
      keterangan:
        'ABSEN MASUK MULAI PUKUL ' +
        pengaturan.jamMasukSetengahHari +
        ' DICATAT HADIR 1/2 HARI.'
    };
  }

  if (totalMenitMasuk <= batasMasukNormal) {
    return {
      success: true,
      statusKehadiran: 'HADIR PENUH',
      statusJamMasuk: 'NORMAL',
      terlambatMenit: 0,
      keterangan: ''
    };
  }

  if (totalMenitMasuk <= batasAkhirTerlambat) {
    return {
      success: true,
      statusKehadiran: 'HADIR PENUH',
      statusJamMasuk: 'TERLAMBAT',
      terlambatMenit:
        totalMenitMasuk - batasMasukNormal,
      keterangan: ''
    };
  }

  // Pengaman apabila terdapat celah konfigurasi waktu.
  return {
    success: true,
    statusKehadiran: 'HADIR 1/2 HARI',
    statusJamMasuk: '1/2 HARI',
    terlambatMenit: 0,
    keterangan:
      'ABSEN MASUK MELEWATI BATAS AKHIR TERLAMBAT.'
  };
}


/**
 * Mengubah teks waktu HH:mm menjadi total menit.
 */
function absensiUbahWaktuKeMenit_(nilaiWaktu) {
  const teks = String(nilaiWaktu || '').trim();
  const hasil = teks.match(/^(\d{1,2}):(\d{2})$/);

  if (!hasil) {
    throw new Error(
      'Format waktu tidak valid: "' +
      teks +
      '". Gunakan format HH:mm.'
    );
  }

  const jam = Number(hasil[1]);
  const menit = Number(hasil[2]);

  if (
    jam < 0 ||
    jam > 23 ||
    menit < 0 ||
    menit > 59
  ) {
    throw new Error(
      'Nilai waktu tidak valid: "' + teks + '".'
    );
  }

  return jam * 60 + menit;
}


/**
 * Menormalisasi nilai tanggal menjadi yyyy-MM-dd.
 */
function normalisasiTanggalAbsensi_(
  nilaiTanggal,
  zonaWaktu
) {
  if (nilaiTanggal instanceof Date) {
    return Utilities.formatDate(
      nilaiTanggal,
      zonaWaktu,
      'yyyy-MM-dd'
    );
  }

  const teks = String(nilaiTanggal || '').trim();

  if (!teks) {
    return '';
  }

  const tanggalPercobaan = new Date(teks);

  if (!isNaN(tanggalPercobaan.getTime())) {
    return Utilities.formatDate(
      tanggalPercobaan,
      zonaWaktu,
      'yyyy-MM-dd'
    );
  }

  return teks;
}


/**
 * Membuat tanggal tanpa komponen jam.
 */
function buatTanggalTanpaJam_(tanggal) {
  return new Date(
    tanggal.getFullYear(),
    tanggal.getMonth(),
    tanggal.getDate()
  );
}


/**
 * Membuat ID absensi unik.
 */
function buatIdAbsensi_(pin, waktu, zonaWaktu) {
  const pinTeks = String(
    pin !== undefined && pin !== null
      ? pin
      : ''
  ).trim();

  if (!pinTeks) {
    throw new Error(
      'PIN tidak tersedia sehingga ID absensi tidak dapat dibuat.'
    );
  }

  const timestamp = Utilities.formatDate(
    waktu,
    zonaWaktu,
    'yyyyMMddHHmmss'
  );

  const kodeAcak = Utilities
    .getUuid()
    .replace(/-/g, '')
    .substring(0, 6)
    .toUpperCase();

  return [
    'ABS',
    timestamp,
    pinTeks,
    kodeAcak
  ].join('-');
}


/**
 * Mengatur format baris absensi yang baru dibuat.
 */
function formatBarisAbsensi_(
  sheet,
  nomorBaris,
  headerMap
) {
  sheet
    .getRange(
      nomorBaris,
      headerMap['PIN'] + 1
    )
    .setNumberFormat('@');

  sheet
    .getRange(
      nomorBaris,
      headerMap['TANGGAL'] + 1
    )
    .setNumberFormat('dd/MM/yyyy');

  sheet
    .getRange(
      nomorBaris,
      headerMap['JAM MASUK'] + 1
    )
    .setNumberFormat('HH:mm:ss');

  sheet
    .getRange(
      nomorBaris,
      headerMap['TERLAMBAT MENIT'] + 1
    )
    .setNumberFormat('0');
}


/**
 * Menguji normalisasi data pengguna tanpa menyimpan ke sheet.
 */
function ujiNormalisasiPenggunaAbsensi() {
  const PIN_UJI = '5184';
  const hasilLogin = validasiLoginPin(PIN_UJI);

  if (!hasilLogin.success) {
    console.log(JSON.stringify(hasilLogin, null, 2));
    return;
  }

  const pengguna = absensiNormalisasiPengguna_(
    hasilLogin.data,
    PIN_UJI
  );

  absensiValidasiDataPengguna_(pengguna);

  console.log(
    JSON.stringify(
      {
        success: true,
        code: 'DATA_PENGGUNA_ABSENSI_VALID',
        data: pengguna
      },
      null,
      2
    )
  );
}


/**
 * Menguji pembacaan MASTER_SETTING tanpa menyimpan data.
 */
function ujiMasterSettingAbsensi() {
  const hasil = absensiAmbilPengaturan_();

  console.log(
    JSON.stringify(
      {
        success: true,
        code: 'MASTER_SETTING_ABSENSI_VALID',
        data: hasil
      },
      null,
      2
    )
  );
}


/**
 * Pengujian logika status tanpa menyimpan ke sheet.
 */
function ujiLogikaStatusAbsensi() {
  const zonaWaktu =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSpreadsheetTimeZone();

  const pengaturan = {
    jamMasukNormal: '08:15',
    toleransiTerlambat: 5,
    batasAkhirTerlambat: '09:00',
    jamMasukSetengahHari: '09:01',
    batasPulangSetengahHari: '14:01',
    jamPulangNormal: '20:30',
    warningPulangTerlaluAwal: ''
  };

  const tanggalDasar = new Date(2026, 6, 29);

  const daftarUji = [
    '08:15',
    '08:20',
    '08:21',
    '08:35',
    '09:00',
    '09:01',
    '12:00',
    '14:00'
  ];

  const hasilUji = daftarUji.map(jamUji => {
    const bagian = jamUji.split(':');

    const waktuUji = new Date(tanggalDasar);
    waktuUji.setHours(
      Number(bagian[0]),
      Number(bagian[1]),
      0,
      0
    );

    return {
      jam: jamUji,
      hasil: hitungStatusAbsensi_(
        waktuUji,
        pengaturan,
        zonaWaktu
      )
    };
  });

  console.log(
    JSON.stringify(hasilUji, null, 2)
  );
}


/**
 * Pengujian simpan absensi.
 * PERHATIAN: fungsi ini benar-benar menambah data.
 */
function ujiSimpanAbsensiMasuk() {
  const PIN_UJI = '9616';

  const hasil = simpanAbsensiMasuk(PIN_UJI);

  console.log(
    JSON.stringify(hasil, null, 2)
  );
}

/**
 * Pengujian simpan absen pulang.
 * PERHATIAN: fungsi ini benar-benar mengubah data.
 */
function ujiSimpanAbsensiPulang() {
  const PIN_UJI = '9616';
  const hasil = simpanAbsensiPulang(PIN_UJI);
  console.log(JSON.stringify(hasil, null, 2));
}


/**
 * Membaca status Break dengan pemeriksaan dependency yang jelas.
 * GTT-08 (pulang saat Break aktif) tetap menjadi backlog integrasi;
 * versi ini masih mengunci tombol Pulang sampai Break selesai.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function absensiAmbilStatusBreakAman_(pinInput) {
  if (typeof ambilStatusBreakHariIni !== 'function') {
    return {
      success: false,
      code: 'DEPENDENCY_BREAK_BELUM_SIAP',
      message:
        'Fungsi ambilStatusBreakHariIni belum tersedia di BreakTime.gs.'
    };
  }

  return ambilStatusBreakHariIni(pinInput);
}


/**
 * Self-test logika inti Absensi.gs tanpa menulis ke sheet.
 *
 * @return {Object}
 */
function ujiAbsensiGs() {
  const zonaWaktu =
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  const pengaturan = {
    jamMasukNormal: '08:15',
    toleransiTerlambat: 5,
    batasAkhirTerlambat: '09:00',
    jamMasukSetengahHari: '09:01',
    batasPulangSetengahHari: '14:01',
    jamPulangNormal: '20:30',
    warningPulangTerlaluAwal: ''
  };

  function waktu_(jam) {
    const bagian = jam.split(':');
    const nilai = new Date(2026, 7, 2, 0, 0, 0, 0);
    nilai.setHours(Number(bagian[0]), Number(bagian[1]), 0, 0);
    return nilai;
  }

  const hasil0815 = hitungStatusAbsensi_(
    waktu_('08:15'), pengaturan, zonaWaktu
  );
  const hasil0820 = hitungStatusAbsensi_(
    waktu_('08:20'), pengaturan, zonaWaktu
  );
  const hasil0821 = hitungStatusAbsensi_(
    waktu_('08:21'), pengaturan, zonaWaktu
  );
  const hasil0901 = hitungStatusAbsensi_(
    waktu_('09:01'), pengaturan, zonaWaktu
  );

  const pulang1400 = hitungStatusPulang_(
    waktu_('14:00'), waktu_('08:15'), 'HADIR PENUH',
    pengaturan, zonaWaktu
  );
  const pulang1401 = hitungStatusPulang_(
    waktu_('14:01'), waktu_('08:15'), 'HADIR PENUH',
    pengaturan, zonaWaktu
  );
  const pulang2030 = hitungStatusPulang_(
    waktu_('20:30'), waktu_('08:15'), 'HADIR PENUH',
    pengaturan, zonaWaktu
  );

  const pemeriksaan = {
    masuk0815Normal:
      hasil0815.statusJamMasuk === 'NORMAL' &&
      hasil0815.statusKehadiran === 'HADIR PENUH',
    masuk0820MasihNormal:
      hasil0820.statusJamMasuk === 'NORMAL',
    masuk0821TerlambatSatuMenit:
      hasil0821.statusJamMasuk === 'TERLAMBAT' &&
      hasil0821.terlambatMenit === 1,
    masuk0901SetengahHari:
      hasil0901.statusKehadiran === 'HADIR 1/2 HARI',
    pulang1400Alpa:
      pulang1400.statusKehadiranAkhir === 'ALPA',
    pulang1401SetengahHari:
      pulang1401.statusKehadiranAkhir === 'HADIR 1/2 HARI',
    pulang2030HadirPenuh:
      pulang2030.statusKehadiranAkhir === 'HADIR PENUH',
    masterSettingTerbaca: false,
    dependencyBreakTersedia:
      typeof ambilStatusBreakHariIni === 'function'
  };

  try {
    absensiAmbilPengaturan_();
    pemeriksaan.masterSettingTerbaca = true;
  } catch (errorSetting) {
    pemeriksaan.masterSettingTerbaca = false;
  }

  const logikaLulus = Object.keys(pemeriksaan)
    .filter(function (kunci) {
      return kunci !== 'dependencyBreakTersedia';
    })
    .every(function (kunci) {
      return pemeriksaan[kunci] === true;
    });

  const hasil = {
    success: logikaLulus,
    code: logikaLulus
      ? 'UJI_ABSENSI_PASS'
      : 'UJI_ABSENSI_FAIL',
    hasil: pemeriksaan,
    catatan:
      pemeriksaan.dependencyBreakTersedia
        ? 'Dependency Break tersedia.'
        : 'Dependency Break belum tersedia; akan diselesaikan pada review BreakTime.gs.'
  };

  console.log(JSON.stringify(hasil, null, 2));
  return hasil;
}
