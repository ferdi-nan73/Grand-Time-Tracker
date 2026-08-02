// ===== API.GS - PART 1 START =====

/**
 * =====================================================
 * GRAND TIME TRACKER — GTT
 * Module : API Gateway
 * Version: 2.2.0
 * Sprint : GTT-002
 * =====================================================
 */


/**
 * Mengecek kesiapan backend dan sheet wajib GTT.
 *
 * @return {Object}
 */
function apiHealthCheck() {
  return apiJalankan_('HEALTH_CHECK', function () {
    var spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();

    var zonaWaktu =
      spreadsheet.getSpreadsheetTimeZone();

    var sekarang = new Date();

    var sheetWajib = [
      'MASTER_SA',
      'MASTER_OUTLET',
      'MASTER_SETTING',
      'ABSENSI_HARIAN',
      'LOG_ISTIRAHAT'
    ];

    var statusSheet = {};

    sheetWajib.forEach(function (namaSheet) {
      statusSheet[namaSheet] = Boolean(
        spreadsheet.getSheetByName(namaSheet)
      );
    });

    var seluruhSheetTersedia =
      Object.keys(statusSheet).every(
        function (namaSheet) {
          return statusSheet[namaSheet];
        }
      );

    return {
      success: seluruhSheetTersedia,
      code: seluruhSheetTersedia
        ? 'API_HEALTHY'
        : 'API_DEPENDENCY_ERROR',
      message: seluruhSheetTersedia
        ? 'Backend GTT berjalan normal.'
        : 'Ada sheet wajib yang tidak ditemukan.',
      data: {
        aplikasi: 'GTT',
        versiApi: '2.2.0',
        waktuServer: {
          tanggal: Utilities.formatDate(
            sekarang,
            zonaWaktu,
            'dd/MM/yyyy'
          ),
          jam: Utilities.formatDate(
            sekarang,
            zonaWaktu,
            'HH:mm:ss'
          ),
          timestamp: sekarang.getTime(),
          zonaWaktu: zonaWaktu
        },
        sheet: statusSheet
      }
    };
  });
}


/**
 * Login SA menggunakan PIN.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiLogin(pinInput) {
  return apiJalankan_('LOGIN', function () {
    return validasiLoginPin(pinInput);
  });
}


/**
 * Menyimpan absen masuk.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiAbsenMasuk(pinInput) {
  return apiJalankan_(
    'ABSEN_MASUK',
    function () {
      return simpanAbsensiMasuk(pinInput);
    }
  );
}


/**
 * Menyimpan absen pulang.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiAbsenPulang(pinInput) {
  return apiJalankan_(
    'ABSEN_PULANG',
    function () {
      return simpanAbsensiPulang(pinInput);
    }
  );
}


/**
 * Memulai Break 1.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiMulaiBreak1(pinInput) {
  return apiJalankan_(
    'MULAI_BREAK_1',
    function () {
      return mulaiBreak1(pinInput);
    }
  );
}


/**
 * Menyelesaikan Break 1.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiSelesaiBreak1(pinInput) {
  return apiJalankan_(
    'SELESAI_BREAK_1',
    function () {
      return selesaiBreak1(pinInput);
    }
  );
}


/**
 * Memulai Break 2.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiMulaiBreak2(pinInput) {
  return apiJalankan_(
    'MULAI_BREAK_2',
    function () {
      return mulaiBreak2(pinInput);
    }
  );
}


/**
 * Menyelesaikan Break 2.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiSelesaiBreak2(pinInput) {
  return apiJalankan_(
    'SELESAI_BREAK_2',
    function () {
      return selesaiBreak2(pinInput);
    }
  );
}


/**
 * Mengambil status Break hari ini.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiStatusBreakHariIni(pinInput) {
  return apiJalankan_(
    'STATUS_BREAK_HARI_INI',
    function () {
      return ambilStatusBreakHariIni(
        pinInput
      );
    }
  );
}


/**
 * Mengambil ringkasan dashboard outlet.
 *
 * @param {string=} outletInput
 * @return {Object}
 */
function apiRingkasanDashboard(outletInput) {
  return apiJalankan_(
    'RINGKASAN_DASHBOARD',
    function () {
      return ambilRingkasanDashboard(
        outletInput || ''
      );
    }
  );
}


/**
 * Bootstrap data awal SA setelah login.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiBootstrapSA(pinInput) {
  return apiJalankan_(
    'BOOTSTRAP_SA',
    function () {
      var hasil =
        apiBangunDataSA_(pinInput);

      if (!hasil.success) {
        return hasil;
      }

      return {
        success: true,
        code: 'BOOTSTRAP_SA_BERHASIL',
        message:
          'Data awal aplikasi berhasil dimuat.',
        data: hasil.data
      };
    }
  );
}


/**
 * Memperbarui status SA.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiRefreshStatusSA(pinInput) {
  return apiJalankan_(
    'REFRESH_STATUS_SA',
    function () {
      var hasil =
        apiBangunDataSA_(pinInput);

      if (!hasil.success) {
        return hasil;
      }

      return {
        success: true,
        code: 'REFRESH_STATUS_SA_BERHASIL',
        message:
          'Status SA berhasil diperbarui.',
        data: hasil.data
      };
    }
  );
}

// ===== API.GS - PART 1 END =====
// ===== API.GS - PART 2 START =====

/**
 * Builder utama data SA untuk bootstrap dan refresh.
 *
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiBangunDataSA_(pinInput) {
  var hasilLogin =
    validasiLoginPin(pinInput);

  if (!hasilLogin.success) {
    return hasilLogin;
  }

  var pengguna =
    apiNormalisasiPengguna_(
      hasilLogin.data,
      pinInput
    );

  var hasilDashboard =
    ambilRingkasanDashboard(
      pengguna.outlet,
      pengguna.pin
    );

  if (!hasilDashboard.success) {
    return hasilDashboard;
  }

  var daftarStatus =
    hasilDashboard.data &&
    hasilDashboard.data.daftarStatus
      ? hasilDashboard.data.daftarStatus
      : [];

  var statusSA =
    apiCariStatusSaDashboard_(
      daftarStatus,
      pengguna.pin
    );

  if (!statusSA) {
    return {
      success: false,
      code: 'STATUS_SA_TIDAK_DITEMUKAN',
      message:
        'Status SA tidak ditemukan pada Dashboard Outlet ' +
        pengguna.outlet +
        '.'
    };
  }

  var hasilStatusBreak =
    ambilStatusBreakHariIni(
      pengguna.pin
    );

  if (!hasilStatusBreak.success) {
    return hasilStatusBreak;
  }

  var spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  var zonaWaktu =
    spreadsheet.getSpreadsheetTimeZone();

  var infoWaktu =
    gttInfoWaktu_(pengguna.pin);

  var sekarang =
    infoWaktu.sekarang;

  var breakInfo =
    apiAmbilBreakInfoPribadi_(
      spreadsheet,
      pengguna.pin,
      sekarang,
      zonaWaktu
    );

  var kapasitasBreak =
    apiAmbilKapasitasBreakSetting_(
      spreadsheet
    );

  var ringkasanDashboard =
    hasilDashboard.data &&
    hasilDashboard.data.ringkasan
      ? hasilDashboard.data.ringkasan
      : {};

  var jumlahSedangBreak =
    Number(
      ringkasanDashboard.sedangBreak || 0
    );

  var kapasitasPenuh =
    jumlahSedangBreak >=
    kapasitasBreak.maksimalBreak;

  var kehadiran =
    apiBangunKehadiran_(
      statusSA
    );

  var statusBreakData =
    hasilStatusBreak.data || {};

  var tombolUtama =
    apiTentukanTombolUtama_(
      kehadiran,
      statusSA,
      statusBreakData,
      breakInfo,
      kapasitasPenuh,
      kapasitasBreak
    );

  var statusOperasional =
    String(
      statusSA.statusOperasional ||
      statusBreakData.statusOperasional ||
      'BELUM ABSEN'
    ).trim();

  var breakData = {
    statusOperasional:
      statusOperasional,

    break1:
      breakInfo.break1,

    break2:
      breakInfo.break2,

    overbreakMenit: {
      break1:
        Number(
          breakInfo.break1.overtimeMenit
        ) || 0,

      break2:
        Number(
          breakInfo.break2.overtimeMenit
        ) || 0
    },

    seluruhBreakSelesai:
      Boolean(
        statusBreakData.seluruhBreakSelesai
      ) ||
      (
        breakInfo.break1.status ===
          'SELESAI' &&
        breakInfo.break2.status ===
          'SELESAI'
      ),

    break1Dilewati:
      Boolean(
        statusBreakData.break1Dilewati
      ),

    aksiDiizinkan:
      statusBreakData.aksiDiizinkan !==
      false,

    alasanTerkunci:
      String(
        statusBreakData.alasanTerkunci ||
        ''
      ).trim(),

    kodeTerkunci:
      String(
        statusBreakData.kodeTerkunci ||
        ''
      ).trim(),

    tersediaPada:
      String(
        statusBreakData.tersediaPada ||
        ''
      ).trim(),

    sisaTungguMenit:
      Number(
        statusBreakData.sisaTungguMenit ||
        0
      ),

    jumlahSedangBreak:
      jumlahSedangBreak,

    maksimalBreak:
      kapasitasBreak.maksimalBreak,

    kapasitasPenuh:
      kapasitasPenuh,

    warningKapasitas:
      kapasitasPenuh
        ? kapasitasBreak.warning
        : ''
  };

  var ringkasan = {
    breakOutlet: {
      jumlahSedangBreak:
        jumlahSedangBreak,

      maksimalBreak:
        kapasitasBreak.maksimalBreak,

      kapasitasPenuh:
        kapasitasPenuh,

      teks:
        jumlahSedangBreak +
        ' dari ' +
        kapasitasBreak.maksimalBreak +
        ' orang sedang break'
    },

    breakInfo: {
      break1:
        breakInfo.break1.status,

      break2:
        breakInfo.break2.status,

      teks:
        'B1 ' +
        apiLabelStatusBreak_(
          breakInfo.break1.status
        ) +
        ' · B2 ' +
        apiLabelStatusBreak_(
          breakInfo.break2.status
        )
    },

    tugasLuar: {
      aktif: 0,
      pulang: 0,
      tersedia: false
    },

    pelanggaran: {
      jumlah:
        Number(
          breakInfo.pelanggaran &&
          breakInfo.pelanggaran.jumlah
        ) || 0,

      tersedia: true
    },

    overbreak: {
      break1Menit:
        Number(
          breakInfo.break1.overtimeMenit
        ) || 0,

      break2Menit:
        Number(
          breakInfo.break2.overtimeMenit
        ) || 0,

      tersedia: true
    }
  };

  var dataUtama = {
    pengguna:
      pengguna,

    waktuServer: {
      tanggal:
        Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'dd/MM/yyyy'
        ),

      jam:
        Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'HH:mm:ss'
        ),

      timestamp:
        sekarang.getTime(),

      zonaWaktu:
        zonaWaktu,

      modeUji:
        Boolean(
          infoWaktu.modeUjiAktif
        ),

      tambahMenitUji:
        Number(
          infoWaktu
            .tambahMenitDiterapkan || 0
        )
    },

    kehadiran:
      kehadiran,

    break:
      breakData,

    operasional: {
      statusOperasional:
        statusOperasional,

      tombolUtama:
        tombolUtama
    },

    ringkasan:
      ringkasan,

    dashboard: {
      statusOutlet:
        hasilDashboard.data
          .statusOutlet || {},

      ringkasanOutlet:
        ringkasanDashboard,

      statusSA:
        statusSA
    }
  };

  return {
    success: true,
    code: 'DATA_SA_BERHASIL_DIBANGUN',
    message:
      'Data SA berhasil dibangun.',
    data: dataUtama
  };
}


/**
 * Menormalisasi data pengguna dari hasil login.
 *
 * @param {Object} dataLogin
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiNormalisasiPengguna_(
  dataLogin,
  pinInput
) {
  var sumber =
    dataLogin || {};

  var penggunaSumber =
    sumber.pengguna || sumber;

  return {
    pin:
      String(
        penggunaSumber.pin ||
        penggunaSumber.PIN ||
        pinInput ||
        ''
      ).trim(),

    namaSA:
      String(
        penggunaSumber.namaSA ||
        penggunaSumber.nama ||
        penggunaSumber.NAMA_SA ||
        penggunaSumber['NAMA SA'] ||
        ''
      ).trim(),

    outlet:
      String(
        penggunaSumber.outlet ||
        penggunaSumber.OUTLET ||
        ''
      ).trim(),

    jabatan:
      String(
        penggunaSumber.jabatan ||
        penggunaSumber.JABATAN ||
        penggunaSumber.role ||
        ''
      ).trim(),

    statusAktif:
      penggunaSumber.statusAktif !== false
  };
}


/**
 * Menyusun data status kehadiran.
 *
 * @param {Object} statusSA
 * @return {Object}
 */
function apiBangunKehadiran_(
  statusSA
) {
  var sumber =
    statusSA || {};

  return {
    sudahAbsenMasuk:
      Boolean(
        sumber.sudahAbsenMasuk ||
        sumber.jamMasuk
      ),

    sudahPulang:
      Boolean(
        sumber.sudahPulang ||
        sumber.jamPulang
      ),

    statusKehadiran:
      String(
        sumber.statusKehadiran ||
        sumber.status ||
        'BELUM ABSEN'
      ).trim(),

    statusJamMasuk:
      String(
        sumber.statusJamMasuk ||
        ''
      ).trim(),

    jamMasuk:
      apiFormatJam_(
        sumber.jamMasuk
      ),

    jamPulang:
      apiFormatJam_(
        sumber.jamPulang
      ),

    terlambatMenit:
      Number(
        sumber.terlambatMenit || 0
      ),

    keterangan:
      String(
        sumber.keterangan || ''
      ).trim()
  };
}

// ===== API.GS - PART 2 END =====
// ===== API.GS - PART 3 START =====

/**
 * Menentukan tombol utama berdasarkan status SA.
 *
 * @param {Object} kehadiran
 * @param {Object} statusSA
 * @param {Object} statusBreakEngine
 * @param {Object} breakInfo
 * @param {boolean} kapasitasPenuh
 * @param {Object} kapasitasBreak
 * @return {Object}
 */
function apiTentukanTombolUtama_(
  kehadiran,
  statusSA,
  statusBreakEngine,
  breakInfo,
  kapasitasPenuh,
  kapasitasBreak
) {
  var statusKehadiran =
    apiNormalisasiTeks_(
      kehadiran.statusKehadiran || ''
    );

  var jamPulang =
    String(
      kehadiran.jamPulang || ''
    ).trim();

  var statusOperasionalSA =
    apiNormalisasiTeks_(
      statusSA.statusOperasional || ''
    );

  var tombolEngine =
    apiNormalisasiTeks_(
      statusBreakEngine.tombolBerikutnya ||
      statusBreakEngine.tombol ||
      ''
    );

  var aksiEngineDiizinkan =
    statusBreakEngine.aksiDiizinkan !==
    false;

  var alasanEngine =
    String(
      statusBreakEngine.alasanTerkunci ||
      ''
    ).trim();

  if (
    statusKehadiran === 'BELUM ABSEN' ||
    !kehadiran.sudahAbsenMasuk ||
    tombolEngine === 'ABSEN MASUK'
  ) {
    return apiBuatTombol_(
      'ABSEN_MASUK',
      'ABSEN MASUK',
      true,
      ''
    );
  }

  if (
    kehadiran.sudahPulang ||
    jamPulang ||
    statusOperasionalSA ===
      'SUDAH PULANG'
  ) {
    return apiBuatTombol_(
      'SELESAI',
      'SELESAI',
      false,
      'Absensi hari ini sudah selesai.'
    );
  }

  var tombolDariEngine = {
    'MULAI BREAK 1': {
      aksi: 'MULAI_BREAK_1',
      label: 'MULAI BREAK 1'
    },

    'SELESAI BREAK 1': {
      aksi: 'SELESAI_BREAK_1',
      label: 'SELESAI BREAK 1'
    },

    'MULAI BREAK 2': {
      aksi: 'MULAI_BREAK_2',
      label: 'MULAI BREAK 2'
    },

    'SELESAI BREAK 2': {
      aksi: 'SELESAI_BREAK_2',
      label: 'SELESAI BREAK 2'
    },

    'ABSEN PULANG': {
      aksi: 'ABSEN_PULANG',
      label: 'PULANG'
    },

    'PULANG': {
      aksi: 'ABSEN_PULANG',
      label: 'PULANG'
    }
  };

  if (tombolDariEngine[tombolEngine]) {
    var model =
      tombolDariEngine[tombolEngine];

    var aksiMulaiBreak =
      model.aksi === 'MULAI_BREAK_1' ||
      model.aksi === 'MULAI_BREAK_2';

    var kapasitasMengunci =
      aksiMulaiBreak &&
      kapasitasPenuh;

    var aktif =
      aksiMulaiBreak
        ? (
            aksiEngineDiizinkan &&
            !kapasitasMengunci
          )
        : true;

    var alasan = '';

    if (!aksiEngineDiizinkan) {
      alasan = alasanEngine;
    } else if (kapasitasMengunci) {
      alasan =
        String(
          kapasitasBreak.warning || ''
        ).trim();
    }

    return apiBuatTombol_(
      model.aksi,
      model.label,
      aktif,
      alasan
    );
  }

  if (
    breakInfo.break1.status ===
    'SEDANG BREAK'
  ) {
    return apiBuatTombol_(
      'SELESAI_BREAK_1',
      'SELESAI BREAK 1',
      true,
      ''
    );
  }

  if (
    breakInfo.break2.status ===
    'SEDANG BREAK'
  ) {
    return apiBuatTombol_(
      'SELESAI_BREAK_2',
      'SELESAI BREAK 2',
      true,
      ''
    );
  }

  if (
    breakInfo.break1.status ===
      'BELUM DIAMBIL' &&
    breakInfo.break2.status ===
      'BELUM DIAMBIL'
  ) {
    return apiBuatTombol_(
      'MULAI_BREAK_1',
      'MULAI BREAK 1',
      !kapasitasPenuh,
      kapasitasPenuh
        ? String(
            kapasitasBreak.warning || ''
          ).trim()
        : ''
    );
  }

  if (
    breakInfo.break1.status ===
      'SELESAI' &&
    breakInfo.break2.status ===
      'BELUM DIAMBIL'
  ) {
    return apiBuatTombol_(
      'MULAI_BREAK_2',
      'MULAI BREAK 2',
      aksiEngineDiizinkan &&
      !kapasitasPenuh,
      !aksiEngineDiizinkan
        ? alasanEngine
        : (
            kapasitasPenuh
              ? String(
                  kapasitasBreak.warning ||
                  ''
                ).trim()
              : ''
          )
    );
  }

  if (
    breakInfo.break1.status ===
      'SELESAI' &&
    breakInfo.break2.status ===
      'SELESAI'
  ) {
    return apiBuatTombol_(
      'ABSEN_PULANG',
      'PULANG',
      true,
      ''
    );
  }

  return apiBuatTombol_(
    'REFRESH',
    'PERBARUI STATUS',
    true,
    ''
  );
}


/**
 * Membuat model tombol utama.
 *
 * @param {string} aksi
 * @param {string} label
 * @param {boolean} aktif
 * @param {string} alasan
 * @return {Object}
 */
function apiBuatTombol_(
  aksi,
  label,
  aktif,
  alasan
) {
  return {
    aksi:
      String(aksi || '').trim(),

    label:
      String(label || '').trim(),

    aktif:
      Boolean(aktif),

    alasan:
      String(alasan || '').trim()
  };
}


/**
 * Mencari status SA pada daftar dashboard.
 *
 * @param {Array} daftarStatus
 * @param {string|number} pinInput
 * @return {Object|null}
 */
function apiCariStatusSaDashboard_(
  daftarStatus,
  pinInput
) {
  var targetPin =
    String(pinInput || '').trim();

  if (
    !Array.isArray(daftarStatus) ||
    !targetPin
  ) {
    return null;
  }

  for (
    var index = 0;
    index < daftarStatus.length;
    index++
  ) {
    var item =
      daftarStatus[index] || {};

    var pinItem =
      String(
        item.pin ||
        item.PIN ||
        ''
      ).trim();

    if (pinItem === targetPin) {
      return item;
    }
  }

  return null;
}


/**
 * Mengubah status Break menjadi label UI.
 *
 * @param {*} status
 * @return {string}
 */
function apiLabelStatusBreak_(status) {
  var nilai =
    apiNormalisasiTeks_(status);

  if (nilai === 'SELESAI') {
    return 'Selesai';
  }

  if (nilai === 'SEDANG BREAK') {
    return 'Sedang Break';
  }

  if (nilai === 'DILEWATI') {
    return 'Dilewati';
  }

  return 'Belum diambil';
}

// ===== API.GS - PART 3 END =====
