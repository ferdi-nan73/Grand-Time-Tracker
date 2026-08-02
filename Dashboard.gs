/**
 * =====================================================
 * GRAND TIME TRACKER — GTT
 * Module : GTT-08 Dashboard Data Engine
 * Version: 1.1.0
 * Status : Development — TERMINOLOGI BREAK & ABSENSI TERBARU
 * =====================================================
 */


/**
 * Menghasilkan ringkasan dashboard hari ini.
 *
 * @param {string=} outletInput
 *   Kosong = seluruh outlet.
 *   Contoh = GP.
 *
 * @return {Object}
 */
function ambilRingkasanDashboard(
  outletInput,
  pinWaktuInput
) {
  try {
    const spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();

    const zonaWaktu =
      spreadsheet.getSpreadsheetTimeZone();

    const sekarang = gttSekarang_(pinWaktuInput);

    const tanggalHariIni = Utilities.formatDate(
      sekarang,
      zonaWaktu,
      'yyyy-MM-dd'
    );

    const filterOutlet =
      normalisasiDashboard_(outletInput);

    const daftarSA = ambilSaAktifDashboard_(
      spreadsheet,
      filterOutlet
    );

    const absensiMap = ambilAbsensiDashboard_(
      spreadsheet,
      tanggalHariIni,
      zonaWaktu
    );

    const breakMap = ambilBreakDashboard_(
      spreadsheet,
      tanggalHariIni,
      zonaWaktu
    );

    const ringkasan = buatRingkasanAwalDashboard_();

    const daftarStatus = [];

    daftarSA.forEach(sa => {
      const absensi =
        absensiMap[sa.pin] || null;

      const breakHariIni =
        breakMap[sa.pin] || null;

      const statusKehadiran =
        tentukanStatusKehadiranDashboard_(
          absensi
        );

      const statusOperasional =
  tentukanStatusOperasionalDashboard_(
    statusKehadiran,
    breakHariIni,
    absensi
        );

      ringkasan.totalSAAktif++;

      hitungStatusKehadiranDashboard_(
        ringkasan,
        statusKehadiran
      );

      hitungStatusOperasionalDashboard_(
        ringkasan,
        statusOperasional
      );

daftarStatus.push({
  pin: sa.pin,
  namaSA: sa.namaSA,
  outlet: sa.outlet,
  statusKehadiran: statusKehadiran,

  statusJamMasuk: absensi
    ? absensi.statusJamMasuk
    : '',

  statusJamPulang: absensi
    ? absensi.statusJamPulang
    : '',

  jamMasuk: absensi
    ? absensi.jamMasuk
    : '',

  jamPulang: absensi
    ? absensi.jamPulang
    : '',

  terlambatMenit: absensi
    ? absensi.terlambatMenit
    : 0,

  statusOperasional:
    statusOperasional
});
    });

    ringkasan.hadirTotal =
      ringkasan.hadirPenuh +
      ringkasan.hadirSetengahHari;

    ringkasan.sedangBreak =
      ringkasan.break1 +
      ringkasan.break2;

    ringkasan.sudahTercatat =
      ringkasan.totalSAAktif -
      ringkasan.belumAbsen;

    ringkasan.aktifBekerja =
  Math.max(
    0,
    ringkasan.hadirTotal -
    ringkasan.sedangBreak -
    ringkasan.sudahPulang
  );

    const statusOutlet =
      tentukanStatusOutletDashboard_(
        ringkasan
      );

    return {
      success: true,
      code: 'RINGKASAN_DASHBOARD_BERHASIL',
      message:
        'Ringkasan dashboard berhasil dibuat.',
      data: {
        tanggal: Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'dd/MM/yyyy'
        ),
        jamServer: Utilities.formatDate(
          sekarang,
          zonaWaktu,
          'HH:mm:ss'
        ),
        outlet: filterOutlet || 'SEMUA OUTLET',
        statusOutlet: statusOutlet,
        ringkasan: ringkasan,
        daftarStatus: daftarStatus
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
  }
}


/**
 * Mengambil seluruh SA aktif dari MASTER_SA.
 */
function ambilSaAktifDashboard_(
  spreadsheet,
  filterOutlet
) {
  const sheet =
    spreadsheet.getSheetByName('MASTER_SA');

  if (!sheet) {
    throw new Error(
      'Sheet MASTER_SA tidak ditemukan.'
    );
  }

  const headerMap =
    buatHeaderMapDashboard_(
      sheet,
      [
        'NAMA SA',
        'PIN',
        'OUTLET',
        'STATUS'
      ]
    );

  const barisTerakhir =
    sheet.getLastRow();

  if (barisTerakhir <= 1) {
    return [];
  }

  const data = sheet
    .getRange(
      2,
      1,
      barisTerakhir - 1,
      sheet.getLastColumn()
    )
    .getValues();

  return data
    .map(baris => {
      return {
        namaSA: String(
          baris[headerMap['NAMA SA']] || ''
        ).trim(),

        pin: String(
          baris[headerMap['PIN']] || ''
        ).trim(),

        outlet: String(
          baris[headerMap['OUTLET']] || ''
        ).trim(),

        status:
          baris[headerMap['STATUS']]
      };
    })
    .filter(sa => {
      if (
        !sa.namaSA ||
        !sa.pin ||
        !sa.outlet
      ) {
        return false;
      }

      if (
        !statusAktifDashboard_(sa.status)
      ) {
        return false;
      }

      if (
        filterOutlet &&
        normalisasiDashboard_(sa.outlet) !==
          filterOutlet
      ) {
        return false;
      }

      return true;
    });
}


/**
 * Mengambil absensi hari ini menjadi map berdasarkan PIN.
 * Data pertama yang ditemukan tetap digunakan.
 */
function ambilAbsensiDashboard_(
  spreadsheet,
  tanggalHariIni,
  zonaWaktu
) {
  const sheet =
    spreadsheet.getSheetByName(
      'ABSENSI_HARIAN'
    );

  if (!sheet) {
    throw new Error(
      'Sheet ABSENSI_HARIAN tidak ditemukan.'
    );
  }

  const headerMap =
    buatHeaderMapDashboard_(
      sheet,
      [
        'TANGGAL',
        'PIN',
        'JAM MASUK',
        'STATUS JAM MASUK',
        'TERLAMBAT MENIT',
        'JAM PULANG',
        'STATUS JAM PULANG',
        'STATUS KEHADIRAN'
      ]
    );

  const hasil = {};
  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    return hasil;
  }

  const data = sheet
    .getRange(
      2,
      1,
      barisTerakhir - 1,
      sheet.getLastColumn()
    )
    .getValues();

  /*
   * Baca dari bawah ke atas agar baris absensi terbaru digunakan.
   * Ini mencegah status BELUM ABSEN apabila terdapat lebih dari satu
   * baris untuk PIN dan tanggal yang sama.
   */
  for (let index = data.length - 1; index >= 0; index--) {
    const baris = data[index];

    const tanggalData =
      formatTanggalDashboard_(
        baris[headerMap['TANGGAL']],
        zonaWaktu
      );

    if (tanggalData !== tanggalHariIni) {
      continue;
    }

    const pin = String(
      baris[headerMap['PIN']] || ''
    ).trim();

    if (!pin || hasil[pin]) {
      continue;
    }

    const jamMasukValue =
      baris[headerMap['JAM MASUK']];

    /*
     * Abaikan baris kosong/tidak valid. Dashboard hanya menganggap
     * sudah absen apabila JAM MASUK benar-benar tersedia.
     */
    const jamMasukAda =
      jamMasukValue instanceof Date ||
      String(jamMasukValue || '').trim() !== '';

    if (!jamMasukAda) {
      continue;
    }

    const jamPulangValue =
      baris[headerMap['JAM PULANG']];

    hasil[pin] = {
      pin: pin,

      jamMasuk:
        jamMasukValue instanceof Date
          ? Utilities.formatDate(
              jamMasukValue,
              zonaWaktu,
              'HH:mm:ss'
            )
          : String(
              jamMasukValue || ''
            ).trim(),

      jamPulang:
        jamPulangValue instanceof Date
          ? Utilities.formatDate(
              jamPulangValue,
              zonaWaktu,
              'HH:mm:ss'
            )
          : String(
              jamPulangValue || ''
            ).trim(),

      statusJamMasuk:
        normalisasiDashboard_(
          baris[
            headerMap['STATUS JAM MASUK']
          ]
        ),

      statusJamPulang:
        normalisasiDashboard_(
          baris[
            headerMap['STATUS JAM PULANG']
          ]
        ),

      terlambatMenit:
        Number(
          baris[
            headerMap['TERLAMBAT MENIT']
          ]
        ) || 0,

      statusKehadiran:
        normalisasiDashboard_(
          baris[
            headerMap['STATUS KEHADIRAN']
          ]
        ),

      nomorBaris: index + 2
    };
  }

  return hasil;
}


/**
 * Mengambil data break hari ini menjadi map berdasarkan PIN.
 */
function ambilBreakDashboard_(
  spreadsheet,
  tanggalHariIni,
  zonaWaktu
) {
  const sheet =
    spreadsheet.getSheetByName(
      'LOG_ISTIRAHAT'
    );

  if (!sheet) {
    throw new Error(
      'Sheet LOG_ISTIRAHAT tidak ditemukan.'
    );
  }

  const headerMap =
    buatHeaderMapDashboard_(
      sheet,
      [
        'TANGGAL',
        'PIN',
        'S1 MULAI',
        'S1 SELESAI',
        'S1 OVERTIME',
        'S1 STATUS',
        'S2 MULAI',
        'S2 SELESAI',
        'S2 OVERTIME',
        'S2 STATUS',
        'TOTAL OVERTIME',
        'TOTAL SANKSI'
      ]
    );

  const hasil = {};
  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= 1) {
    return hasil;
  }

  const data = sheet
    .getRange(
      2,
      1,
      barisTerakhir - 1,
      sheet.getLastColumn()
    )
    .getValues();

  data.forEach(baris => {
    const tanggalData =
      formatTanggalDashboard_(
        baris[headerMap['TANGGAL']],
        zonaWaktu
      );

    if (tanggalData !== tanggalHariIni) {
      return;
    }

    const pin = String(
      baris[headerMap['PIN']] || ''
    ).trim();

    if (!pin || hasil[pin]) {
      return;
    }

    hasil[pin] = {
      break1Mulai:
        baris[headerMap['S1 MULAI']],

      break1Selesai:
        baris[headerMap['S1 SELESAI']],

      break1Overtime:
        Number(
          baris[headerMap['S1 OVERTIME']]
        ) || 0,

      break1Status:
        normalisasiDashboard_(
          baris[headerMap['S1 STATUS']]
        ),

      break2Mulai:
        baris[headerMap['S2 MULAI']],

      break2Selesai:
        baris[headerMap['S2 SELESAI']],

      break2Overtime:
        Number(
          baris[headerMap['S2 OVERTIME']]
        ) || 0,

      break2Status:
        normalisasiDashboard_(
          baris[headerMap['S2 STATUS']]
        ),

      totalOvertime:
        Number(
          baris[
            headerMap['TOTAL OVERTIME']
          ]
        ) || 0,

      totalSanksi:
        Number(
          baris[
            headerMap['TOTAL SANKSI']
          ]
        ) || 0
    };
  });

  return hasil;
}


/**
 * Menentukan status kehadiran SA.
 */
function tentukanStatusKehadiranDashboard_(
  absensi
) {
  if (!absensi) {
    return 'BELUM ABSEN';
  }

  const status =
    normalisasiDashboard_(
      absensi.statusKehadiran
    );

  if (
    status === 'HADIR' ||
    status === 'HADIR PENUH'
  ) {
    return 'HADIR PENUH';
  }

  if (
    status === '1/2 HARI' ||
    status === 'HADIR 1/2 HARI'
  ) {
    return 'HADIR 1/2 HARI';
  }

  if (
    status === 'IZIN' ||
    status === 'IJIN'
  ) {
    return 'IJIN';
  }

  const statusDiizinkan = [
    'BELUM ABSEN',
    'OFF',
    'IJIN',
    'ALPA'
  ];

  if (statusDiizinkan.includes(status)) {
    /*
     * BELUM ABSEN hanya valid jika memang belum ada JAM MASUK.
     * Baris absensi lama atau hasil input yang kolom statusnya kosong/tidak
     * konsisten tidak boleh membuat Dashboard menawarkan absen masuk lagi.
     */
    if (status !== 'BELUM ABSEN') {
      return status;
    }
  }

  const jamMasukAda =
    absensi.jamMasuk instanceof Date ||
    String(absensi.jamMasuk || '').trim() !== '';

  if (jamMasukAda) {
    const statusJamMasuk =
      normalisasiDashboard_(
        absensi.statusJamMasuk
      );

    if (
      statusJamMasuk === '1/2 HARI' ||
      statusJamMasuk === 'HADIR 1/2 HARI'
    ) {
      return 'HADIR 1/2 HARI';
    }

    return 'HADIR PENUH';
  }

  return 'BELUM ABSEN';
}


/**
 * Menentukan status operasional SA.
 */
function tentukanStatusOperasionalDashboard_(
  statusKehadiran,
  dataBreak,
  absensi
) {
  const statusTidakBekerja = [
    'BELUM ABSEN',
    'OFF',
    'IJIN',
    'ALPA'
  ];

  if (
    statusTidakBekerja.includes(
      statusKehadiran
    )
  ) {
    return statusKehadiran;
  }
if (
  absensi &&
  String(absensi.jamPulang || '').trim() !== ''
) {
  return 'SUDAH PULANG';
}
  if (!dataBreak) {
    return 'AKTIF BEKERJA';
  }


  const break1Aktif =
    dataBreak.break1Mulai instanceof Date &&
    !(dataBreak.break1Selesai instanceof Date);

  if (break1Aktif) {
    return 'BREAK 1';
  }

  const break2Aktif =
    dataBreak.break2Mulai instanceof Date &&
    !(dataBreak.break2Selesai instanceof Date);

  if (break2Aktif) {
    return 'BREAK 2';
  }

  return 'AKTIF BEKERJA';
}



/**
 * Menambah hitungan berdasarkan status kehadiran.
 */
function hitungStatusKehadiranDashboard_(
  ringkasan,
  status
) {
  switch (status) {
    case 'HADIR PENUH':
      ringkasan.hadirPenuh++;
      break;

    case 'HADIR 1/2 HARI':
      ringkasan.hadirSetengahHari++;
      break;

    case 'OFF':
      ringkasan.off++;
      break;

    case 'IJIN':
      ringkasan.ijin++;
      break;

    case 'ALPA':
      ringkasan.alpa++;
      break;

    default:
      ringkasan.belumAbsen++;
      break;
  }
}


/**
 * Menambah hitungan status operasional.
 */
function hitungStatusOperasionalDashboard_(
  ringkasan,
  status
) {
  if (status === 'BREAK 1') {
    ringkasan.break1++;
  }

  if (status === 'BREAK 2') {
    ringkasan.break2++;
  }

if (status === 'SUDAH PULANG') {
  ringkasan.sudahPulang++;
}

}

/**
 * Membuat struktur ringkasan awal.
 */
function buatRingkasanAwalDashboard_() {
  return {
    totalSAAktif: 0,
    sudahTercatat: 0,
    hadirTotal: 0,
    hadirPenuh: 0,
    hadirSetengahHari: 0,
    belumAbsen: 0,
    off: 0,
    ijin: 0,
    alpa: 0,
    aktifBekerja: 0,
    break1: 0,
    break2: 0,
    sedangBreak: 0,
    sudahPulang: 0,
  };
}


/**
 * Menentukan kondisi operasional outlet.
 */
function tentukanStatusOutletDashboard_(
  ringkasan
) {
  if (    ringkasan.alpa > 0
  ) {
    return 'PERLU PERHATIAN';
  }

  if (
    ringkasan.belumAbsen > 0 ||
    ringkasan.sedangBreak > 0
  ) {
    return 'ADA AKTIVITAS';
  }

  return 'OPERASIONAL NORMAL';
}


/**
 * Membuat map header dan memvalidasi kolom.
 */
function buatHeaderMapDashboard_(
  sheet,
  kolomWajib
) {
  const jumlahKolom =
    sheet.getLastColumn();

  if (jumlahKolom === 0) {
    throw new Error(
      'Header sheet ' +
      sheet.getName() +
      ' belum tersedia.'
    );
  }

  const daftarHeader = sheet
    .getRange(
      1,
      1,
      1,
      jumlahKolom
    )
    .getDisplayValues()[0]
    .map(normalisasiDashboard_);

  const headerMap = {};

  daftarHeader.forEach(
    (namaHeader, index) => {
      if (namaHeader) {
        headerMap[namaHeader] = index;
      }
    }
  );

  const kolomTidakAda =
    kolomWajib.filter(
      namaKolom =>
        headerMap[
          normalisasiDashboard_(namaKolom)
        ] === undefined
    );

  if (kolomTidakAda.length > 0) {
    throw new Error(
      'Kolom berikut tidak ditemukan di ' +
      sheet.getName() +
      ': ' +
      kolomTidakAda.join(', ')
    );
  }

  return headerMap;
}


/**
 * Mengecek nilai aktif.
 */
function statusAktifDashboard_(nilai) {
  if (nilai === true) {
    return true;
  }

  if (
    nilai === false ||
    nilai === null ||
    nilai === undefined
  ) {
    return false;
  }

  return [
    'TRUE',
    'AKTIF',
    'YA',
    'YES',
    '1'
  ].includes(
    normalisasiDashboard_(nilai)
  );
}


/**
 * Menormalisasi teks.
 */
function normalisasiDashboard_(nilai) {
  return String(nilai || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}


/**
 * Mengubah tanggal menjadi yyyy-MM-dd.
 */
function formatTanggalDashboard_(
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

  const teks =
    String(nilaiTanggal || '').trim();

  if (!teks) {
    return '';
  }

  const polaIndonesia =
    teks.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
    );

  if (polaIndonesia) {
    return [
      polaIndonesia[3],
      String(
        Number(polaIndonesia[2])
      ).padStart(2, '0'),
      String(
        Number(polaIndonesia[1])
      ).padStart(2, '0')
    ].join('-');
  }

  const tanggalPercobaan =
    new Date(teks);

  if (
    !isNaN(
      tanggalPercobaan.getTime()
    )
  ) {
    return Utilities.formatDate(
      tanggalPercobaan,
      zonaWaktu,
      'yyyy-MM-dd'
    );
  }

  return teks;
}


/**
 * =====================================================
 * FUNGSI PENGUJIAN
 * =====================================================
 */


/**
 * Menguji dashboard seluruh outlet.
 * Tidak mengubah data.
 */
function ujiRingkasanDashboardSemuaOutlet() {
  const hasil =
    ambilRingkasanDashboard('');

  console.log(
    JSON.stringify(hasil, null, 2)
  );
}


/**
 * Menguji dashboard Outlet GP.
 * Tidak mengubah data.
 */
function ujiRingkasanDashboardGP() {
  const hasil =
    ambilRingkasanDashboard('GP');

  console.log(
    JSON.stringify(hasil, null, 2)
  );
}

/**
 * Menguji status dashboard satu PIN tanpa mengubah data.
 */
function ujiStatusDashboardPin() {
  const PIN_UJI = '9616';
  const hasil = ambilRingkasanDashboard('');

  if (!hasil.success) {
    console.log(JSON.stringify(hasil, null, 2));
    return;
  }

  const status = hasil.data.daftarStatus.find(function (item) {
    return String(item.pin).trim() === String(PIN_UJI).trim();
  }) || null;

  console.log(
    JSON.stringify(
      {
        success: Boolean(status),
        code: status
          ? 'STATUS_PIN_DITEMUKAN'
          : 'STATUS_PIN_TIDAK_DITEMUKAN',
        data: status
      },
      null,
      2
    )
  );
}
