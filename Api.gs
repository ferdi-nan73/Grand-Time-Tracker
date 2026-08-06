/**
 * GRAND TIME TRACKER
 * File: Api.gs
 * Build: GTT v0.9.3-dev(2) Rev (3.6.1)
 * Build Date: 2026-08-03
 * Module: Activity State, Duplicate Lock, Dashboard Summary & Favicon Test
 */

function apiHealthCheck() {
  return apiJalankan_('HEALTH_CHECK', function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var now = new Date();
    var required = [
      'MASTER_SA',
      'MASTER_OUTLET',
      'MASTER_SETTING',
      'ABSENSI_HARIAN',
      'LOG_ISTIRAHAT'
    ];
    var sheets = {};
    required.forEach(function (name) {
      sheets[name] = Boolean(ss.getSheetByName(name));
    });
    var ok = required.every(function (name) { return sheets[name]; });
    return {
      success: ok,
      code: ok ? 'API_HEALTHY' : 'API_DEPENDENCY_ERROR',
      message: ok ? 'Backend GTT berjalan normal.' : 'Ada sheet wajib yang tidak ditemukan.',
      data: {
        aplikasi: 'GTT',
        versiApi: '2.3.1',
        waktuServer: {
          tanggal: Utilities.formatDate(now, tz, 'dd/MM/yyyy'),
          jam: Utilities.formatDate(now, tz, 'HH:mm:ss'),
          timestamp: now.getTime(),
          zonaWaktu: tz
        },
        sheet: sheets
      }
    };
  });
}

function apiLogin(pinInput) {
  return apiJalankan_('LOGIN', function () {
    return validasiLoginPin(pinInput);
  });
}

function apiAbsenMasuk(pinInput, lokasiInput) {
  return apiJalankan_('ABSEN_MASUK', function () {
    var validasiLokasi =
      apiValidasiLokasiAbsensi_(pinInput, lokasiInput, 'MASUK');

    if (!validasiLokasi.success) {
      return validasiLokasi;
    }

    return simpanAbsensiMasuk(pinInput);
  });
}

function apiAbsenPulang(pinInput, lokasiInput) {
  return apiJalankan_('ABSEN_PULANG', function () {
    var validasiLokasi =
      apiValidasiLokasiAbsensi_(pinInput, lokasiInput, 'PULANG');

    if (!validasiLokasi.success) {
      return validasiLokasi;
    }

    return simpanAbsensiPulang(pinInput);
  });
}

/**
 * Rev (3.6.1)
 * Validasi geofence untuk Absen Masuk dan Absen Pulang.
 * Titik acuan selalu outlet SA pada MASTER_LOKASI.
 */
function apiValidasiLokasiAbsensi_(pinInput, lokasiInput, jenisAksi) {
  var login = validasiLoginPin(pinInput);

  if (!login || login.success !== true) {
    return login || {
      success: false,
      code: 'LOGIN_TIDAK_VALID',
      message: 'Sesi login tidak valid.'
    };
  }

  var user =
    apiNormalisasiPengguna_(
      login.data,
      pinInput
    );

  var outlet =
    String(user.outlet || '').trim();

  if (!outlet) {
    return {
      success: false,
      code: 'OUTLET_TIDAK_DITEMUKAN',
      message: 'Outlet SA belum terdaftar pada MASTER_SA.'
    };
  }

  var hasil =
    gttValidasiLokasiOutletTugasLuar_(
      outlet,
      lokasiInput || {}
    );

  if (hasil && hasil.success === true) {
    return hasil;
  }

  var aksi =
    String(jenisAksi || '').toUpperCase() === 'PULANG'
      ? 'Absen Pulang'
      : 'Absen Masuk';

  var data =
    hasil && hasil.data
      ? hasil.data
      : {};

  if (
    hasil &&
    hasil.code === 'DI_LUAR_RADIUS_OUTLET'
  ) {
    return {
      success: false,
      code: 'ABSEN_DI_LUAR_RADIUS',
      message:
        aksi +
        ' hanya dapat dilakukan di area ' +
        (data.outlet || outlet) +
        '. Jarak Anda sekitar ' +
        (data.jarakMeter || 0) +
        ' meter, sedangkan radius yang diizinkan ' +
        (data.radiusMeter || 0) +
        ' meter.',
      data: data
    };
  }

  return {
    success: false,
    code:
      hasil && hasil.code
        ? hasil.code
        : 'LOKASI_TIDAK_VALID',
    message:
      hasil && hasil.message
        ? hasil.message
        : 'Lokasi perangkat belum dapat divalidasi.',
    data: data
  };
}

function apiMulaiBreak1(pinInput) {
  return apiJalankan_('MULAI_BREAK_1', function () {
    return mulaiBreak1(pinInput);
  });
}

function apiSelesaiBreak1(pinInput) {
  return apiJalankan_('SELESAI_BREAK_1', function () {
    return selesaiBreak1(pinInput);
  });
}

function apiMulaiBreak2(pinInput) {
  return apiJalankan_('MULAI_BREAK_2', function () {
    return mulaiBreak2(pinInput);
  });
}

function apiSelesaiBreak2(pinInput) {
  return apiJalankan_('SELESAI_BREAK_2', function () {
    return selesaiBreak2(pinInput);
  });
}


function apiMulaiTugasLuar(pinInput, formInput, lokasiInput) {
  return apiJalankan_('MULAI_TUGAS_LUAR', function () {
    return gttMulaiTugasLuar_(pinInput, formInput, lokasiInput);
  });
}

function apiSelesaiTugasLuar(pinInput, lokasiInput) {
  return apiJalankan_('SELESAI_TUGAS_LUAR', function () {
    return gttSelesaiTugasLuar_(pinInput, lokasiInput);
  });
}

function apiStatusBreakHariIni(pinInput) {
  return apiJalankan_('STATUS_BREAK_HARI_INI', function () {
    var login = validasiLoginPin(pinInput);
    if (!login.success) return login;

    var user = apiNormalisasiPengguna_(login.data, pinInput);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var infoWaktu = gttInfoWaktu_(user.pin);
    var now = infoWaktu.sekarang;
    var tanggalOperasional = Utilities.formatDate(
    infoWaktu.waktuServerAsli,
    tz,
    'yyyy-MM-dd'
  );
  var breakInfo = apiAmbilBreakInfoPribadi_(
    ss,
    user.pin,
    now,
    tz,
    tanggalOperasional
  );
    var settings = apiAmbilSettingBreak_(ss);
    var state = apiBangunStatusBreak_(breakInfo, settings, now, tz);

    return {
      success: true,
      code: 'STATUS_BREAK_BERHASIL',
      message: 'Status break berhasil dimuat.',
      data: state
    };
  });
}

function apiRingkasanDashboard(outletInput) {
  return apiJalankan_('RINGKASAN_DASHBOARD', function () {
    return ambilRingkasanDashboard(outletInput || '');
  });
}

function apiBootstrapSA(pinInput) {
  return apiJalankan_('BOOTSTRAP_SA', function () {
    return apiBangunDataSA_(pinInput, 'BOOTSTRAP_SA_BERHASIL');
  });
}

function apiRefreshStatusSA(pinInput) {
  return apiJalankan_('REFRESH_STATUS_SA', function () {
    return apiBangunDataSA_(pinInput, 'REFRESH_STATUS_SA_BERHASIL');
  });
}

function apiBangunDataSA_(pinInput, successCode) {
  var login = validasiLoginPin(pinInput);
  if (!login.success) return login;

  var user = apiNormalisasiPengguna_(login.data, pinInput);
  var dashboard = ambilRingkasanDashboard(user.outlet, user.pin);
  if (!dashboard.success) return dashboard;

  var list = dashboard.data && Array.isArray(dashboard.data.daftarStatus)
    ? dashboard.data.daftarStatus
    : [];
  var statusSA = apiCariStatusSaDashboard_(list, user.pin);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();

  // Dashboard outlet dapat terlambat memperbarui daftarStatus sesaat
  // setelah transaksi. Jangan menggagalkan sesi pengguna. Ambil data
  // kehadiran langsung dari ABSENSI_HARIAN sebagai fallback.
  if (!statusSA) {
    statusSA = apiBangunStatusSaFallback_(
      ss,
      user,
      tz
    );
  }
  var infoWaktu = gttInfoWaktu_(user.pin);
  var now = infoWaktu.sekarang;
  var tanggalOperasional = Utilities.formatDate(
    infoWaktu.waktuServerAsli,
    tz,
    'yyyy-MM-dd'
  );
  var breakInfo = apiAmbilBreakInfoPribadi_(
    ss,
    user.pin,
    now,
    tz,
    tanggalOperasional
  );
  var settings = apiAmbilSettingBreak_(ss);
  var breakState = apiBangunStatusBreak_(breakInfo, settings, now, tz);
  var attendance = apiBangunKehadiran_(statusSA);
  var summary = dashboard.data && dashboard.data.ringkasan
    ? dashboard.data.ringkasan
    : {};
  var jumlahSedangBreak = Number(summary.sedangBreak || 0);
  var capacityFull = jumlahSedangBreak >= settings.maxBreakBersamaan;
  var button = apiTentukanTombolUtama_(
    attendance,
    statusSA,
    breakState,
    breakInfo,
    capacityFull,
    settings
  );

  return {
    success: true,
    code: successCode,
    message: 'Data dashboard berhasil dimuat.',
    data: {
      pengguna: user,
      waktuServer: {
        tanggal: Utilities.formatDate(now, tz, 'dd/MM/yyyy'),
        jam: Utilities.formatDate(now, tz, 'HH:mm:ss'),
        timestamp: now.getTime(),
        zonaWaktu: tz,
        modeUji: Boolean(infoWaktu.modeUjiAktif),
        tambahMenitUji: Number(infoWaktu.tambahMenitDiterapkan || 0)
      },
      kehadiran: attendance,
      break: {
        statusOperasional: String(statusSA.statusOperasional || breakState.statusOperasional || 'BELUM ABSEN'),
        break1: breakInfo.break1,
        break2: breakInfo.break2,
        overbreakMenit: {
          break1: Number(breakInfo.break1.overtimeMenit || 0),
          break2: Number(breakInfo.break2.overtimeMenit || 0)
        },
        seluruhBreakSelesai: breakInfo.break1.status === 'SELESAI' && breakInfo.break2.status === 'SELESAI',
        break1Dilewati: breakInfo.break1.status === 'DILEWATI',
        aksiDiizinkan: breakState.aksiDiizinkan,
        alasanTerkunci: breakState.alasanTerkunci,
        kodeTerkunci: breakState.kodeTerkunci,
        tersediaPada: breakState.tersediaPada,
        sisaTungguMenit: breakState.sisaTungguMenit,
        jumlahSedangBreak: jumlahSedangBreak,
        maksimalBreak: settings.maxBreakBersamaan,
        kapasitasPenuh: capacityFull,
        warningKapasitas: capacityFull ? settings.warningKapasitas : ''
      },
      tugasLuar:
        gttAmbilDataTugasLuar_(user),

      operasional: {
        statusOperasional: String(statusSA.statusOperasional || breakState.statusOperasional || 'BELUM ABSEN'),
        tombolUtama: button
      },
      ringkasan: {
        breakOutlet: {
          jumlahSedangBreak: jumlahSedangBreak,
          maksimalBreak: settings.maxBreakBersamaan,
          kapasitasPenuh: capacityFull,
          teks: jumlahSedangBreak + ' dari ' + settings.maxBreakBersamaan + ' orang sedang break'
        },
        breakInfo: {
          break1: breakInfo.break1.status,
          break2: breakInfo.break2.status,
          teks: 'B1 ' + apiLabelStatusBreak_(breakInfo.break1.status) +
            ' · B2 ' + apiLabelStatusBreak_(breakInfo.break2.status)
        },
        tugasLuar: {
          aktif: Number(
            gttHitungTugasLuarAktifOutlet_(user.outlet)
          ),
          pulang: 0,
          tersedia: true
        },
        pelanggaran: {
          jumlah: Number(breakInfo.pelanggaran.jumlah || 0),
          tersedia: true
        },
        overbreak: {
          break1Menit: Number(breakInfo.break1.overtimeMenit || 0),
          break2Menit: Number(breakInfo.break2.overtimeMenit || 0),
          tersedia: true
        }
      },
      dashboard: {
        statusOutlet: dashboard.data.statusOutlet || {},
        ringkasanOutlet: summary,
        statusSA: statusSA
      }
    }
  };
}

function apiBangunStatusBreak_(breakInfo, settings, now, tz) {
  var result = {
    statusOperasional: 'AKTIF BEKERJA',
    tombolBerikutnya: 'MULAI BREAK 1',
    aksiDiizinkan: true,
    alasanTerkunci: '',
    kodeTerkunci: '',
    tersediaPada: '',
    sisaTungguMenit: 0
  };

  if (breakInfo.break1.status === 'SEDANG BREAK') {
    result.statusOperasional = 'BREAK 1';
    result.tombolBerikutnya = 'SELESAI BREAK 1';
    return result;
  }
  if (breakInfo.break2.status === 'SEDANG BREAK') {
    result.statusOperasional = 'BREAK 2';
    result.tombolBerikutnya = 'SELESAI BREAK 2';
    return result;
  }
  if (breakInfo.break1.status === 'BELUM DIAMBIL') {
    result.tombolBerikutnya = 'MULAI BREAK 1';
    return result;
  }
  if (breakInfo.break1.status === 'SELESAI' && breakInfo.break2.status === 'BELUM DIAMBIL') {
    result.tombolBerikutnya = 'MULAI BREAK 2';
    var selesai1 = apiNilaiKeDate_(breakInfo.break1.selesaiRaw, now, tz);
    if (selesai1) {
      var available = new Date(selesai1.getTime() + settings.jedaBreak1KeBreak2 * 60000);
      var diff = Math.ceil((available.getTime() - now.getTime()) / 60000);
      if (diff > 0) {
        result.aksiDiizinkan = false;
        result.kodeTerkunci = 'JEDA_BREAK_2_BELUM_TERPENUHI';
        result.tersediaPada = Utilities.formatDate(available, tz, 'HH:mm');
        result.sisaTungguMenit = diff;
        result.alasanTerkunci = apiIsiTemplate_(settings.warningJedaBreak2, {
          JAM: result.tersediaPada,
          MENIT: String(diff)
        });
      }
    }
    return result;
  }
  if (breakInfo.break1.status === 'SELESAI' && breakInfo.break2.status === 'SELESAI') {
    result.tombolBerikutnya = 'PULANG';
  }
  return result;
}

function apiTentukanTombolUtama_(attendance, statusSA, breakState, breakInfo, capacityFull, settings) {
  if (!attendance.sudahAbsenMasuk) {
    return apiBuatTombol_('ABSEN_MASUK', 'ABSEN MASUK', true, '');
  }
  if (attendance.sudahPulang) {
    return apiBuatTombol_('SELESAI', 'SELESAI', false, 'Absensi hari ini sudah selesai.');
  }

  var map = {
    'MULAI BREAK 1': ['MULAI_BREAK_1', 'MULAI BREAK 1'],
    'SELESAI BREAK 1': ['SELESAI_BREAK_1', 'SELESAI BREAK 1'],
    'MULAI BREAK 2': ['MULAI_BREAK_2', 'MULAI BREAK 2'],
    'SELESAI BREAK 2': ['SELESAI_BREAK_2', 'SELESAI BREAK 2'],
    'PULANG': ['ABSEN_PULANG', 'PULANG']
  };
  var key = apiNormalisasiTeks_(breakState.tombolBerikutnya);
  var model = map[key] || ['REFRESH', 'PERBARUI STATUS'];
  var isStartBreak = model[0] === 'MULAI_BREAK_1' || model[0] === 'MULAI_BREAK_2';
  var enabled = breakState.aksiDiizinkan !== false && !(isStartBreak && capacityFull);
  var reason = breakState.alasanTerkunci || '';

  if (isStartBreak && capacityFull) {
    reason = apiIsiTemplate_(
      settings.warningKapasitas,
      {
        JUMLAH:
          String(
            settings.maxBreakBersamaan
          ),
        MAKSIMAL:
          String(
            settings.maxBreakBersamaan
          )
      }
    );
  }

  return apiBuatTombol_(
    model[0],
    model[1],
    enabled,
    reason
  );
}

function apiBuatTombol_(action, label, enabled, reason) {
  return {
    aksi: String(action || ''),
    label: String(label || ''),
    aktif: Boolean(enabled),
    alasan: String(reason || '')
  };
}

function apiNormalisasiPengguna_(data, pinInput) {
  var source = data || {};
  return {
    saId: String(source.saId || source.SA_ID || '').trim(),
    pin: String(source.pin || source.PIN || pinInput || '').trim(),
    namaSA: String(source.namaSA || source.nama || source['NAMA SA'] || '').trim(),
    outlet: String(source.outlet || source.OUTLET || '').trim(),
    status: String(source.status || source.STATUS || '').trim().toUpperCase(),
    jabatan: String(source.jabatan || source.JABATAN || '').trim(),
    role: String(source.role || source.ROLE || 'SA').trim().toUpperCase()
  };
}

function apiBangunKehadiran_(statusSA) {
  var source = statusSA || {};
  var jamMasuk = apiFormatJam_(source.jamMasuk);
  var jamPulang = apiFormatJam_(source.jamPulang);
  return {
    sudahAbsenMasuk: Boolean(jamMasuk),
    sudahPulang: Boolean(jamPulang),
    statusKehadiran: String(source.statusKehadiran || 'BELUM ABSEN'),
    statusJamMasuk: String(source.statusJamMasuk || ''),
    statusJamPulang: String(source.statusJamPulang || ''),
    jamMasuk: jamMasuk,
    jamPulang: jamPulang,
    terlambatMenit: Number(source.terlambatMenit || 0),
    keterangan: String(source.keterangan || '')
  };
}

function apiCariStatusSaDashboard_(list, pinInput) {
  var pin = String(pinInput || '').trim();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].pin || '').trim() === pin) return list[i];
  }
  return null;
}


/**
 * Membuat status SA langsung dari ABSENSI_HARIAN ketika daftarStatus
 * Dashboard belum tersedia atau belum selesai diperbarui.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} user
 * @param {string} tz
 * @return {Object}
 */
function apiBangunStatusSaFallback_(ss, user, tz) {
  var fallback = {
    pin: String(user.pin || '').trim(),
    namaSA: String(user.namaSA || '').trim(),
    outlet: String(user.outlet || '').trim(),
    statusKehadiran: 'BELUM ABSEN',
    statusJamMasuk: '',
    statusJamPulang: '',
    jamMasuk: '',
    jamPulang: '',
    terlambatMenit: 0,
    statusOperasional: 'BELUM ABSEN',
    keterangan: '',
    sumberData: 'FALLBACK_ABSENSI_HARIAN'
  };

  if (!ss) return fallback;

  var sheet = ss.getSheetByName('ABSENSI_HARIAN');
  if (!sheet || sheet.getLastRow() < 2) return fallback;

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return fallback;

  var headers = apiHeaderMap_(values[0]);
  var pinIndex = apiHeaderIndex_(headers, ['PIN']);
  var dateIndex = apiHeaderIndex_(headers, ['TANGGAL']);

  if (pinIndex < 0 || dateIndex < 0) return fallback;

  var targetPin = String(user.pin || '').trim();
  var targetDate = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i];
    var rowPin = String(row[pinIndex] || '').trim();
    if (rowPin !== targetPin) continue;

    var rowDate = apiFormatTanggalIso_(row[dateIndex], tz);
    if (rowDate !== targetDate) continue;

    var jamMasukIndex = apiHeaderIndex_(headers, ['JAM MASUK']);
    var jamPulangIndex = apiHeaderIndex_(headers, ['JAM PULANG']);
    var statusKehadiranIndex = apiHeaderIndex_(headers, ['STATUS KEHADIRAN']);
    var statusJamMasukIndex = apiHeaderIndex_(headers, ['STATUS JAM MASUK']);
    var statusJamPulangIndex = apiHeaderIndex_(headers, ['STATUS JAM PULANG']);
    var terlambatIndex = apiHeaderIndex_(headers, ['TERLAMBAT MENIT']);
    var keteranganIndex = apiHeaderIndex_(headers, ['KETERANGAN']);

    fallback.jamMasuk = jamMasukIndex >= 0
      ? apiFormatJam_(row[jamMasukIndex])
      : '';
    fallback.jamPulang = jamPulangIndex >= 0
      ? apiFormatJam_(row[jamPulangIndex])
      : '';
    fallback.statusKehadiran = statusKehadiranIndex >= 0
      ? String(row[statusKehadiranIndex] || 'BELUM ABSEN')
      : (fallback.jamMasuk ? 'HADIR' : 'BELUM ABSEN');
    fallback.statusJamMasuk = statusJamMasukIndex >= 0
      ? String(row[statusJamMasukIndex] || '')
      : '';
    fallback.statusJamPulang = statusJamPulangIndex >= 0
      ? String(row[statusJamPulangIndex] || '')
      : '';
    fallback.terlambatMenit = terlambatIndex >= 0
      ? Number(row[terlambatIndex] || 0)
      : 0;
    fallback.keterangan = keteranganIndex >= 0
      ? String(row[keteranganIndex] || '')
      : '';
    fallback.statusOperasional = fallback.jamPulang
      ? 'SUDAH PULANG'
      : (fallback.jamMasuk ? 'AKTIF BEKERJA' : 'BELUM ABSEN');

    return fallback;
  }

  return fallback;
}

function apiAmbilBreakInfoPribadi_(ss, pinInput, now, tz, tanggalOperasional) {
  var empty = apiBreakKosong_();
  var sheet = ss.getSheetByName('LOG_ISTIRAHAT');
  if (!sheet || sheet.getLastRow() < 2) return empty;

  var values = sheet.getDataRange().getValues();
  var headers = apiHeaderMap_(values[0]);
  var pinIndex = apiHeaderIndex_(headers, ['PIN']);
  var dateIndex = apiHeaderIndex_(headers, ['TANGGAL']);
  if (pinIndex < 0 || dateIndex < 0) return empty;

  var targetPin = String(pinInput || '').trim();
  var targetDate = String(tanggalOperasional || '').trim() ||
    Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var row = null;
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][pinIndex] || '').trim() !== targetPin) continue;
    if (apiFormatTanggalIso_(values[i][dateIndex], tz) !== targetDate) continue;
    row = values[i];
    break;
  }
  if (!row) return empty;

  var b1 = apiBreakDariBaris_(row, headers, 'S1', tz);
  var b2 = apiBreakDariBaris_(row, headers, 'S2', tz);
  var violations = (b1.overtimeMenit > 0 ? 1 : 0) + (b2.overtimeMenit > 0 ? 1 : 0);
  return {
    break1: b1,
    break2: b2,
    pelanggaran: {
      jumlah: violations,
      totalOvertimeMenit: b1.overtimeMenit + b2.overtimeMenit,
      totalSanksi: b1.sanksi + b2.sanksi
    }
  };
}

function apiBreakKosong_() {
  return {
    break1: { status: 'BELUM DIAMBIL', mulai: '', selesai: '', mulaiRaw: '', selesaiRaw: '', durasiMenit: 0, overtimeMenit: 0, sanksi: 0 },
    break2: { status: 'BELUM DIAMBIL', mulai: '', selesai: '', mulaiRaw: '', selesaiRaw: '', durasiMenit: 0, overtimeMenit: 0, sanksi: 0 },
    pelanggaran: { jumlah: 0, totalOvertimeMenit: 0, totalSanksi: 0 }
  };
}

function apiBreakDariBaris_(row, headers, prefix, tz) {
  var mulaiRaw = apiCell_(row, headers, prefix + ' MULAI');
  var selesaiRaw = apiCell_(row, headers, prefix + ' SELESAI');
  var storedStatus = apiNormalisasiTeks_(apiCell_(row, headers, prefix + ' STATUS'));
  var status = 'BELUM DIAMBIL';
  if (mulaiRaw && !selesaiRaw) status = 'SEDANG BREAK';
  if (mulaiRaw && selesaiRaw) status = 'SELESAI';
  if (storedStatus === 'DILEWATI') status = 'DILEWATI';
  return {
    status: status,
    mulai: apiFormatJamDenganZona_(mulaiRaw, tz),
    selesai: apiFormatJamDenganZona_(selesaiRaw, tz),

    // google.script.run tidak boleh menerima objek Date mentah
    // di dalam payload. Konversi ke string ISO agar bootstrap
    // dan refresh dashboard tetap berhasil saat Break aktif.
    mulaiRaw: apiNilaiSerializable_(mulaiRaw),
    selesaiRaw: apiNilaiSerializable_(selesaiRaw),

    durasiMenit: Number(apiCell_(row, headers, prefix + ' DURASI') || 0),
    overtimeMenit: Number(apiCell_(row, headers, prefix + ' OVERTIME') || 0),
    statusTersimpan: storedStatus,
    sanksi: Number(apiCell_(row, headers, prefix + ' SANKSI') || 0)
  };
}

function apiAmbilSettingBreak_(ss) {
  var result = {
    maxBreakBersamaan: 5,
    jedaBreak1KeBreak2: 120,
    warningKapasitas: 'Kapasitas break sedang penuh. Silakan menunggu SA lain menyelesaikan break.',
    warningJedaBreak2: 'ANDA BARU MENYELESAIKAN BREAK 1. LANJUTKAN BEKERJA TERLEBIH DAHULU. BREAK 2 DAPAT DIMULAI PUKUL {{JAM}}.'
  };
  var sheet = ss.getSheetByName('MASTER_SETTING');
  if (!sheet || sheet.getLastRow() < 2) return result;
  var rows = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    var key = apiNormalisasiTeks_(rows[i][0]);
    var value = rows[i][1];
    var warning = String(rows[i][2] || '').trim();
    if (key === 'BREAK - MAKSIMAL SA BREAK BERSAMAAN') {
      result.maxBreakBersamaan = apiPositiveNumber_(value, result.maxBreakBersamaan);
      if (warning) result.warningKapasitas = warning;
    }
    if (key === 'JEDA BREAK 1 KE BREAK 2 (MENIT)') {
      result.jedaBreak1KeBreak2 = apiPositiveNumber_(value, result.jedaBreak1KeBreak2);
      if (warning) result.warningJedaBreak2 = warning;
    }
  }
  return result;
}

function apiHeaderMap_(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var key = apiNormalisasiTeks_(headerRow[i]);
    if (key) map[key] = i;
  }
  return map;
}

function apiHeaderIndex_(map, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var key = apiNormalisasiTeks_(candidates[i]);
    if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
  }
  return -1;
}

function apiCell_(row, map, header) {
  var index = apiHeaderIndex_(map, [header]);
  return index >= 0 ? row[index] : '';
}

function apiNormalisasiTeks_(value) {
  return String(value === null || value === undefined ? '' : value).trim().toUpperCase();
}

function apiFormatJam_(value) {
  return apiFormatJamDenganZona_(value, Session.getScriptTimeZone());
}

/**
 * Mengubah nilai menjadi tipe yang aman dikirim melalui google.script.run.
 * Date dikonversi ke ISO string; tipe primitif dipertahankan.
 *
 * @param {*} value
 * @return {*}
 */
function apiNilaiSerializable_(value) {
  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return value.toISOString();
  }

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return value;
}

function apiFormatJamDenganZona_(value, tz) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'HH:mm:ss');
  }
  var text = String(value).trim();
  var match = text.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return text;
  return String(match[1]).padStart(2, '0') + ':' + match[2] + ':' + String(match[3] || '00').padStart(2, '0');
}

function apiFormatTanggalIso_(value, tz) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  var text = String(value).trim();
  var ymd = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (ymd) return ymd[1] + '-' + String(ymd[2]).padStart(2, '0') + '-' + String(ymd[3]).padStart(2, '0');
  var dmy = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (dmy) return dmy[3] + '-' + String(dmy[2]).padStart(2, '0') + '-' + String(dmy[1]).padStart(2, '0');
  var date = new Date(text);
  return isNaN(date.getTime()) ? text : Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function apiNilaiKeDate_(value, referenceDate, tz) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  var time = apiFormatJamDenganZona_(value, tz);
  var match = time.match(/^(\d{2}):(\d{2}):(\d{2})$/);
  if (!match) return null;
  var date = new Date(referenceDate.getTime());
  date.setHours(Number(match[1]), Number(match[2]), Number(match[3]), 0);
  return date;
}

function apiPositiveNumber_(value, fallback) {
  var n = Number(value);
  return !isNaN(n) && n > 0 ? n : fallback;
}

function apiIsiTemplate_(template, replacements) {
  var text = String(template || '');
  Object.keys(replacements || {}).forEach(function (key) {
    text = text.replace(new RegExp('\\{\\{' + key + '\\}\\}', 'g'), String(replacements[key]));
  });
  return text;
}

function apiLabelStatusBreak_(status) {
  var value = apiNormalisasiTeks_(status);
  if (value === 'SELESAI') return 'Selesai';
  if (value === 'SEDANG BREAK') return 'Sedang Break';
  if (value === 'DILEWATI') return 'Dilewati';
  return 'Belum diambil';
}

function apiJalankan_(processName, callback) {
  var started = new Date();
  try {
    if (typeof callback !== 'function') throw new Error('Callback API tidak valid.');
    var result = callback();
    if (!result || typeof result !== 'object') {
      result = { success: false, code: 'API_RESPONSE_TIDAK_VALID', message: 'Respons API tidak valid.' };
    }
    result.meta = apiBuatMeta_(processName, started);
    return result;
  } catch (error) {
    console.error('[' + processName + ']', error && error.stack ? error.stack : error);
    return {
      success: false,
      code: 'API_SYSTEM_ERROR',
      message: 'Terjadi kesalahan sistem: ' + (error && error.message ? error.message : String(error)),
      meta: apiBuatMeta_(processName, started)
    };
  }
}

function apiBuatMeta_(processName, started) {
  var finished = new Date();
  return {
    proses: String(processName || ''),
    durasiMs: finished.getTime() - started.getTime(),
    timestamp: finished.toISOString()
  };
}

function testApiGs() {
  var health = apiHealthCheck();
  var result = {
    success: Boolean(health && typeof health.success === 'boolean'),
    code: health && typeof health.success === 'boolean' ? 'TEST_API_PASS' : 'TEST_API_FAIL',
    message: health && typeof health.success === 'boolean' ? 'Struktur Api.gs berhasil diuji.' : 'Struktur Api.gs belum valid.',
    data: { healthCheck: health }
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}


function gttHitungTugasLuarAktifOutlet_(outlet) {
  var sheet =
    gttPastikanSheetTugasLuar_();

  var values =
    sheet.getDataRange().getValues();

  if (values.length < 2) {
    return 0;
  }

  var map =
    gttHeaderMapAktivitas_(
      values[0]
    );

  var uniquePin = {};

  values.slice(1).forEach(function(row) {
    var rowOutlet =
      String(
        row[map['OUTLET']] || ''
      ).trim();

    var status =
      String(
        row[map['STATUS']] || ''
      ).trim().toUpperCase();

    var pin =
      String(
        row[map['PIN']] || ''
      ).trim();

    if (
      rowOutlet ===
        String(outlet || '').trim() &&
      status === 'AKTIF' &&
      pin
    ) {
      uniquePin[pin] = true;
    }
  });

  return Object.keys(
    uniquePin
  ).length;
}
