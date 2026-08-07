/**
 * GRAND TIME TRACKER
 * File: Activity.gs
 * Build: GTT v0.9.3-dev(2) Rev (3.6.1)
 * Build Date: 2026-08-03
 * Module: Single Active Task, Active Panel, Home Badge & Return Geofence
 */

function gttPastikanSheetTugasLuar_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'LOG_AKTIVITAS';
  var sheet = ss.getSheetByName(name);

  var headers = [
    'ID',
    'TANGGAL',
    'PIN',
    'NAMA SA',
    'OUTLET',
    'JABATAN',
    'TUJUAN',
    'KEPERLUAN',
    'DISETUJUI PIN',
    'DISETUJUI OLEH',
    'DISETUJUI JABATAN',
    'JAM KELUAR',
    'JAM KEMBALI',
    'DURASI MENIT',
    'STATUS',
    'CATATAN',
    'JARAK KELUAR (M)',
    'AKURASI GPS KELUAR (M)',
    'JARAK KEMBALI (M)',
    'AKURASI GPS KEMBALI (M)'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    var existingHeaders =
      sheet.getRange(
        1,
        1,
        1,
        Math.max(1, sheet.getLastColumn())
      ).getDisplayValues()[0];

    var existingMap =
      gttHeaderMapAktivitas_(existingHeaders);

    headers.forEach(function(header) {
      if (existingMap[header] === undefined) {
        sheet
          .getRange(1, sheet.getLastColumn() + 1)
          .setValue(header);
      }
    });
  }

  return sheet;
}

function gttMulaiTugasLuar_(pinInput, formInput, lokasiInput) {
  var login = validasiLoginPin(pinInput);
  if (!login.success) return login;

  var form = formInput || {};
  var lokasi = lokasiInput || {};
  var tujuan = String(form.tujuan || '').trim();
  var keperluan = String(form.keperluan || '').trim();
  var pemberiIzinPin = String(form.pemberiIzinPin || '').trim();
  var pemberiIzinNama = String(form.pemberiIzinNama || '').trim();
  var pemberiIzinJabatan = String(form.pemberiIzinJabatan || '').trim();

  if (!tujuan) return gttResponTugasLuar_(false, 'TUJUAN_WAJIB', 'Tujuan wajib diisi.');
  if (!keperluan) return gttResponTugasLuar_(false, 'KEPERLUAN_WAJIB', 'Keperluan wajib diisi.');
  if (!pemberiIzinPin) return gttResponTugasLuar_(false, 'IZIN_WAJIB', 'Pemberi izin wajib dipilih.');

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    return gttResponTugasLuar_(
      false,
      'TRANSAKSI_SEDANG_DIPROSES',
      'Transaksi Tugas Luar sedang diproses. Silakan tunggu beberapa detik.'
    );
  }

  try {
    var user = apiNormalisasiPengguna_(login.data, pinInput);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var info = gttInfoWaktu_(user.pin);
    var now = info.sekarang;
    var tanggalTampilan =
      Utilities.formatDate(info.waktuServerAsli, tz, 'dd/MM/yyyy');

    var absen = gttCariAbsensiMasukTugasLuar_(
      ss,
      user.pin,
      info.waktuServerAsli,
      tz
    );

    if (!absen.ditemukan) {
      return gttResponTugasLuar_(
        false,
        'BELUM_ABSEN',
        'Anda belum melakukan absensi masuk hari ini.'
      );
    }

    var breakInfo = apiAmbilBreakInfoPribadi_(
      ss,
      user.pin,
      now,
      tz,
      Utilities.formatDate(info.waktuServerAsli, tz, 'yyyy-MM-dd')
    );

    if (
      breakInfo.break1.status === 'SEDANG BREAK' ||
      breakInfo.break2.status === 'SEDANG BREAK'
    ) {
      return gttResponTugasLuar_(
        false,
        'SEDANG_BREAK',
        'Selesaikan Break terlebih dahulu sebelum memulai tugas luar.'
      );
    }

    // Rev (3.6): kunci berdasarkan PIN saja, tanpa bergantung format tanggal.
    var current = gttAmbilTugasLuarAktif_(user.pin);
    if (current.aktif) {
      return gttResponTugasLuar_(
        false,
        'TUGAS_LUAR_SUDAH_AKTIF',
        'Anda masih menjalankan Tugas Luar sejak pukul ' +
          Utilities.formatDate(current.mulai, tz, 'HH:mm') +
          '. Selesaikan Tugas Luar terlebih dahulu sebelum membuat tugas baru.',
        {
          jamKeluar: Utilities.formatDate(current.mulai, tz, 'HH:mm:ss')
        }
      );
    }

    var validasiLokasi =
      gttValidasiLokasiOutletTugasLuar_(user.outlet, lokasi);

    if (!validasiLokasi.success) {
      return validasiLokasi;
    }

    var sheet = gttPastikanSheetTugasLuar_();
    var headers =
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    var map = gttHeaderMapAktivitas_(headers);

    var id =
      'AKT-' +
      Utilities.formatDate(now, tz, 'yyyyMMddHHmmss') +
      '-' +
      user.pin +
      '-' +
      Utilities.getUuid().slice(0, 6).toUpperCase();

    var row = new Array(headers.length).fill('');
    gttSetRowAktivitas_(row, map, 'ID', id);
    gttSetRowAktivitas_(row, map, 'TANGGAL', tanggalTampilan);
    gttSetRowAktivitas_(row, map, 'PIN', user.pin);
    gttSetRowAktivitas_(row, map, 'NAMA SA', user.namaSA);
    gttSetRowAktivitas_(row, map, 'OUTLET', user.outlet);
    gttSetRowAktivitas_(row, map, 'JABATAN', user.jabatan);
    gttSetRowAktivitas_(row, map, 'TUJUAN', tujuan);
    gttSetRowAktivitas_(row, map, 'KEPERLUAN', keperluan);
    gttSetRowAktivitas_(row, map, 'DISETUJUI PIN', pemberiIzinPin);
    gttSetRowAktivitas_(row, map, 'DISETUJUI OLEH', pemberiIzinNama);
    gttSetRowAktivitas_(row, map, 'DISETUJUI JABATAN', pemberiIzinJabatan);
    gttSetRowAktivitas_(row, map, 'JAM KELUAR', now);
    gttSetRowAktivitas_(row, map, 'STATUS', 'AKTIF');
    gttSetRowAktivitas_(row, map, 'JARAK KELUAR (M)', validasiLokasi.data.jarakMeter);
    gttSetRowAktivitas_(row, map, 'AKURASI GPS KELUAR (M)', validasiLokasi.data.akurasiMeter);

    sheet.appendRow(row);

    return gttResponTugasLuar_(
      true,
      'TUGAS_LUAR_DIMULAI',
      'Tugas luar berhasil dimulai.',
      {
        id: id,
        jamKeluar: Utilities.formatDate(now, tz, 'HH:mm:ss'),
        jarakKeluarMeter: validasiLokasi.data.jarakMeter
      }
    );
  } finally {
    lock.releaseLock();
  }
}

function gttSelesaiTugasLuar_(pinInput, lokasiInput) {
  var login = validasiLoginPin(pinInput);
  if (!login.success) return login;

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    return gttResponTugasLuar_(
      false,
      'TRANSAKSI_SEDANG_DIPROSES',
      'Transaksi Tugas Luar sedang diproses. Silakan tunggu beberapa detik.'
    );
  }

  try {
    var user = apiNormalisasiPengguna_(login.data, pinInput);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var info = gttInfoWaktu_(user.pin);
    var now = info.sekarang;

    var active = gttAmbilTugasLuarAktif_(user.pin);

    if (!active.aktif) {
      return gttResponTugasLuar_(
        false,
        'TUGAS_LUAR_TIDAK_AKTIF',
        'Tidak ada tugas luar aktif yang dapat diselesaikan.'
      );
    }

    // Acuan kembali selalu outlet asal yang tersimpan pada transaksi aktif.
    var outletAsal = String(active.outlet || user.outlet || '').trim();
    var validasiLokasi =
      gttValidasiLokasiOutletTugasLuar_(outletAsal, lokasiInput || {});

    if (!validasiLokasi.success) {
      return gttResponTugasLuar_(
        false,
        validasiLokasi.code || 'DI_LUAR_RADIUS_OUTLET',
        validasiLokasi.message ||
          'Anda belum berada di area outlet asal. Silakan kembali ke outlet untuk menyelesaikan Tugas Luar.',
        validasiLokasi.data || {}
      );
    }

    var sheet = gttPastikanSheetTugasLuar_();
    var matches = active.matches && active.matches.length
      ? active.matches
      : [{ row: active.row, mulai: active.mulai }];

    var durationLatest = 0;

    matches.forEach(function(item, index) {
      var mulai = item.mulai instanceof Date
        ? item.mulai
        : new Date(item.mulai);

      var duration = Math.max(
        0,
        Math.ceil((now.getTime() - mulai.getTime()) / 60000)
      );

      if (index === 0) durationLatest = duration;

      gttSetCellAktivitas_(sheet, item.row, active.map, 'JAM KEMBALI', now);
      gttSetCellAktivitas_(sheet, item.row, active.map, 'DURASI MENIT', duration);
      gttSetCellAktivitas_(
        sheet,
        item.row,
        active.map,
        'STATUS',
        index === 0 ? 'SELESAI' : 'DUPLIKAT DITUTUP'
      );
      gttSetCellAktivitas_(
        sheet,
        item.row,
        active.map,
        'JARAK KEMBALI (M)',
        validasiLokasi.data.jarakMeter
      );
      gttSetCellAktivitas_(
        sheet,
        item.row,
        active.map,
        'AKURASI GPS KEMBALI (M)',
        validasiLokasi.data.akurasiMeter
      );

      if (index > 0) {
        gttSetCellAktivitas_(
          sheet,
          item.row,
          active.map,
          'CATATAN',
          'DUPLIKAT TRANSAKSI DITUTUP OTOMATIS OLEH REV (3.6)'
        );
      }
    });

    return gttResponTugasLuar_(
      true,
      'TUGAS_LUAR_SELESAI',
      'Tugas luar berhasil diselesaikan.',
      {
        jamKembali: Utilities.formatDate(now, tz, 'HH:mm:ss'),
        durasiMenit: durationLatest,
        jarakKembaliMeter: validasiLokasi.data.jarakMeter,
        duplikatDitutup: Math.max(0, matches.length - 1)
      }
    );
  } finally {
    lock.releaseLock();
  }
}

function gttAmbilDataTugasLuar_(user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var info = gttInfoWaktu_(user.pin);
  var active = gttAmbilTugasLuarAktif_(user.pin);

  var tujuan = gttAmbilDaftarTujuan_();
  var approvers = gttAmbilDaftarPemberiIzin_(user.outlet);
  var minimum = gttAmbilDurasiMinimalTugasLuar_();

  return {
    aktif: Boolean(active.aktif),
    tujuan: active.aktif ? active.tujuan : '',
    keperluan: active.aktif ? active.keperluan : '',
    disetujuiOleh: active.aktif ? active.disetujuiOleh : '',
    jamKeluar: active.aktif
      ? Utilities.formatDate(active.mulai, tz, 'HH:mm:ss')
      : '',
    mulaiTimestamp: active.aktif ? active.mulai.getTime() : 0,
    durasiMinimalMenit: minimum,
    daftarTujuan: tujuan,
    daftarPemberiIzin: approvers
  };
}



function gttCariAbsensiMasukTugasLuar_(
  spreadsheet,
  pinInput,
  waktuOperasional,
  zonaWaktu
) {
  var sheet =
    spreadsheet.getSheetByName(
      'ABSENSI_HARIAN'
    );

  if (!sheet || sheet.getLastRow() <= 1) {
    return { ditemukan: false };
  }

  var range =
    sheet.getDataRange();

  var rawValues =
    range.getValues();

  var displayValues =
    range.getDisplayValues();

  var map =
    gttHeaderMapAktivitas_(
      displayValues[0]
    );

  var pinCol =
    gttCariKolomAktivitas_(
      map,
      ['PIN']
    );

  var tanggalCol =
    gttCariKolomAktivitas_(
      map,
      [
        'TANGGAL',
        'TANGGAL ABSENSI'
      ]
    );

  var jamMasukCol =
    gttCariKolomAktivitas_(
      map,
      [
        'JAM MASUK',
        'ABSEN MASUK'
      ]
    );

  var statusCol =
    gttCariKolomAktivitas_(
      map,
      ['STATUS KEHADIRAN']
    );

  if (
    pinCol === undefined ||
    tanggalCol === undefined ||
    jamMasukCol === undefined
  ) {
    return { ditemukan: false };
  }

  var tanggalTarget =
    Utilities.formatDate(
      waktuOperasional,
      zonaWaktu,
      'yyyy-MM-dd'
    );

  var pinTarget =
    String(pinInput || '').trim();

  for (
    var i = rawValues.length - 1;
    i >= 1;
    i--
  ) {
    var rawRow =
      rawValues[i];

    var displayRow =
      displayValues[i];

    var pinData =
      String(
        displayRow[pinCol] ||
        rawRow[pinCol] ||
        ''
      ).trim();

    var tanggalData =
      gttNormalisasiTanggalTugasLuar_(
        rawRow[tanggalCol],
        displayRow[tanggalCol],
        zonaWaktu
      );

    var jamMasukDisplay =
      String(
        displayRow[jamMasukCol] || ''
      ).trim();

    var jamMasukRaw =
      rawRow[jamMasukCol];

    var jamMasukTerisi =
      (
        jamMasukRaw !== '' &&
        jamMasukRaw !== null &&
        jamMasukRaw !== undefined
      ) ||
      (
        jamMasukDisplay !== '' &&
        jamMasukDisplay !== '--:--:--' &&
        jamMasukDisplay !== '--:--'
      );

    if (
      pinData === pinTarget &&
      tanggalData === tanggalTarget &&
      jamMasukTerisi
    ) {
      return {
        ditemukan: true,
        nomorBaris: i + 1,
        jamMasuk:
          jamMasukDisplay ||
          jamMasukRaw,
        statusKehadiran:
          statusCol === undefined
            ? ''
            : String(
                displayRow[statusCol] ||
                rawRow[statusCol] ||
                ''
              ).trim()
      };
    }
  }

  return { ditemukan: false };
}

function gttNormalisasiTanggalTugasLuar_(
  rawValue,
  displayValue,
  zonaWaktu
) {
  if (rawValue instanceof Date) {
    return Utilities.formatDate(
      rawValue,
      zonaWaktu,
      'yyyy-MM-dd'
    );
  }

  var text =
    String(
      displayValue ||
      rawValue ||
      ''
    ).trim();

  if (!text) return '';

  var cleanText =
    text.split(' ')[0];

  var parts =
    cleanText.split(/[\/\-\.]/);

  if (parts.length !== 3) {
    var parsed =
      new Date(text);

    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(
        parsed,
        zonaWaktu,
        'yyyy-MM-dd'
      );
    }

    return cleanText;
  }

  // yyyy-MM-dd
  if (parts[0].length === 4) {
    return [
      parts[0],
      String(parts[1]).padStart(2, '0'),
      String(parts[2]).padStart(2, '0')
    ].join('-');
  }

  // dd/MM/yyyy
  return [
    parts[2],
    String(parts[1]).padStart(2, '0'),
    String(parts[0]).padStart(2, '0')
  ].join('-');
}

function gttAmbilTugasLuarAktif_(pin) {
  var sheet = gttPastikanSheetTugasLuar_();
  var values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return { aktif: false, matches: [] };
  }

  var map = gttHeaderMapAktivitas_(values[0]);
  var matches = [];
  var targetPin = String(pin || '').trim();

  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i];

    if (
      String(row[map['PIN']] || '').trim() === targetPin &&
      String(row[map['STATUS']] || '').trim().toUpperCase() === 'AKTIF'
    ) {
      var mulai =
        row[map['JAM KELUAR']] instanceof Date
          ? row[map['JAM KELUAR']]
          : new Date(row[map['JAM KELUAR']]);

      if (isNaN(mulai.getTime())) continue;

      matches.push({
        row: i + 1,
        mulai: mulai,
        outlet: String(row[map['OUTLET']] || '').trim(),
        tujuan: String(row[map['TUJUAN']] || ''),
        keperluan: String(row[map['KEPERLUAN']] || ''),
        disetujuiOleh: String(row[map['DISETUJUI OLEH']] || '')
      });
    }
  }

  if (!matches.length) {
    return { aktif: false, matches: [] };
  }

  var latest = matches[0];

  return {
    aktif: true,
    row: latest.row,
    rows: matches.map(function(item) { return item.row; }),
    matches: matches,
    map: map,
    outlet: latest.outlet,
    tujuan: latest.tujuan,
    keperluan: latest.keperluan,
    disetujuiOleh: latest.disetujuiOleh,
    mulai: latest.mulai
  };
}

function gttAmbilDaftarTujuan_() {
  var result = [];
  var seen = {};

  gttTambahTujuanUnik_(
    result,
    seen,
    gttAmbilTujuanDariMasterOutlet_()
  );

  // Fallback utama: semua kode OUTLET unik yang sudah
  // terbukti terbaca pada MASTER_SA.
  gttTambahTujuanUnik_(
    result,
    seen,
    gttAmbilTujuanDariMasterSA_()
  );

  return result;
}

function gttTambahTujuanUnik_(
  target,
  seen,
  source
) {
  (source || []).forEach(function(item) {
    var kode =
      String(item.kode || '').trim();

    var nama =
      String(item.nama || kode).trim();

    if (!kode) return;

    var key =
      kode.toUpperCase();

    if (seen[key]) return;

    seen[key] = true;

    target.push({
      kode: kode,
      nama: nama || kode
    });
  });
}

function gttAmbilTujuanDariMasterOutlet_() {
  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  // MASTER_LOKASI adalah sumber utama tujuan GTT.
  var candidateNames = [
    'MASTER_LOKASI',
    'MASTER LOKASI',
    'MASTERLOKASI',
    'MASTER_OUTLET',
    'MASTER OUTLET',
    'MASTEROUTLET',
    'OUTLET'
  ];

  var sheet = null;

  for (
    var i = 0;
    i < candidateNames.length;
    i++
  ) {
    sheet =
      ss.getSheetByName(
        candidateNames[i]
      );

    if (sheet) break;
  }

  if (!sheet) {
    var allSheets =
      ss.getSheets();

    for (
      var s = 0;
      s < allSheets.length;
      s++
    ) {
      var normalized =
        String(
          allSheets[s].getName() || ''
        )
          .trim()
          .toUpperCase()
          .replace(/[\s_-]+/g, '');

      if (
        normalized.indexOf('MASTER') >= 0 &&
        (
          normalized.indexOf('LOKASI') >= 0 ||
          normalized.indexOf('OUTLET') >= 0
        )
      ) {
        sheet = allSheets[s];
        break;
      }
    }
  }

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  var values =
    sheet.getDataRange().getDisplayValues();

  var headerRowIndex =
    gttCariBarisHeaderOutlet_(
      values
    );

  if (headerRowIndex < 0) {
    return [];
  }

  var map =
    gttHeaderMapAktivitas_(
      values[headerRowIndex]
    );

  var kodeCol =
    gttCariKolomAktivitas_(
      map,
      [
        'KODE LOKASI',
        'KODE OUTLET',
        'KODE TOKO',
        'KODE CABANG',
        'ID LOKASI',
        'ID OUTLET',
        'KODE',
        'OUTLET'
      ]
    );

  var namaCol =
    gttCariKolomAktivitas_(
      map,
      [
        'NAMA LOKASI',
        'NAMA OUTLET',
        'NAMA TOKO',
        'NAMA CABANG',
        'LOKASI',
        'OUTLET',
        'TOKO',
        'CABANG',
        'NAMA'
      ]
    );

  var statusCol =
    gttCariKolomAktivitas_(
      map,
      [
        'STATUS',
        'STATUS LOKASI',
        'STATUS OUTLET',
        'STATUS TOKO',
        'STATUS AKUN',
        'AKTIF'
      ]
    );

  // Bila hanya ada satu kolom nama lokasi/outlet,
  // gunakan kolom itu sebagai kode sekaligus nama.
  if (
    kodeCol === undefined &&
    namaCol !== undefined
  ) {
    kodeCol = namaCol;
  }

  if (
    namaCol === undefined &&
    kodeCol !== undefined
  ) {
    namaCol = kodeCol;
  }

  if (kodeCol === undefined) {
    return [];
  }

  return values
    .slice(headerRowIndex + 1)
    .filter(function(row) {
      var kode =
        String(
          row[kodeCol] || ''
        ).trim();

      if (!kode) return false;

      if (statusCol === undefined) {
        return true;
      }

      var status =
        String(
          row[statusCol] || ''
        )
          .trim()
          .toUpperCase();

      return (
        status !== 'NONAKTIF' &&
        status !== 'TIDAK AKTIF' &&
        status !== 'FALSE' &&
        status !== 'NO' &&
        status !== '0'
      );
    })
    .map(function(row) {
      var kode =
        String(
          row[kodeCol] || ''
        ).trim();

      var nama =
        String(
          row[namaCol] || kode
        ).trim();

      return {
        kode: kode,
        nama:
          nama &&
          nama.toUpperCase() !==
            kode.toUpperCase()
            ? kode + ' - ' + nama
            : kode
      };
    });
}

function gttCariBarisHeaderOutlet_(
  values
) {
  var maxRows =
    Math.min(
      values.length,
      10
    );

  for (
    var r = 0;
    r < maxRows;
    r++
  ) {
    var normalizedRow =
      values[r].map(function(value) {
        return String(value || '')
          .trim()
          .toUpperCase();
      });

    var joined =
      normalizedRow.join(' | ');

    if (
      joined.indexOf('LOKASI') >= 0 ||
      joined.indexOf('OUTLET') >= 0 ||
      joined.indexOf('CABANG') >= 0 ||
      joined.indexOf('KODE TOKO') >= 0 ||
      joined.indexOf('NAMA TOKO') >= 0
    ) {
      return r;
    }
  }

  return 0;
}

function gttAmbilTujuanDariMasterSA_() {
  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var sheet =
    ss.getSheetByName(
      'MASTER_SA'
    );

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  var values =
    sheet.getDataRange().getValues();

  var map =
    gttHeaderMapAktivitas_(
      values[0]
    );

  var outletCol =
    gttCariKolomAktivitas_(
      map,
      [
        'OUTLET',
        'KODE OUTLET',
        'TOKO'
      ]
    );

  var statusCol =
    gttCariKolomAktivitas_(
      map,
      [
        'STATUS AKUN',
        'STATUS'
      ]
    );

  if (outletCol === undefined) {
    return [];
  }

  var seen = {};

  return values.slice(1)
    .filter(function(row) {
      var outlet =
        String(row[outletCol] || '')
          .trim();

      if (!outlet) return false;

      if (statusCol !== undefined) {
        var status =
          String(row[statusCol] || '')
            .trim()
            .toUpperCase();

        if (
          status === 'NONAKTIF' ||
          status === 'TIDAK AKTIF'
        ) {
          return false;
        }
      }

      var key =
        outlet.toUpperCase();

      if (seen[key]) return false;

      seen[key] = true;
      return true;
    })
    .map(function(row) {
      var outlet =
        String(row[outletCol] || '')
          .trim();

      return {
        kode: outlet,
        nama: outlet
      };
    });
}

function gttCariKolomAktivitas_(map, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    if (map[candidates[i]] !== undefined) {
      return map[candidates[i]];
    }
  }

  var keys = Object.keys(map);

  for (var k = 0; k < keys.length; k++) {
    var normalizedKey =
      String(keys[k] || '')
        .trim()
        .toUpperCase()
        .replace(/[\s_-]+/g, '');

    for (var j = 0; j < candidates.length; j++) {
      var normalizedCandidate =
        String(candidates[j] || '')
          .trim()
          .toUpperCase()
          .replace(/[\s_-]+/g, '');

      if (
        normalizedKey === normalizedCandidate ||
        normalizedKey.indexOf(normalizedCandidate) >= 0 ||
        normalizedCandidate.indexOf(normalizedKey) >= 0
      ) {
        return map[keys[k]];
      }
    }
  }

  return undefined;
}

function gttAmbilDaftarPemberiIzin_(outlet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MASTER_SA');
  if (!sheet) return [];

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var map = gttHeaderMapAktivitas_(values[0]);

  return values.slice(1)
    .map(function(row) {
      var jabatan = String(
        row[
          map['JABATAN/POSISI'] !== undefined
            ? map['JABATAN/POSISI']
            : map['JABATAN']
        ] || ''
      ).trim().toUpperCase();

      var rowOutlet = String(row[map['OUTLET']] || '').trim();
      var aktif = map['STATUS AKUN'] === undefined ||
        String(row[map['STATUS AKUN']] || '').trim().toUpperCase() !== 'NONAKTIF';

      return {
        valid:
          aktif &&
          rowOutlet === String(outlet || '').trim() &&
          (
            jabatan === 'SL' ||
            jabatan === 'SPV' ||
            jabatan.indexOf('MANAGER') >= 0
          ),
        id: String(row[map['ID SA']] || '').trim(),
        pin: String(row[map['PIN']] || '').trim(),
        nama: String(row[map['NAMA SA']] || '').trim(),
        jabatan: jabatan
      };
    })
    .filter(function(item) { return item.valid; })
    .map(function(item) {
      delete item.valid;
      return item;
    });
}

function gttAmbilDurasiMinimalTugasLuar_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MASTER_SETTING');
  if (!sheet) return 15;

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (
      String(values[i][0] || '').trim().toUpperCase() ===
      'TUGAS LUAR - DURASI MINIMAL (MENIT)'
    ) {
      return Number(values[i][1] || 15);
    }
  }

  return 15;
}

function gttHeaderMapAktivitas_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    map[String(header || '').trim().toUpperCase()] = index;
  });
  return map;
}

function gttFormatTanggalAktivitas_(value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz, 'dd/MM/yyyy');
  }

  var text = String(value || '').trim();
  var parts = text.split(/[\/\-]/);
  if (parts.length === 3 && parts[0].length === 4) {
    return parts[2] + '/' + parts[1] + '/' + parts[0];
  }
  return text;
}

function gttResponTugasLuar_(success, code, message, data) {
  return {
    success: Boolean(success),
    code: String(code || ''),
    message: String(message || ''),
    data: data || {}
  };
}


/**
 * Rev (3.6) — helper penulisan kolom dinamis.
 */
function gttSetRowAktivitas_(row, map, header, value) {
  if (map[header] !== undefined) {
    row[map[header]] = value;
  }
}

function gttSetCellAktivitas_(sheet, rowNumber, map, header, value) {
  if (map[header] !== undefined) {
    sheet.getRange(rowNumber, map[header] + 1).setValue(value);
  }
}

/**
 * Membaca titik outlet dari MASTER_LOKASI dan memvalidasi posisi HP.
 */
function gttValidasiLokasiOutletTugasLuar_(outletInput, lokasiInput) {
  var lokasi = lokasiInput || {};
  var lat = Number(lokasi.latitude);
  var lng = Number(lokasi.longitude);
  var accuracy = Number(lokasi.accuracy);

  if (!isFinite(lat) || !isFinite(lng)) {
    return gttResponTugasLuar_(
      false,
      'LOKASI_TIDAK_TERSEDIA',
      'Lokasi perangkat belum tersedia. Aktifkan GPS dan izinkan akses lokasi, lalu coba kembali.'
    );
  }

  if (!isFinite(accuracy) || accuracy <= 0) {
    accuracy = 9999;
  }

  var outlet = gttCariLokasiOutlet_(outletInput);

  if (!outlet.ditemukan) {
    return gttResponTugasLuar_(
      false,
      'OUTLET_TIDAK_DITEMUKAN',
      'Koordinat outlet belum ditemukan di MASTER_LOKASI.'
    );
  }

  var jarak = Math.round(
    gttHitungJarakMeter_(lat, lng, outlet.latitude, outlet.longitude)
  );

  var batasAkurasi = Math.max(100, Number(outlet.radius) * 2);

  if (accuracy > batasAkurasi) {
    return gttResponTugasLuar_(
      false,
      'AKURASI_GPS_RENDAH',
      'Akurasi GPS masih ±' + Math.round(accuracy) +
        ' meter. Tunggu beberapa detik di area terbuka, lalu coba kembali.',
      {
        akurasiMeter: Math.round(accuracy),
        batasAkurasiMeter: batasAkurasi
      }
    );
  }

  if (jarak > outlet.radius) {
    return gttResponTugasLuar_(
      false,
      'DI_LUAR_RADIUS_OUTLET',
      'Anda belum berada di area ' + outlet.nama +
        '. Jarak saat ini sekitar ' + jarak +
        ' meter, sedangkan radius yang diizinkan ' +
        outlet.radius + ' meter.',
      {
        outlet: outlet.nama,
        jarakMeter: jarak,
        radiusMeter: outlet.radius,
        akurasiMeter: Math.round(accuracy)
      }
    );
  }

  return gttResponTugasLuar_(
    true,
    'LOKASI_VALID',
    'Lokasi berada dalam radius outlet.',
    {
      outlet: outlet.nama,
      jarakMeter: jarak,
      radiusMeter: outlet.radius,
      akurasiMeter: Math.round(accuracy)
    }
  );
}

function gttCariLokasiOutlet_(outletInput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet =
    ss.getSheetByName('MASTER_LOKASI') ||
    ss.getSheetByName('MASTER LOKASI');

  if (!sheet || sheet.getLastRow() <= 1) {
    return { ditemukan: false };
  }

  var values = sheet.getDataRange().getDisplayValues();
  var map = gttHeaderMapAktivitas_(values[0]);

  var kodeCol = gttCariKolomAktivitas_(map, [
    'KODE', 'KODE OUTLET', 'KODE LOKASI', 'OUTLET'
  ]);
  var namaCol = gttCariKolomAktivitas_(map, [
    'NAMA OUTLET', 'NAMA LOKASI', 'LOKASI'
  ]);
  var latCol = gttCariKolomAktivitas_(map, ['LATITUDE', 'LAT']);
  var lngCol = gttCariKolomAktivitas_(map, ['LONGITUDE', 'LNG', 'LONG']);
  var radiusCol = gttCariKolomAktivitas_(map, [
    'RADIUS ABSEN (M)', 'RADIUS (M)', 'RADIUS'
  ]);
  var aktifCol = gttCariKolomAktivitas_(map, ['AKTIF', 'STATUS']);

  if (latCol === undefined || lngCol === undefined) {
    return { ditemukan: false };
  }

  var target = String(outletInput || '').trim().toUpperCase();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var kode = kodeCol === undefined ? '' : String(row[kodeCol] || '').trim();
    var nama = namaCol === undefined ? kode : String(row[namaCol] || '').trim();
    var aktif = aktifCol === undefined
      ? true
      : !['TIDAK', 'FALSE', '0', 'NONAKTIF', 'TIDAK AKTIF'].includes(
          String(row[aktifCol] || '').trim().toUpperCase()
        );

    if (!aktif) continue;

    if (
      String(kode).toUpperCase() === target ||
      String(nama).toUpperCase() === target
    ) {
      var latitude = Number(String(row[latCol] || '').replace(',', '.'));
      var longitude = Number(String(row[lngCol] || '').replace(',', '.'));
      var radius = radiusCol === undefined
        ? 50
        : Number(String(row[radiusCol] || '').replace(',', '.'));

      if (!isFinite(latitude) || !isFinite(longitude)) {
        return { ditemukan: false };
      }

      if (!isFinite(radius) || radius <= 0) radius = 50;

      return {
        ditemukan: true,
        kode: kode,
        nama: nama || kode || target,
        latitude: latitude,
        longitude: longitude,
        radius: Math.round(radius)
      };
    }
  }

  return { ditemukan: false };
}

function gttHitungJarakMeter_(lat1, lng1, lat2, lng2) {
  var earthRadius = 6371000;
  var toRad = function(value) {
    return value * Math.PI / 180;
  };

  var dLat = toRad(lat2 - lat1);
  var dLng = toRad(lng2 - lng1);
  var a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) *
      Math.cos(toRad(lat2)) *
      Math.sin(dLng / 2) *
      Math.sin(dLng / 2);

  return earthRadius * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
