/**
 * =====================================================
 * GRAND TIME TRACKER — GTT
 * Module : GTT-08 Break Time Engine
 * Version: 2.1.1
 * Status : Sprint 2 — Break Engine Revision
 * =====================================================
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
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (![1, 2].includes(nomorBreak)) {
      throw new Error('Nomor break tidak valid.');
    }

    const konteks = ambilKonteksBreak_(pinInput);

    if (!konteks.success) {
      return konteks;
    }

    const {
      pengguna,
      spreadsheet,
      sheetLog,
      headerMap,
      zonaWaktu,
      sekarang,
      tanggalHariIni,
      pengaturan
    } = konteks.data;

    const hasilAbsensi = cariAbsensiAktifHariIni_(
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

    const dataLog = cariLogIstirahatHariIni_(
      sheetLog,
      headerMap,
      pengguna.pin,
      tanggalHariIni,
      zonaWaktu
    );

    if (nomorBreak === 1) {
      const validasiBreak1 = validasiMulaiBreak1_(
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
      const validasiBreak2 = validasiMulaiBreak2_(
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

    const jumlahSedangBreak = hitungSedangBreakOutlet_(
      sheetLog,
      headerMap,
      pengguna.outlet,
      tanggalHariIni,
      zonaWaktu
    );

    const maxIstirahat = pengaturan.maxBreakBersamaan;

    if (jumlahSedangBreak >= maxIstirahat) {
      return {
        success: false,
        code: 'KAPASITAS_BREAK_PENUH',
        message:
          'Kapasitas istirahat Outlet ' +
          pengguna.outlet +
          ' sedang penuh. Maksimal ' +
          maxIstirahat +
          ' orang dapat beristirahat bersamaan.',
        data: {
          outlet: pengguna.outlet,
          sedangBreak: jumlahSedangBreak,
          maxIstirahat: maxIstirahat
        }
      };
    }

    let nomorBaris;

    if (!dataLog.ditemukan) {
      nomorBaris = buatBarisLogIstirahatBaru_(
        sheetLog,
        headerMap,
        pengguna,
        sekarang,
        zonaWaktu
      );
    } else {
      nomorBaris = dataLog.nomorBaris;
    }

    const namaKolomMulai =
      nomorBreak === 1 ? 'S1 MULAI' : 'S2 MULAI';

    sheetLog
      .getRange(
        nomorBaris,
        headerMap[namaKolomMulai] + 1
      )
      .setValue(sekarang)
      .setNumberFormat('HH:mm:ss');

    SpreadsheetApp.flush();

    const durasiBreak =
      nomorBreak === 1
        ? pengaturan.durasiBreak1
        : pengaturan.durasiBreak2;

    const estimasiSelesai = new Date(
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
