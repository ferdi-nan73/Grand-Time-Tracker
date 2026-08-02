/**
 * Memvalidasi login berdasarkan PIN pada sheet MASTER_SA.
 *
 * @param {string|number} pinInput PIN yang dimasukkan pengguna.
 * @return {Object} Hasil validasi login.
 */
function validasiLoginPin(pinInput) {
  const NAMA_SHEET = 'MASTER_SA';
  const BARIS_HEADER = 1;

  try {
    const pin = String(pinInput || '').trim();

    // PIN GTT wajib terdiri dari tepat 4 digit.
    if (!/^\d{4}$/.test(pin)) {
      return {
        success: false,
        code: 'FORMAT_PIN_TIDAK_VALID',
        message: 'PIN harus terdiri dari 4 angka.'
      };
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(NAMA_SHEET);

    if (!sheet) {
      return {
        success: false,
        code: 'SHEET_TIDAK_DITEMUKAN',
        message: 'Sheet MASTER_SA tidak ditemukan.'
      };
    }

    const barisTerakhir = sheet.getLastRow();
    const kolomTerakhir = sheet.getLastColumn();

    if (barisTerakhir <= BARIS_HEADER) {
      return {
        success: false,
        code: 'MASTER_SA_KOSONG',
        message: 'Data MASTER_SA masih kosong.'
      };
    }

    const seluruhData = sheet
      .getRange(
        BARIS_HEADER,
        1,
        barisTerakhir,
        kolomTerakhir
      )
      .getDisplayValues();

    const header = seluruhData[0].map(normalisasiHeader_);
    const dataSA = seluruhData.slice(1);

    const indexNama = header.indexOf('NAMA SA');
    const indexPin = header.indexOf('PIN');
    const indexOutlet = header.indexOf('OUTLET');
    const indexStatus = header.indexOf('STATUS');
    const indexSaId = header.indexOf('SA_ID');
    const indexJabatan = header.indexOf('JABATAN');
    const indexRole = header.indexOf('ROLE');

    const kolomWajib = [
      { nama: 'NAMA SA', index: indexNama },
      { nama: 'PIN', index: indexPin },
      { nama: 'OUTLET', index: indexOutlet },
      { nama: 'STATUS', index: indexStatus },
      { nama: 'SA_ID', index: indexSaId },
      { nama: 'JABATAN', index: indexJabatan },
      { nama: 'ROLE', index: indexRole }
    ];

    const kolomTidakDitemukan = kolomWajib
      .filter(item => item.index === -1)
      .map(item => item.nama);

    if (kolomTidakDitemukan.length > 0) {
      return {
        success: false,
        code: 'KOLOM_TIDAK_LENGKAP',
        message:
          'Kolom berikut tidak ditemukan di MASTER_SA: ' +
          kolomTidakDitemukan.join(', ')
      };
    }

    const hasilPencarian = dataSA
      .map((baris, index) => {
        return {
          nomorBaris: index + BARIS_HEADER + 1,
          namaSA: String(baris[indexNama] || '').trim(),
          pin: String(baris[indexPin] || '').trim(),
          outlet: String(baris[indexOutlet] || '').trim(),
          status: String(baris[indexStatus] || '')
            .trim()
            .toUpperCase(),
          saId: String(baris[indexSaId] || '')
            .trim()
            .toUpperCase(),
          jabatan: String(baris[indexJabatan] || '').trim(),
          role: String(baris[indexRole] || '')
            .trim()
            .toUpperCase()
        };
      })
      .filter(sa => sa.pin === pin);

    if (hasilPencarian.length === 0) {
      return {
        success: false,
        code: 'PIN_TIDAK_DITEMUKAN',
        message: 'PIN tidak terdaftar.'
      };
    }

    const penggunaAktif = hasilPencarian.filter(
      sa => sa.status === 'AKTIF'
    );

    if (penggunaAktif.length === 0) {
      return {
        success: false,
        code: 'SA_TIDAK_AKTIF',
        message: 'PIN terdaftar, tetapi status pengguna tidak aktif.'
      };
    }

    if (penggunaAktif.length > 1) {
      return {
        success: false,
        code: 'PIN_DUPLIKAT',
        message:
          'PIN digunakan oleh lebih dari satu pengguna aktif. ' +
          'Hubungi administrator.'
      };
    }

    const pengguna = penggunaAktif[0];

    const validasiData = [
      {
        nilai: pengguna.saId,
        code: 'SA_ID_KOSONG',
        message: 'SA_ID pengguna pada MASTER_SA masih kosong.'
      },
      {
        nilai: pengguna.namaSA,
        code: 'NAMA_SA_KOSONG',
        message: 'Nama pengguna pada MASTER_SA masih kosong.'
      },
      {
        nilai: pengguna.outlet,
        code: 'OUTLET_KOSONG',
        message: 'Outlet pengguna pada MASTER_SA masih kosong.'
      },
      {
        nilai: pengguna.jabatan,
        code: 'JABATAN_KOSONG',
        message: 'Jabatan pengguna pada MASTER_SA masih kosong.'
      },
      {
        nilai: pengguna.role,
        code: 'ROLE_KOSONG',
        message: 'Role pengguna pada MASTER_SA masih kosong.'
      }
    ];

    const dataTidakLengkap = validasiData.find(
      item => !item.nilai
    );

    if (dataTidakLengkap) {
      return {
        success: false,
        code: dataTidakLengkap.code,
        message: dataTidakLengkap.message
      };
    }

    if (!['SA', 'ADMIN'].includes(pengguna.role)) {
      return {
        success: false,
        code: 'ROLE_TIDAK_VALID',
        message:
          'Role pengguna harus bernilai SA atau ADMIN.'
      };
    }

    return {
      success: true,
      code: 'LOGIN_BERHASIL',
      message: 'Login berhasil.',
      data: {
        saId: pengguna.saId,
        namaSA: pengguna.namaSA,
        outlet: pengguna.outlet,
        status: pengguna.status,
        jabatan: pengguna.jabatan,
        role: pengguna.role
      }
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      code: 'SYSTEM_ERROR',
      message: 'Terjadi kesalahan sistem: ' + error.message
    };
  }
}


function normalisasiHeader_(nilaiHeader) {
  return String(nilaiHeader || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}
function ujiValidasiLoginPin() {
  const PIN_UJI = '9616';

  const hasil = validasiLoginPin(PIN_UJI);

  Logger.log(JSON.stringify(hasil, null, 2));
}
/**
 * Entry point Web App Grand Time Tracker.
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Grand Time Tracker')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1, maximum-scale=1'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}

/**
 * Memuat file HTML parsial seperti Styles.html dan Scripts.html.
 *
 * @param {string} filename Nama file HTML tanpa ekstensi.
 * @return {string} Isi file HTML.
 */
function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}
