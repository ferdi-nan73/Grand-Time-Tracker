function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  var NAMA_SHEET = 'MASTER_SA';
  var KOLOM_NAMA = 1;
  var KOLOM_PIN = 2;
  var BARIS_HEADER = 1;

  var range = e.range;
  var sheet = range.getSheet();

  if (sheet.getName() !== NAMA_SHEET) {
    return;
  }

  var kolomAwal = range.getColumn();
  var kolomAkhir = range.getLastColumn();

  if (
    KOLOM_NAMA < kolomAwal ||
    KOLOM_NAMA > kolomAkhir
  ) {
    return;
  }

  var barisAwal = Math.max(
    range.getRow(),
    BARIS_HEADER + 1
  );

  var barisAkhir = range.getLastRow();

  if (barisAwal > barisAkhir) {
    return;
  }

  var jumlahBaris =
    barisAkhir - barisAwal + 1;

  var rangeNama = sheet.getRange(
    barisAwal,
    KOLOM_NAMA,
    jumlahBaris,
    1
  );

  var rangePin = sheet.getRange(
    barisAwal,
    KOLOM_PIN,
    jumlahBaris,
    1
  );

  var daftarNama =
    rangeNama.getDisplayValues();

  var daftarPin =
    rangePin.getDisplayValues();

  var pinTerpakai = ambilSemuaPin_(
    sheet,
    KOLOM_PIN,
    BARIS_HEADER
  );

  var hasilPin = [];

  for (
    var index = 0;
    index < daftarNama.length;
    index++
  ) {
    var namaSA = String(
      daftarNama[index][0] || ''
    ).trim();

    var pinSaatIni = String(
      daftarPin[index][0] || ''
    ).trim();

    if (!namaSA) {
      hasilPin.push(['']);
      continue;
    }

    if (pinSaatIni) {
      hasilPin.push([pinSaatIni]);
      continue;
    }

    var pinBaru =
      buatPinUnikDariDaftar_(
        pinTerpakai
      );

    pinTerpakai.add(pinBaru);
    hasilPin.push([pinBaru]);
  }

  rangePin.setNumberFormat('@');
  rangePin.setValues(hasilPin);
}


function ambilSemuaPin_(
  sheet,
  kolomPin,
  barisHeader
) {
  var pinTerpakai = new Set();
  var barisTerakhir = sheet.getLastRow();

  if (barisTerakhir <= barisHeader) {
    return pinTerpakai;
  }

  var daftarPin = sheet
    .getRange(
      barisHeader + 1,
      kolomPin,
      barisTerakhir - barisHeader,
      1
    )
    .getDisplayValues();

  for (
    var index = 0;
    index < daftarPin.length;
    index++
  ) {
    var nilaiPin = String(
      daftarPin[index][0] || ''
    ).trim();

    if (nilaiPin) {
      pinTerpakai.add(nilaiPin);
    }
  }

  return pinTerpakai;
}


function buatPinUnikDariDaftar_(
  pinTerpakai
) {
  var pinBaru = '';
  var percobaan = 0;

  do {
    pinBaru = String(
      Math.floor(
        1000 + Math.random() * 9000
      )
    );

    percobaan++;

    if (percobaan > 10000) {
      throw new Error(
        'Tidak dapat membuat PIN unik.'
      );
    }
  } while (
    pinTerpakai.has(pinBaru)
  );

  return pinBaru;
}
