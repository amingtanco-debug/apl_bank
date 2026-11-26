function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Aplikasi Mutasi Rekening Bank')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Konfigurasi Daftar Bank
const DAFTAR_BANK = {
  'B01': 'Bank Mandiri_01',
  'B02': 'Bank Mandiri_02',
  'B03': 'Bank BRI'
};

// Konfigurasi Kode Transaksi
const KODE_TRANSAKSI = {
  '01': { nama: 'Penerimaan dari Penjualan Tunai', tipe: 'MASUK' },
  '02': { nama: 'Penerimaan dari Penjualan QRIS', tipe: 'MASUK' },
  '03': { nama: 'Penerimaan dari Penjualan E-Wallet', tipe: 'MASUK' },
  '04': { nama: 'Mutasi Rekening (Transfer Antar Bank)', tipe: 'MUTASI' },
  '05': { nama: 'Pengeluaran Kas Operasional Usaha', tipe: 'KELUAR' },
  '06': { nama: 'Pengeluaran Kas Lainnya', tipe: 'KELUAR' }
};

function getDataDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transaksi');
  const data = sheet.getDataRange().getValues();
  
  let saldo = { 'B01': 0, 'B02': 0, 'B03': 0 };
  let riwayat = [];

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let kodeBank = String(row[2]).trim();
    let debit = Number(row[6]) || 0;
    let kredit = Number(row[7]) || 0;

    if (saldo.hasOwnProperty(kodeBank)) {
      saldo[kodeBank] += (debit - kredit);
    }

    let tglFormatted = '';
    let tglInput = '';
    if (row[1]) {
      let d = new Date(row[1]);
      tglFormatted = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      tglInput = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    riwayat.push({
      rowIndex: i + 1, // Menyimpan posisi baris asli di Google Sheets
      tanggal: tglFormatted,
      tglInput: tglInput,
      kodeBank: kodeBank,
      namaBank: row[3],
      kodeTrx: String(row[4]).padStart(2, '0'),
      ketTrx: row[5],
      debit: debit,
      kredit: kredit,
      nominal: debit > 0 ? debit : kredit,
      catatan: row[8] || ''
    });
  }

  return {
    saldo: saldo,
    bankInfo: DAFTAR_BANK,
    riwayat: riwayat.reverse().slice(0, 15)
  };
}

function simpanTransaksi(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transaksi');
  const timestamp = new Date();
  const tanggal = new Date(form.tanggal);
  const nominal = Number(form.nominal);
  const kodeTrx = form.kodeTransaksi;
  const infoTrx = KODE_TRANSAKSI[kodeTrx];

  if (!infoTrx) throw new Error('Kode transaksi tidak valid.');

  if (kodeTrx === '04') {
    if (form.kodeBank === form.bankTujuan) {
      throw new Error('Bank asal dan bank tujuan mutasi tidak boleh sama.');
    }
    sheet.appendRow([timestamp, tanggal, form.kodeBank, DAFTAR_BANK[form.kodeBank], kodeTrx, infoTrx.nama + ' (Ke ' + DAFTAR_BANK[form.bankTujuan] + ')', 0, nominal, form.catatan]);
    sheet.appendRow([timestamp, tanggal, form.bankTujuan, DAFTAR_BANK[form.bankTujuan], kodeTrx, infoTrx.nama + ' (Dari ' + DAFTAR_BANK[form.kodeBank] + ')', nominal, 0, form.catatan]);
  } else {
    let debit = infoTrx.tipe === 'MASUK' ? nominal : 0;
    let kredit = infoTrx.tipe === 'KELUAR' ? nominal : 0;
    sheet.appendRow([timestamp, tanggal, form.kodeBank, DAFTAR_BANK[form.kodeBank], kodeTrx, infoTrx.nama, debit, kredit, form.catatan]);
  }

  return getDataDashboard();
}

// Fungsi Baru: Koreksi / Update Baris Transaksi
function updateTransaksi(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transaksi');
  const rowIndex = Number(form.rowIndex);
  const tanggal = new Date(form.tanggal);
  const nominal = Number(form.nominal);
  const kodeTrx = form.kodeTransaksi;
  const infoTrx = KODE_TRANSAKSI[kodeTrx];

  let debit = 0;
  let kredit = 0;

  if (kodeTrx === '04') {
    // Pertahankan arah debit/kredit sebelumnya pada baris mutasi
    let oldDebit = Number(sheet.getRange(rowIndex, 7).getValue()) || 0;
    if (oldDebit > 0) {
      debit = nominal;
    } else {
      kredit = nominal;
    }
  } else {
    debit = infoTrx.tipe === 'MASUK' ? nominal : 0;
    kredit = infoTrx.tipe === 'KELUAR' ? nominal : 0;
  }

  let ketLama = sheet.getRange(rowIndex, 6).getValue();
  let ketBaru = kodeTrx === '04' ? ketLama : infoTrx.nama;

  // Update kolom B (Tanggal) sampai kolom I (Catatan)
  sheet.getRange(rowIndex, 2, 1, 8).setValues([[
    tanggal,
    form.kodeBank,
    DAFTAR_BANK[form.kodeBank],
    kodeTrx,
    ketBaru,
    debit,
    kredit,
    form.catatan
  ]]);

  return getDataDashboard();
}

// Fungsi Baru: Hapus Baris Transaksi yang Salah Input
function hapusTransaksi(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transaksi');
  sheet.deleteRow(Number(rowIndex));
  return getDataDashboard();
}
