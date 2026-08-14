/**
 * Google Apps Script
 * Kirim PDF dari Google Drive ke Flask API,
 * lalu tulis hasilnya ke Google Sheets.
 *
 * CARA PAKAI:
 * 1. Buka Google Sheets → Extensions → Apps Script
 * 2. Paste seluruh kode ini
 * 3. Ubah API_URL dan FILE_ID di bawah
 * 4. Klik Run → convertPdfToSheet
 */

// ── KONFIGURASI ──────────────────────────────────────────────────────────────

// URL Flask API kamu (ganti dengan IP/domain server)
const API_URL = "http://YOUR_IP:5000/convert";

// ID file PDF di Google Drive
// Cara dapat FILE_ID: buka file di Drive, lihat URL:
// https://drive.google.com/file/d/FILE_ID_ADA_DI_SINI/view
const FILE_ID = "GANTI_DENGAN_FILE_ID_PDF_KAMU";

// ─────────────────────────────────────────────────────────────────────────────

function convertPdfToSheet() {
  // 1. Ambil file PDF dari Google Drive
  const file     = DriveApp.getFileById(FILE_ID);
  const pdfBlob  = file.getBlob().setContentType("application/pdf");

  // 2. Kirim ke Flask API
  const options = {
    method      : "post",
    payload     : { file: pdfBlob },
    muteHttpExceptions: true,
  };

  Logger.log("Mengirim PDF ke API...");
  const response = UrlFetchApp.fetch(API_URL, options);

  if (response.getResponseCode() !== 200) {
    Logger.log("Error dari API: " + response.getContentText());
    SpreadsheetApp.getUi().alert("Gagal: " + response.getContentText());
    return;
  }

  const result = JSON.parse(response.getContentText());
  Logger.log("Total tabel ditemukan: " + result.total_tables);

  // 3. Tulis ke sheet
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = file.getName().replace(".pdf", "");

  // Hapus sheet lama jika ada, buat baru
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  let currentRow = 1; // Google Sheets 1-indexed

  result.tables.forEach((tbl) => {
    // ── Label pemisah ──
    const labelCell = sheet.getRange(currentRow, 1);
    const labelText = `[ Halaman ${tbl.page} — Tabel ${tbl.table_index} ]`;
    labelCell.setValue(labelText);

    // Merge label sepanjang jumlah kolom
    const numCols = tbl.headers.length || 1;
    if (numCols > 1) {
      sheet.getRange(currentRow, 1, 1, numCols).merge();
    }

    // Style label
    labelCell
      .setBackground("#2F5496")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(11);

    currentRow++;

    // ── Header ──
    if (tbl.headers.length > 0) {
      const headerRange = sheet.getRange(currentRow, 1, 1, tbl.headers.length);
      headerRange.setValues([tbl.headers]);
      headerRange
        .setBackground("#4472C4")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      currentRow++;
    }

    // ── Data rows ──
    if (tbl.rows.length > 0) {
      tbl.rows.forEach((row, rowIdx) => {
        const dataRange = sheet.getRange(currentRow, 1, 1, row.length);
        dataRange.setValues([row]);

        // Warna selang-seling
        const bgColor = rowIdx % 2 === 0 ? "#FFFFFF" : "#D9E1F2";
        dataRange.setBackground(bgColor);

        // Format angka: rata kanan, format ribuan
        row.forEach((val, colIdx) => {
          const cell = sheet.getRange(currentRow, colIdx + 1);
          if (typeof val === "number") {
            cell.setHorizontalAlignment("right");
            // Format dengan pemisah ribuan
            const fmt = Number.isInteger(val) ? "#,##0" : "#,##0.00";
            cell.setNumberFormat(fmt);
          }
        });

        currentRow++;
      });
    }

    // 2 baris kosong antar tabel
    currentRow += 2;
  });

  // Auto resize semua kolom
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    sheet.autoResizeColumns(1, lastCol);
  }

  Logger.log("Selesai! Data ditulis ke sheet: " + sheetName);
  SpreadsheetApp.getUi().alert("Selesai! " + result.total_tables + " tabel berhasil diimpor ke sheet '" + sheetName + "'.");
}
