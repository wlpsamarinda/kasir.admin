/**
 * =====================================================================
 * WONDERLAND SAMARINDA - DAILY CASHIER REPORTING SYSTEM
 * =====================================================================
 * @file Code.gs (Combined Backend Engine) - v3.0.1
 *
 * PENTING: File ini membuat SPREADSHEET DATABASE BARU (bukan menyambung
 * ke database lama) supaya tidak ada lagi masalah skema tidak sinkron.
 * Nama spreadsheet baru: "Wonderland_Samarinda_DB_v3".
 * Database lama Anda (jika ada) TIDAK dihapus otomatis - aman di Drive,
 * silakan hapus manual nanti kalau sudah yakin tidak dibutuhkan.
 *
 * CATATAN PERUBAHAN v3.1.0 (KARTU STOK):
 * - FITUR BARU "Kartu Stok": setiap produk yang memakai stok kini punya
 *   kartu stok sendiri di tiap lokasi - Kasir Loket Masuk (WH02), Kasir
 *   Loket Wahana (WH03), Kasir Resto (WH04), Gudang Utama (WH01), plus
 *   gudang internal (WH05-WH08). Kartu stok = buku besar mutasi (stok
 *   masuk, keluar, saldo) per produk per lokasi, disimpan di sheet baru
 *   "Stock_Card" (lihat STOCK_CARD_HEADERS & bagian 7d).
 * - Pencatatan OTOMATIS dari mesin stok yang sudah ada (updateStock_ &
 *   setStockAbsolute_) + semua fungsi yang menulis Warehouse_Stock langsung
 *   (Set Stok Manual, Grid Simpan Semua, Hapus Catatan Stok), sehingga
 *   penjualan, edit/hapus penjualan, transfer antar gudang, stok masuk
 *   gudang utama, stok opname, pemakaian gudang internal, & stok awal
 *   produk semuanya masuk kartu stok tanpa input tambahan dari kasir.
 * - Kartu stok dimulai dari saat versi ini di-deploy (riwayat sebelumnya
 *   tidak ikut ter-backfill karena tidak ada catatan per-produk yang
 *   lengkap sebelumnya).
 * - Setelah update ini, WAJIB buat deployment/versi BARU di Apps Script
 *   Editor (Deploy > Manage deployments > Edit > New version > Deploy).
 *
 * CATATAN PERUBAHAN v3.0.1:
 * - FIX KRITIS: getSheetDataAsObjects_() sekarang mengonversi semua nilai
 *   bertipe Date ke string ISO sebelum dikembalikan ke client. Sebelumnya,
 *   objek Date mentah yang ikut terkirim lewat google.script.run membuat
 *   browser menerima `null` tanpa error apapun - inilah penyebab gejala
 *   "data tersimpan di spreadsheet tapi tidak tampil di web app/dashboard".
 *   Lihat komentar di fungsi getSheetDataAsObjects_() untuk detail lengkap.
 * - Setelah update ini, WAJIB buat deployment/versi BARU di Apps Script
 *   Editor (Deploy > Manage deployments > Edit > New version > Deploy),
 *   karena sekadar menyimpan file tidak memperbarui Web App yang sudah
 *   dideploy sebelumnya.
 * =====================================================================
 */

// =====================================================================
// 1. CONFIG
// =====================================================================
const CONFIG = {
  APP_NAME: "Wonderland Samarinda - Cashier Report System",
  COMPANY_NAME: "PT. Wisata Taman Ceria",
  DB_PROPERTY_KEY: "WONDERLAND_DB_ID_V3", // sengaja key baru -> memaksa DB baru dibuat
  SPREADSHEET_NAME: "Wonderland_Samarinda_DB_v3",
  DRIVE_FOLDERS: {
    REPORT: "Wonderland_Report_Folder",
    BACKUP: "Wonderland_Backup_Folder",
    LOGO: "Wonderland_Logo_Folder"
  },
  SHEETS: {
    USERS: "Users",
    ROLES: "Roles",
    SETTINGS: "Settings",
    PRODUCTS: "Products",
    CATEGORIES: "Categories",
    PAYMENT_METHODS: "Payment_Methods",
    PRODUCT_TYPES: "Product_Types",
    APPLICABLE_UNITS: "Applicable_Units",
    UNITS: "Units",
    WAREHOUSES: "Warehouses",
    WAREHOUSE_STOCK: "Warehouse_Stock",
    STOCK_OPNAME: "Stock_Opname_Fisik",
    ENTRANCE_SALES: "Entrance_Sales",
    ENTRANCE_PAYMENT: "Entrance_Payment",
    ENTRANCE_STOCK: "Entrance_Stock",
    RIDE_SALES: "Ride_Sales",
    RIDE_PAYMENT: "Ride_Payment",
    RIDE_STOCK: "Ride_Stock",
    RESTO_SALES: "Resto_Sales",
    RESTO_PAYMENT: "Resto_Payment",
    RESTO_STOCK: "Resto_Stock",
    // Sheet riwayat pemakaian/mutasi stok KHUSUS masing-masing gudang
    // internal (tanpa kasir/POS) - Kolam Renang/Teknisi/Cleaning Service/
    // Dapur, mirip Entrance_Stock/Ride_Stock/Resto_Stock untuk outlet kasir.
    // Sebelumnya riwayat ke-4 gudang ini cuma bercampur di Stock_Mutations
    // (satu sheet raksasa untuk SEMUA gudang) - lihat INTERNAL_WAREHOUSE_LOG_SHEET_.
    KOLAM_STOCK: "Kolam_Stock",
    TEKNISI_STOCK: "Teknisi_Stock",
    CLEANING_STOCK: "Cleaning_Stock",
    DAPUR_STOCK: "Dapur_Stock",
    STOCK_MUTATIONS: "Stock_Mutations",
    // Kartu Stok: 1 baris per pergerakan stok (masuk/keluar) per produk per
    // gudang/outlet, lengkap dengan saldo sebelum & sesudah - lihat
    // STOCK_CARD_HEADERS dan bagian 7d (KARTU STOK).
    STOCK_CARD: "Stock_Card",
    ACTIVITY_LOG: "Activity_Log",
    LOGIN_LOG: "Login_Log",
    BACKUP_LOG: "Backup_Log",
    SESSIONS: "Sessions",
    FINANCE_BALANCE: "Finance_Balance",
    FINANCE_BALANCE_MUTATIONS: "Finance_Balance_Mutations",
    FINANCE_EXPENSES: "Finance_Expenses",
    FINANCE_COA: "Finance_COA",
    FINANCE_PAYMENT_STATUS: "Finance_Payment_Status",
    FINANCE_POS_PENGGUNAAN: "Finance_Pos_Penggunaan",
    FINANCE_REPORT_EMAIL: "Finance_Report_Email_Recipients"
  }
};

// Kategori/"Bucket" pengeluaran pada modul Admin Keuangan, sesuai 3 bagian
// di form laporan harian: pengeluaran yang dipotong langsung dari omset
// harian (sebelum dihitung Pendapatan Bersih), pengeluaran dari Brangkas,
// pengeluaran dari Rekening PT, dan BelumLunas = tagihan/hutang supplier
// yang belum dibayar dari manapun (ditampilkan sebagai widget "Tagihan").
const FINANCE_EXPENSE_CATEGORY = { OMSET: "Omset", BRANGKAS: "Brangkas", REKENING: "RekeningPT", BELUM_LUNAS: "BelumLunas" };
const FINANCE_EXPENSE_LABEL_ = {
  Omset: "Pengeluaran dari Omset Harian", Brangkas: "Pengeluaran dari Brangkas",
  RekeningPT: "Pengeluaran dari Rekening PT", BelumLunas: "Belum Lunas (Tagihan)"
};

// Sesi login: cache cepat (maks 6 jam bawaan Google) + dicadangkan di sheet
// Sessions supaya TIDAK otomatis logout saat halaman di-refresh atau server
// Apps Script "tidur". Selama dipakai (sliding expiration), sesi diperpanjang
// otomatis tiap kali diverifikasi.
const SESSION_HEADERS = ["Token", "UserID", "Username", "FullName", "Role", "AssignedUnit", "CreatedAt", "ExpiresAt"];
const SESSION_DURATION_MS = 12 * 60 * 60 * 1000; // 12 jam sejak aktivitas terakhir

// Modul Admin Keuangan: 1 baris per tanggal (Saldo Awal/Akhir Brangkas &
// Rekening, Saldo THR, Hutang/Piutang & catatan tambahan disimpan sebagai
// JSON string supaya jumlah baris hutang/piutang bisa fleksibel).
// SaldoAkhirBrangkas/SaldoAkhirRekening (2 kolom terakhir) adalah SNAPSHOT
// PERMANEN hasil resolusi Saldo Akhir hari itu (Aktual kalau diisi manual,
// atau hasil Hitungan otomatis kalau Aktual dikosongkan) - lihat
// resolvedSaldoAkhir_(). Nilai inilah yang jadi acuan Saldo Awal besok,
// BUKAN kolom Aktual saja, supaya rantai saldo antar hari tidak putus
// hanya karena Admin membiarkan Aktual kosong (pakai Estimasi otomatis).
const FINANCE_BALANCE_HEADERS = ["BalanceID", "Date", "SaldoAwalBrangkas", "SaldoAwalRekening", "SaldoTHR",
  "SaldoAkhirBrangkasAktual", "SaldoAkhirRekeningAktual", "CatatanTambahan", "HutangJSON", "PiutangJSON", "CreatedBy", "UpdatedAt",
  "SaldoAkhirBrangkas", "SaldoAkhirRekening"];
// Mutasi Saldo (pindah dana internal Brangkas <-> Rekening, BUKAN transaksi
// kasir/pengeluaran) - misal Owner setor tunai dari Brangkas ke Rekening PT
// (atau tarik tunai sebaliknya). Setiap baris HANYA memindahkan 1 arah:
// FromBucket berkurang, ToBucket bertambah sebesar Amount, di TANGGAL yang
// sama (bukan cuma hari ini - bisa dicatat utk tanggal manapun). Dampaknya
// otomatis masuk ke perhitungan Saldo Akhir Brangkas/Rekening tanggal itu
// lewat computeFinanceCashFlow_() - lihat addFinanceMutation() &
// recalcFinanceBalanceSaldoAkhirSnapshot_().
const FINANCE_BALANCE_MUTATIONS_HEADERS = ["MutationID", "Date", "FromBucket", "ToBucket", "Amount", "Notes", "CreatedBy", "Timestamp"];
// Category di sini berfungsi sebagai "Bucket" pembayaran (Omset/Brangkas/
// RekeningPT/BelumLunas) - dipertahankan namanya untuk kompatibilitas data
// lama. Kolom baru (Supplier, UnitPrice, Qty, Unit, StatusPembayaran,
// COACode, COAName, PosPenggunaan) ditambahkan di akhir supaya sheet lama
// otomatis bisa dimigrasi tanpa kehilangan data (lihat ensureFinanceExpensesSchema_).
const FINANCE_EXPENSES_HEADERS = ["ExpenseID", "Date", "Category", "Description", "Amount", "CreatedBy", "Timestamp",
  "Supplier", "UnitPrice", "Qty", "Unit", "StatusPembayaran", "COACode", "COAName", "PosPenggunaan"];
const FINANCE_COA_HEADERS = ["COAID", "COACode", "COAName", "Type", "Status"]; // Type: Pendapatan/Pengeluaran
const FINANCE_PAYMENT_STATUS_HEADERS = ["StatusID", "StatusName", "Bucket", "Status"]; // Bucket: Omset/Brangkas/RekeningPT/BelumLunas
const FINANCE_POS_PENGGUNAAN_HEADERS = ["PosID", "PosName", "Status"];
// Daftar tujuan email penerima Laporan Harian Admin Keuangan (fitur kirim
// laporan PDF/XLSX langsung ke email) - bisa ditambah/edit/hapus dari web app.
const FINANCE_REPORT_EMAIL_HEADERS = ["RecipientID", "Email", "Name", "Status"];

const ROLE = { OWNER: "Owner", ADMIN: "Admin", KASIR: "Kasir", STAFF: "Staff" };
const CENTRAL_WAREHOUSE_ID = "WH01";
// Skema sheet Warehouse_Stock: selain ID mentah (WarehouseID/ProductCode)
// juga menyimpan WarehouseName, ProductName, & Status (Active/Inactive)
// produk ybs supaya sheet-nya enak dibaca langsung tanpa VLOOKUP manual.
// Kolom "manusiawi" ini diisi/disinkronkan otomatis lewat warehouseProductInfo_()
// setiap kali baris dibuat/diupdate - lihat updateStock_/setOutletStock/setWarehouseStock.
const WAREHOUSE_STOCK_HEADERS = ["StockID", "WarehouseID", "WarehouseName", "ProductCode", "ProductName", "Quantity", "MinStock", "Status", "LastUpdated"];
// Header sheet riwayat pemakaian/mutasi KHUSUS tiap gudang internal (lihat
// CONFIG.SHEETS.KOLAM_STOCK/TEKNISI_STOCK/CLEANING_STOCK/DAPUR_STOCK &
// INTERNAL_WAREHOUSE_LOG_SHEET_). ActivityType: "Terpakai/Keluar" (dicatat
// manual lewat recordWarehouseUsage), "Transfer Masuk"/"Transfer Keluar"
// (dicatat otomatis lewat transferStock kalau salah satu ujungnya gudang ini).
const INTERNAL_WH_LOG_HEADER = ["LogID", "Date", "ProductCode", "ProductName", "Qty", "ActivityType", "FromWarehouse", "ToWarehouse", "Notes", "By"];
// Kartu Stok (Stock Card): buku besar pergerakan stok. Setiap kali stok
// sebuah produk di sebuah gudang/outlet berubah, SATU baris dicatat di sini
// (bukan ditimpa) berisi QtyIn (masuk), QtyOut (keluar), BalanceBefore &
// BalanceAfter (saldo sebelum & sesudah), jenis transaksi, lawan transaksi
// (Counterpart: gudang asal/tujuan transfer, atau "Supplier/Manual"), nomor
// referensi (RefID: SaleID/OpnameID/TransferRef), catatan, & pelaku. Saldo
// selalu dibaca LANGSUNG dari Warehouse_Stock pada saat transaksi terjadi,
// jadi rantai saldo kartu stok selalu sinkron dengan stok riil.
// ViaProduct terisi kalau yang terjual/ditransfer adalah produk ber-Acuan
// Stok (StockRefCode) - kartunya dicatat di produk acuan, dan kolom ini
// menyimpan kode produk aslinya supaya tetap terlacak.
const STOCK_CARD_HEADERS = ["EntryID", "Timestamp", "WarehouseID", "WarehouseName", "ProductCode", "ProductName",
  "TransactionType", "QtyIn", "QtyOut", "BalanceBefore", "BalanceAfter", "Counterpart", "RefID", "Notes", "User", "ViaProduct"];
const STOCK_CARD_TYPE = {
  OPENING: "Stok Awal Produk",
  SALE: "Penjualan",
  SALE_EDIT: "Koreksi Edit Penjualan",
  SALE_DELETE: "Hapus Penjualan (Stok Kembali)",
  TRANSFER_IN: "Transfer Masuk",
  TRANSFER_OUT: "Transfer Keluar",
  CENTRAL_IN: "Stok Masuk (Supplier/Manual)",
  OPNAME: "Stok Opname",
  MANUAL: "Penyesuaian Manual",
  USAGE: "Pemakaian / Keluar",
  RECORD_DELETE: "Hapus Catatan Stok",
  UNKNOWN: "Perubahan Stok"
};
// Cek Stok Fisik Barang (stok opname harian per outlet): 1 baris per produk
// per tanggal per outlet. StokAwal = stok tercatat di Warehouse_Stock SAAT
// disimpan (idealnya diisi kasir SEBELUM mulai berjualan, jadi ini = stok
// akhir hari sebelumnya). Menyimpan ulang di tanggal & outlet yang sama utk
// produk yang sama akan MENIMPA baris lama (bukan duplikat) - lihat
// savePhysicalStockCheck().
const STOCK_OPNAME_HEADERS = ["OpnameID", "Date", "Unit", "WarehouseID", "ProductCode", "ProductName", "Category",
  "StokAwal", "StokFisik", "Selisih", "SellPrice", "NilaiSelisih", "CheckedBy", "Timestamp"];
const UNIT_LABEL_SERVER_ = { Entrance: "Kasir Loket Masuk", Ride: "Kasir Loket Wahana", Resto: "Kasir Resto", General: "Semua Outlet" };
// Daftar Tipe Produk (dulu WAREHOUSE_ITEM_TYPES yang hardcode 4 pilihan)
// sekarang jadi master data yang bisa diatur sendiri (tambah/edit/hapus)
// lewat menu Kategori & Metode Pembayaran, disimpan di sheet Product_Types.
const PRODUCT_TYPE_HEADERS = ["TypeID", "TypeName", "Status"];
// Daftar Unit/Outlet Berlaku (dulu 4 pilihan hardcode General/Entrance/
// Ride/Resto di dropdown ApplicableUnit Master Produk) - sekarang jadi
// master data yang bisa diatur sendiri (tambah/edit/hapus) lewat menu
// Kategori & Metode Pembayaran, disimpan di sheet Applicable_Units.
// UnitKey = nilai mentah yang disimpan di kolom ApplicableUnit Products
// (dan dicocokkan oleh productUnits_/applicableUnitByWarehouseId_), harus
// unik & tanpa koma/spasi supaya aman dipakai sebagai bagian dari string
// gabungan "Resto,Ride". UnitLabel = nama yang ditampilkan ke pengguna.
const APPLICABLE_UNIT_HEADERS = ["UnitID", "UnitKey", "UnitLabel", "Status"];
// UnitKey "General" tidak boleh dihapus lewat UI (lihat deleteApplicableUnit)
// karena banyak logika inti (productUnits_, getProducts, dst) memakainya
// sebagai fallback "berlaku di semua unit".
const GENERAL_APPLICABLE_UNIT_KEY_ = "General";

// Peta generik tiap outlet kasir. Menambah outlet baru cukup tambah 1 entri di sini.
const UNIT_MAP = {
  Entrance: {
    label: "Kasir Loket Masuk",
    salesSheet: CONFIG.SHEETS.ENTRANCE_SALES,
    paySheet: CONFIG.SHEETS.ENTRANCE_PAYMENT,
    stockLogSheet: CONFIG.SHEETS.ENTRANCE_STOCK,
    warehouseId: "WH02",
    hasSerial: true
  },
  Ride: {
    label: "Kasir Loket Wahana",
    salesSheet: CONFIG.SHEETS.RIDE_SALES,
    paySheet: CONFIG.SHEETS.RIDE_PAYMENT,
    stockLogSheet: CONFIG.SHEETS.RIDE_STOCK,
    warehouseId: "WH03",
    hasSerial: true
  },
  Resto: {
    label: "Kasir Resto",
    salesSheet: CONFIG.SHEETS.RESTO_SALES,
    paySheet: CONFIG.SHEETS.RESTO_PAYMENT,
    stockLogSheet: CONFIG.SHEETS.RESTO_STOCK,
    warehouseId: "WH04",
    hasSerial: false
  }
};

function unitCfg_(unit) {
  const cfg = UNIT_MAP[unit];
  if (!cfg) throw new Error("Unit tidak dikenali: " + unit);
  return cfg;
}

function fmtD_(d, tz) {
  return d ? Utilities.formatDate(new Date(d), tz, "dd/MM/yyyy") : "-";
}

// =====================================================================
// 2. ROUTER
// =====================================================================
function doGet(e) {
  let template = HtmlService.createTemplateFromFile('Index');
  template.appName = CONFIG.APP_NAME;
  template.companyName = CONFIG.COMPANY_NAME;
  return template.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================================
// API ENDPOINT UNTUK PWA (GitHub Pages) - dipanggil oleh gas-bridge.js
// =====================================================================
const API_WHITELIST_ = ['addCentralStock', 'addFinanceMutation', 'addStockFromGudang', 'bulkDeleteProducts', 'bulkImportFinanceCOA', 'bulkImportFinanceExpenses', 'bulkImportFinancePaymentStatus', 'bulkImportFinancePosPenggunaan', 'bulkImportProducts', 'bulkSetOutletStock', 'deleteApplicableUnit', 'deleteCategory', 'deleteFinanceBalance', 'deleteFinanceCOA', 'deleteFinanceExpense', 'deleteFinanceMutation', 'deleteFinancePaymentStatus', 'deleteFinancePosPenggunaan', 'deleteFinanceReportEmailRecipient', 'deleteOutletStock', 'deletePaymentMethod', 'deletePaymentTransaction', 'deleteProduct', 'deleteProductType', 'deleteSalesTransaction', 'deleteUser', 'deleteWarehouseStockRecord', 'getAllKasirReportData', 'getApplicableUnits', 'getBalanceCheckReport', 'getCategories', 'getDailyReportData', 'getDashboardData', 'getDatabaseInfo', 'getFinanceBalance', 'getFinanceBalanceHistory', 'getFinanceCOA', 'getFinanceDailyReport', 'getFinanceExpenses', 'getFinanceExpensesRange', 'getFinanceMutationHistory', 'getFinancePaymentStatusList', 'getFinancePosPenggunaanList', 'getFinanceReportEmailRecipients', 'getMenuConfig', 'getPaymentData', 'getPaymentMethods', 'getPhysicalStockCheckView', 'getProductStock', 'getProductTypes', 'getProducts', 'getProductsWithStock', 'getSalesData', 'getSalesFormConfig', 'getStockCard', 'getStockCardList', 'getStockCardsBulk', 'getStockView', 'getSystemSettings', 'getUserPassword', 'getUsers', 'getWarehouseList', 'getWarehouseMonitor', 'getWarehouseOverview', 'getWarehouseStockDetail', 'getWarehouseTransferHistory', 'loginUser', 'logoutUser', 'recordWarehouseUsage', 'repairDatabaseAuto', 'resetMenuConfig', 'runDatabaseDiagnostic', 'saveApplicableUnit', 'saveCategory', 'saveFinanceBalance', 'saveFinanceCOA', 'saveFinanceExpense', 'saveFinancePaymentStatus', 'saveFinancePosPenggunaan', 'saveFinanceReportEmailRecipient', 'saveMenuConfig', 'savePaymentMethod', 'savePaymentTransaction', 'savePhysicalStockCheck', 'saveProduct', 'saveProductType', 'saveSalesFormConfig', 'saveSalesTransaction', 'saveSystemSettings', 'saveUser', 'sendFinanceReportEmail', 'setDatabaseId', 'setOutletStock', 'setWarehouseStock', 'transferStock', 'triggerManualBackup', 'uploadLogo', 'verifySession'];

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (API_WHITELIST_.indexOf(req.fn) === -1) throw new Error('Fungsi tidak diizinkan: ' + req.fn);
    const fn = this[req.fn];
    if (typeof fn !== 'function') throw new Error('Fungsi tidak ditemukan');
    const result = fn.apply(null, req.args || []);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, result: (result === undefined ? null : result) }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =====================================================================
// 3. DATABASE - DIBUAT BARU, SKEMA FINAL (TIDAK ADA MIGRASI)
// =====================================================================
function getDatabase() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let ssId = scriptProperties.getProperty(CONFIG.DB_PROPERTY_KEY);
  let ss;

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }

  if (!ss) {
    const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
      initializeNewDatabase(ss);
    }
    scriptProperties.setProperty(CONFIG.DB_PROPERTY_KEY, ss.getId());
  }

  ensureUsersPasswordColumn_(ss);
  ensureSheet_(ss, CONFIG.SHEETS.FINANCE_BALANCE, FINANCE_BALANCE_HEADERS);
  ensureFinanceBalanceSchema_(ss);
  ensureSheet_(ss, CONFIG.SHEETS.FINANCE_BALANCE_MUTATIONS, FINANCE_BALANCE_MUTATIONS_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.FINANCE_EXPENSES, FINANCE_EXPENSES_HEADERS);
  ensureFinanceExpensesSchema_(ss);
  ensureSheet_(ss, CONFIG.SHEETS.FINANCE_COA, FINANCE_COA_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.FINANCE_PAYMENT_STATUS, FINANCE_PAYMENT_STATUS_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN, FINANCE_POS_PENGGUNAAN_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.FINANCE_REPORT_EMAIL, FINANCE_REPORT_EMAIL_HEADERS);
  seedFinanceMasterDataIfEmpty_(ss);
  ensureWarehousesSeed_(ss);
  ensureProductTypesSeed_(ss);
  ensureApplicableUnitsSeed_(ss);
  ensureProductsItemTypeColumn_(ss);
  ensureProductsHargaFleksibelColumn_(ss);
  ensureProductsUsesStockColumn_(ss);
  ensureProductsUsesSerialColumn_(ss);
  ensureProductsStockRefColumn_(ss);
  ensureWarehouseStockColumns_(ss);
  ensureSheet_(ss, CONFIG.SHEETS.STOCK_OPNAME, STOCK_OPNAME_HEADERS);
  // Sheet riwayat KHUSUS tiap gudang internal (lihat INTERNAL_WH_LOG_HEADER) -
  // dibuat otomatis kalau belum ada, database lama pun langsung dapat
  // sheet ini tanpa perlu klik "Perbaiki Database" (tapi tetap terdeteksi
  // di sana juga lewat getFullSchemaMap_() kalau untuk alasan tertentu
  // gagal dibuat di sini, mis. kuota).
  ensureSheet_(ss, CONFIG.SHEETS.KOLAM_STOCK, INTERNAL_WH_LOG_HEADER);
  ensureSheet_(ss, CONFIG.SHEETS.TEKNISI_STOCK, INTERNAL_WH_LOG_HEADER);
  ensureSheet_(ss, CONFIG.SHEETS.CLEANING_STOCK, INTERNAL_WH_LOG_HEADER);
  ensureSheet_(ss, CONFIG.SHEETS.DAPUR_STOCK, INTERNAL_WH_LOG_HEADER);
  ensureSheet_(ss, CONFIG.SHEETS.STOCK_CARD, STOCK_CARD_HEADERS);
  return ss;
}

// MIGRASI: kalau sheet Finance_Expenses sudah ada dari versi lama (skema
// 7 kolom: ExpenseID, Date, Category, Description, Amount, CreatedBy,
// Timestamp), kolom-kolom baru untuk fitur "Pengeluaran Harian" (Supplier,
// UnitPrice, Qty, Unit, StatusPembayaran, COACode, COAName, PosPenggunaan)
// otomatis ditambahkan di akhir tanpa mengubah/menghapus data lama sama
// sekali - data lama tetap valid, hanya kolom detail barunya kosong.
function ensureFinanceExpensesSchema_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    let missing = FINANCE_EXPENSES_HEADERS.filter(h => headers.indexOf(h) === -1);
    if (missing.length) {
      sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing])
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  } catch (e) {}
}

// MIGRASI: samakan kolom sheet Finance_Balance yang sudah ada (dibuat dari
// versi kode lama, mis. urutan/isi kolom belum termasuk SaldoAkhirBrangkas
// Aktual & SaldoAkhirRekeningAktual) dengan FINANCE_BALANCE_HEADERS versi
// terbaru. Kolom yang belum ada ditambahkan di akhir tanpa menghapus data
// lama - dijalankan oleh saveFinanceBalance/getFinanceBalance lewat
// header-name mapping (bukan posisi kolom tetap), supaya Saldo Akhir yang
// diisi Admin Keuangan selalu tersimpan ke kolom yang benar walau urutan
// kolom fisik di sheet berbeda dari FINANCE_BALANCE_HEADERS.
function ensureFinanceBalanceSchema_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    let missing = FINANCE_BALANCE_HEADERS.filter(h => headers.indexOf(h) === -1);
    if (missing.length) {
      sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing])
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  } catch (e) {}
}

// Isi data awal (default) untuk master Kode Akun / Status Pembayaran /
// Pos Penggunaan HANYA kalau sheet-nya masih kosong (baris 2 belum ada) -
// supaya aman dipanggil berulang tiap kali getDatabase() jalan tanpa
// menimpa perubahan yang sudah dilakukan Admin Keuangan.
function seedFinanceMasterDataIfEmpty_(ss) {
  try {
    const coaSheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_COA);
    if (coaSheet && coaSheet.getLastRow() < 2) {
      coaSheet.getRange(2, 1, 6, 5).setValues([
        ["COA001", "5101", "Operasional Harian", "Pengeluaran", "Active"],
        ["COA002", "5102", "Perlengkapan & Maintenance", "Pengeluaran", "Active"],
        ["COA003", "5103", "Gaji & Upah", "Pengeluaran", "Active"],
        ["COA004", "5104", "Marketing & Promosi", "Pengeluaran", "Active"],
        ["COA005", "5105", "Bahan Baku F&B", "Pengeluaran", "Active"],
        ["COA006", "4101", "Pendapatan Tiket & Penjualan", "Pendapatan", "Active"]
      ]);
    }
    const psSheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_PAYMENT_STATUS);
    if (psSheet && psSheet.getLastRow() < 2) {
      psSheet.getRange(2, 1, 4, 4).setValues([
        ["PS001", "Lunas - Omset Harian", FINANCE_EXPENSE_CATEGORY.OMSET, "Active"],
        ["PS002", "Lunas - Brangkas", FINANCE_EXPENSE_CATEGORY.BRANGKAS, "Active"],
        ["PS003", "Lunas - Rekening PT", FINANCE_EXPENSE_CATEGORY.REKENING, "Active"],
        ["PS004", "Belum Lunas", FINANCE_EXPENSE_CATEGORY.BELUM_LUNAS, "Active"]
      ]);
    }
    const posSheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN);
    if (posSheet && posSheet.getLastRow() < 2) {
      posSheet.getRange(2, 1, 5, 3).setValues([
        ["POS001", "Loket Masuk", "Active"],
        ["POS002", "Loket Wahana", "Active"],
        ["POS003", "Resto", "Active"],
        ["POS004", "Kantor / Umum", "Active"],
        ["POS005", "Maintenance / Teknisi", "Active"]
      ]);
    }
  } catch (e) {}
}

// MIGRASI: kalau database lama (dibuat sebelum fitur "Lihat Password" oleh
// Owner ditambahkan) belum punya kolom PasswordPlain di sheet Users, kolom
// ini otomatis ditambahkan di sini supaya tidak perlu reset database.
// Password user-user LAMA tidak bisa dipulihkan dalam bentuk asli (karena
// sebelumnya hanya hash yang tersimpan) - kolom akan kosong sampai Owner
// mengatur ulang password user tsb lewat Manajemen User.
function ensureUsersPasswordColumn_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headers.indexOf("PasswordPlain") === -1) {
      sheet.getRange(1, headers.length + 1).setValue("PasswordPlain")
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  } catch (e) {}
}

/**
 * UTILITY: hapus pointer database supaya lain kali getDatabase() dipanggil,
 * spreadsheet BARU akan dibuat lagi dari nol. Jalankan manual dari editor
 * Apps Script (pilih fungsi ini di dropdown toolbar > Run) kalau suatu saat
 * perlu reset total lagi. Spreadsheet lama TIDAK dihapus otomatis dari Drive.
 */
function forceRecreateDatabase() {
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.DB_PROPERTY_KEY);
  Logger.log("Pointer database dihapus. Buka web app lagi untuk membuat database baru.");
}

// Peta LENGKAP nama sheet -> header yang seharusnya ada, dipakai baik oleh
// initializeNewDatabase() (database baru) MAUPUN oleh Diagnostic Database
// (database lama yang sudah jalan) supaya keduanya selalu sinkron dari SATU
// sumber kebenaran - tidak ada lagi daftar sheet/header yang didefinisikan
// dobel di tempat berbeda dan bisa saling berbeda seiring waktu.
function getFullSchemaMap_() {
  const salesHeaderFull = ["SaleID", "Date", "ProductCode", "ProductName", "Variant", "Category", "Qty", "Unit",
    "SerialStart", "SerialEnd", "TotalSerial", "StockStart", "StockIn", "StockOut", "StockEnd",
    "UnitPrice", "TotalPrice", "Cashier", "InputTime", "Notes"];
  const salesHeaderResto = ["SaleID", "Date", "ProductCode", "ProductName", "Variant", "Category", "Qty", "Unit",
    "StockStart", "StockIn", "StockOut", "StockEnd", "UnitPrice", "TotalPrice", "Cashier", "InputTime", "Notes"];
  const paymentHeader = ["PaymentID", "Date", "Time", "PaymentMethod", "Amount", "Reference", "Cashier", "Notes"];
  const stockLogHeader = ["ReqID", "Date", "ProductCode", "ProductName", "QtyRequested", "FromWarehouse", "ToWarehouse", "Status", "RequestedBy", "Timestamp"];

  return {
    [CONFIG.SHEETS.USERS]: ["UserID", "Username", "FullName", "Role", "AssignedUnit", "PasswordHash", "Status", "CreatedAt", "PasswordPlain"],
    [CONFIG.SHEETS.ROLES]: ["RoleID", "RoleName", "Description", "Permissions"],
    [CONFIG.SHEETS.SETTINGS]: ["Key", "Value", "Description"],
    [CONFIG.SHEETS.PRODUCTS]: ["ProductCode", "ProductName", "Variant", "Category", "Unit", "BuyPrice", "SellPrice", "Status", "ApplicableUnit", "HargaFleksibel", "ItemType", "UsesStock", "UsesSerial", "StockRefCode"],
    [CONFIG.SHEETS.CATEGORIES]: ["CategoryID", "CategoryName", "Type", "Status"],
    [CONFIG.SHEETS.PAYMENT_METHODS]: ["MethodID", "MethodName", "Status"],
    [CONFIG.SHEETS.PRODUCT_TYPES]: PRODUCT_TYPE_HEADERS,
    [CONFIG.SHEETS.APPLICABLE_UNITS]: APPLICABLE_UNIT_HEADERS,
    [CONFIG.SHEETS.UNITS]: ["UnitID", "UnitName", "Symbol"],
    [CONFIG.SHEETS.WAREHOUSES]: ["WarehouseID", "WarehouseName", "Location", "UnitType"],
    [CONFIG.SHEETS.WAREHOUSE_STOCK]: WAREHOUSE_STOCK_HEADERS,
    [CONFIG.SHEETS.ENTRANCE_SALES]: salesHeaderFull,
    [CONFIG.SHEETS.ENTRANCE_PAYMENT]: paymentHeader,
    [CONFIG.SHEETS.ENTRANCE_STOCK]: stockLogHeader,
    [CONFIG.SHEETS.RIDE_SALES]: salesHeaderFull,
    [CONFIG.SHEETS.RIDE_PAYMENT]: paymentHeader,
    [CONFIG.SHEETS.RIDE_STOCK]: stockLogHeader,
    [CONFIG.SHEETS.RESTO_SALES]: salesHeaderResto,
    [CONFIG.SHEETS.RESTO_PAYMENT]: paymentHeader,
    [CONFIG.SHEETS.RESTO_STOCK]: stockLogHeader,
    [CONFIG.SHEETS.KOLAM_STOCK]: INTERNAL_WH_LOG_HEADER,
    [CONFIG.SHEETS.TEKNISI_STOCK]: INTERNAL_WH_LOG_HEADER,
    [CONFIG.SHEETS.CLEANING_STOCK]: INTERNAL_WH_LOG_HEADER,
    [CONFIG.SHEETS.DAPUR_STOCK]: INTERNAL_WH_LOG_HEADER,
    [CONFIG.SHEETS.STOCK_CARD]: STOCK_CARD_HEADERS,
    [CONFIG.SHEETS.STOCK_MUTATIONS]: ["MutationID", "Timestamp", "ProductCode", "FromLocation", "ToLocation", "Qty", "ReferenceType", "ReferenceID", "User"],
    [CONFIG.SHEETS.ACTIVITY_LOG]: ["LogID", "Timestamp", "User", "Action", "Module", "Details", "IP"],
    [CONFIG.SHEETS.LOGIN_LOG]: ["LoginID", "Timestamp", "User", "Status", "IP", "Browser"],
    [CONFIG.SHEETS.BACKUP_LOG]: ["BackupID", "Timestamp", "FileName", "Status", "DriveURL"],
    [CONFIG.SHEETS.SESSIONS]: SESSION_HEADERS,
    [CONFIG.SHEETS.FINANCE_BALANCE]: FINANCE_BALANCE_HEADERS,
    [CONFIG.SHEETS.FINANCE_BALANCE_MUTATIONS]: FINANCE_BALANCE_MUTATIONS_HEADERS,
    [CONFIG.SHEETS.FINANCE_EXPENSES]: FINANCE_EXPENSES_HEADERS,
    [CONFIG.SHEETS.FINANCE_COA]: FINANCE_COA_HEADERS,
    [CONFIG.SHEETS.FINANCE_PAYMENT_STATUS]: FINANCE_PAYMENT_STATUS_HEADERS,
    [CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN]: FINANCE_POS_PENGGUNAAN_HEADERS,
    [CONFIG.SHEETS.FINANCE_REPORT_EMAIL]: FINANCE_REPORT_EMAIL_HEADERS,
    [CONFIG.SHEETS.STOCK_OPNAME]: STOCK_OPNAME_HEADERS
  };
}

function initializeNewDatabase(ss) {
  createDriveFolders_();

  const schema = getFullSchemaMap_();
  let defaultSheet = ss.getSheetByName("Sheet1");

  for (let sheetName in schema) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    const headers = schema[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  if (defaultSheet) { try { ss.deleteSheet(defaultSheet); } catch (e) {} }

  seedDefaultData_(ss);
  seedFinanceMasterDataIfEmpty_(ss);
}

function createDriveFolders_() {
  for (let key in CONFIG.DRIVE_FOLDERS) {
    let folderName = CONFIG.DRIVE_FOLDERS[key];
    let folders = DriveApp.getFoldersByName(folderName);
    if (!folders.hasNext()) DriveApp.createFolder(folderName);
  }
}

function seedDefaultData_(ss) {
  ss.getSheetByName(CONFIG.SHEETS.ROLES).getRange(2, 1, 4, 4).setValues([
    ["R001", "Owner", "Akses penuh + pengaturan aplikasi & password user", "ALL"],
    ["R002", "Admin", "Akses penuh operasional, tidak bisa lihat/ubah password user & pengaturan", "ADMIN"],
    ["R003", "Kasir", "Akses input & laporan pada unit yang ditugaskan saja", "CASHIER"],
    ["R004", "Staff", "Akses baca (read-only) untuk laporan", "READ_ONLY"]
  ]);

  const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  usersSheet.getRange(2, 1, 6, 9).setValues([
    ["U001", "owner", "Super Administrator", ROLE.OWNER, "All", hashPassword_("owner123"), "Active", new Date(), "owner123"],
    ["U002", "admin", "Administrator", ROLE.ADMIN, "All", hashPassword_("admin123"), "Active", new Date(), "admin123"],
    ["U003", "kasir.masuk", "Kasir Loket Masuk", ROLE.KASIR, "Entrance", hashPassword_("kasir123"), "Active", new Date(), "kasir123"],
    ["U004", "kasir.wahana", "Kasir Loket Wahana", ROLE.KASIR, "Ride", hashPassword_("kasir123"), "Active", new Date(), "kasir123"],
    ["U005", "kasir.resto", "Kasir Resto", ROLE.KASIR, "Resto", hashPassword_("kasir123"), "Active", new Date(), "kasir123"],
    ["U006", "staff", "Staff Pelaporan", ROLE.STAFF, "All", hashPassword_("staff123"), "Active", new Date(), "staff123"]
  ]);

  ss.getSheetByName(CONFIG.SHEETS.SETTINGS).getRange(2, 1, 10, 3).setValues([
    ["APP_NAME", CONFIG.APP_NAME, "Application Name"],
    ["COMPANY_NAME", CONFIG.COMPANY_NAME, "Company Name"],
    ["ADDRESS", "Jl. Poros Samarinda", "Company Address"],
    ["PHONE", "+62 811 5555 777", "Contact Phone"],
    ["EMAIL", "info@wonderlandsamarinda.com", "Contact Email"],
    ["CURRENCY", "IDR", "Default Currency"],
    ["TIMEZONE", "Asia/Makassar", "System Timezone"],
    ["LOGO_URL", "", "URL Logo Aplikasi"],
    ["ENABLE_STOCK_VALIDATION", "No", "Yes = tolak input penjualan jika stok kurang. No = fleksibel, stok boleh minus"],
    ["SHOW_PAYMENT_METHOD_IN_SALES", "No", "Yes = tampilkan & catat metode pembayaran di form Input Penjualan"]
  ]);

  ss.getSheetByName(CONFIG.SHEETS.UNITS).getRange(2, 1, 4, 3).setValues([
    ["UN01", "Pcs", "pcs"], ["UN02", "Tiket", "tkt"], ["UN03", "Cup", "cup"], ["UN04", "Porsi", "prs"]
  ]);

  ss.getSheetByName(CONFIG.SHEETS.CATEGORIES).getRange(2, 1, 4, 4).setValues([
    ["CAT01", "Tiket Masuk", "Entrance", "Active"],
    ["CAT02", "Tiket Wahana", "Ride", "Active"],
    ["CAT03", "Makanan & Minuman", "Resto", "Active"],
    ["CAT04", "Merchandise", "General", "Active"]
  ]);

  ss.getSheetByName(CONFIG.SHEETS.PAYMENT_METHODS).getRange(2, 1, 5, 3).setValues([
    ["PM01", "Uang Tunai", "Active"],
    ["PM02", "QRIS", "Active"],
    ["PM03", "Transfer", "Active"],
    ["PM04", "Kartu Debit", "Active"],
    ["PM05", "Kartu Kredit", "Active"]
  ]);

  ss.getSheetByName(CONFIG.SHEETS.PRODUCT_TYPES).getRange(2, 1, 4, 3).setValues([
    ["TYP01", "Konsumsi", "Active"],
    ["TYP02", "Peralatan", "Active"],
    ["TYP03", "Bahan", "Active"],
    ["TYP04", "Lainnya", "Active"]
  ]);

  // Daftar Unit/Outlet Berlaku (ApplicableUnit) bawaan: 3 outlet kasir POS
  // + General (berlaku di semua unit) + 4 gudang internal (Kolam/Teknisi/
  // Cleaning Service/Dapur). Bisa ditambah sendiri lewat menu Kategori &
  // Metode Pembayaran > Unit/Outlet Berlaku - lihat APPLICABLE_UNIT_HEADERS.
  ss.getSheetByName(CONFIG.SHEETS.APPLICABLE_UNITS).getRange(2, 1, 8, 4).setValues([
    ["APU01", "General", "Semua Outlet", "Active"],
    ["APU02", "Entrance", "Kasir Loket Masuk", "Active"],
    ["APU03", "Ride", "Kasir Loket Wahana", "Active"],
    ["APU04", "Resto", "Kasir Resto", "Active"],
    ["APU05", "Kolam", "Gudang Kolam Renang", "Active"],
    ["APU06", "Teknisi", "Gudang Teknisi", "Active"],
    ["APU07", "Cleaning", "Gudang Cleaning Service", "Active"],
    ["APU08", "Dapur", "Gudang Dapur", "Active"]
  ]);

  // 8 gudang: 4 gudang yang terhubung ke outlet kasir (Utama/Entrance/
  // Ride/Resto - dipakai juga oleh mesin stok penjualan lewat UNIT_MAP)
  // + 4 gudang internal (Kolam Renang/Teknisi/Dapur/Cleaning Service) yang
  // murni dipakai lewat modul Warehouse (tidak ada kasir/POS terhubung).
  // Kolom ke-4 (UnitType) memakai UnitKey Applicable_Units yang sama supaya
  // getWarehouseStockDetail() bisa memfilter produk tiap gudang internal
  // sesuai ApplicableUnit-nya sendiri (lihat WAREHOUSE_APPLICABLE_UNIT_).
  ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES).getRange(2, 1, 8, 4).setValues([
    ["WH01", "Gudang Utama", "Area Kantor Utama", "Central"],
    ["WH02", "Gudang Kantor dan Kasir Loket Masuk", "Gerbang Depan", "Entrance"],
    ["WH03", "Gudang Kasir Loket Wahana", "Area Wahana", "Ride"],
    ["WH04", "Gudang Kasir Resto", "Food Court", "Resto"],
    ["WH05", "Gudang Kolam Renang", "Area Kolam Renang", "Kolam"],
    ["WH06", "Gudang Teknisi", "Area Teknisi/Maintenance", "Teknisi"],
    ["WH07", "Gudang Cleaning Service", "Area Cleaning Service", "Cleaning"],
    ["WH08", "Gudang Dapur", "Area Dapur", "Dapur"]
  ]);

  // Kolom ItemType = klasifikasi barang untuk modul Warehouse (Konsumsi/
  // Peralatan/Bahan/Lainnya, terpisah dari Category penjualan/POS). Kolom
  // HargaFleksibel (Ya/Tidak) menandai produk yang harga jualnya boleh
  // diubah manual saat transaksi. Kolom terakhir UsesStock (Ya/Tidak)
  // menandai apakah produk memakai pelacakan stok gudang otomatis -
  // default "Ya"/"Tidak" untuk semua produk contoh ini.
  ss.getSheetByName(CONFIG.SHEETS.PRODUCTS).getRange(2, 1, 5, 12).setValues([
    ["PRD001", "Tiket Masuk Dewasa", "-", "Tiket Masuk", "Tiket", 30000, 50000, "Active", "Entrance", "Tidak", "Lainnya", "Ya"],
    ["PRD002", "Tiket Masuk Anak", "-", "Tiket Masuk", "Tiket", 20000, 35000, "Active", "Entrance", "Tidak", "Lainnya", "Ya"],
    ["PRD003", "Tiket Wahana Kolam Ombak", "-", "Tiket Wahana", "Tiket", 15000, 25000, "Active", "Ride", "Tidak", "Lainnya", "Ya"],
    ["PRD004", "Es Teh Manis", "Regular", "Makanan & Minuman", "Cup", 4000, 10000, "Active", "Resto,Ride", "Tidak", "Konsumsi", "Ya"],
    ["PRD005", "Nasi Goreng Spesial", "-", "Makanan & Minuman", "Porsi", 14000, 25000, "Active", "Resto", "Tidak", "Konsumsi", "Ya"]
  ]);

  // Seed stok awal di gudang pusat supaya modul stok bisa langsung diuji.
  // WarehouseName/ProductName/Status diisi manual di sini (khusus data
  // contoh) - selanjutnya kolom ini otomatis disinkronkan oleh
  // warehouseProductInfo_() setiap kali stok berubah lewat aplikasi.
  ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK).getRange(2, 1, 5, 9).setValues([
    ["STK001", "WH01", "Gudang Utama", "PRD001", "Tiket Masuk Dewasa", 500, 50, "Active", new Date()],
    ["STK002", "WH01", "Gudang Utama", "PRD002", "Tiket Masuk Anak", 500, 50, "Active", new Date()],
    ["STK003", "WH01", "Gudang Utama", "PRD003", "Tiket Wahana Kolam Ombak", 300, 30, "Active", new Date()],
    ["STK004", "WH01", "Gudang Utama", "PRD004", "Es Teh Manis", 200, 20, "Active", new Date()],
    ["STK005", "WH01", "Gudang Utama", "PRD005", "Nasi Goreng Spesial", 150, 15, "Active", new Date()]
  ]);
}

function hashPassword_(password) {
  let rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + "WonderlandSecret2026", Utilities.Charset.UTF_8);
  return rawHash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function genId_(prefix) {
  return prefix + Utilities.getUuid().substring(0, 8).toUpperCase();
}

// Buat sheet baru kalau belum ada (dipakai supaya database LAMA yang sudah
// terlanjur dipakai tetap otomatis mendapat sheet baru seperti Sessions,
// tanpa perlu reset/rebuild database dari nol).
function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// =====================================================================
// 4. AUTH - LOGIN, SESSION & PERMISSION MIDDLEWARE
// =====================================================================
function loginUser(username, password, clientIp, userAgent) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    let userRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) { userRow = data[i]; break; }
    }

    if (!userRow || userRow[5] !== hashPassword_(password)) {
      logLoginAttempt_(username, "FAILED - Invalid Credentials", clientIp, userAgent);
      return { success: false, message: "Username atau password salah." };
    }
    if (userRow[6] !== "Active") {
      logLoginAttempt_(username, "FAILED - Account Inactive", clientIp, userAgent);
      return { success: false, message: "Akun Anda dinonaktifkan. Hubungi Owner." };
    }

    let sessionToken = Utilities.getUuid();
    let cache = CacheService.getUserCache();
    let sessionData = {
      userId: userRow[0], username: userRow[1], fullName: userRow[2],
      role: userRow[3], assignedUnit: userRow[4], token: sessionToken
    };
    cache.put(sessionToken, JSON.stringify(sessionData), 21600);

    // Cadangkan sesi ke sheet supaya bertahan lebih lama dari batas cache
    // Google (maks 6 jam) dan tidak hilang saat halaman di-refresh.
    let sessionSheet = ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, SESSION_HEADERS);
    let now = new Date();
    let expiresAt = new Date(now.getTime() + SESSION_DURATION_MS);
    sessionSheet.appendRow([sessionToken, userRow[0], userRow[1], userRow[2], userRow[3], userRow[4], now, expiresAt]);

    logLoginAttempt_(username, "SUCCESS", clientIp, userAgent);
    logActivity_(userRow[1], "LOGIN", "Authentication", "User logged in successfully", clientIp);

    return {
      success: true, token: sessionToken,
      user: { userId: userRow[0], username: userRow[1], fullName: userRow[2], role: userRow[3], assignedUnit: userRow[4] }
    };
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan sistem: " + error.message };
  }
}

function verifySession(token) {
  if (!token) return null;
  let cache = CacheService.getUserCache();
  let cached = cache.get(token);
  if (cached) return JSON.parse(cached);

  // Fallback: cache sudah kedaluwarsa/kepakai ulang server -> cek sheet Sessions.
  // Ini yang membuat sesi TIDAK hilang hanya karena refresh halaman.
  try {
    const ss = getDatabase();
    const sheet = ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, SESSION_HEADERS);
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === token) {
        let expiresAt = new Date(data[i][7]);
        if (isNaN(expiresAt.getTime()) || expiresAt.getTime() < Date.now()) {
          sheet.deleteRow(i + 1);
          return null;
        }
        let sessionData = {
          userId: data[i][1], username: data[i][2], fullName: data[i][3],
          role: data[i][4], assignedUnit: data[i][5], token: token
        };
        // sliding expiration: perpanjang masa berlaku selama masih dipakai
        cache.put(token, JSON.stringify(sessionData), 21600);
        sheet.getRange(i + 1, 8).setValue(new Date(Date.now() + SESSION_DURATION_MS));
        return sessionData;
      }
    }
  } catch (e) {}
  return null;
}

function logoutUser(token) {
  try {
    if (!token) return { success: true };
    CacheService.getUserCache().remove(token);
    const ss = getDatabase();
    const sheet = ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, SESSION_HEADERS);
    let data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === token) sheet.deleteRow(i + 1);
    }
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function checkPermission_(token, allowedRoles, unit) {
  let user = verifySession(token);
  if (!user) throw new Error("Sesi tidak valid atau telah berakhir. Silakan login kembali.");
  if (allowedRoles && allowedRoles.indexOf(user.role) === -1) {
    throw new Error("Anda tidak memiliki akses untuk melakukan aksi ini.");
  }
  if (unit && user.role === ROLE.KASIR && user.assignedUnit !== unit) {
    throw new Error("Anda hanya memiliki akses ke unit " + user.assignedUnit + ".");
  }
  return user;
}

function logLoginAttempt_(username, status, ip, browser) {
  try {
    const ss = getDatabase();
    ss.getSheetByName(CONFIG.SHEETS.LOGIN_LOG).appendRow([genId_("LOG"), new Date(), username, status, ip || "Unknown", browser || "Unknown"]);
  } catch (e) {}
}

function logActivity_(user, action, module, details, ip) {
  try {
    const ss = getDatabase();
    ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG).appendRow([genId_("ACT"), new Date(), user, action, module, details, ip || "System"]);
  } catch (e) {}
}

// =====================================================================
// 5. DASHBOARD
// =====================================================================
// scope: opsional, dikirim dari client. "All" = paksa tampilkan transaksi
// terbaru dari SEMUA outlet (dipakai tombol filter khusus di dashboard);
// selain itu, Kasir hanya melihat transaksi terbaru outlet yang ditugaskan
// padanya saja (Owner/Admin/Staff selalu melihat semua outlet).
// startDate/endDate: opsional, format "yyyy-MM-dd". Kalau kosong, dipakai
// tanggal hari ini saja (perilaku lama/default). Dipakai oleh toggle
// filter tanggal di dashboard (Hari Ini/Kemarin/Minggu Ini/Bulan Ini/
// Tahun Ini/rentang custom).
function getDashboardData(token, scope, startDate, endDate) {
  try {
    let user = checkPermission_(token, null);
    const ss = getDatabase();
    let tz = Session.getScriptTimeZone();
    let todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    let rangeStart = startDate || todayStr;
    let rangeEnd = endDate || todayStr;
    if (rangeStart > rangeEnd) { let t = rangeStart; rangeStart = rangeEnd; rangeEnd = t; }

    let entranceSales = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.ENTRANCE_SALES));
    let rideSales = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.RIDE_SALES));
    let restoSales = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.RESTO_SALES));

    let entrancePay = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.ENTRANCE_PAYMENT));
    let ridePay = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.RIDE_PAYMENT));
    let restoPay = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.RESTO_PAYMENT));

    function inRange_(dateVal) {
      if (!dateVal) return false;
      let d = Utilities.formatDate(new Date(dateVal), tz, "yyyy-MM-dd");
      return d >= rangeStart && d <= rangeEnd;
    }

    function sumRange_(rows) {
      let t = 0;
      rows.forEach(row => { if (inRange_(row.Date)) t += Number(row.TotalPrice) || 0; });
      return t;
    }

    // Rincian metode pembayaran (tunai, QRIS, transfer, debit, kredit) per
    // outlet - dipakai widget dashboard per kasir (Loket Masuk/Wahana/Resto).
    function paymentBreakdown_(rows) {
      let b = { total: 0, cash: 0, qris: 0, transfer: 0, debit: 0, kredit: 0 };
      rows.forEach(p => {
        if (inRange_(p.Date)) {
          let amt = Number(p.Amount) || 0;
          let method = (p.PaymentMethod || "").toLowerCase();
          b.total += amt;
          if (method.includes("debit")) b.debit += amt;
          else if (method.includes("kredit") || method.includes("credit")) b.kredit += amt;
          else if (method.includes("cash") || method.includes("tunai")) b.cash += amt;
          else if (method.includes("qris")) b.qris += amt;
          else if (method.includes("transfer")) b.transfer += amt;
        }
      });
      return b;
    }

    function emptyBreakdown_() { return { total: 0, cash: 0, qris: 0, transfer: 0, debit: 0, kredit: 0 }; }

    // Kasir (selain Owner/Admin/Staff) secara default hanya melihat
    // ANGKA (total penjualan, tunai, QRIS, transfer, debit, kredit) dari
    // outlet-nya sendiri saja (mis. kasir resto tidak perlu melihat angka
    // gabungan kasir loket masuk / wahana). scope="All" membuka tampilan
    // gabungan semua outlet lewat filter di dashboard. Owner/Admin/Staff
    // selalu melihat angka gabungan semua outlet.
    let unitSalesMap = { Entrance: entranceSales, Ride: rideSales, Resto: restoSales };
    let visibleUnits = ["Entrance", "Ride", "Resto"];
    let isRestrictedByDefault = (user.role === ROLE.KASIR);
    if (isRestrictedByDefault && scope !== "All") {
      visibleUnits = [user.assignedUnit];
    }

    let totalMasuk = visibleUnits.indexOf("Entrance") !== -1 ? sumRange_(entranceSales) : 0;
    let totalWahana = visibleUnits.indexOf("Ride") !== -1 ? sumRange_(rideSales) : 0;
    let totalResto = visibleUnits.indexOf("Resto") !== -1 ? sumRange_(restoSales) : 0;
    let totalSalesToday = totalMasuk + totalWahana + totalResto;

    let payMasuk = visibleUnits.indexOf("Entrance") !== -1 ? paymentBreakdown_(entrancePay) : emptyBreakdown_();
    let payWahana = visibleUnits.indexOf("Ride") !== -1 ? paymentBreakdown_(ridePay) : emptyBreakdown_();
    let payResto = visibleUnits.indexOf("Resto") !== -1 ? paymentBreakdown_(restoPay) : emptyBreakdown_();

    let totalCash = payMasuk.cash + payWahana.cash + payResto.cash;
    let totalQris = payMasuk.qris + payWahana.qris + payResto.qris;
    let totalTransfer = payMasuk.transfer + payWahana.transfer + payResto.transfer;
    let totalDebit = payMasuk.debit + payWahana.debit + payResto.debit;
    let totalKredit = payMasuk.kredit + payWahana.kredit + payResto.kredit;

    let stockData = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK));
    let lowStockItems = stockData.filter(s => Number(s.Quantity) <= Number(s.MinStock));

    let allSales = [];
    visibleUnits.forEach(u => { allSales = allSales.concat(unitSalesMap[u] || []); });
    // "Transaksi Terbaru" HARUS ikut rentang tanggal yang sedang difilter -
    // supaya begitu tanggal berganti (mis. sudah lewat tengah malam),
    // transaksi kemarin otomatis tidak tampil lagi di dashboard.
    let recentSales = allSales.filter(r => inRange_(r.Date));
    recentSales.sort((a, b) => new Date(b.InputTime || b.Date) - new Date(a.InputTime || a.Date));

    // Widget Dashboard "Grafik Penjualan per Kategori" & "Grafik Produk
    // Terlaris" - dihitung dari SELURUH penjualan yang sesuai rentang
    // tanggal & outlet yang sedang ditampilkan (recentSales, sebelum
    // dipotong ke 10 baris untuk tabel Transaksi Terbaru). Tersedia untuk
    // semua role (termasuk Kasir, otomatis mengikuti scope outlet-nya).
    let categoryAgg_ = {}, productAgg_ = {};
    recentSales.forEach(r => {
      let cat = (r.Category || "").toString().trim() || "Tanpa Kategori";
      categoryAgg_[cat] = (categoryAgg_[cat] || 0) + (Number(r.TotalPrice) || 0);
      let prod = (r.ProductName || "").toString().trim() || "-";
      productAgg_[prod] = (productAgg_[prod] || 0) + (Number(r.TotalPrice) || 0);
    });
    let categoryEntries_ = Object.keys(categoryAgg_).map(k => [k, categoryAgg_[k]]).sort((a, b) => b[1] - a[1]);
    let topProductEntries_ = Object.keys(productAgg_).map(k => [k, productAgg_[k]]).sort((a, b) => b[1] - a[1]).slice(0, 8);
    let categoryChart = { labels: categoryEntries_.map(e => e[0]), values: categoryEntries_.map(e => e[1]) };
    let topProductsChart = { labels: topProductEntries_.map(e => e[0]), values: topProductEntries_.map(e => e[1]) };

    // Widget Keuangan (Total Pengeluaran & Tagihan) + grafik - hanya
    // dihitung untuk Owner/Admin/Staff (Kasir tidak perlu lihat data
    // keuangan gabungan perusahaan).
    let finance = null;
    let chart = null;
    if (user.role === ROLE.OWNER || user.role === ROLE.ADMIN || user.role === ROLE.STAFF) {
      let allExpenses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES));
      allExpenses.forEach(e => { e.Date = normalizeDateStr_(e.Date); });
      let rangeExpenses = allExpenses.filter(e => e.Date && e.Date >= rangeStart && e.Date <= rangeEnd);
      function expSum_(cat) { return rangeExpenses.filter(e => e.Category === cat).reduce((t, e) => t + (Number(e.Amount) || 0), 0); }
      let pengeluaranOmset = expSum_(FINANCE_EXPENSE_CATEGORY.OMSET);
      let pengeluaranBrangkas = expSum_(FINANCE_EXPENSE_CATEGORY.BRANGKAS);
      let pengeluaranRekening = expSum_(FINANCE_EXPENSE_CATEGORY.REKENING);
      let totalPengeluaran = pengeluaranOmset + pengeluaranBrangkas + pengeluaranRekening;

      // Tagihan/hutang belum lunas: ditampilkan APA ADANYA (semua yang
      // berstatus BelumLunas, tanpa batas tanggal) supaya Owner selalu
      // tahu total tagihan yang masih outstanding, bukan cuma yang ada
      // di rentang tanggal filter dashboard saat ini.
      let unpaidRows = allExpenses.filter(e => e.Category === FINANCE_EXPENSE_CATEGORY.BELUM_LUNAS);
      let totalTagihanBelumLunas = unpaidRows.reduce((t, e) => t + (Number(e.Amount) || 0), 0);

      finance = {
        totalPengeluaran: totalPengeluaran,
        pengeluaranOmset: pengeluaranOmset, pengeluaranBrangkas: pengeluaranBrangkas, pengeluaranRekening: pengeluaranRekening,
        totalTagihanBelumLunas: totalTagihanBelumLunas, tagihanCount: unpaidRows.length,
        tagihanList: unpaidRows.sort((a, b) => b.Date < a.Date ? -1 : (b.Date > a.Date ? 1 : 0)).slice(0, 8)
      };

      // Data grafik harian (Pendapatan vs Pengeluaran) untuk rentang yang
      // dipilih. Dibatasi maksimal 92 hari (~3 bulan) supaya perhitungan
      // tetap ringan; kalau rentang lebih panjang, tetap dihitung tapi
      // dibatasi ke 92 hari terakhir dalam rentang tsb.
      let d0 = new Date(rangeStart + "T00:00:00");
      let d1 = new Date(rangeEnd + "T00:00:00");
      let dayCount = Math.min(92, Math.round((d1 - d0) / 86400000) + 1);
      let dayLabels = [];
      for (let i = dayCount - 1; i >= 0; i--) {
        let d = new Date(d1.getTime() - i * 86400000);
        dayLabels.push(Utilities.formatDate(d, tz, "yyyy-MM-dd"));
      }
      let salesByDay = {}, expByDay = {};
      dayLabels.forEach(d => { salesByDay[d] = 0; expByDay[d] = 0; });
      allSales.forEach(r => {
        if (!r.Date) return;
        let d = Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd");
        if (salesByDay.hasOwnProperty(d)) salesByDay[d] += Number(r.TotalPrice) || 0;
      });
      rangeExpenses.forEach(e => {
        if (e.Category === FINANCE_EXPENSE_CATEGORY.BELUM_LUNAS) return; // belum keluar kas, tidak masuk grafik pengeluaran
        let d = e.Date;
        if (expByDay.hasOwnProperty(d)) expByDay[d] += Number(e.Amount) || 0;
      });
      chart = {
        labels: dayLabels,
        sales: dayLabels.map(d => salesByDay[d]),
        expenses: dayLabels.map(d => expByDay[d]),
        expenseBreakdown: { Omset: pengeluaranOmset, Brangkas: pengeluaranBrangkas, RekeningPT: pengeluaranRekening }
      };
    }

    // Total Penerimaan Kasir = jumlah SELURUH uang yang benar-benar diterima
    // kasir (Tunai+QRIS+Transfer+Debit+Kredit) dari 3 outlet - dipakai widget
    // dashboard "Total Penerimaan Kasir" (terpisah dari "Total Penjualan"
    // yang dihitung dari data Penjualan/Sales, bukan dari data Penerimaan).
    let totalPenerimaanKasir = totalCash + totalQris + totalTransfer + totalDebit + totalKredit;

    return {
      success: true,
      rangeStart: rangeStart, rangeEnd: rangeEnd,
      metrics: { totalSalesToday, totalCash, totalQris, totalTransfer, totalDebit, totalKredit, totalMasuk, totalWahana, totalResto, totalPenerimaanKasir },
      paymentBreakdown: { Entrance: payMasuk, Ride: payWahana, Resto: payResto },
      lowStockCount: lowStockItems.length,
      lowStockItems: lowStockItems.slice(0, 5),
      recentTransactions: recentSales.slice(0, 10),
      canFilterAllOutlets: isRestrictedByDefault,
      currentScope: (isRestrictedByDefault && scope !== "All") ? user.assignedUnit : "All",
      finance: finance,
      chart: chart,
      categoryChart: categoryChart,
      topProductsChart: topProductsChart
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// PERBAIKAN PENTING (v3.0.1): google.script.run TIDAK BISA mengirim objek
// Date mentah balik ke browser - kalau tetap dikirim, bukan error yang
// muncul, tapi browser diam-diam menerima `null` di withSuccessHandler
// (ini perilaku resmi Apps Script, lihat dokumentasi HTML Service). Inilah
// akar dari bug "data sudah tersimpan di spreadsheet tapi tidak tampil di
// web app" yang sempat terjadi - kolom seperti Date/InputTime/CreatedAt
// terbaca sebagai objek Date asli lewat getValues(), lalu ikut dikirim
// mentah ke client dan membuat SEMUA data di halaman itu gagal tampil
// tanpa pesan error yang jelas.
//
// Solusinya: setiap nilai bertipe Date dikonversi ke string ISO di sini,
// SATU tempat, sehingga otomatis berlaku untuk semua fungsi yang memakai
// helper ini (getDashboardData, getSalesData, getPaymentData, getStockView,
// getUsers, dll). Sisi client (fmtDate(), new Date(...)) tetap bisa membaca
// string ISO ini dengan normal, jadi tidak perlu ada perubahan di Index.html.
function getSheetDataAsObjects_(sheet) {
  if (!sheet) return [];
  let rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  let headers = rows[0];
  let result = [];
  for (let i = 1; i < rows.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = rows[i][j];
      // PENTING - kolom yang seharusnya angka (mis. Quantity) kadang tidak
      // sengaja ikut ter-format sebagai Tanggal di Google Sheets (misalnya
      // gara-gara fill-handle/format painter menular dari kolom tanggal di
      // sebelahnya). Kalau itu terjadi, Apps Script mengembalikan nilainya
      // sebagai objek Date, dan kalau Date itu somehow tidak valid,
      // toISOString() akan melempar "Invalid time value" dan bikin SELURUH
      // request gagal tanpa pesan yang jelas. Ditangkap di sini supaya satu
      // sel bermasalah tidak bikin seluruh halaman error - nilai mentahnya
      // tetap dikembalikan (masalah format sebenarnya ditangani lewat
      // Diagnostic Database di menu Pengaturan).
      if (val instanceof Date) {
        try { val = val.toISOString(); } catch (e) { /* Invalid Date - biarkan objek Date apa adanya, jangan sampai melempar error */ }
      }
      obj[headers[j]] = val;
    }
    obj._row = i + 1;
    result.push(obj);
  }
  return result;
}

// PENTING - AKAR MASALAH TANGGAL: kolom "Date" di Google Sheets otomatis
// dikonversi jadi tipe Date murni oleh Sheets begitu diisi string
// "yyyy-MM-dd" (mis. lewat saveFinanceExpense/saveFinanceBalance). Saat
// dibaca lagi lewat getSheetDataAsObjects_(), nilai Date itu diubah jadi
// string ISO UTC (val.toISOString(), mis. "2026-08-05T16:00:00.000Z" untuk
// tanggal lokal 6 Agustus WIB) - kalau ini dibandingkan pakai "===" atau
// ">="/"<=" langsung terhadap string polos "yyyy-MM-dd" dari client, hasil
// perbandingannya SELALU SALAH (beda tipe & bisa geser 1 hari akibat
// offset UTC). normalizeDateStr_() mengembalikan representasi "yyyy-MM-dd"
// yang benar & konsisten (memakai timezone project) untuk SEGALA bentuk
// input tanggal (Date asli, string ISO hasil getSheetDataAsObjects_,
// ataupun string "yyyy-MM-dd" polos) - dipakai di SEMUA perbandingan
// tanggal modul Admin Keuangan supaya data yang diinput di web app selalu
// konsisten dengan yang tersimpan & tampil kembali di web app/report.
function normalizeDateStr_(val) {
  if (!val) return "";
  let tz = Session.getScriptTimeZone();
  if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy-MM-dd");
  let s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) { // string ISO datetime (dari getSheetDataAsObjects_)
    return Utilities.formatDate(new Date(s), tz, "yyyy-MM-dd");
  }
  return s.slice(0, 10); // sudah berupa "yyyy-MM-dd" polos
}

// =====================================================================
// 6. PRODUCTS - MASTER PRODUK
// =====================================================================
// ApplicableUnit disimpan sebagai string, bisa berisi lebih dari 1 outlet
// dipisah koma, contoh: "Resto,Ride" artinya produk itu dijual di kasir
// Resto DAN kasir Loket Wahana sekaligus. "General" = tampil di semua outlet.
function productUnits_(applicableUnit) {
  return String(applicableUnit || "").split(",").map(u => u.trim()).filter(Boolean);
}

// unit: dipakai oleh kasir (dropdown input penjualan) -> hanya produk Active
// dan yang memang ditugaskan ke outlet tsb (atau General) yang ditampilkan.
function getProducts(unit) {
  const ss = getDatabase();
  let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
  products = products.filter(p => p.Status === "Active");
  if (unit) {
    products = products.filter(p => {
      let units = productUnits_(p.ApplicableUnit);
      return units.indexOf(unit) !== -1 || units.indexOf("General") !== -1;
    });
  }
  return products;
}

function getProductsWithStock(token) {
  try {
    checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let stockRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK)).filter(s => s.WarehouseID === CENTRAL_WAREHOUSE_ID);
    let stockMap = {};
    stockRows.forEach(s => stockMap[s.ProductCode] = s.Quantity);
    // Produk dengan Acuan Stok (StockRefCode) tidak pernah punya baris
    // sendiri di Warehouse_Stock - qty yang ditampilkan di Master Produk
    // untuk produk ini HARUS ikut qty produk acuannya (bukan 0), supaya
    // Owner/Admin tidak salah kira produk itu kehabisan stok.
    let productMap = {};
    products.forEach(p => productMap[p.ProductCode] = p);
    function resolveLocal_(code) {
      let visited = {};
      while (productMap[code] && String(productMap[code].StockRefCode || "").trim() && productMap[code].StockRefCode !== code) {
        if (visited[code]) break;
        visited[code] = true;
        code = String(productMap[code].StockRefCode).trim();
      }
      return code;
    }
    let result = products.map(p => {
      let resolvedCode = resolveLocal_(p.ProductCode);
      p.StokAwal = stockMap[resolvedCode] !== undefined ? stockMap[resolvedCode] : 0;
      return p;
    });
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Peta nama header -> nomor kolom (1-based) untuk sheet Products. Dipakai
// supaya saveProduct()/bulkImportProducts() menulis tiap field ke KOLOM
// YANG BENAR berdasarkan judul header di baris 1, bukan asumsi posisi
// tetap - ini mencegah data "geser" kalau ada kolom tambahan yang pernah
// disisipkan manual di spreadsheet (mis. kolom HargaFleksibel).
function productHeaderIndex_(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  let idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i + 1; });
  return idx;
}

function saveProduct(product, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const idx = productHeaderIndex_(sheet);

    // UsesStock: "Ya" (default, produk memakai pelacakan stok gudang seperti
    // biasa) atau "Tidak" (khusus produk tanpa fisik/tanpa stok - stok akhir
    // di laporan penjualan akan selalu 0 & stok gudang tidak pernah disentuh,
    // lihat saveSalesTransaction/deleteSalesTransaction).
    let usesStock = product.usesStock === "Tidak" ? "Tidak" : "Ya";
    // UsesSerial: "Ya" (default, produk menampilkan field Nomor Seri
    // Awal/Akhir di Form Input Penjualan kalau outletnya memang punya
    // kolom seri) atau "Tidak" (produk ini tidak pakai nomor seri, field
    // seri tidak ditampilkan sama sekali walau outletnya hasSerial).
    let usesSerial = product.usesSerial === "Tidak" ? "Tidak" : "Ya";
    // HargaFleksibel: "Ya" kalau harga jual produk ini boleh diubah manual
    // saat transaksi (mis. barang re-sell/konsinyasi), default "Tidak".
    let hargaFleksibel = product.hargaFleksibel === "Ya" ? "Ya" : "Tidak";
    let code = product.isEdit ? product.productCode : (product.productCode || genId_("PRD"));

    // StockRefCode ("Acuan Stok"): kode produk LAIN yang stok fisiknya
    // dipakai bersama - dipakai untuk kasus mis. TERP002 "Tiket Terusan
    // Plus Diskon 20%" yang tidak punya stok sendiri, tapi setiap kali
    // terjual harus memotong stok TER001 "Tiket Terusan Plus" (lihat
    // resolveStockProductCode_). Hanya 1 tingkat referensi yang didukung -
    // produk acuan itu sendiri harus punya stok berdiri sendiri.
    let stockRefCode = String(product.stockRefCode || "").trim();
    if (stockRefCode) {
      if (stockRefCode === code) {
        return { success: false, message: "Kode Barang Acuan Stok tidak boleh sama dengan Kode Barang produk itu sendiri." };
      }
      let allProductsForRef = getSheetDataAsObjects_(sheet);
      let refProduct = allProductsForRef.find(p => p.ProductCode === stockRefCode);
      if (!refProduct) {
        return { success: false, message: "Kode Barang Acuan Stok \"" + stockRefCode + "\" tidak ditemukan di Master Produk." };
      }
      if (String(refProduct.StockRefCode || "").trim()) {
        return { success: false, message: "Produk acuan (\"" + stockRefCode + "\") juga memakai Acuan Stok ke produk lain - rantai bertingkat tidak didukung. Pilih produk acuan yang stoknya berdiri sendiri." };
      }
      // Produk yang berbagi stok tetap dianggap "Pakai Stok" (stoknya
      // memang dipotong, hanya saja lewat produk acuan).
      usesStock = "Ya";
    }

    if (product.isEdit) {
      let data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === product.productCode) {
          let rowNum = i + 1;
          let fields = {
            ProductName: product.productName, Variant: product.variant || "-", Category: product.category,
            Unit: product.unit, BuyPrice: Number(product.buyPrice) || 0, SellPrice: Number(product.sellPrice) || 0,
            Status: product.status, ApplicableUnit: product.applicableUnit,
            HargaFleksibel: hargaFleksibel, ItemType: product.itemType || "Lainnya", UsesStock: usesStock,
            UsesSerial: usesSerial, StockRefCode: stockRefCode
          };
          Object.keys(fields).forEach(key => { if (idx[key]) sheet.getRange(rowNum, idx[key]).setValue(fields[key]); });
          logActivity_(user.username, "EDIT_PRODUCT", "Master Produk", "Edit produk " + product.productCode, "");
          return { success: true, message: "Produk berhasil diperbarui." };
        }
      }
      return { success: false, message: "Produk tidak ditemukan." };
    } else {
      let data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === code) return { success: false, message: "Kode Barang sudah digunakan: " + code };
      }
      let rowArr = new Array(sheet.getLastColumn()).fill("");
      let fields = {
        ProductCode: code, ProductName: product.productName, Variant: product.variant || "-", Category: product.category,
        Unit: product.unit, BuyPrice: Number(product.buyPrice) || 0, SellPrice: Number(product.sellPrice) || 0,
        Status: product.status || "Active", ApplicableUnit: product.applicableUnit,
        HargaFleksibel: hargaFleksibel, ItemType: product.itemType || "Lainnya", UsesStock: usesStock,
        UsesSerial: usesSerial, StockRefCode: stockRefCode
      };
      Object.keys(fields).forEach(key => { if (idx[key]) rowArr[idx[key] - 1] = fields[key]; });
      sheet.appendRow(rowArr);

      // Stok Awal hanya relevan untuk produk yang stoknya berdiri sendiri.
      // Produk dengan Acuan Stok (stockRefCode) ikut stok yang SUDAH ADA di
      // produk acuannya - tidak punya Stok Awal terpisah sendiri.
      let initialStock = Number(product.initialStock) || 0;
      if (usesStock === "Ya" && !stockRefCode && initialStock > 0) updateStock_(ss, CENTRAL_WAREHOUSE_ID, code, initialStock, 10, { type: STOCK_CARD_TYPE.OPENING, user: user.fullName, notes: "Stok awal saat produk didaftarkan" });

      logActivity_(user.username, "CREATE_PRODUCT", "Master Produk", "Tambah produk " + code, "");
      return { success: true, message: "Produk berhasil ditambahkan." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteProduct(productCode, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productCode) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_PRODUCT", "Master Produk", "Hapus produk " + productCode, "");
        return { success: true, message: "Produk berhasil dihapus." };
      }
    }
    return { success: false, message: "Produk tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Hapus banyak produk sekaligus (dipakai fitur cengtang semua + hapus terpilih)
function bulkDeleteProducts(productCodes, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    if (!productCodes || !productCodes.length) return { success: false, message: "Tidak ada produk yang dipilih." };
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    let codesSet = {};
    productCodes.forEach(c => codesSet[c] = true);
    let data = sheet.getDataRange().getValues();
    let deleted = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (codesSet[data[i][0]]) { sheet.deleteRow(i + 1); deleted++; }
    }
    logActivity_(user.username, "BULK_DELETE_PRODUCT", "Master Produk", "Hapus " + deleted + " produk sekaligus", "");
    return { success: true, message: deleted + " produk berhasil dihapus.", deleted: deleted };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function bulkImportProducts(products, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const idx = productHeaderIndex_(sheet);
    let data = sheet.getDataRange().getValues();
    let existingRow = {};
    for (let i = 1; i < data.length; i++) existingRow[data[i][0]] = i + 1;

    let inserted = 0, updated = 0, errors = [];
    products.forEach((p, idx2) => {
      try {
        if (!p.productCode || !p.productName) { errors.push("Baris " + (idx2 + 2) + ": Kode Barang / Nama Produk kosong, dilewati."); return; }
        let usesStock = String(p.usesStock || "Ya").trim() === "Tidak" ? "Tidak" : "Ya";
        let usesSerial = String(p.usesSerial || "Ya").trim() === "Tidak" ? "Tidak" : "Ya";
        // StockRefCode via import: tidak divalidasi seketat saveProduct
        // (mis. rantai bertingkat) supaya proses import massal tidak
        // gagal di tengah jalan - cukup dipaksa UsesStock="Ya" kalau diisi,
        // karena stoknya memang ikut dipotong lewat produk acuan.
        let stockRefCode = String(p.stockRefCode || "").trim();
        if (stockRefCode) usesStock = "Ya";
        let fields = {
          ProductCode: p.productCode, ProductName: p.productName, Variant: p.variant || "-", Category: p.category || "-",
          Unit: p.unit || "Pcs", BuyPrice: Number(p.buyPrice) || 0, SellPrice: Number(p.sellPrice) || 0,
          Status: p.status || "Active", ApplicableUnit: p.applicableUnit || "General",
          HargaFleksibel: p.hargaFleksibel === "Ya" ? "Ya" : "Tidak", ItemType: p.itemType || "Lainnya", UsesStock: usesStock,
          UsesSerial: usesSerial, StockRefCode: stockRefCode
        };

        if (existingRow[p.productCode]) {
          let rowNum = existingRow[p.productCode];
          Object.keys(fields).forEach(key => { if (key !== "ProductCode" && idx[key]) sheet.getRange(rowNum, idx[key]).setValue(fields[key]); });
          updated++;
        } else {
          let rowArr = new Array(sheet.getLastColumn()).fill("");
          Object.keys(fields).forEach(key => { if (idx[key]) rowArr[idx[key] - 1] = fields[key]; });
          sheet.appendRow(rowArr);
          existingRow[p.productCode] = sheet.getLastRow();
          inserted++;
          let initialStock = Number(p.initialStock) || 0;
          if (usesStock === "Ya" && !stockRefCode && initialStock > 0) updateStock_(ss, CENTRAL_WAREHOUSE_ID, p.productCode, initialStock, 10, { type: STOCK_CARD_TYPE.OPENING, user: user.fullName, notes: "Stok awal (import produk)" });
        }
      } catch (rowErr) {
        errors.push("Baris " + (idx2 + 2) + ": " + rowErr.message);
      }
    });

    logActivity_(user.username, "IMPORT_PRODUCTS", "Master Produk", "Import " + inserted + " baru, " + updated + " update", "");
    return { success: true, inserted: inserted, updated: updated, errors: errors };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 6b. CATEGORIES & PAYMENT METHODS - MASTER DATA YANG BISA DIATUR SENDIRI
// =====================================================================
function getCategories() {
  const ss = getDatabase();
  return getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.CATEGORIES));
}

function saveCategory(data, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CATEGORIES);
    if (data.isEdit) {
      let rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.categoryId) {
          sheet.getRange(i + 1, 2, 1, 3).setValues([[data.categoryName, data.type, data.status]]);
          logActivity_(user.username, "EDIT_CATEGORY", "Master Kategori", "Edit kategori " + data.categoryName, "");
          return { success: true, message: "Kategori berhasil diperbarui." };
        }
      }
      return { success: false, message: "Kategori tidak ditemukan." };
    } else {
      sheet.appendRow([genId_("CAT"), data.categoryName, data.type || "General", data.status || "Active"]);
      logActivity_(user.username, "CREATE_CATEGORY", "Master Kategori", "Tambah kategori " + data.categoryName, "");
      return { success: true, message: "Kategori berhasil ditambahkan." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteCategory(categoryId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CATEGORIES);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === categoryId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_CATEGORY", "Master Kategori", "Hapus kategori " + categoryId, "");
        return { success: true, message: "Kategori berhasil dihapus." };
      }
    }
    return { success: false, message: "Kategori tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getPaymentMethods() {
  const ss = getDatabase();
  return getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PAYMENT_METHODS));
}

function savePaymentMethod(data, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_METHODS);
    if (data.isEdit) {
      let rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.methodId) {
          sheet.getRange(i + 1, 2, 1, 2).setValues([[data.methodName, data.status]]);
          logActivity_(user.username, "EDIT_PAYMENT_METHOD", "Master Metode Pembayaran", "Edit metode " + data.methodName, "");
          return { success: true, message: "Metode pembayaran berhasil diperbarui." };
        }
      }
      return { success: false, message: "Metode pembayaran tidak ditemukan." };
    } else {
      sheet.appendRow([genId_("PM"), data.methodName, data.status || "Active"]);
      logActivity_(user.username, "CREATE_PAYMENT_METHOD", "Master Metode Pembayaran", "Tambah metode " + data.methodName, "");
      return { success: true, message: "Metode pembayaran berhasil ditambahkan." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deletePaymentMethod(methodId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_METHODS);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === methodId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_PAYMENT_METHOD", "Master Metode Pembayaran", "Hapus metode " + methodId, "");
        return { success: true, message: "Metode pembayaran berhasil dihapus." };
      }
    }
    return { success: false, message: "Metode pembayaran tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Tipe Produk (dulu WAREHOUSE_ITEM_TYPES yang hardcode 4 pilihan tetap) -
// sekarang bisa diatur sendiri (tambah/edit/hapus) sama seperti Kategori &
// Metode Pembayaran, disimpan di sheet Product_Types sendiri.
function getProductTypes() {
  const ss = getDatabase();
  return getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCT_TYPES));
}

function saveProductType(data, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCT_TYPES);
    if (data.isEdit) {
      let rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.typeId) {
          sheet.getRange(i + 1, 2, 1, 2).setValues([[data.typeName, data.status]]);
          logActivity_(user.username, "EDIT_PRODUCT_TYPE", "Master Tipe Produk", "Edit tipe produk " + data.typeName, "");
          return { success: true, message: "Tipe produk berhasil diperbarui." };
        }
      }
      return { success: false, message: "Tipe produk tidak ditemukan." };
    } else {
      sheet.appendRow([genId_("TYP"), data.typeName, data.status || "Active"]);
      logActivity_(user.username, "CREATE_PRODUCT_TYPE", "Master Tipe Produk", "Tambah tipe produk " + data.typeName, "");
      return { success: true, message: "Tipe produk berhasil ditambahkan." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteProductType(typeId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCT_TYPES);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === typeId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_PRODUCT_TYPE", "Master Tipe Produk", "Hapus tipe produk " + typeId, "");
        return { success: true, message: "Tipe produk berhasil dihapus." };
      }
    }
    return { success: false, message: "Tipe produk tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Unit/Outlet Berlaku (ApplicableUnit) - dulu 4 pilihan hardcode (General/
// Entrance/Ride/Resto) di form Master Produk, sekarang master data sendiri
// (bisa tambah/edit/hapus lewat menu Kategori & Metode Pembayaran) supaya
// Owner bisa menambah unit baru (mis. Teknisi/Kolam/Dapur/Cleaning) tanpa
// perlu ubah kode. Dipakai untuk mengisi checkbox "Outlet Penjualan" di
// Master Produk & dropdown "Outlet Berlaku" di Kategori Produk, dan untuk
// memfilter produk yang tampil di tiap Gudang Internal (lihat
// applicableUnitByWarehouseId_/getWarehouseStockDetail).
function getApplicableUnits() {
  const ss = getDatabase();
  return getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.APPLICABLE_UNITS));
}

function saveApplicableUnit(data, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.APPLICABLE_UNITS);
    let unitKey = String(data.unitKey || "").trim();
    if (!unitKey) return { success: false, message: "Kode Unit wajib diisi." };
    if (/[,\s]/.test(unitKey)) return { success: false, message: "Kode Unit tidak boleh mengandung spasi atau koma." };
    let rows = sheet.getDataRange().getValues();

    if (data.isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.unitId) {
          let oldKey = rows[i][1];
          // Cek duplikat UnitKey (selain baris ini sendiri).
          for (let j = 1; j < rows.length; j++) {
            if (j !== i && String(rows[j][1]).trim() === unitKey) {
              return { success: false, message: "Kode Unit \"" + unitKey + "\" sudah dipakai." };
            }
          }
          if (oldKey === GENERAL_APPLICABLE_UNIT_KEY_ && unitKey !== GENERAL_APPLICABLE_UNIT_KEY_) {
            return { success: false, message: "Kode Unit \"General\" tidak boleh diubah karena dipakai sebagai acuan \"berlaku di semua unit\" di seluruh sistem." };
          }
          sheet.getRange(i + 1, 2, 1, 3).setValues([[unitKey, data.unitLabel || unitKey, data.status || "Active"]]);
          logActivity_(user.username, "EDIT_APPLICABLE_UNIT", "Master Unit Berlaku", "Edit unit " + unitKey, "");
          return { success: true, message: "Unit berhasil diperbarui." };
        }
      }
      return { success: false, message: "Unit tidak ditemukan." };
    } else {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === unitKey) return { success: false, message: "Kode Unit \"" + unitKey + "\" sudah dipakai." };
      }
      sheet.appendRow([genId_("APU"), unitKey, data.unitLabel || unitKey, data.status || "Active"]);
      logActivity_(user.username, "CREATE_APPLICABLE_UNIT", "Master Unit Berlaku", "Tambah unit " + unitKey, "");
      return { success: true, message: "Unit berhasil ditambahkan." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteApplicableUnit(unitId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.APPLICABLE_UNITS);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === unitId) {
        let unitKey = rows[i][1];
        if (unitKey === GENERAL_APPLICABLE_UNIT_KEY_) {
          return { success: false, message: "Unit \"General\" tidak boleh dihapus karena dipakai sebagai acuan \"berlaku di semua unit\" di seluruh sistem." };
        }
        // Cek dulu apakah masih ada produk yang memakai unit ini, supaya
        // tidak diam-diam membuat produk jadi tidak tampil di outlet manapun.
        let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
        let inUse = products.filter(p => productUnits_(p.ApplicableUnit).indexOf(unitKey) !== -1).length;
        if (inUse > 0) {
          return { success: false, message: "Unit \"" + unitKey + "\" masih dipakai oleh " + inUse + " produk. Ubah dulu Outlet Penjualan produk-produk tsb sebelum menghapus unit ini." };
        }
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_APPLICABLE_UNIT", "Master Unit Berlaku", "Hapus unit " + unitKey, "");
        return { success: true, message: "Unit berhasil dihapus." };
      }
    }
    return { success: false, message: "Unit tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// MIGRASI: pastikan sheet Applicable_Units ada (database lama yang dibuat
// sebelum fitur "Unit Berlaku bisa diatur sendiri" ditambahkan) & sudah
// terisi 8 unit bawaan (General/Entrance/Ride/Resto/Kolam/Teknisi/
// Cleaning/Dapur) kalau sheetnya masih kosong, supaya checkbox "Outlet
// Penjualan" di Master Produk tidak pernah kosong. Aman & idempoten,
// dijalankan tiap kali getDatabase() dipanggil.
function ensureApplicableUnitsSeed_(ss) {
  try {
    const sheet = ensureSheet_(ss, CONFIG.SHEETS.APPLICABLE_UNITS, APPLICABLE_UNIT_HEADERS);
    if (sheet.getLastRow() < 2) {
      sheet.getRange(2, 1, 8, 4).setValues([
        ["APU01", "General", "Semua Outlet", "Active"],
        ["APU02", "Entrance", "Kasir Loket Masuk", "Active"],
        ["APU03", "Ride", "Kasir Loket Wahana", "Active"],
        ["APU04", "Resto", "Kasir Resto", "Active"],
        ["APU05", "Kolam", "Gudang Kolam Renang", "Active"],
        ["APU06", "Teknisi", "Gudang Teknisi", "Active"],
        ["APU07", "Cleaning", "Gudang Cleaning Service", "Active"],
        ["APU08", "Dapur", "Gudang Dapur", "Active"]
      ]);
    }
  } catch (e) {}
}

// =====================================================================
// 7. STOCK ENGINE - dipakai oleh Sales & Manajemen Stok
// =====================================================================
// =====================================================================
// 8b. SALES FORM CONFIG - kontrol field mana yang tampil & bisa diedit
// manual oleh kasir di form Input Penjualan, per unit (Owner only untuk
// mengubah). Disimpan sebagai JSON di Settings dengan key
// SALESFORM_CONFIG_<Unit>. Field Serial hanya bisa diaktifkan untuk unit
// yang memang punya kolom Nomor Seri di database (cfg.hasSerial).
// =====================================================================
function getSalesFormConfig(unit) {
  try {
    const cfg = unitCfg_(unit);
    let settings = getSystemSettings();
    let raw = settings['SALESFORM_CONFIG_' + unit];
    let saved = {};
    if (raw) { try { saved = JSON.parse(raw); } catch (e) { saved = {}; } }
    let result = {};
    ['StockStart', 'StockIn', 'StockOut', 'StockEnd'].forEach(function(k) {
      result[k] = Object.assign({ show: false, editable: false }, saved[k] || {});
    });
    ['SerialStart', 'SerialEnd'].forEach(function(k) {
      if (!cfg.hasSerial) { result[k] = { show: false, editable: false }; return; }
      result[k] = Object.assign({ show: true, editable: true }, saved[k] || {});
    });
    return { success: true, hasSerial: !!cfg.hasSerial, config: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveSalesFormConfig(unit, configObj, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    const cfg = unitCfg_(unit);
    let clean = {};
    ['StockStart', 'StockIn', 'StockOut', 'StockEnd'].forEach(function(k) {
      let f = configObj[k] || {};
      clean[k] = { show: !!f.show, editable: !!f.show && !!f.editable };
    });
    ['SerialStart', 'SerialEnd'].forEach(function(k) {
      if (!cfg.hasSerial) { clean[k] = { show: false, editable: false }; return; }
      let f = configObj[k] || {};
      clean[k] = { show: !!f.show, editable: !!f.show && !!f.editable };
    });
    let payload = {};
    payload['SALESFORM_CONFIG_' + unit] = JSON.stringify(clean);
    saveSystemSettings(payload, token);
    logActivity_(user.username, "EDIT_SETTINGS", cfg.label, "Ubah pengaturan form input penjualan", "");
    return { success: true, message: "Pengaturan form input " + cfg.label + " berhasil disimpan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Stok terkini produk di gudang outlet ybs (dipakai form Input Penjualan
// untuk menampilkan Stok Awal/Stok Akhir secara live saat produk dipilih).
function getProductStock(unit, productCode, token) {
  try {
    checkPermission_(token, null, unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    let stock = getCurrentStock_(ss, cfg.warehouseId, productCode);
    return { success: true, stock: stock };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// FITUR "ACUAN STOK" (StockRefCode): sebagian produk (mis. TERP002 "Tiket
// Terusan Plus Diskon 20%") tidak punya stok fisik sendiri - stoknya SAMA
// dengan produk lain (mis. TER001 "Tiket Terusan Plus"), karena keduanya
// sebetulnya menjual tiket fisik yang sama, cuma beda harga/kode. Kolom
// Master Produk "StockRefCode" menyimpan kode produk acuan tsb. Fungsi ini
// menerjemahkan SEMBARANG kode produk menjadi kode yang BENAR-BENAR dipakai
// untuk membaca/menulis Warehouse_Stock - dipanggil otomatis di dalam
// getCurrentStock_/updateStock_/setStockAbsolute_ sehingga SEMUA alur (jual/
// edit/hapus transaksi, opname, mutasi antar gudang, tambah/isi stok) yang
// menyentuh produk ber-StockRefCode otomatis memotong/menambah stok produk
// acuannya, TANPA perlu mengubah kode di tempat lain. Hanya 1 tingkat
// referensi yang didukung (divalidasi di saveProduct) - loop dengan
// `visited` di bawah ini sekadar jaring pengaman kalau ada data lama yang
// (secara tidak sengaja) membentuk rantai/lingkaran referensi.
function resolveStockProductCode_(ss, productCode) {
  let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
  let map = {};
  products.forEach(p => map[p.ProductCode] = p);
  let code = productCode, visited = {};
  while (map[code] && String(map[code].StockRefCode || "").trim() && map[code].StockRefCode !== code) {
    if (visited[code]) break; // cegah lingkaran referensi tak berujung
    visited[code] = true;
    code = String(map[code].StockRefCode).trim();
  }
  return code;
}

// Kolom Warehouse_Stock (lihat WAREHOUSE_STOCK_HEADERS): StockID(0)
// WarehouseID(1) WarehouseName(2) ProductCode(3) ProductName(4) Quantity(5)
// MinStock(6) Status(7) LastUpdated(8).
function getCurrentStock_(ss, warehouseId, productCode) {
  productCode = resolveStockProductCode_(ss, productCode);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === warehouseId && data[i][3] === productCode) return Number(data[i][5]) || 0;
  }
  return 0;
}

// PERBAIKAN BUG "stok duplikat" (lihat catatan panjang di bawah fungsi ini,
// findDuplicateWarehouseStock_/manualMergeDuplicateWarehouseStock): fungsi
// ini SEKARANG mengunci script (LockService) selama baca-hitung-tulis,
// supaya 2 eksekusi yang nyaris bersamaan (mis. 2 kasir jual produk yang
// sama detik yang sama, atau tombol admin ke-klik 2x/retry jaringan) tidak
// bisa sama-sama gagal menemukan baris yang sudah ada lalu sama-sama
// appendRow baris BARU untuk kombinasi WarehouseID+ProductCode yang sama -
// itulah yang menyebabkan baris "duplikat" yang salah satunya tidak pernah
// ikut ter-update lagi (lihat Warehouse_Stock, WH04, banyak produk Resto).
function updateStock_(ss, warehouseId, productCode, deltaQty, minStockDefault, meta) {
  // meta (opsional) = konteks untuk KARTU STOK: { type, counterpart, refId,
  // notes, user, silent, balanceBefore }. Tanpa meta, perubahan tetap
  // tercatat di kartu stok dengan jenis generik "Perubahan Stok".
  // silent:true = jangan catat di kartu stok (dipakai langkah antara pada
  // edit penjualan, yang dicatat SATU baris bersih lewat balanceBefore).
  const requestedCode = productCode;
  productCode = resolveStockProductCode_(ss, productCode);
  const viaCode = (requestedCode !== productCode) ? requestedCode : "";
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    let data = sheet.getDataRange().getValues();
    let info = warehouseProductInfo_(ss, warehouseId, productCode);
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === warehouseId && data[i][3] === productCode) {
        let oldQty = Number(data[i][5]) || 0;
        let newQty = Number(data[i][5]) + deltaQty;
        // Ikut menyegarkan WarehouseName/ProductName/Status (kalau produk
        // sempat diganti nama/status di Master Produk sejak baris ini terakhir
        // diupdate), bukan cuma Quantity & LastUpdated.
        sheet.getRange(i + 1, 3, 1, 7).setValues([[info.warehouseName, productCode, info.productName, newQty, data[i][6], info.status, new Date()]]);
        logStockCardMovement_(ss, warehouseId, info, productCode, oldQty, newQty, meta, viaCode);
        return newQty;
      }
    }
    let newQty = deltaQty;
    sheet.appendRow([genId_("STK"), warehouseId, info.warehouseName, productCode, info.productName, newQty, minStockDefault || 10, info.status, new Date()]);
    logStockCardMovement_(ss, warehouseId, info, productCode, 0, newQty, meta, viaCode);
    return newQty;
  } finally {
    lock.releaseLock();
  }
}

// Set stok outlet LANGSUNG ke nilai absolut (bukan tambah/kurang seperti
// updateStock_) - dipakai savePhysicalStockCheck() supaya hasil stok
// opname selalu = angka yang benar-benar dihitung fisik oleh kasir,
// terlepas dari akurasi StokAwal yang dipakai untuk pelaporan Selisih.
// Sama seperti updateStock_ di atas: SEKARANG dikunci (LockService) supaya
// tidak bisa membuat baris duplikat di Warehouse_Stock kalau terpanggil
// nyaris bersamaan (mis. savePhysicalStockCheck menyimpan banyak produk
// dalam satu klik "Simpan Semua").
function setStockAbsolute_(ss, warehouseId, productCode, newQty, minStockDefault, meta) {
  // meta (opsional): konteks KARTU STOK, lihat updateStock_.
  const requestedCode = productCode;
  productCode = resolveStockProductCode_(ss, productCode);
  const viaCode = (requestedCode !== productCode) ? requestedCode : "";
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    let data = sheet.getDataRange().getValues();
    let info = warehouseProductInfo_(ss, warehouseId, productCode);
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === warehouseId && data[i][3] === productCode) {
        let oldQty = Number(data[i][5]) || 0;
        sheet.getRange(i + 1, 3, 1, 7).setValues([[info.warehouseName, productCode, info.productName, newQty, data[i][6], info.status, new Date()]]);
        logStockCardMovement_(ss, warehouseId, info, productCode, oldQty, newQty, meta, viaCode);
        return newQty;
      }
    }
    sheet.appendRow([genId_("STK"), warehouseId, info.warehouseName, productCode, info.productName, newQty, minStockDefault || 10, info.status, new Date()]);
    logStockCardMovement_(ss, warehouseId, info, productCode, 0, newQty, meta, viaCode);
    return newQty;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
// Penulis KARTU STOK (lihat STOCK_CARD_HEADERS). Dipanggil dari DALAM
// updateStock_/setStockAbsolute_ (saat lock script sudah dipegang) dan dari
// fungsi yang menulis Warehouse_Stock langsung (Set Stok Manual dst).
// Pencatatan kartu stok bersifat "best effort": kalau gagal menulis (mis.
// kuota Sheets), error hanya dicatat di Logger dan TIDAK menggagalkan
// transaksi penjualan/transfer yang sedang berjalan.
// ---------------------------------------------------------------------
function buildStockCardRow_(warehouseId, warehouseName, productCode, productName, before, after, meta, viaCode) {
  meta = meta || {};
  before = Number(before) || 0;
  after = Number(after) || 0;
  const diff = Math.round((after - before) * 10000) / 10000;
  return [genId_("KRT"), meta.timestamp || new Date(), warehouseId, warehouseName, productCode, productName,
    meta.type || STOCK_CARD_TYPE.UNKNOWN,
    diff > 0 ? diff : 0, diff < 0 ? -diff : 0, before, after,
    meta.counterpart || "", meta.refId || "", meta.notes || "", meta.user || "", viaCode || ""];
}

// Baris tanpa perubahan (QtyIn = QtyOut = 0) tidak dicatat - kartu stok
// hanya berisi pergerakan stok yang nyata.
function appendStockCardRows_(ss, rows) {
  try {
    rows = (rows || []).filter(r => r[7] !== 0 || r[8] !== 0);
    if (!rows.length) return;
    const sheet = ss.getSheetByName(CONFIG.SHEETS.STOCK_CARD) || ensureSheet_(ss, CONFIG.SHEETS.STOCK_CARD, STOCK_CARD_HEADERS);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, STOCK_CARD_HEADERS.length).setValues(rows);
  } catch (e) {
    Logger.log("Gagal mencatat kartu stok: " + e.message);
  }
}

function logStockCardMovement_(ss, warehouseId, info, productCode, before, after, meta, viaCode) {
  if (meta && meta.silent) return;
  if (meta && meta.balanceBefore !== undefined && meta.balanceBefore !== null) before = Number(meta.balanceBefore) || 0;
  appendStockCardRows_(ss, [buildStockCardRow_(warehouseId, info.warehouseName, productCode, info.productName, before, after, meta, viaCode)]);
}

// =====================================================================
// 7c. MAINTENANCE - CEK & BERSIHKAN BARIS DUPLIKAT DI WAREHOUSE_STOCK
// -----------------------------------------------------------------------
// LATAR BELAKANG BUG (ditemukan 2/9/2026): updateStock_/setStockAbsolute_
// dulu TIDAK memakai LockService (sudah diperbaiki di atas). Akibatnya,
// kalau 2 eksekusi sempat berjalan nyaris bersamaan (mis. tombol admin
// ke-klik 2x, retry jaringan, atau 2 kasir transaksi produk sama di detik
// yang sama), keduanya bisa SAMA-SAMA tidak menemukan baris Warehouse_Stock
// yang sudah ada untuk kombinasi WarehouseID+ProductCode ybs, lalu
// SAMA-SAMA appendRow baris BARU -> jadilah 2 baris untuk 1 produk yang
// sama. Efeknya: updateStock_ (dipanggil tiap ada penjualan) hanya
// mengupdate baris PERTAMA yang ditemukan di sheet, sedangkan beberapa
// fungsi TAMPILAN (mis. getPhysicalStockCheckView) membangun peta produk
// dengan pola "forEach menimpa" sehingga baris TERAKHIR yang ditemukan
// yang menang ditampilkan - kalau baris pertama & terakhir itu BEDA baris
// (karena duplikat), stok yang ditampilkan ke user jadi tidak pernah
// sinkron dengan stok yang sebenarnya diupdate (kasus nyata: Kasir Resto,
// WH04, 15 produk kena, termasuk Populaire Shaky Shake yang stuck di 19
// padahal harusnya sudah 18 sejak transaksi kemarin).
//
// CARA PAKAI: buka project ini di Apps Script Editor (Extensions > Apps
// Script kalau dari Sheets, atau langsung di script.google.com), pilih
// fungsi "manualMergeDuplicateWarehouseStock" di dropdown sebelah tombol
// Run, lalu klik Run. Cek hasilnya di View > Logs (Ctrl+Enter). TIDAK
// dipanggil dari web app manapun - sengaja dibuat manual supaya Owner bisa
// tinjau dulu log-nya sebelum baris dihapus, dan supaya tidak butuh token
// login web app untuk menjalankannya. Aman dijalankan berkali-kali -
// kalau sudah tidak ada duplikat, tidak melakukan apa-apa.
function findDuplicateWarehouseStock_() {
  const ss = getDatabase();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
  const data = sheet.getDataRange().getValues();
  const seen = {}; // key -> row index pertama yang ditemukan (0-based di `data`)
  const duplicates = []; // { row (1-based di sheet), keptRow, warehouseId, productCode, keptQty, dupQty }
  for (let i = 1; i < data.length; i++) {
    const warehouseId = data[i][1], productCode = data[i][3];
    if (!warehouseId || !productCode) continue; // baris kosong di bawah data, lewati
    const key = warehouseId + "||" + productCode;
    if (seen[key] !== undefined) {
      duplicates.push({
        row: i + 1, keptRow: seen[key] + 1,
        warehouseId: warehouseId, productCode: productCode, productName: data[i][4],
        keptQty: data[seen[key]][5], dupQty: data[i][5]
      });
    } else {
      seen[key] = i;
    }
  }
  return duplicates;
}

// Jalankan fungsi ini secara manual dari Apps Script Editor (lihat catatan
// di atas). Mempertahankan baris PERTAMA tiap kombinasi WarehouseID+
// ProductCode (baris itulah yang selama ini benar-benar diupdate oleh
// updateStock_/setStockAbsolute_ setiap ada transaksi/opname), menghapus
// baris duplikat sesudahnya. Kalau qty di baris yang dipertahankan vs yang
// dihapus BERBEDA (artinya sudah ada transaksi yang jalan sejak baris
// duplikat itu tercipta), qty yang dipertahankan SELALU yang dipakai -
// itu yang paling akurat karena sudah mengikuti riwayat transaksi asli.
function manualMergeDuplicateWarehouseStock() {
  const ss = getDatabase();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
  const duplicates = findDuplicateWarehouseStock_();
  if (!duplicates.length) {
    Logger.log("Tidak ada baris duplikat ditemukan di Warehouse_Stock. Tidak ada yang perlu dibersihkan.");
    return duplicates;
  }
  Logger.log("Ditemukan " + duplicates.length + " baris duplikat. Menghapus (baris pertama tiap produk dipertahankan)...");
  // Hapus dari baris PALING BAWAH dulu supaya nomor baris yang belum
  // dihapus tidak ikut bergeser saat proses berjalan.
  duplicates.slice().sort((a, b) => b.row - a.row).forEach(d => sheet.deleteRow(d.row));
  duplicates.forEach(d => Logger.log(
    "[" + d.warehouseId + "/" + d.productCode + "] " + d.productName +
    " -> baris duplikat #" + d.row + " (qty " + d.dupQty + ") DIHAPUS, baris #" + d.keptRow +
    " (qty " + d.keptQty + ") dipertahankan sebagai stok yang benar."));
  logActivity_("SYSTEM", "MERGE_DUPLICATE_STOCK", "Warehouse", "Bersihkan " + duplicates.length + " baris duplikat Warehouse_Stock", "");
  return duplicates;
}

// =====================================================================
// 7d. KARTU STOK (STOCK CARD) - buku besar pergerakan stok per produk
// -----------------------------------------------------------------------
// Sumber data: sheet Stock_Card (lihat STOCK_CARD_HEADERS), yang diisi
// OTOMATIS oleh mesin stok (updateStock_/setStockAbsolute_ dan fungsi Set
// Stok Manual/Hapus Catatan Stok). Fungsi di bagian ini HANYA MEMBACA &
// merangkum data itu - tidak pernah mengubah stok.
//
// Lokasi yang didukung: SEMUA gudang di sheet Warehouses -
//   WH01 Gudang Utama, WH02 Kasir Loket Masuk, WH03 Kasir Loket Wahana,
//   WH04 Kasir Resto, WH05-WH08 gudang internal.
// Produk yang punya kartu: Active, UsesStock = "Ya", tidak ber-Acuan Stok
// (produk ber-Acuan Stok memakai kartu produk acuannya), dan berlaku di
// lokasi tsb - aturan yang SAMA dengan daftar stok di Manajemen Stok &
// modul Warehouse.
//
// Rumus tiap kartu untuk periode [startDate, endDate]:
//   Saldo Awal  = saldo akhir transaksi terakhir SEBELUM startDate
//                 (kalau belum ada: saldo sebelum transaksi pertama; kalau
//                  produk belum pernah bergerak: stok saat ini)
//   Masuk/Keluar= jumlah QtyIn/QtyOut di dalam periode
//   Saldo Akhir = saldo setelah transaksi terakhir di dalam periode
// Deteksi selisih: kalau BalanceBefore sebuah baris berbeda dari
// BalanceAfter baris sebelumnya, berarti stok diubah DI LUAR aplikasi
// (mis. edit langsung di Google Sheets) - ditandai lewat field Gap.
// =====================================================================
function stockCardAccess_(token, warehouseId) {
  const unit = unitByWarehouseId_(warehouseId); // WH02-WH04 -> Kasir hanya boleh outlet-nya sendiri
  if (unit) return checkPermission_(token, null, unit);
  return checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
}

function getWarehouseOrThrow_(ss, warehouseId) {
  const w = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES)).find(x => x.WarehouseID === warehouseId);
  if (!w) throw new Error("Gudang/lokasi tidak ditemukan: " + warehouseId);
  return { WarehouseID: w.WarehouseID, WarehouseName: w.WarehouseName, Location: w.Location };
}

function stockCardProducts_(ss, warehouseId) {
  const unitOfWarehouse = applicableUnitByWarehouseId_(warehouseId); // null utk Gudang Utama (semua produk)
  return getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS)).filter(p => {
    if (p.Status !== "Active") return false;
    if (String(p.UsesStock || "Ya").trim() === "Tidak") return false;
    if (String(p.StockRefCode || "").trim()) return false;
    if (!unitOfWarehouse) return true;
    const units = productUnits_(p.ApplicableUnit);
    return units.indexOf(unitOfWarehouse) !== -1 || units.indexOf("General") !== -1;
  });
}

// Baca Stock_Card SATU kali, kelompokkan per ProductCode untuk 1 gudang,
// urut kronologis. Memakai array mentah (bukan objek per baris) supaya tetap
// cepat walau sheet-nya besar.
function readStockCardEntries_(ss, warehouseId, onlyCode) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.STOCK_CARD);
  const byProduct = {};
  if (!sheet || sheet.getLastRow() < 2) return byProduct;
  const values = sheet.getDataRange().getValues();
  const h = {};
  values[0].forEach((name, i) => { h[String(name).trim()] = i; });
  const need = ["Timestamp", "WarehouseID", "ProductCode", "TransactionType", "QtyIn", "QtyOut", "BalanceBefore", "BalanceAfter"];
  for (let k = 0; k < need.length; k++) { if (h[need[k]] === undefined) return byProduct; }
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r[h.WarehouseID] !== warehouseId) continue;
    const code = r[h.ProductCode];
    if (!code || (onlyCode && code !== onlyCode)) continue;
    const ts = r[h.Timestamp] instanceof Date ? r[h.Timestamp] : new Date(r[h.Timestamp]);
    if (isNaN(ts.getTime())) continue;
    (byProduct[code] = byProduct[code] || []).push({
      _row: i + 1, ts: ts,
      type: r[h.TransactionType],
      qtyIn: Number(r[h.QtyIn]) || 0, qtyOut: Number(r[h.QtyOut]) || 0,
      before: Number(r[h.BalanceBefore]) || 0, after: Number(r[h.BalanceAfter]) || 0,
      counterpart: h.Counterpart !== undefined ? String(r[h.Counterpart] || "") : "",
      refId: h.RefID !== undefined ? String(r[h.RefID] || "") : "",
      notes: h.Notes !== undefined ? String(r[h.Notes] || "") : "",
      user: h.User !== undefined ? String(r[h.User] || "") : "",
      via: h.ViaProduct !== undefined ? String(r[h.ViaProduct] || "") : ""
    });
  }
  Object.keys(byProduct).forEach(code => byProduct[code].sort((a, b) => (a.ts - b.ts) || (a._row - b._row)));
  return byProduct;
}

function round4_(n) { return Math.round((Number(n) || 0) * 10000) / 10000; }

function buildStockCards_(ss, warehouse, products, startDate, endDate, withRows, onlyCode) {
  const tz = Session.getScriptTimeZone();
  const warehouseId = warehouse.WarehouseID;
  const whNameMap = {};
  getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES)).forEach(w => whNameMap[w.WarehouseID] = w.WarehouseName);
  const stockMap = {};
  getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK))
    .filter(s => s.WarehouseID === warehouseId).forEach(s => stockMap[s.ProductCode] = s);
  const entriesByProduct = readStockCardEntries_(ss, warehouseId, onlyCode);

  return products.map(p => {
    const s = stockMap[p.ProductCode];
    const currentQty = s ? (Number(s.Quantity) || 0) : 0;
    const entries = entriesByProduct[p.ProductCode] || [];

    let prevAfter = null;
    entries.forEach(e => {
      e.date = Utilities.formatDate(e.ts, tz, "yyyy-MM-dd");
      e.gap = (prevAfter !== null && Math.abs(e.before - prevAfter) > 0.0001) ? round4_(e.before - prevAfter) : 0;
      prevAfter = e.after;
    });

    let lastBefore = null;
    if (startDate) entries.forEach(e => { if (e.date < startDate) lastBefore = e; });
    let opening;
    if (lastBefore) opening = lastBefore.after;
    else if (entries.length) opening = entries[0].before;
    else opening = currentQty;

    const inRange = entries.filter(e => (!startDate || e.date >= startDate) && (!endDate || e.date <= endDate));
    let masuk = 0, keluar = 0, gapTotal = 0;
    inRange.forEach(e => { masuk += e.qtyIn; keluar += e.qtyOut; gapTotal += e.gap; });
    const closing = inRange.length ? inRange[inRange.length - 1].after : opening;
    const latest = entries.length ? entries[entries.length - 1] : null;

    const card = {
      ProductCode: p.ProductCode, ProductName: p.ProductName, Category: p.Category, Unit: p.Unit || "",
      ItemType: p.ItemType || "",
      Quantity: currentQty, MinStock: s ? s.MinStock : 10, HasRecord: !!s,
      Opening: round4_(opening), Masuk: round4_(masuk), Keluar: round4_(keluar), Closing: round4_(closing),
      GapTotal: round4_(gapTotal), MovementCount: inRange.length,
      HasHistory: entries.length > 0,
      LatestBalance: latest ? latest.after : null,
      Mismatch: latest ? (Math.abs(latest.after - currentQty) > 0.0001) : false,
      LastMovement: latest ? latest.ts.toISOString() : null
    };
    if (withRows) {
      card.rows = inRange.map(e => ({
        Timestamp: e.ts.toISOString(), Date: e.date, TransactionType: e.type,
        Masuk: e.qtyIn, Keluar: e.qtyOut, Before: e.before, After: e.after,
        Counterpart: whNameMap[e.counterpart] || e.counterpart, CounterpartID: e.counterpart,
        RefID: e.refId, Notes: e.notes, User: e.user, ViaProduct: e.via, Gap: e.gap
      }));
    }
    return card;
  });
}

// Daftar ringkas semua kartu stok di 1 lokasi untuk 1 periode (tabel utama
// menu Kartu Stok): Saldo Awal, Masuk, Keluar, Saldo Akhir, stok saat ini.
// startDate/endDate format "yyyy-MM-dd" (inklusif), boleh kosong = semua.
function getStockCardList(warehouseId, startDate, endDate, token) {
  try {
    stockCardAccess_(token, warehouseId);
    const ss = getDatabase();
    const warehouse = getWarehouseOrThrow_(ss, warehouseId);
    const cards = buildStockCards_(ss, warehouse, stockCardProducts_(ss, warehouseId), startDate || "", endDate || "", false, null);
    return { success: true, warehouse: warehouse, startDate: startDate || "", endDate: endDate || "", cards: cards };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Kartu stok LENGKAP (dengan seluruh baris transaksi) untuk 1 produk di 1
// lokasi. Kalau produk ber-Acuan Stok, otomatis dialihkan ke kartu produk
// acuannya (field viaRef terisi kode produk yang diminta).
function getStockCard(warehouseId, productCode, startDate, endDate, token) {
  try {
    stockCardAccess_(token, warehouseId);
    const ss = getDatabase();
    const warehouse = getWarehouseOrThrow_(ss, warehouseId);
    const resolved = resolveStockProductCode_(ss, productCode);
    const product = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS)).find(p => p.ProductCode === resolved);
    if (!product) return { success: false, message: "Produk tidak ditemukan: " + productCode };
    const cards = buildStockCards_(ss, warehouse, [product], startDate || "", endDate || "", true, resolved);
    return {
      success: true, warehouse: warehouse, startDate: startDate || "", endDate: endDate || "",
      card: cards[0], viaRef: resolved !== productCode ? productCode : ""
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// SEMUA kartu (lengkap dengan baris transaksi) di 1 lokasi - dipakai tombol
// "Semua Kartu Stok" (PDF/XLSX 1 sheet per produk) di frontend.
function getStockCardsBulk(warehouseId, startDate, endDate, token) {
  try {
    stockCardAccess_(token, warehouseId);
    const ss = getDatabase();
    const warehouse = getWarehouseOrThrow_(ss, warehouseId);
    const cards = buildStockCards_(ss, warehouse, stockCardProducts_(ss, warehouseId), startDate || "", endDate || "", true, null);
    return { success: true, warehouse: warehouse, startDate: startDate || "", endDate: endDate || "", cards: cards };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 8. SALES - LAPORAN PENJUALAN (generik per unit)
// =====================================================================
function getSalesData(unit, token) {
  try {
    checkPermission_(token, null, unit);
    const ss = getDatabase();
    return { success: true, data: getSheetDataAsObjects_(ss.getSheetByName(unitCfg_(unit).salesSheet)) };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveSalesTransaction(unit, data, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.KASIR], unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    const sheet = ss.getSheetByName(cfg.salesSheet);
    const paySheet = ss.getSheetByName(cfg.paySheet);
    let qty = Number(data.qty);
    let unitPrice = Number(data.unitPrice);
    // Total Harga: pakai nilai dari form (bisa diedit manual oleh kasir) kalau
    // dikirim dan valid. Kalau kosong/tidak dikirim (mis. dari pemanggilan lain
    // yang belum diupdate), fallback ke perhitungan otomatis qty x harga satuan.
    let totalPrice = (data.totalPrice !== undefined && data.totalPrice !== null && data.totalPrice !== "" && !isNaN(Number(data.totalPrice)))
      ? Number(data.totalPrice)
      : qty * unitPrice;
    let cashierName = user.fullName;

    // Produk "Tidak Pakai Stok" (UsesStock = "Tidak" di Master Produk, mis.
    // jasa/tiket tanpa fisik): stok gudang TIDAK PERNAH dikurangi/ditambah,
    // dan Stok Awal/Masuk/Keluar/Akhir yang tercatat di laporan penjualan
    // SELALU 0 - tidak bisa ditimpa manual oleh kasir. Produk yang memakai
    // stok (default, UsesStock = "Ya") tetap berjalan seperti sebelumnya.
    let usesStock = productUsesStock_(ss, data.productCode);
    // Field Nomor Seri Awal/Akhir cuma relevan kalau produknya memang
    // pakai stok DAN ditandai UsesSerial = "Ya" di Master Produk - produk
    // "Tidak Pakai Stok" tidak pernah menampilkan/menyimpan nomor seri.
    let usesSerial = usesStock && productUsesSerial_(ss, data.productCode);

    // Fleksibel: stok boleh kosong/minus saat input penjualan (stok awal
    // bisa menyusul ditambahkan lewat Manajemen Stok). Validasi stok hanya
    // dijalankan jika Owner mengaktifkan ENABLE_STOCK_VALIDATION di Pengaturan.
    let settings = getSystemSettings();
    let enforceStock = settings.ENABLE_STOCK_VALIDATION === "Yes";
    // Metode pembayaran per-item hanya dicatat jika fitur ini diaktifkan di
    // Pengaturan (SHOW_PAYMENT_METHOD_IN_SALES). Default: uang diterima
    // dicatat terpisah lewat menu Penerimaan Uang (rekap harian saja).
    let recordPaymentOnSale = settings.SHOW_PAYMENT_METHOD_IN_SALES === "Yes";

    // Nilai Stok Awal/Masuk/Keluar/Akhir yang DICATAT di laporan boleh
    // ditimpa manual oleh kasir HANYA untuk field yang Owner tandai
    // "Bisa Diedit Kasir" di Pengaturan > Pengaturan Form Input Penjualan.
    // Stok gudang sesungguhnya (WAREHOUSE_STOCK) TETAP dihitung otomatis
    // dari Qty apa adanya - override manual di sini hanya memengaruhi
    // angka yang tercatat di baris laporan penjualan, bukan saldo stok
    // riil di gudang, supaya otomasi inventori tidak pernah rusak walau
    // Owner mengizinkan koreksi manual pada laporan.
    let formCfgRaw = settings['SALESFORM_CONFIG_' + unit];
    let formCfg = {};
    if (formCfgRaw) { try { formCfg = JSON.parse(formCfgRaw); } catch (e) { formCfg = {}; } }
    function manualOverride_(fieldKey, autoVal) {
      let fc = formCfg[fieldKey];
      let manualField = fieldKey.charAt(0).toLowerCase() + fieldKey.slice(1) + "Manual";
      let raw = data[manualField];
      if (fc && fc.editable && raw !== undefined && raw !== null && raw !== "" && !isNaN(Number(raw))) {
        return Number(raw);
      }
      return autoVal;
    }

    if (data.isEdit) {
      let rows = sheet.getDataRange().getValues();
      let headers = rows[0];
      let saleIdCol = headers.indexOf("SaleID");
      let qtyCol = headers.indexOf("Qty");
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][saleIdCol] === data.saleId) {
          let stockStart, stockIn, stockOut, stockEnd;
          if (usesStock) {
            let oldQty = Number(rows[i][qtyCol]);
            // Kartu stok: langkah "kembalikan stok lama" & "potong stok baru"
            // dicatat sebagai SATU baris bersih (Koreksi Edit Penjualan) -
            // saldo sebelum = stok sebelum edit dimulai.
            let balanceBeforeEdit = getCurrentStock_(ss, cfg.warehouseId, data.productCode);
            updateStock_(ss, cfg.warehouseId, data.productCode, oldQty, undefined, { silent: true });
            let stockStartAuto = getCurrentStock_(ss, cfg.warehouseId, data.productCode);
            if (enforceStock && stockStartAuto < qty) {
              updateStock_(ss, cfg.warehouseId, data.productCode, -oldQty, undefined, { silent: true });
              return { success: false, message: "Stok tidak mencukupi. Sisa stok saat ini: " + stockStartAuto };
            }
            let stockEndAuto = stockStartAuto - qty;
            updateStock_(ss, cfg.warehouseId, data.productCode, -qty, undefined, {
              type: STOCK_CARD_TYPE.SALE_EDIT, refId: data.saleId, user: user.fullName,
              balanceBefore: balanceBeforeEdit, notes: "Edit qty penjualan: " + oldQty + " menjadi " + qty
            });

            stockStart = manualOverride_("StockStart", stockStartAuto);
            stockIn = manualOverride_("StockIn", 0);
            stockOut = manualOverride_("StockOut", qty);
            stockEnd = manualOverride_("StockEnd", stockEndAuto);
          } else {
            // Produk tidak pakai stok: stok gudang tidak disentuh sama
            // sekali, dan angka di laporan penjualan selalu 0.
            stockStart = 0; stockIn = 0; stockOut = 0; stockEnd = 0;
          }

          let rowValues = buildSalesRow_(cfg, data.saleId, rows[i][headers.indexOf("Date")], data, qty, stockStart, stockIn, stockOut, stockEnd, totalPrice, cashierName, rows[i][headers.indexOf("InputTime")], usesSerial);
          sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
          logActivity_(user.username, "EDIT_SALES", cfg.label, "Edit transaksi " + data.saleId, "");
          return { success: true, message: "Transaksi berhasil diperbarui." };
        }
      }
      return { success: false, message: "Transaksi tidak ditemukan." };
    } else {
      // ID transaksi dibuat lebih awal supaya bisa dicatat sebagai referensi
      // di kartu stok saat stok dipotong.
      const newSaleId = genId_("TRX");
      let stockStart, stockIn, stockOut, stockEnd;
      if (usesStock) {
        let stockStartAuto = getCurrentStock_(ss, cfg.warehouseId, data.productCode);
        if (enforceStock && stockStartAuto < qty) {
          return { success: false, message: "Stok tidak mencukupi. Sisa stok saat ini: " + stockStartAuto };
        }
        let stockEndAuto = stockStartAuto - qty;
        updateStock_(ss, cfg.warehouseId, data.productCode, -qty, undefined, { type: STOCK_CARD_TYPE.SALE, refId: newSaleId, user: user.fullName });

        stockStart = manualOverride_("StockStart", stockStartAuto);
        stockIn = manualOverride_("StockIn", 0);
        stockOut = manualOverride_("StockOut", qty);
        stockEnd = manualOverride_("StockEnd", stockEndAuto);
      } else {
        // Produk tidak pakai stok: stok gudang tidak disentuh sama sekali,
        // dan angka di laporan penjualan selalu 0.
        stockStart = 0; stockIn = 0; stockOut = 0; stockEnd = 0;
      }

      let saleId = newSaleId;
      let now = new Date();
      let rowValues = buildSalesRow_(cfg, saleId, now, data, qty, stockStart, stockIn, stockOut, stockEnd, totalPrice, cashierName, now, usesSerial);
      sheet.appendRow(rowValues);

      if (recordPaymentOnSale && data.paymentMethod) {
        paySheet.appendRow([genId_("PAY"), now, Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss"),
          data.paymentMethod, totalPrice, saleId, cashierName, "Pembayaran untuk " + saleId]);
      }

      logActivity_(user.username, "CREATE_SALES", cfg.label, "Transaksi baru " + saleId, "");
      return { success: true, message: "Transaksi berhasil disimpan." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function buildSalesRow_(cfg, saleId, date, data, qty, stockStart, stockIn, stockOut, stockEnd, totalPrice, cashierName, inputTime, includeSerial) {
  if (cfg.hasSerial) {
    // includeSerial = false untuk produk yang UsesStock = "Tidak" atau
    // UsesSerial = "Tidak" di Master Produk - kolom seri tetap ada di
    // sheet (skema outletnya hasSerial), tapi nilainya dipaksa "-" / qty
    // apa pun yang (mungkin) terkirim dari client, konsisten dengan
    // field seri yang memang disembunyikan di Form Input Penjualan.
    let serialStart = includeSerial ? (data.serialStart || "-") : "-";
    let serialEnd = includeSerial ? (data.serialEnd || "-") : "-";
    let totalSerial = (includeSerial && Number(data.serialEnd) && Number(data.serialStart))
      ? (Number(data.serialEnd) - Number(data.serialStart) + 1) : qty;
    return [saleId, date, data.productCode, data.productName, data.variant || "-", data.category,
      qty, data.unit, serialStart, serialEnd, totalSerial,
      stockStart, stockIn, stockOut, stockEnd, Number(data.unitPrice), totalPrice, cashierName, inputTime, data.notes || "-"];
  }
  return [saleId, date, data.productCode, data.productName, data.variant || "-", data.category,
    qty, data.unit, stockStart, stockIn, stockOut, stockEnd, Number(data.unitPrice), totalPrice, cashierName, inputTime, data.notes || "-"];
}

function deleteSalesTransaction(unit, saleId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.KASIR], unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    const sheet = ss.getSheetByName(cfg.salesSheet);
    let rows = sheet.getDataRange().getValues();
    let headers = rows[0];
    let saleIdCol = headers.indexOf("SaleID");
    let productCol = headers.indexOf("ProductCode");
    let qtyCol = headers.indexOf("Qty");

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][saleIdCol] === saleId) {
        // Kembalikan stok gudang HANYA kalau produk ini memang memakai
        // pelacakan stok (UsesStock = "Ya") - produk tanpa stok tidak
        // pernah mengurangi stok gudang saat dijual, jadi tidak perlu
        // "dikembalikan" saat transaksinya dihapus.
        if (productUsesStock_(ss, rows[i][productCol])) {
          updateStock_(ss, cfg.warehouseId, rows[i][productCol], Number(rows[i][qtyCol]), undefined, { type: STOCK_CARD_TYPE.SALE_DELETE, refId: saleId, user: user.fullName });
        }
        sheet.deleteRow(i + 1);

        const paySheet = ss.getSheetByName(cfg.paySheet);
        let payRows = paySheet.getDataRange().getValues();
        let payHeaders = payRows[0];
        let refCol = payHeaders.indexOf("Reference");
        for (let j = payRows.length - 1; j >= 1; j--) {
          if (payRows[j][refCol] === saleId) paySheet.deleteRow(j + 1);
        }

        logActivity_(user.username, "DELETE_SALES", cfg.label, "Hapus transaksi " + saleId, "");
        return { success: true, message: "Transaksi berhasil dihapus dan stok dikembalikan." };
      }
    }
    return { success: false, message: "Transaksi tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 9. PAYMENT - PENERIMAAN UANG (generik per unit)
// =====================================================================
function getPaymentData(unit, token) {
  try {
    checkPermission_(token, null, unit);
    const ss = getDatabase();
    return { success: true, data: getSheetDataAsObjects_(ss.getSheetByName(unitCfg_(unit).paySheet)) };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function savePaymentTransaction(unit, data, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.KASIR], unit);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(unitCfg_(unit).paySheet);
    let now = new Date();

    if (data.isEdit) {
      let rows = sheet.getDataRange().getValues();
      let headers = rows[0];
      let idCol = headers.indexOf("PaymentID");
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][idCol] === data.paymentId) {
          sheet.getRange(i + 1, 4, 1, 5).setValues([[data.paymentMethod, Number(data.amount), data.reference || "-", user.fullName, data.notes || "-"]]);
          return { success: true, message: "Data penerimaan berhasil diperbarui." };
        }
      }
      return { success: false, message: "Data tidak ditemukan." };
    } else {
      sheet.appendRow([genId_("PAY"), data.date ? new Date(data.date) : now,
        Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss"),
        data.paymentMethod, Number(data.amount), data.reference || "Manual Entry", user.fullName, data.notes || "-"]);
      logActivity_(user.username, "CREATE_PAYMENT", unit, "Input penerimaan manual", "");
      return { success: true, message: "Penerimaan berhasil dicatat." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deletePaymentTransaction(unit, paymentId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.KASIR], unit);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(unitCfg_(unit).paySheet);
    let rows = sheet.getDataRange().getValues();
    let headers = rows[0];
    let idCol = headers.indexOf("PaymentID");
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idCol] === paymentId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Data penerimaan berhasil dihapus." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 10. STOCK - MANAJEMEN STOK (generik per unit)
// =====================================================================
function getStockView(unit, token) {
  try {
    checkPermission_(token, null, unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    let stockRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK)).filter(s => s.WarehouseID === cfg.warehouseId);
    let stockMap = {};
    stockRows.forEach(s => stockMap[s.ProductCode] = s);

    let allProducts = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let productMap = {};
    allProducts.forEach(p => productMap[p.ProductCode] = p);

    // Tampilkan SEMUA produk AKTIF yang dijual outlet ini (berlaku untuk
    // unit ini atau "General"), bukan cuma yang sudah punya baris stok.
    // Produk yang belum pernah dicatat stoknya akan tampil dengan qty 0.
    // Produk berstatus Inactive TIDAK ditampilkan sama sekali di sini,
    // termasuk kalau kebetulan masih punya baris riwayat stok lama.
    // Hanya produk yang (1) Status = "Active" DAN (2) UsesStock = "Ya"
    // (default kalau kolomnya kosong) yang relevan di submenu Manajemen
    // Stok - produk "Tidak Pakai Stok" tidak perlu dilacak/ditampilkan di
    // sini sama sekali (baik di tabel utama, dropdown Set Stok Manual,
    // dropdown Tambah Stok dari Gudang Pusat, maupun panel Stok Gudang Pusat).
    // Produk dengan Acuan Stok (StockRefCode terisi) tidak pernah punya
    // baris stok sendiri (stoknya ikut produk acuan) - tidak ditampilkan
    // sebagai baris terpisah di Manajemen Stok/dropdown Set Stok Manual/
    // Tambah Stok dari Gudang, supaya angka stok di sini tidak dobel/
    // membingungkan. Kelola stoknya lewat produk acuannya saja.
    let applicableProducts = allProducts.filter(p => {
      if (p.Status !== "Active") return false;
      if (String(p.UsesStock || "Ya").trim() === "Tidak") return false;
      if (String(p.StockRefCode || "").trim()) return false;
      let units = productUnits_(p.ApplicableUnit);
      return units.indexOf(unit) !== -1 || units.indexOf("General") !== -1;
    });

    let result = applicableProducts.map(p => {
      let s = stockMap[p.ProductCode];
      return {
        ProductCode: p.ProductCode, ProductName: p.ProductName, Category: p.Category,
        Quantity: s ? s.Quantity : 0, MinStock: s ? s.MinStock : 10,
        LastUpdated: s ? s.LastUpdated : null, HasRecord: !!s
      };
    });

    // "Tambah Stok dari Gudang": tampilkan SEMUA produk yang berlaku untuk
    // outlet ini (bukan cuma yang sudah pernah diisi stok gudangnya), supaya
    // dropdown-nya tidak pernah kosong. Qty gudang = 0 kalau belum pernah
    // diisi - bisa diisi lewat panel "Isi Stok Gudang Pusat" di sebelahnya.
    let centralMap = {};
    getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK))
      .filter(s => s.WarehouseID === CENTRAL_WAREHOUSE_ID)
      .forEach(s => centralMap[s.ProductCode] = s.Quantity);
    let centralStock = applicableProducts.map(p => ({
      ProductCode: p.ProductCode, ProductName: p.ProductName,
      Quantity: centralMap[p.ProductCode] || 0
    }));

    let history = getSheetDataAsObjects_(ss.getSheetByName(cfg.stockLogSheet)).slice(-50).reverse();

    return { success: true, stock: result, centralStock: centralStock, history: history };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Set langsung stok outlet untuk sebuah produk (buat baru kalau belum ada,
// update kalau sudah ada). Dipakai fitur Edit/Tambah manual di Manajemen Stok.
// Dikunci untuk role Kasir - Set Stok Manual cuma boleh Owner/Admin, supaya
// koreksi stok manual selalu terverifikasi (Kasir tetap bisa lihat & minta
// transfer dari Gudang Pusat lewat addStockFromGudang di bawah).
function setOutletStock(unit, productCode, quantity, minStock, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN], unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    quantity = Number(quantity) || 0;
    minStock = Number(minStock) || 0;

    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    let data = sheet.getDataRange().getValues();
    let found = false;
    let oldQty = 0;
    let info = warehouseProductInfo_(ss, cfg.warehouseId, productCode);
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === cfg.warehouseId && data[i][3] === productCode) {
        oldQty = Number(data[i][5]) || 0;
        sheet.getRange(i + 1, 3, 1, 7).setValues([[info.warehouseName, productCode, info.productName, quantity, minStock, info.status, new Date()]]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([genId_("STK"), cfg.warehouseId, info.warehouseName, productCode, info.productName, quantity, minStock, info.status, new Date()]);
    }

    ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS).appendRow([genId_("MUT"), new Date(), productCode,
      "-", cfg.warehouseId, quantity, "Manual Adjustment", "", user.fullName]);
    appendStockCardRows_(ss, [buildStockCardRow_(cfg.warehouseId, info.warehouseName, productCode, info.productName, oldQty, quantity,
      { type: STOCK_CARD_TYPE.MANUAL, user: user.fullName, notes: "Set stok manual (Manajemen Stok " + cfg.label + ")" }, "")]);
    logActivity_(user.username, "SET_STOCK", cfg.label, "Set stok manual " + productCode + " = " + quantity, "");
    return { success: true, message: "Stok berhasil disimpan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Versi "grid" dari setOutletStock: simpan Qty & Min. Stok untuk BANYAK
// produk sekaligus dalam satu panggilan - dipakai tombol "Simpan Semua" di
// tabel grid Manajemen Stok, supaya Owner/Admin tidak perlu edit/simpan
// baris satu per satu. rows: [{ productCode, quantity, minStock }, ...].
// Baris dengan productCode kosong atau quantity tidak valid dilewati saja
// (tidak menggagalkan seluruh batch).
function bulkSetOutletStock(unit, rows, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN], unit);
    if (!rows || !rows.length) return { success: false, message: "Tidak ada data stok untuk disimpan." };

    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    const now = new Date();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    let data = sheet.getDataRange().getValues();
    let rowIndexByCode = {}; // ProductCode -> nomor baris sheet (1-based), khusus gudang outlet ini
    let qtyByCode = {};      // ProductCode -> qty terakhir (untuk saldo sebelum di kartu stok)
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === cfg.warehouseId) {
        rowIndexByCode[data[i][3]] = i + 1;
        qtyByCode[data[i][3]] = Number(data[i][5]) || 0;
      }
    }

    let savedCount = 0;
    let mutationRows = [];
    let cardRows = [];
    rows.forEach(r => {
      let productCode = r.productCode;
      let quantity = Number(r.quantity);
      let minStock = Number(r.minStock);
      if (!productCode || isNaN(quantity) || quantity < 0) return;
      if (isNaN(minStock) || minStock < 0) minStock = 10;

      let info = warehouseProductInfo_(ss, cfg.warehouseId, productCode);
      let rIdx = rowIndexByCode[productCode];
      let oldQty = qtyByCode[productCode] !== undefined ? qtyByCode[productCode] : 0;
      if (rIdx) {
        sheet.getRange(rIdx, 3, 1, 7).setValues([[info.warehouseName, productCode, info.productName, quantity, minStock, info.status, now]]);
      } else {
        sheet.appendRow([genId_("STK"), cfg.warehouseId, info.warehouseName, productCode, info.productName, quantity, minStock, info.status, now]);
        rowIndexByCode[productCode] = sheet.getLastRow();
      }
      qtyByCode[productCode] = quantity;
      mutationRows.push([genId_("MUT"), now, productCode, "-", cfg.warehouseId, quantity, "Manual Adjustment (Grid)", "", user.fullName]);
      cardRows.push(buildStockCardRow_(cfg.warehouseId, info.warehouseName, productCode, info.productName, oldQty, quantity,
        { type: STOCK_CARD_TYPE.MANUAL, user: user.fullName, notes: "Set stok manual (grid Manajemen Stok " + cfg.label + ")", timestamp: now }, ""));
      savedCount++;
    });

    if (!savedCount) return { success: false, message: "Tidak ada data valid untuk disimpan." };

    let mutSheet = ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS);
    if (mutationRows.length) {
      mutSheet.getRange(mutSheet.getLastRow() + 1, 1, mutationRows.length, mutationRows[0].length).setValues(mutationRows);
    }
    appendStockCardRows_(ss, cardRows);
    logActivity_(user.username, "BULK_SET_STOCK", cfg.label, "Simpan semua stok (grid) - " + savedCount + " produk", "");
    return { success: true, message: "Stok berhasil disimpan untuk " + savedCount + " produk." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Hapus pencatatan stok untuk sebuah produk di outlet ini (produk tetap ada
// di Master Produk, hanya baris stok di outlet tsb yang dihapus).
function deleteOutletStock(unit, productCode, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.KASIR], unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === cfg.warehouseId && data[i][3] === productCode) {
        let oldQty = Number(data[i][5]) || 0;
        let info = warehouseProductInfo_(ss, cfg.warehouseId, productCode);
        sheet.deleteRow(i + 1);
        // Kartu stok: saldo dinolkan supaya rantai saldo tetap sinkron
        // dengan kondisi sheet (baris stok sudah tidak ada).
        appendStockCardRows_(ss, [buildStockCardRow_(cfg.warehouseId, info.warehouseName, productCode, info.productName, oldQty, 0,
          { type: STOCK_CARD_TYPE.RECORD_DELETE, user: user.fullName, notes: "Pencatatan stok dihapus (saldo " + oldQty + " menjadi 0)" }, "")]);
        logActivity_(user.username, "DELETE_STOCK", cfg.label, "Hapus pencatatan stok " + productCode, "");
        return { success: true, message: "Pencatatan stok berhasil dihapus." };
      }
    }
    return { success: false, message: "Data stok tidak ditemukan (produk ini belum punya catatan stok)." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 7b. CEK STOK FISIK BARANG (STOK OPNAME HARIAN PER OUTLET)
// -----------------------------------------------------------------------
// Dipakai kasir untuk memonitor stok harian: sebelum mulai berjualan,
// kasir menghitung stok fisik barangnya dan membandingkannya dengan stok
// yang tercatat di sistem. Hanya produk AKTIF & yang memakai pelacakan
// stok (UsesStock != "Tidak") yang relevan untuk dicek di sini.
//
// CATATAN PERBAIKAN BUG "Stok Awal tidak terbaca" (v3.0.2): versi
// sebelumnya mengambil "Stok Awal" dari StockEnd transaksi TERAKHIR di
// sheet Sales (lihat getLatestStockEndMap_ di bawah, sekarang TIDAK
// dipakai lagi sebagai sumber utama). Itu ternyata bisa salah/basi kalau:
// (a) kolom StockEnd di form Laporan Penjualan diizinkan diedit manual
//     oleh kasir (Pengaturan > Form Input Penjualan) - nilai yang
//     tersimpan di Sales BISA BEDA dari stok gudang sesungguhnya
//     (Warehouse_Stock), karena override itu sengaja hanya memengaruhi
//     angka di laporan, bukan saldo stok riil (lihat manualOverride_ di
//     saveSalesTransaction), atau
// (b) stok outlet ditambah/dikoreksi setelah transaksi penjualan
//     terakhir (mis. lewat "Tambah Stok dari Gudang Pusat" atau "Set
//     Stok Manual" di Manajemen Stok) - Warehouse_Stock sudah update,
//     tapi baris Sales terakhir tidak tahu-menahu soal itu.
// Akibatnya sebagian produk tampil "Stok Awal 0" padahal Warehouse_Stock
// sudah punya angka yang benar. SEKARANG: Stok Awal SELALU diambil
// langsung dari Warehouse_Stock.Quantity (satu-satunya sumber kebenaran
// stok outlet, selalu disinkronkan setiap transaksi/koreksi/opname),
// supaya konsisten dengan logika penyimpanan di savePhysicalStockCheck().
// =====================================================================
// [DEPRECATED - tidak dipakai lagi sebagai sumber Stok Awal, lihat catatan
// di atas. Dibiarkan ada (tidak dihapus) untuk referensi/kompatibilitas
// kalau suatu saat dibutuhkan lagi untuk keperluan lain.]
function getLatestStockEndMap_(salesSheet) {
  let map = {};
  if (!salesSheet || salesSheet.getLastRow() < 2) return map;
  let data = salesSheet.getDataRange().getValues();
  let headers = data[0].map(h => String(h).trim());
  let codeIdx = headers.indexOf("ProductCode");
  let dateIdx = headers.indexOf("Date");
  let stockEndIdx = headers.indexOf("StockEnd");
  if (codeIdx === -1 || stockEndIdx === -1) return map;
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let code = row[codeIdx];
    if (!code) continue;
    let rowDate = dateIdx !== -1 ? row[dateIdx] : null;
    let existing = map[code];
    // Data Sales SEHARUSNYA sudah urut kronologis (append-only), tapi tetap
    // dijaga: hanya timpa kalau baris ini tanggalnya sama/lebih baru,
    // supaya aman kalau suatu saat ada baris yang disisip manual.
    if (!existing || !(rowDate instanceof Date) || !(existing.date instanceof Date) || rowDate >= existing.date) {
      map[code] = { stockEnd: Number(row[stockEndIdx]) || 0, date: rowDate };
    }
  }
  return map;
}

// PERUBAHAN v3.1: sekarang menerima parameter `dateStr` opsional supaya hasil
// Cek Stok Fisik bisa DILIHAT per tanggal (riwayat), bukan cuma hari ini.
// - Kalau dateStr kosong ATAU sama dengan hari ini: perilaku SAMA seperti
//   sebelumnya - menampilkan SEMUA produk aktif+pelacakan stok outlet ini,
//   StokAwal live dari Warehouse_Stock, StokFisik kosong kalau belum pernah
//   disimpan hari ini (form siap diisi & disimpan).
// - Kalau dateStr adalah tanggal LAIN (masa lalu): mode "riwayat" (read-only
//   di sisi client) - hanya menampilkan produk yang BENAR-BENAR punya baris
//   tersimpan di Stock_Opname_Fisik pada tanggal tsb (StokAwal/StokFisik
//   diambil apa adanya dari baris yang tersimpan, tidak dihitung ulang).
// Flag `isToday` dikirim ke client supaya UI tahu kapan harus mengunci form
// & menyembunyikan tombol "Simpan Semua".
function getPhysicalStockCheckView(unit, dateStr, token) {
  try {
    checkPermission_(token, null, unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const targetDate = normalizeDateStr_(dateStr) || todayStr;
    // Kalau tanggal yang diminta ada di masa depan (tidak mungkin ada
    // datanya), perlakukan saja sebagai hari ini supaya tidak membingungkan.
    const effectiveDate = (targetDate > todayStr) ? todayStr : targetDate;
    const effectiveIsToday = (effectiveDate === todayStr);

    // Kalau kasir sudah simpan cek stok fisik pada tanggal ybs, kembalikan
    // lagi nilai StokFisik yang sudah tersimpan supaya bisa dilihat/
    // dikoreksi, bukan form kosong.
    // CATATAN PERFORMA: sheet Stock_Opname_Fisik terus bertambah setiap
    // hari selama-lamanya (tidak ada arsip otomatis), jadi lama-lama bisa
    // jadi puluhan ribu baris. Di sini SENGAJA tidak pakai
    // getSheetDataAsObjects_() (yang mengubah SEMUA baris jadi object)
    // karena itu bikin halaman ini makin lambat/kadang gagal tanpa respon
    // seiring sheet membesar - cukup scan array mentah dan hanya bikin
    // object untuk baris yang benar-benar cocok tanggal & outlet ini.
    let opnameMap = {};
    const opnameSheet = ss.getSheetByName(CONFIG.SHEETS.STOCK_OPNAME);
    if (opnameSheet && opnameSheet.getLastRow() > 1) {
      const opnameData = opnameSheet.getRange(2, 1, opnameSheet.getLastRow() - 1, opnameSheet.getLastColumn()).getValues();
      // Kolom sesuai STOCK_OPNAME_HEADERS: OpnameID,Date,Unit,WarehouseID,ProductCode,ProductName,Category,StokAwal,StokFisik,Selisih,SellPrice,NilaiSelisih,...
      for (let i = 0; i < opnameData.length; i++) {
        const row = opnameData[i];
        if (normalizeDateStr_(row[1]) === effectiveDate && row[3] === cfg.warehouseId) {
          opnameMap[row[4]] = {
            ProductName: row[5], Category: row[6], StokAwal: row[7], StokFisik: row[8], SellPrice: row[10]
          };
        }
      }
    }

    let result;
    if (effectiveIsToday) {
      let allProducts = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
      let applicableProducts = allProducts.filter(p => {
        if (p.Status !== "Active") return false;
        if (String(p.UsesStock || "Ya").trim() === "Tidak") return false;
        // Produk ber-Acuan Stok ikut stok produk acuannya - cek fisik cukup
        // dilakukan lewat produk acuan itu saja, tidak perlu baris terpisah.
        if (String(p.StockRefCode || "").trim()) return false;
        let units = productUnits_(p.ApplicableUnit);
        return units.indexOf(unit) !== -1 || units.indexOf("General") !== -1;
      });

      let stockRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK)).filter(s => s.WarehouseID === cfg.warehouseId);
      let stockMap = {};
      stockRows.forEach(s => stockMap[s.ProductCode] = s);

      // StokAwal = Quantity SAAT INI di Warehouse_Stock untuk outlet ini -
      // satu-satunya sumber kebenaran stok outlet (lihat catatan perbaikan
      // bug di atas fungsi getLatestStockEndMap_). Kalau kasir/Owner/Admin
      // sudah pernah simpan cek stok fisik HARI INI, StokAwal yang dipakai
      // tetap StokAwal yang TERSIMPAN di baris opname tsb (bukan dihitung
      // ulang dari Warehouse_Stock, yang mungkin sudah berubah lagi karena
      // penyesuaian stok fisik tadi) - supaya nilai yang tampil konsisten
      // dengan yang sudah disimpan sampai kasir klik "Refresh Data" di hari
      // berikutnya.
      result = applicableProducts.map(p => {
        let s = stockMap[p.ProductCode];
        let o = opnameMap[p.ProductCode];
        let stokAwal = o ? Number(o.StokAwal) : (s ? Number(s.Quantity) || 0 : 0);
        return {
          ProductCode: p.ProductCode, ProductName: p.ProductName, Category: p.Category,
          SellPrice: Number(p.SellPrice) || 0,
          StokAwal: stokAwal,
          StokFisik: o ? Number(o.StokFisik) : null,
          AlreadyChecked: !!o
        };
      });
    } else {
      // Mode riwayat (tanggal lampau): hanya produk yang benar-benar punya
      // catatan opname di tanggal tsb, apa adanya (tidak digabung dengan
      // Master Produk/Warehouse_Stock terkini).
      result = Object.keys(opnameMap).map(code => {
        let o = opnameMap[code];
        return {
          ProductCode: code, ProductName: o.ProductName, Category: o.Category,
          SellPrice: Number(o.SellPrice) || 0,
          StokAwal: Number(o.StokAwal) || 0,
          StokFisik: Number(o.StokFisik),
          AlreadyChecked: true
        };
      }).sort((a, b) => String(a.ProductName).localeCompare(String(b.ProductName)));
    }

    return { success: true, date: effectiveDate, isToday: effectiveIsToday, products: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// rows: [{ productCode, stokFisik, stokAwal }, ...] dari client (hasil klik
// "Simpan Semua"). `stokAwal` OPSIONAL - dikirim hanya kalau Owner/Admin
// mengoreksi kolom "Stok Awal Hari Ini" secara manual di frontend (lihat
// getPhysicalStockCheckView). Kalau tidak dikirim/bukan Owner-Admin, StokAwal
// dihitung otomatis dari Warehouse_Stock.Quantity saat ini (satu-satunya
// sumber kebenaran stok outlet, lihat catatan di getLatestStockEndMap_).
// Menyimpan 1 baris per produk ke Stock_Opname_Fisik (menimpa baris hari ini
// kalau sudah pernah disimpan sebelumnya - supaya tidak dobel), SEKALIGUS
// menyesuaikan stok outlet di Warehouse_Stock supaya sama dengan hasil stok
// fisik (stok opname) - Manajemen Stok & cek stok fisik besok akan memakai
// angka yang sudah benar ini sebagai Stok Awal.
// PENTING: stok outlet akhir SELALU di-set LANGSUNG ke nilai StokFisik yang
// diinput (bukan ditambah/dikurangi sebesar "Selisih" dari StokAwal seperti
// versi sebelumnya) - stok hasil hitung fisik adalah kebenaran yang paling
// akurat, dan cara ini membuat hasil akhirnya tidak tergantung sama sekali
// pada apakah StokAwal (otomatis ataupun koreksi manual Owner/Admin) akurat
// atau tidak. StokAwal & Selisih yang tersimpan tetap dipakai murni untuk
// keperluan pelaporan (menunjukkan seberapa besar penyesuaian yang terjadi).
function savePhysicalStockCheck(unit, rows, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.KASIR, ROLE.STAFF], unit);
    if (!rows || !rows.length) return { success: false, message: "Tidak ada data stok fisik untuk disimpan." };

    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const now = new Date();
    let canOverrideAwal = (user.role === ROLE.OWNER || user.role === ROLE.ADMIN);

    let productMap = {};
    getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS)).forEach(p => productMap[p.ProductCode] = p);

    const sheet = ss.getSheetByName(CONFIG.SHEETS.STOCK_OPNAME);
    let data = sheet.getDataRange().getValues();
    let existingRow = {}; // ProductCode -> nomor baris sheet (1-based)
    for (let i = 1; i < data.length; i++) {
      if (normalizeDateStr_(data[i][1]) === todayStr && data[i][3] === cfg.warehouseId) {
        existingRow[data[i][4]] = i + 1;
      }
    }

    let savedCount = 0, totalStokAwal = 0, totalSelisih = 0, totalNilaiSelisih = 0;

    rows.forEach(r => {
      let productCode = r.productCode;
      let stokFisik = Number(r.stokFisik);
      let product = productMap[productCode];
      if (!productCode || !product || isNaN(stokFisik) || stokFisik < 0) return;

      // StokAwal "resmi" (dipakai untuk menghitung delta stok riil di
      // Warehouse_Stock) SELALU dari stok outlet saat ini - lihat catatan
      // di atas fungsi. Kalau Owner/Admin mengirim koreksi manual & itu
      // angka valid, koreksi itu dipakai untuk kolom StokAwal & Selisih
      // yang DITAMPILKAN/disimpan sebagai laporan saja.
      let actualCurrentQty = getCurrentStock_(ss, cfg.warehouseId, productCode);
      let stokAwal = actualCurrentQty;
      if (canOverrideAwal && r.stokAwal !== undefined && r.stokAwal !== null && r.stokAwal !== "" && !isNaN(Number(r.stokAwal))) {
        stokAwal = Number(r.stokAwal);
      }
      let selisih = stokFisik - stokAwal;
      let sellPrice = Number(product.SellPrice) || 0;
      let nilaiSelisih = selisih * sellPrice;

      let existingRowNum = existingRow[productCode];
      let opnameId = existingRowNum ? data[existingRowNum - 1][0] : genId_("OPN");
      let rowValues = [opnameId, todayStr, unit, cfg.warehouseId, productCode, product.ProductName, product.Category,
        stokAwal, stokFisik, selisih, sellPrice, nilaiSelisih, user.fullName, now];

      if (existingRowNum) {
        sheet.getRange(existingRowNum, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }

      // Stok outlet akhir di-set LANGSUNG ke hasil hitung fisik (lihat
      // catatan di atas fungsi) - delta riil yang dicatat di Stock_Mutations
      // dihitung dari stok AKTUAL saat ini, bukan dari StokAwal (yang bisa
      // saja sudah dikoreksi manual & tidak selalu sama dengan stok aktual).
      let realDelta = stokFisik - actualCurrentQty;
      if (realDelta !== 0) {
        setStockAbsolute_(ss, cfg.warehouseId, productCode, stokFisik, undefined, {
          type: STOCK_CARD_TYPE.OPNAME, refId: opnameId, user: user.fullName, timestamp: now,
          notes: "Stok fisik " + stokFisik + " (tercatat sistem " + actualCurrentQty + ")"
        });
        ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS).appendRow([genId_("MUT"), now, productCode,
          "-", cfg.warehouseId, stokFisik, "Stock Opname Harian", opnameId, user.fullName]);
      }

      savedCount++;
      totalStokAwal += stokAwal; totalSelisih += selisih; totalNilaiSelisih += nilaiSelisih;
    });

    if (!savedCount) return { success: false, message: "Tidak ada data valid untuk disimpan." };

    logActivity_(user.username, "STOCK_OPNAME", cfg.label, "Simpan cek stok fisik " + savedCount + " produk (tanggal " + todayStr + ")", "");
    return {
      success: true, message: "Stok fisik berhasil disimpan (" + savedCount + " produk).",
      summary: { totalStokAwal: totalStokAwal, totalSelisih: totalSelisih, totalNilaiSelisih: totalNilaiSelisih }
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function addStockFromGudang(unit, productCode, qty, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.KASIR], unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    qty = Number(qty);
    if (!qty || qty <= 0) return { success: false, message: "Jumlah harus lebih dari 0." };

    let centralStock = getCurrentStock_(ss, CENTRAL_WAREHOUSE_ID, productCode);
    if (centralStock < qty) return { success: false, message: "Stok gudang pusat tidak mencukupi (sisa: " + centralStock + "). Isi dulu lewat panel \"Isi Stok Gudang Pusat\"." };

    const trfRef = genId_("TRF");
    updateStock_(ss, CENTRAL_WAREHOUSE_ID, productCode, -qty, undefined, { type: STOCK_CARD_TYPE.TRANSFER_OUT, counterpart: cfg.warehouseId, refId: trfRef, user: user.fullName });
    updateStock_(ss, cfg.warehouseId, productCode, qty, undefined, { type: STOCK_CARD_TYPE.TRANSFER_IN, counterpart: CENTRAL_WAREHOUSE_ID, refId: trfRef, user: user.fullName });

    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let product = products.find(p => p.ProductCode === productCode);
    let now = new Date();

    ss.getSheetByName(cfg.stockLogSheet).appendRow([genId_("REQ"), now, productCode,
      product ? product.ProductName : productCode, qty, CENTRAL_WAREHOUSE_ID, cfg.warehouseId, "Selesai", user.fullName, now]);
    ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS).appendRow([genId_("MUT"), now, productCode,
      CENTRAL_WAREHOUSE_ID, cfg.warehouseId, qty, "Stock Transfer", "", user.fullName]);

    logActivity_(user.username, "STOCK_TRANSFER", cfg.label, "Tambah stok " + productCode + " qty " + qty, "");
    return { success: true, message: "Stok berhasil ditambahkan dari gudang." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Isi/top-up stok Gudang Pusat (sumber utama sebelum ditransfer ke outlet
// lewat "Tambah Stok dari Gudang"). Owner/Admin saja.
function addCentralStock(productCode, qty, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN]);
    const ss = getDatabase();
    qty = Number(qty);
    if (!qty || qty <= 0) return { success: false, message: "Jumlah harus lebih dari 0." };

    let newQty = updateStock_(ss, CENTRAL_WAREHOUSE_ID, productCode, qty, undefined, { type: STOCK_CARD_TYPE.CENTRAL_IN, counterpart: "Supplier/Manual", user: user.fullName });

    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let product = products.find(p => p.ProductCode === productCode);
    ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS).appendRow([genId_("MUT"), new Date(), productCode,
      "Supplier/Manual", CENTRAL_WAREHOUSE_ID, qty, "Central Stock In", "", user.fullName]);

    logActivity_(user.username, "CENTRAL_STOCK_IN", "Gudang Pusat", "Isi stok " + productCode + " qty " + qty, "");
    return { success: true, message: "Stok gudang pusat berhasil diisi. Stok sekarang: " + newQty };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 10b. WAREHOUSE MODULE - manajemen multi-gudang & monitoring stok
// -----------------------------------------------------------------------
// Menggantikan menu "Dapur" (placeholder) - modul ini mengelola SEMUA
// gudang penyimpanan (bukan cuma Gudang Pusat), termasuk gudang yang
// terhubung ke outlet kasir (Utama/Loket Masuk/Loket Wahana/Resto) MAUPUN
// gudang internal yang tidak punya kasir/POS sama sekali (Kolam Renang,
// Teknisi, Cleaning Service). Alur transfer & pengurangan stok saat
// penjualan TETAP memakai mesin stok yang sudah ada (updateStock_/
// getCurrentStock_ di WAREHOUSE_STOCK) - jadi begitu Kasir Resto menjual,
// stoknya otomatis berkurang dan langsung terlihat real-time di modul ini.
// =====================================================================
const INTERNAL_WAREHOUSE_IDS = ["WH05", "WH06", "WH07", "WH08"]; // gudang tanpa kasir/POS
// Peta WarehouseID gudang internal -> nama sheet riwayat KHUSUS gudang itu
// (lihat INTERNAL_WH_LOG_HEADER) - dipakai recordWarehouseUsage(),
// transferStock(), & getWarehouseStockDetail() supaya riwayat Gudang
// Teknisi/Kolam/Cleaning Service/Dapur tercatat di sheet masing-masing,
// bukan cuma bercampur di Stock_Mutations.
const INTERNAL_WAREHOUSE_LOG_SHEET_ = {
  WH05: CONFIG.SHEETS.KOLAM_STOCK,
  WH06: CONFIG.SHEETS.TEKNISI_STOCK,
  WH07: CONFIG.SHEETS.CLEANING_STOCK,
  WH08: CONFIG.SHEETS.DAPUR_STOCK
};
const WAREHOUSE_ITEM_TYPES = ["Konsumsi", "Peralatan", "Bahan", "Lainnya"];

// Kebalikan dari UNIT_MAP: dari WarehouseID, cari tahu apakah gudang itu
// terhubung ke sebuah outlet kasir (dipakai supaya transfer stok ke/dari
// gudang outlet tetap tercatat di riwayat "Manajemen Stok" outlet ybs).
function unitByWarehouseId_(warehouseId) {
  for (let key in UNIT_MAP) { if (UNIT_MAP[key].warehouseId === warehouseId) return key; }
  return null;
}

// Peta WarehouseID -> UnitKey Applicable_Units yang berlaku UNTUK FILTER
// PRODUK gudang tsb (dipakai getWarehouseStockDetail() SAJA). Beda dengan
// unitByWarehouseId_() di atas (yang cuma tahu 3 outlet kasir ber-POS dan
// dipakai untuk mencari stockLogSheet lewat unitCfg_) - peta ini mencakup
// SEMUA 7 gudang non-Utama, termasuk 4 gudang internal (Kolam/Teknisi/
// Cleaning Service/Dapur), supaya tiap gudang internal juga hanya
// menampilkan produk yang ApplicableUnit-nya memang ditugaskan ke gudang
// itu (atau "General"). Gudang Utama (WH01) sengaja TIDAK dipetakan
// (null) karena tetap menampilkan SEMUA produk aktif seperti sebelumnya.
const WAREHOUSE_APPLICABLE_UNIT_ = {
  WH02: "Entrance", WH03: "Ride", WH04: "Resto",
  WH05: "Kolam", WH06: "Teknisi", WH07: "Cleaning", WH08: "Dapur"
};
function applicableUnitByWarehouseId_(warehouseId) {
  return WAREHOUSE_APPLICABLE_UNIT_[warehouseId] || null;
}

// Daftar tunggal 8 gudang yang seharusnya selalu ada (dipakai baik oleh
// migrasi otomatis ensureWarehousesSeed_ MAUPUN oleh Diagnostic Database -
// supaya keduanya tidak pernah beda pendapat soal gudang apa saja yang
// seharusnya ada di sheet Warehouses).
const CANONICAL_WAREHOUSES_ = [
  ["WH01", "Gudang Utama", "Area Kantor Utama", "Central"],
  ["WH02", "Gudang Kantor dan Kasir Loket Masuk", "Gerbang Depan", "Entrance"],
  ["WH03", "Gudang Kasir Loket Wahana", "Area Wahana", "Ride"],
  ["WH04", "Gudang Kasir Resto", "Food Court", "Resto"],
  ["WH05", "Gudang Kolam Renang", "Area Kolam Renang", "Kolam"],
  ["WH06", "Gudang Teknisi", "Area Teknisi/Maintenance", "Teknisi"],
  ["WH07", "Gudang Cleaning Service", "Area Cleaning Service", "Cleaning"],
  ["WH08", "Gudang Dapur", "Area Dapur", "Dapur"]
];

// MIGRASI: tambahkan gudang-gudang baru (Kolam Renang/Teknisi/Cleaning
// Service) kalau belum ada di sheet Warehouses, dan rapikan nama 4 gudang
// lama supaya konsisten dengan penamaan resmi terbaru. Aman & idempoten,
// dijalankan tiap kali getDatabase() dipanggil.
function ensureWarehousesSeed_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES);
    if (!sheet) return;
    let data = sheet.getDataRange().getValues();
    let idRow = {};
    for (let i = 1; i < data.length; i++) idRow[data[i][0]] = i + 1;

    CANONICAL_WAREHOUSES_.forEach(row => {
      if (idRow[row[0]]) {
        sheet.getRange(idRow[row[0]], 2, 1, 3).setValues([[row[1], row[2], row[3]]]);
      } else {
        sheet.appendRow(row);
      }
    });
  } catch (e) {}
}

// MIGRASI: pastikan sheet Product_Types ada (database lama yang dibuat
// sebelum fitur "Tipe Produk bisa diatur sendiri" ditambahkan) & sudah
// terisi 4 tipe bawaan (Konsumsi/Peralatan/Bahan/Lainnya) kalau sheetnya
// masih kosong, supaya dropdown "Tipe Barang" di Master Produk & tab
// kategori di modul Warehouse tidak pernah kosong. Aman & idempoten,
// dijalankan tiap kali getDatabase() dipanggil.
function ensureProductTypesSeed_(ss) {
  try {
    const sheet = ensureSheet_(ss, CONFIG.SHEETS.PRODUCT_TYPES, PRODUCT_TYPE_HEADERS);
    if (sheet.getLastRow() < 2) {
      sheet.getRange(2, 1, 4, 3).setValues([
        ["TYP01", "Konsumsi", "Active"],
        ["TYP02", "Peralatan", "Active"],
        ["TYP03", "Bahan", "Active"],
        ["TYP04", "Lainnya", "Active"]
      ]);
    }
  } catch (e) {}
}

// Ambil semua nama Tipe Produk (Active maupun Inactive) dari sheet
// Product_Types - dipakai untuk klasifikasi ItemType produk di modul
// Warehouse (Ringkasan/Detail/Monitor Gudang). "Lainnya" selalu disertakan
// sebagai fallback default kalau belum ada di daftar, supaya produk dengan
// ItemType yang tidak/ belum dikenal (mis. tipe lama yang sudah dihapus
// dari master) tetap punya kategori tampung yang konsisten.
function getProductTypeNamesAll_(ss) {
  let rows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCT_TYPES));
  let names = rows.map(r => String(r.TypeName || "").trim()).filter(Boolean);
  if (names.indexOf("Lainnya") === -1) names.push("Lainnya");
  return names;
}

// MIGRASI: tambahkan kolom ItemType (Konsumsi/Peralatan/Bahan/Lainnya) di
// sheet Products kalau belum ada. Ini klasifikasi barang untuk modul
// Warehouse, terpisah dari Category (yang tetap dipakai untuk kategori
// penjualan/POS seperti "Tiket Masuk"/"Makanan & Minuman").
function ensureProductsItemTypeColumn_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headers.indexOf("ItemType") === -1) {
      sheet.getRange(1, headers.length + 1).setValue("ItemType")
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  } catch (e) {}
}

// MIGRASI: tambahkan kolom UsesStock (Ya/Tidak) di sheet Products kalau
// belum ada. Kolom ini menandai produk mana yang memakai pelacakan stok
// gudang otomatis (dipakai saveSalesTransaction/deleteSalesTransaction).
// Data lama otomatis diisi "Ya" supaya perilaku stok yang sudah berjalan
// untuk produk yang sudah ada TIDAK berubah sama sekali.
function ensureProductsUsesStockColumn_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headers.indexOf("UsesStock") === -1) {
      let col = headers.length + 1;
      sheet.getRange(1, col).setValue("UsesStock")
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
      let lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        let fill = [];
        for (let i = 0; i < lastRow - 1; i++) fill.push(["Ya"]);
        sheet.getRange(2, col, fill.length, 1).setValues(fill);
      }
    }
  } catch (e) {}
}

// MIGRASI: tambahkan kolom UsesSerial (Ya/Tidak) di sheet Products kalau
// belum ada. Kolom ini menandai produk mana yang perlu menampilkan field
// Nomor Seri Awal/Akhir di Form Input Penjualan (khusus outlet yang memang
// punya kolom seri di database, lihat UNIT_MAP.hasSerial). Data lama
// otomatis diisi "Ya" supaya perilaku yang sudah berjalan (field seri
// selalu tampil di outlet Entrance/Ride) TIDAK berubah sama sekali.
function ensureProductsUsesSerialColumn_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headers.indexOf("UsesSerial") === -1) {
      let col = headers.length + 1;
      sheet.getRange(1, col).setValue("UsesSerial")
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
      let lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        let fill = [];
        for (let i = 0; i < lastRow - 1; i++) fill.push(["Ya"]);
        sheet.getRange(2, col, fill.length, 1).setValues(fill);
      }
    }
  } catch (e) {}
}

// MIGRASI: tambahkan kolom StockRefCode (Kode Barang Acuan Stok) di sheet
// Products kalau belum ada. Kolom ini dipakai untuk produk yang TIDAK
// punya stok fisik sendiri, melainkan berbagi/memotong stok dari produk
// LAIN (mis. "Tiket Terusan Plus Diskon 20%" berbagi stok fisik dengan
// "Tiket Terusan Plus") - lihat resolveStockProductCode_() yang dipakai
// otomatis di getCurrentStock_/updateStock_/setStockAbsolute_ setiap ada
// transaksi/opname/mutasi stok. Data lama otomatis dikosongkan (semua
// produk lama tetap memakai stoknya sendiri seperti biasa, perilaku yang
// sudah berjalan TIDAK berubah).
function ensureProductsStockRefColumn_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headers.indexOf("StockRefCode") === -1) {
      sheet.getRange(1, headers.length + 1).setValue("StockRefCode")
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  } catch (e) {}
}

// MIGRASI: tambahkan kolom HargaFleksibel (Ya/Tidak) di sheet Products
// kalau belum ada - menandai produk yang harga jualnya boleh diubah
// manual saat transaksi (mis. barang re-sell/konsinyasi). Data lama
// otomatis diisi "Tidak" (default aman, harga tetap mengikuti Master
// Produk seperti perilaku yang sudah berjalan selama ini).
function ensureProductsHargaFleksibelColumn_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headers.indexOf("HargaFleksibel") === -1) {
      let col = headers.length + 1;
      sheet.getRange(1, col).setValue("HargaFleksibel")
        .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
      let lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        let fill = [];
        for (let i = 0; i < lastRow - 1; i++) fill.push(["Tidak"]);
        sheet.getRange(2, col, fill.length, 1).setValues(fill);
      }
    }
  } catch (e) {}
}

// Cek apakah sebuah produk memakai pelacakan stok gudang (kolom UsesStock
// di Master Produk). Default TRUE (Ya) kalau produk tidak ditemukan atau
// kolomnya kosong, supaya kompatibel dengan data lama.
function productUsesStock_(ss, productCode) {
  let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
  let product = products.find(p => p.ProductCode === productCode);
  if (!product) return true;
  return String(product.UsesStock || "Ya").trim() !== "Tidak";
}

// Cek apakah sebuah produk memakai Nomor Seri Awal/Akhir (kolom UsesSerial
// di Master Produk). Default TRUE (Ya) kalau produk tidak ditemukan atau
// kolomnya kosong, supaya kompatibel dengan data lama. Ini hanya relevan
// untuk outlet yang memang punya kolom seri di database (cfg.hasSerial) -
// dan tidak berlaku sama sekali untuk produk yang UsesStock = "Tidak"
// (lihat buildSalesRow_/saveSalesTransaction).
function productUsesSerial_(ss, productCode) {
  let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
  let product = products.find(p => p.ProductCode === productCode);
  if (!product) return true;
  return String(product.UsesSerial || "Ya").trim() !== "Tidak";
}

// Ambil nama gudang & info produk (nama, status) - dipakai mengisi kolom
// "manusiawi" (WarehouseName/ProductName/Status) di sheet Warehouse_Stock
// setiap kali baris dibuat/diupdate, supaya sheet-nya enak dibaca langsung
// tanpa harus VLOOKUP manual ke sheet Warehouses/Products.
function warehouseProductInfo_(ss, warehouseId, productCode) {
  let warehouses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES));
  let warehouse = warehouses.find(w => w.WarehouseID === warehouseId);
  let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
  let product = products.find(p => p.ProductCode === productCode);
  return {
    warehouseName: warehouse ? warehouse.WarehouseName : warehouseId,
    productName: product ? product.ProductName : productCode,
    status: product ? (product.Status || "Active") : "Active"
  };
}

// MIGRASI: skema Warehouse_Stock lama hanya berisi ID mentah (StockID,
// WarehouseID, ProductCode, Quantity, MinStock, LastUpdated) sehingga sulit
// dibaca langsung dari Google Sheets. Migrasi ini menata ulang sheet supaya
// juga menampilkan WarehouseName, ProductName, & Status (Active/Inactive)
// produk ybs di kolomnya sendiri, TANPA mengubah data stok (Quantity/
// MinStock) itu sendiri. Aman & idempoten - dideteksi dari ada/tidaknya
// header "WarehouseName".
function ensureWarehouseStockColumns_(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    if (!sheet) return;
    let lastCol = sheet.getLastColumn();
    let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (headers.indexOf("WarehouseName") !== -1) return; // sudah dimigrasi

    let lastRow = sheet.getLastRow();
    let oldData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];

    let warehouses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES));
    let warehouseNameMap = {};
    warehouses.forEach(w => warehouseNameMap[w.WarehouseID] = w.WarehouseName);
    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let productMap = {};
    products.forEach(p => productMap[p.ProductCode] = p);

    // Kolom lama (urutan tetap): StockID(0) WarehouseID(1) ProductCode(2)
    // Quantity(3) MinStock(4) LastUpdated(5)
    let newRows = oldData.map(row => {
      let warehouseId = row[1], productCode = row[2];
      let prod = productMap[productCode];
      return [
        row[0], warehouseId, warehouseNameMap[warehouseId] || warehouseId,
        productCode, prod ? prod.ProductName : productCode,
        row[3], row[4], prod ? (prod.Status || "Active") : "Active", row[5]
      ];
    });

    sheet.clear();
    sheet.getRange(1, 1, 1, WAREHOUSE_STOCK_HEADERS.length).setValues([WAREHOUSE_STOCK_HEADERS])
      .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    if (newRows.length) sheet.getRange(2, 1, newRows.length, WAREHOUSE_STOCK_HEADERS.length).setValues(newRows);
  } catch (e) {}
}

// Daftar semua gudang - dipakai untuk dropdown/pilihan gudang di modul Warehouse.
function getWarehouseList(token) {
  try {
    checkPermission_(token, null);
    const ss = getDatabase();
    let warehouses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES));
    return { success: true, warehouses: warehouses };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Ringkasan seluruh gudang: total SKU, total qty, breakdown per ItemType
// (Konsumsi/Peralatan/Bahan/Lainnya), & jumlah item stok menipis - dipakai
// kartu ringkasan di tab "Ringkasan Gudang" modul Warehouse.
function getWarehouseOverview(token) {
  try {
    checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    const ss = getDatabase();
    let warehouses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES));
    let stockRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK));
    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let productMap = {};
    products.forEach(p => productMap[p.ProductCode] = p);
    let productTypeNames = getProductTypeNamesAll_(ss);

    let result = warehouses.map(w => {
      let rows = stockRows.filter(s => s.WarehouseID === w.WarehouseID);
      let totalQty = 0, totalSKU = 0, lowStock = 0;
      let byType = {};
      productTypeNames.forEach(t => byType[t] = 0);
      rows.forEach(s => {
        let qty = Number(s.Quantity) || 0;
        totalSKU++;
        totalQty += qty;
        let prod = productMap[s.ProductCode];
        // Produk Inactive di Master Produk tidak perlu memicu peringatan
        // "Stok Menipis" (sama seperti sudah tidak tampil di stok kasir).
        let isActiveProduct = !prod || prod.Status === "Active";
        if (isActiveProduct && qty <= Number(s.MinStock)) lowStock++;
        let type = (prod && productTypeNames.indexOf(prod.ItemType) !== -1) ? prod.ItemType : "Lainnya";
        byType[type] += qty;
      });
      return {
        WarehouseID: w.WarehouseID, WarehouseName: w.WarehouseName, Location: w.Location, UnitType: w.UnitType,
        totalSKU: totalSKU, totalQty: totalQty, lowStock: lowStock, byType: byType,
        isInternal: INTERNAL_WAREHOUSE_IDS.indexOf(w.WarehouseID) !== -1
      };
    });
    return { success: true, warehouses: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Detail stok 1 gudang tertentu (dipakai tab "Detail & Stok Gudang").
// Gudang outlet (WH02-WH04) & gudang internal (WH05-WH08/Kolam-Teknisi-
// Cleaning Service-Dapur): hanya tampilkan produk yang ApplicableUnit-nya
// memang ditugaskan ke gudang tsb (atau "General") - lihat
// WAREHOUSE_APPLICABLE_UNIT_/applicableUnitByWarehouseId_. Dengan ini,
// mis. Gudang Teknisi hanya menampilkan produk/barang yang di Master
// Produk dicentang Outlet Penjualan "Teknisi" (atau "General"). HANYA
// Gudang Utama (WH01) yang tetap menampilkan SEMUA produk aktif, karena
// gudang ini jadi tempat penyimpanan pusat segala kategori barang.
function getWarehouseStockDetail(warehouseId, token) {
  try {
    checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    const ss = getDatabase();
    let warehouses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES));
    let warehouse = warehouses.find(w => w.WarehouseID === warehouseId);
    if (!warehouse) return { success: false, message: "Gudang tidak ditemukan." };

    let stockRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK)).filter(s => s.WarehouseID === warehouseId);
    let stockMap = {};
    stockRows.forEach(s => stockMap[s.ProductCode] = s);

    // Hanya tampilkan produk yang: (1) Status = "Active" di Master Produk,
    // dan (2) UsesStock = "Ya" (produk tanpa fisik/tanpa pelacakan stok
    // seperti jasa TIDAK perlu tampil di modul Warehouse gudang manapun,
    // termasuk Gudang Utama, gudang kasir loket masuk/wahana/resto, gudang
    // teknisi, gudang dapur, & gudang cleaning service).
    // Produk ber-Acuan Stok tidak ditampilkan sebagai baris sendiri di sini
    // (stoknya ikut produk acuan - lihat resolveStockProductCode_), supaya
    // data stok yang tampil di Detail Gudang & Cek Stok Fisik tidak dobel.
    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS))
      .filter(p => p.Status === "Active" && String(p.UsesStock || "Ya").trim() !== "Tidak" && !String(p.StockRefCode || "").trim());
    let productTypeNames = getProductTypeNamesAll_(ss);
    let unitOfWarehouse = applicableUnitByWarehouseId_(warehouseId);
    let applicable = products.filter(p => {
      if (!unitOfWarehouse) return true;
      let units = productUnits_(p.ApplicableUnit);
      return units.indexOf(unitOfWarehouse) !== -1 || units.indexOf("General") !== -1;
    });

    let result = applicable.map(p => {
      let s = stockMap[p.ProductCode];
      return {
        ProductCode: p.ProductCode, ProductName: p.ProductName, Category: p.Category,
        ItemType: productTypeNames.indexOf(p.ItemType) !== -1 ? p.ItemType : "Lainnya",
        Quantity: s ? s.Quantity : 0, MinStock: s ? s.MinStock : 10,
        LastUpdated: s ? s.LastUpdated : null, HasRecord: !!s
      };
    });

    // Untuk gudang internal (Kolam/Teknisi/Cleaning Service/Dapur), baca
    // riwayat dari sheet KHUSUS gudang itu (lebih ringan/cepat karena
    // sheetnya sudah otomatis cuma berisi baris gudang ini - tidak perlu
    // filter dari Stock_Mutations yang isinya SEMUA gudang & terus membesar).
    // Hasilnya dinormalisasi ke bentuk yang sama seperti baris Stock_Mutations
    // (Timestamp/FromLocation/ToLocation/User) supaya tampilan di frontend
    // tidak perlu diubah sama sekali.
    let history;
    let internalLogSheetName = INTERNAL_WAREHOUSE_LOG_SHEET_[warehouseId];
    if (internalLogSheetName) {
      history = getSheetDataAsObjects_(ss.getSheetByName(internalLogSheetName))
        .slice(-50).reverse()
        .map(r => ({ Timestamp: r.Date, ProductCode: r.ProductCode, FromLocation: r.FromWarehouse, ToLocation: r.ToWarehouse, Qty: r.Qty, User: r.By }));
    } else {
      history = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS))
        .filter(m => m.FromLocation === warehouseId || m.ToLocation === warehouseId)
        .slice(-50).reverse();
    }

    return { success: true, warehouse: warehouse, stock: result, history: history };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Set langsung stok sebuah produk di gudang manapun (generalisasi dari
// setOutletStock, tapi berlaku untuk ke-8 gudang, dipakai modul Warehouse
// oleh Owner/Admin/Staff). Kasir tetap memakai setOutletStock lewat menu
// "Manajemen Stok" outletnya masing-masing seperti sebelumnya.
function setWarehouseStock(warehouseId, productCode, quantity, minStock, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    const ss = getDatabase();
    quantity = Number(quantity) || 0;
    minStock = Number(minStock) || 0;

    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    let data = sheet.getDataRange().getValues();
    let found = false;
    let oldQty = 0;
    let info = warehouseProductInfo_(ss, warehouseId, productCode);
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === warehouseId && data[i][3] === productCode) {
        oldQty = Number(data[i][5]) || 0;
        sheet.getRange(i + 1, 3, 1, 7).setValues([[info.warehouseName, productCode, info.productName, quantity, minStock, info.status, new Date()]]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([genId_("STK"), warehouseId, info.warehouseName, productCode, info.productName, quantity, minStock, info.status, new Date()]);
    }

    ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS).appendRow([genId_("MUT"), new Date(), productCode,
      "-", warehouseId, quantity, "Manual Adjustment", "", user.fullName]);
    appendStockCardRows_(ss, [buildStockCardRow_(warehouseId, info.warehouseName, productCode, info.productName, oldQty, quantity,
      { type: STOCK_CARD_TYPE.MANUAL, user: user.fullName, notes: "Set stok manual (modul Warehouse)" }, "")]);
    logActivity_(user.username, "SET_WAREHOUSE_STOCK", "Warehouse", "Set stok manual " + productCode + " di " + warehouseId + " = " + quantity, "");
    return { success: true, message: "Stok gudang berhasil disimpan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Hapus pencatatan stok produk di gudang tertentu (produk tetap ada di
// Master Produk, hanya baris stok di gudang tsb yang dihapus).
function deleteWarehouseStockRecord(warehouseId, productCode, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK);
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === warehouseId && data[i][3] === productCode) {
        let oldQty = Number(data[i][5]) || 0;
        let info = warehouseProductInfo_(ss, warehouseId, productCode);
        sheet.deleteRow(i + 1);
        appendStockCardRows_(ss, [buildStockCardRow_(warehouseId, info.warehouseName, productCode, info.productName, oldQty, 0,
          { type: STOCK_CARD_TYPE.RECORD_DELETE, user: user.fullName, notes: "Pencatatan stok dihapus (saldo " + oldQty + " menjadi 0)" }, "")]);
        logActivity_(user.username, "DELETE_WAREHOUSE_STOCK", "Warehouse", "Hapus pencatatan stok " + productCode + " di " + warehouseId, "");
        return { success: true, message: "Pencatatan stok berhasil dihapus." };
      }
    }
    return { success: false, message: "Data stok tidak ditemukan (produk ini belum punya catatan stok di gudang ini)." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Transfer stok antar 2 gudang MANAPUN (superset dari addStockFromGudang
// yang lama, dulu cuma bisa Gudang Pusat -> outlet). Contoh sesuai
// kebutuhan: Gudang Utama (48) -> Gudang Kasir Resto (24) transfer 24 pcs,
// hasil akhir Gudang Utama 24 & Gudang Kasir Resto 48 - kalau di hari yang
// sama outlet itu menjual 8 pcs (lewat Input Penjualan seperti biasa),
// stok Gudang Kasir Resto otomatis jadi 40 sementara Gudang Utama tetap 24.
// CATATAN: Kasir mengambil stok dari Gudang Utama ke outletnya sendiri
// lewat addStockFromGudang() (panel "Tambah Stok dari Gudang Pusat" di
// Manajemen Stok) - itu sudah cukup & lebih aman (dibatasi ketat ke
// outletnya sendiri). Fungsi transferStock ini TETAP khusus Owner/Admin/
// Staff karena dipakai modul Warehouse yang bisa transfer ANTAR gudang
// manapun (termasuk antar gudang internal), jadi tidak dibuka ke Kasir.
function transferStock(fromWarehouseId, toWarehouseId, productCode, qty, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    if (fromWarehouseId === toWarehouseId) return { success: false, message: "Gudang asal dan tujuan tidak boleh sama." };
    const ss = getDatabase();
    qty = Number(qty);
    if (!qty || qty <= 0) return { success: false, message: "Jumlah harus lebih dari 0." };

    let fromStock = getCurrentStock_(ss, fromWarehouseId, productCode);
    if (fromStock < qty) return { success: false, message: "Stok gudang asal tidak mencukupi (sisa: " + fromStock + ")." };

    const trfRef = genId_("TRF");
    updateStock_(ss, fromWarehouseId, productCode, -qty, undefined, { type: STOCK_CARD_TYPE.TRANSFER_OUT, counterpart: toWarehouseId, refId: trfRef, user: user.fullName });
    updateStock_(ss, toWarehouseId, productCode, qty, undefined, { type: STOCK_CARD_TYPE.TRANSFER_IN, counterpart: fromWarehouseId, refId: trfRef, user: user.fullName });

    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let product = products.find(p => p.ProductCode === productCode);
    let now = new Date();

    // Kalau salah satu ujung transfer adalah gudang outlet kasir, catat
    // juga di riwayat "Manajemen Stok" outlet tsb, supaya tetap konsisten
    // dengan apa yang tampil di menu Operasional Kasir > outlet ybs.
    let toUnit = unitByWarehouseId_(toWarehouseId);
    let fromUnit = unitByWarehouseId_(fromWarehouseId);
    if (toUnit) {
      ss.getSheetByName(unitCfg_(toUnit).stockLogSheet).appendRow([genId_("REQ"), now, productCode,
        product ? product.ProductName : productCode, qty, fromWarehouseId, toWarehouseId, "Selesai", user.fullName, now]);
    } else if (fromUnit) {
      ss.getSheetByName(unitCfg_(fromUnit).stockLogSheet).appendRow([genId_("REQ"), now, productCode,
        product ? product.ProductName : productCode, qty, fromWarehouseId, toWarehouseId, "Selesai", user.fullName, now]);
    }

    // Sama halnya kalau salah satu (atau kedua) ujung transfer adalah gudang
    // internal (Kolam/Teknisi/Cleaning Service/Dapur) - catat di sheet
    // riwayat KHUSUS gudang tsb (lihat INTERNAL_WAREHOUSE_LOG_SHEET_).
    // Pakai 2 "if" terpisah (bukan else-if) supaya transfer ANTAR 2 gudang
    // internal sekaligus tetap tercatat di riwayat kedua-duanya.
    let toInternalLog = INTERNAL_WAREHOUSE_LOG_SHEET_[toWarehouseId];
    let fromInternalLog = INTERNAL_WAREHOUSE_LOG_SHEET_[fromWarehouseId];
    if (toInternalLog) {
      let sheet = ss.getSheetByName(toInternalLog);
      if (sheet) sheet.appendRow([genId_("LOG"), now, productCode, product ? product.ProductName : productCode,
        qty, "Transfer Masuk", fromWarehouseId, toWarehouseId, "", user.fullName]);
    }
    if (fromInternalLog) {
      let sheet = ss.getSheetByName(fromInternalLog);
      if (sheet) sheet.appendRow([genId_("LOG"), now, productCode, product ? product.ProductName : productCode,
        qty, "Transfer Keluar", fromWarehouseId, toWarehouseId, "", user.fullName]);
    }

    ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS).appendRow([genId_("MUT"), now, productCode,
      fromWarehouseId, toWarehouseId, qty, "Stock Transfer", "", user.fullName]);

    logActivity_(user.username, "WAREHOUSE_TRANSFER", "Warehouse", "Transfer " + productCode + " qty " + qty + " dari " + fromWarehouseId + " ke " + toWarehouseId, "");
    return { success: true, message: "Stok berhasil ditransfer." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Riwayat SEMUA transfer stok antar gudang (ReferenceType = "Stock Transfer"
// di sheet Stock_Mutations), dipakai panel "Riwayat Transfer" di tab
// "Transfer Stok Antar Gudang" supaya Owner/Admin/Staff bisa memantau
// langsung gudang mana mengirim/menerima apa saja. startDate/endDate
// (format "yyyy-MM-dd", INKLUSIF, opsional) dipakai untuk filter cepat
// (Hari Ini/Bulan Ini/Tahun Ini) maupun tanggal tunggal/rentang custom di
// frontend - kalau keduanya kosong, kembalikan SEMUA riwayat transfer.
function getWarehouseTransferHistory(startDate, endDate, token) {
  try {
    checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    const ss = getDatabase();
    let warehouses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES));
    let whNameMap = {};
    warehouses.forEach(w => whNameMap[w.WarehouseID] = w.WarehouseName);
    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
    let productNameMap = {};
    products.forEach(p => productNameMap[p.ProductCode] = p.ProductName);

    let mutations = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS))
      .filter(m => m.ReferenceType === "Stock Transfer");

    if (startDate) mutations = mutations.filter(m => normalizeDateStr_(m.Timestamp) >= startDate);
    if (endDate) mutations = mutations.filter(m => normalizeDateStr_(m.Timestamp) <= endDate);

    mutations.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
    // Batasi maksimal 500 baris TERBARU supaya tetap ringan & cepat dimuat -
    // persempit rentang tanggal kalau butuh riwayat spesifik yang lebih tua.
    let totalMatching = mutations.length;
    mutations = mutations.slice(0, 500);

    let result = mutations.map(m => ({
      Timestamp: m.Timestamp, ProductCode: m.ProductCode,
      ProductName: productNameMap[m.ProductCode] || m.ProductCode,
      FromWarehouseID: m.FromLocation, FromWarehouseName: whNameMap[m.FromLocation] || m.FromLocation,
      ToWarehouseID: m.ToLocation, ToWarehouseName: whNameMap[m.ToLocation] || m.ToLocation,
      Qty: m.Qty, User: m.User
    }));

    return { success: true, history: result, totalMatching: totalMatching };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Catat stok keluar/terpakai secara manual - dipakai gudang INTERNAL yang
// tidak punya kasir/POS (Kolam Renang/Teknisi/Cleaning Service), karena
// stok di gudang-gudang ini tidak pernah otomatis berkurang lewat
// transaksi penjualan seperti outlet kasir.
function recordWarehouseUsage(warehouseId, productCode, qty, notes, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    const ss = getDatabase();
    qty = Number(qty);
    if (!qty || qty <= 0) return { success: false, message: "Jumlah harus lebih dari 0." };
    let current = getCurrentStock_(ss, warehouseId, productCode);
    if (current < qty) return { success: false, message: "Stok tidak mencukupi (sisa: " + current + ")." };

    updateStock_(ss, warehouseId, productCode, -qty, undefined, { type: STOCK_CARD_TYPE.USAGE, notes: notes || "", user: user.fullName });
    let now = new Date();
    ss.getSheetByName(CONFIG.SHEETS.STOCK_MUTATIONS).appendRow([genId_("MUT"), now, productCode,
      warehouseId, "Terpakai/Keluar", qty, "Usage", notes || "", user.fullName]);

    // Catat juga di sheet riwayat KHUSUS gudang ini (Kolam/Teknisi/Cleaning
    // Service/Dapur) - lihat INTERNAL_WAREHOUSE_LOG_SHEET_.
    let logSheetName = INTERNAL_WAREHOUSE_LOG_SHEET_[warehouseId];
    if (logSheetName) {
      let logSheet = ss.getSheetByName(logSheetName);
      if (logSheet) {
        let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS));
        let product = products.find(p => p.ProductCode === productCode);
        logSheet.appendRow([genId_("LOG"), now, productCode, product ? product.ProductName : productCode,
          qty, "Terpakai/Keluar", warehouseId, "-", notes || "", user.fullName]);
      }
    }

    logActivity_(user.username, "WAREHOUSE_USAGE", "Warehouse", "Stok keluar/terpakai " + productCode + " qty " + qty + " dari " + warehouseId, "");
    return { success: true, message: "Stok keluar berhasil dicatat." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Matrix monitoring real-time: untuk setiap produk, tampilkan qty di
// SEMUA gudang sekaligus (termasuk gudang outlet kasir) dalam satu
// tabel - inilah yang memenuhi kebutuhan "memantau stok kasir secara
// langsung" dari modul Warehouse.
function getWarehouseMonitor(token) {
  try {
    checkPermission_(token, [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF]);
    const ss = getDatabase();
    let warehouses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSES));
    let stockRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE_STOCK));
    // Sama seperti getWarehouseStockDetail(): hanya produk Active & yang
    // memakai pelacakan stok (UsesStock != "Tidak") yang relevan dipantau
    // di modul Warehouse.
    let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS))
      .filter(p => p.Status === "Active" && String(p.UsesStock || "Ya").trim() !== "Tidak" && !String(p.StockRefCode || "").trim());
    let productTypeNames = getProductTypeNamesAll_(ss);

    let stockLookup = {}; // ProductCode -> { WarehouseID: qty }
    stockRows.forEach(s => {
      if (!stockLookup[s.ProductCode]) stockLookup[s.ProductCode] = {};
      stockLookup[s.ProductCode][s.WarehouseID] = Number(s.Quantity) || 0;
    });

    let result = products.map(p => {
      let stocks = {};
      warehouses.forEach(w => { stocks[w.WarehouseID] = (stockLookup[p.ProductCode] && stockLookup[p.ProductCode][w.WarehouseID]) || 0; });
      let total = Object.keys(stocks).reduce((sum, k) => sum + stocks[k], 0);
      return {
        ProductCode: p.ProductCode, ProductName: p.ProductName,
        ItemType: productTypeNames.indexOf(p.ItemType) !== -1 ? p.ItemType : "Lainnya",
        stocks: stocks, total: total
      };
    });
    return { success: true, warehouses: warehouses, products: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 11. REPORT & EXPORT
// -----------------------------------------------------------------------
// PENTING: pembuatan file PDF/XLSX dilakukan di sisi CLIENT (browser)
// pakai SheetJS + jsPDF, bukan lewat UrlFetchApp/Drive export seperti
// versi sebelumnya. Server di sini hanya menyediakan DATA mentah yang
// sudah difilter/dirapikan. Ini menghindari masalah export gagal akibat
// izin OAuth/Drive pada web app (persis pola yang sudah terbukti jalan
// di aplikasi Manajemen Stok Gudang sebelumnya).
// =====================================================================
function getDailyReportData(unit, dateStr, token) {
  try {
    let user = checkPermission_(token, null, unit);
    const ss = getDatabase();
    const cfg = unitCfg_(unit);
    const tz = Session.getScriptTimeZone();

    let sales = getSheetDataAsObjects_(ss.getSheetByName(cfg.salesSheet))
      .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
    let payments = getSheetDataAsObjects_(ss.getSheetByName(cfg.paySheet))
      .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
    let settings = getSystemSettings();

    let salesTotal = sales.reduce((t, r) => t + (Number(r.TotalPrice) || 0), 0);
    let paymentTotal = payments.reduce((t, p) => t + (Number(p.Amount) || 0), 0);

    sales.forEach(r => { r.Date = fmtD_(r.Date, tz); });
    payments.forEach(p => { p.Date = fmtD_(p.Date, tz); });

    return {
      success: true, unitLabel: cfg.label, hasSerial: cfg.hasSerial, dateStr: dateStr,
      companyName: settings.COMPANY_NAME || CONFIG.COMPANY_NAME, address: settings.ADDRESS || "-",
      logoUrl: settings.LOGO_URL || "",
      cashierName: user.fullName, cashierPosition: cfg.label,
      printedAt: new Date().toISOString(),
      sales: sales, salesTotal: salesTotal, payments: payments, paymentTotal: paymentTotal
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Khusus Owner: rekap gabungan penjualan & penerimaan uang SEMUA outlet
// (Loket Masuk + Loket Wahana + Resto) untuk satu tanggal, dalam satu report.
function getCombinedDailyReportData(dateStr, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    const tz = Session.getScriptTimeZone();
    const units = ["Entrance", "Ride", "Resto"];
    let settings = getSystemSettings();

    let result = {
      success: true, dateStr: dateStr,
      companyName: settings.COMPANY_NAME || CONFIG.COMPANY_NAME, address: settings.ADDRESS || "-",
      logoUrl: settings.LOGO_URL || "",
      cashierName: user.fullName, cashierPosition: "Owner - Laporan Rekapan Gabungan",
      printedAt: new Date().toISOString(),
      units: {}, grandSalesTotal: 0, grandPaymentTotal: 0
    };

    units.forEach(u => {
      const cfg = unitCfg_(u);
      let sales = getSheetDataAsObjects_(ss.getSheetByName(cfg.salesSheet))
        .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
      let payments = getSheetDataAsObjects_(ss.getSheetByName(cfg.paySheet))
        .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);

      let salesTotal = sales.reduce((t, r) => t + (Number(r.TotalPrice) || 0), 0);
      let paymentTotal = payments.reduce((t, p) => t + (Number(p.Amount) || 0), 0);

      sales.forEach(r => { r.Date = fmtD_(r.Date, tz); });
      payments.forEach(p => { p.Date = fmtD_(p.Date, tz); });

      result.units[u] = {
        unitLabel: cfg.label, hasSerial: cfg.hasSerial,
        sales: sales, salesTotal: salesTotal, payments: payments, paymentTotal: paymentTotal
      };
      result.grandSalesTotal += salesTotal;
      result.grandPaymentTotal += paymentTotal;
    });

    return result;
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 11B. ADMIN KEUANGAN - SALDO, PENGELUARAN & REPORT HARIAN OTOMATIS
// -----------------------------------------------------------------------
// Modul ini TIDAK mencatat penjualan sendiri - semua angka pendapatan
// (Tiket Masuk, Tiket Wahana, Resto & F&B, rincian Cash/QRIS/Transfer/
// Debit/Kredit) diambil otomatis dari data Kasir (Entrance/Ride/Resto)
// yang sudah ada, supaya tidak ada input dobel. Admin Keuangan hanya
// perlu mengisi: Saldo Awal (Brangkas/Rekening/THR), Pengeluaran (3
// kategori: dari Omset Harian, dari Brangkas, dari Rekening PT), Saldo
// Akhir Aktual (hasil hitung fisik, untuk rekonsiliasi), dan Hutang/
// Piutang + catatan tambahan. Dari situ, Report harian (mirip form PDF
// & format teks WA) dirakit otomatis lewat getFinanceDailyReport().
// =====================================================================
const FINANCE_ROLES_WRITE = [ROLE.OWNER, ROLE.ADMIN];
const FINANCE_ROLES_READ = [ROLE.OWNER, ROLE.ADMIN, ROLE.STAFF];

// -----------------------------------------------------------------------
// 11A. MASTER KEUANGAN (khusus Owner/Admin) - Kode Akun (COA), Kategori
// Status Pembayaran, dan Pos Penggunaan/Kegiatan. Ketiganya bisa diatur
// & ditambah sendiri oleh Admin Keuangan lewat submenu "Master Keuangan",
// lalu dipakai sebagai dropdown pilihan saat mencatat Pengeluaran Harian.
// -----------------------------------------------------------------------
function getFinanceCOA(token) {
  try {
    if (token) checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    return { success: true, data: getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_COA)) };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveFinanceCOA(data, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    if (!data.COACode || !data.COAName) return { success: false, message: "Kode Akun dan Nama Akun wajib diisi." };
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_COA);
    if (data.isEdit && data.COAID) {
      let rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.COAID) {
          sheet.getRange(i + 1, 2, 1, 4).setValues([[data.COACode, data.COAName, data.Type || "Pengeluaran", data.Status || "Active"]]);
          logActivity_(user.username, "EDIT_COA", "Master Keuangan", "Edit Kode Akun " + data.COACode, "");
          return { success: true, message: "Kode Akun berhasil diperbarui." };
        }
      }
      return { success: false, message: "Kode Akun tidak ditemukan." };
    }
    sheet.appendRow([genId_("COA"), data.COACode, data.COAName, data.Type || "Pengeluaran", data.Status || "Active"]);
    logActivity_(user.username, "CREATE_COA", "Master Keuangan", "Tambah Kode Akun " + data.COACode, "");
    return { success: true, message: "Kode Akun berhasil ditambahkan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteFinanceCOA(coaId, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_COA);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === coaId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_COA", "Master Keuangan", "Hapus Kode Akun " + coaId, "");
        return { success: true, message: "Kode Akun berhasil dihapus." };
      }
    }
    return { success: false, message: "Kode Akun tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Import massal Kode Akun (COA) dari file Excel.
function bulkImportFinanceCOA(rows, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_COA);
    let existing = getSheetDataAsObjects_(sheet);
    let inserted = 0, updated = 0, errors = [];
    rows.forEach((r, idx) => {
      try {
        if (!r.coaCode || !r.coaName) { errors.push("Baris " + (idx + 2) + ": Kode/Nama Akun kosong, dilewati."); return; }
        let found = existing.find(c => String(c.COACode) === String(r.coaCode));
        if (found) {
          sheet.getRange(found._row, 2, 1, 4).setValues([[r.coaCode, r.coaName, r.type || "Pengeluaran", r.status || "Active"]]);
          updated++;
        } else {
          sheet.appendRow([genId_("COA"), r.coaCode, r.coaName, r.type || "Pengeluaran", r.status || "Active"]);
          inserted++;
        }
      } catch (rowErr) { errors.push("Baris " + (idx + 2) + ": " + rowErr.message); }
    });
    logActivity_(user.username, "IMPORT_COA", "Master Keuangan", "Import " + inserted + " baru, " + updated + " update Kode Akun", "");
    return { success: true, inserted: inserted, updated: updated, errors: errors };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getFinancePaymentStatusList(token) {
  try {
    if (token) checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    return { success: true, data: getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_PAYMENT_STATUS)) };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveFinancePaymentStatus(data, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    if (!data.StatusName || !data.Bucket) return { success: false, message: "Nama Status dan Bucket wajib diisi." };
    let allowedBuckets = [FINANCE_EXPENSE_CATEGORY.OMSET, FINANCE_EXPENSE_CATEGORY.BRANGKAS, FINANCE_EXPENSE_CATEGORY.REKENING, FINANCE_EXPENSE_CATEGORY.BELUM_LUNAS];
    if (allowedBuckets.indexOf(data.Bucket) === -1) return { success: false, message: "Bucket tidak dikenali." };
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_PAYMENT_STATUS);
    if (data.isEdit && data.StatusID) {
      let rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.StatusID) {
          sheet.getRange(i + 1, 2, 1, 3).setValues([[data.StatusName, data.Bucket, data.Status || "Active"]]);
          logActivity_(user.username, "EDIT_PAYMENT_STATUS", "Master Keuangan", "Edit Status Pembayaran " + data.StatusName, "");
          return { success: true, message: "Status Pembayaran berhasil diperbarui." };
        }
      }
      return { success: false, message: "Status Pembayaran tidak ditemukan." };
    }
    sheet.appendRow([genId_("PS"), data.StatusName, data.Bucket, data.Status || "Active"]);
    logActivity_(user.username, "CREATE_PAYMENT_STATUS", "Master Keuangan", "Tambah Status Pembayaran " + data.StatusName, "");
    return { success: true, message: "Status Pembayaran berhasil ditambahkan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteFinancePaymentStatus(statusId, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_PAYMENT_STATUS);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === statusId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_PAYMENT_STATUS", "Master Keuangan", "Hapus Status Pembayaran " + statusId, "");
        return { success: true, message: "Status Pembayaran berhasil dihapus." };
      }
    }
    return { success: false, message: "Status Pembayaran tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Import massal Status Pembayaran dari file Excel.
function bulkImportFinancePaymentStatus(rows, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    let allowedBuckets = [FINANCE_EXPENSE_CATEGORY.OMSET, FINANCE_EXPENSE_CATEGORY.BRANGKAS, FINANCE_EXPENSE_CATEGORY.REKENING, FINANCE_EXPENSE_CATEGORY.BELUM_LUNAS];
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_PAYMENT_STATUS);
    let existing = getSheetDataAsObjects_(sheet);
    let inserted = 0, updated = 0, errors = [];
    rows.forEach((r, idx) => {
      try {
        if (!r.statusName || !r.bucket) { errors.push("Baris " + (idx + 2) + ": Nama Status/Bucket kosong, dilewati."); return; }
        if (allowedBuckets.indexOf(r.bucket) === -1) { errors.push("Baris " + (idx + 2) + ": Bucket '" + r.bucket + "' tidak dikenali (harus Omset/Brangkas/RekeningPT/BelumLunas), dilewati."); return; }
        let found = existing.find(s => s.StatusName === r.statusName);
        if (found) {
          sheet.getRange(found._row, 2, 1, 3).setValues([[r.statusName, r.bucket, r.status || "Active"]]);
          updated++;
        } else {
          sheet.appendRow([genId_("PS"), r.statusName, r.bucket, r.status || "Active"]);
          inserted++;
        }
      } catch (rowErr) { errors.push("Baris " + (idx + 2) + ": " + rowErr.message); }
    });
    logActivity_(user.username, "IMPORT_PAYMENT_STATUS", "Master Keuangan", "Import " + inserted + " baru, " + updated + " update Status Pembayaran", "");
    return { success: true, inserted: inserted, updated: updated, errors: errors };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getFinancePosPenggunaanList(token) {
  try {
    if (token) checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    return { success: true, data: getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN)) };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveFinancePosPenggunaan(data, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    if (!data.PosName) return { success: false, message: "Nama Pos Penggunaan wajib diisi." };
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN);
    if (data.isEdit && data.PosID) {
      let rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.PosID) {
          sheet.getRange(i + 1, 2, 1, 2).setValues([[data.PosName, data.Status || "Active"]]);
          logActivity_(user.username, "EDIT_POS_PENGGUNAAN", "Master Keuangan", "Edit Pos Penggunaan " + data.PosName, "");
          return { success: true, message: "Pos Penggunaan berhasil diperbarui." };
        }
      }
      return { success: false, message: "Pos Penggunaan tidak ditemukan." };
    }
    sheet.appendRow([genId_("POS"), data.PosName, data.Status || "Active"]);
    logActivity_(user.username, "CREATE_POS_PENGGUNAAN", "Master Keuangan", "Tambah Pos Penggunaan " + data.PosName, "");
    return { success: true, message: "Pos Penggunaan berhasil ditambahkan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteFinancePosPenggunaan(posId, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === posId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_POS_PENGGUNAAN", "Master Keuangan", "Hapus Pos Penggunaan " + posId, "");
        return { success: true, message: "Pos Penggunaan berhasil dihapus." };
      }
    }
    return { success: false, message: "Pos Penggunaan tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Import massal Pos Penggunaan dari file Excel.
function bulkImportFinancePosPenggunaan(rows, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN);
    let existing = getSheetDataAsObjects_(sheet);
    let inserted = 0, updated = 0, errors = [];
    rows.forEach((r, idx) => {
      try {
        if (!r.posName) { errors.push("Baris " + (idx + 2) + ": Nama Pos Penggunaan kosong, dilewati."); return; }
        let found = existing.find(p => p.PosName === r.posName);
        if (found) {
          sheet.getRange(found._row, 2, 1, 2).setValues([[r.posName, r.status || "Active"]]);
          updated++;
        } else {
          sheet.appendRow([genId_("POS"), r.posName, r.status || "Active"]);
          inserted++;
        }
      } catch (rowErr) { errors.push("Baris " + (idx + 2) + ": " + rowErr.message); }
    });
    logActivity_(user.username, "IMPORT_POS_PENGGUNAAN", "Master Keuangan", "Import " + inserted + " baru, " + updated + " update Pos Penggunaan", "");
    return { success: true, inserted: inserted, updated: updated, errors: errors };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// DAFTAR TUJUAN EMAIL - Report Admin Keuangan (Kirim PDF/XLSX ke Email)
// =====================================================================
function getFinanceReportEmailRecipients(token) {
  try {
    if (token) checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    return { success: true, data: getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_REPORT_EMAIL)) };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveFinanceReportEmailRecipient(data, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    let email = (data.Email || "").toString().trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success: false, message: "Alamat email tidak valid." };
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_REPORT_EMAIL);
    if (data.isEdit && data.RecipientID) {
      let rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.RecipientID) {
          sheet.getRange(i + 1, 2, 1, 3).setValues([[email, data.Name || "", data.Status || "Active"]]);
          logActivity_(user.username, "EDIT_FINANCE_EMAIL_RECIPIENT", "Admin Keuangan", "Edit tujuan email " + email, "");
          return { success: true, message: "Tujuan email berhasil diperbarui." };
        }
      }
      return { success: false, message: "Data tujuan email tidak ditemukan." };
    }
    sheet.appendRow([genId_("MAIL"), email, data.Name || "", data.Status || "Active"]);
    logActivity_(user.username, "ADD_FINANCE_EMAIL_RECIPIENT", "Admin Keuangan", "Tambah tujuan email " + email, "");
    return { success: true, message: "Tujuan email berhasil ditambahkan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteFinanceReportEmailRecipient(id, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_REPORT_EMAIL);
    let rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_FINANCE_EMAIL_RECIPIENT", "Admin Keuangan", "Hapus tujuan email " + id, "");
        return { success: true, message: "Tujuan email berhasil dihapus." };
      }
    }
    return { success: false, message: "Data tujuan email tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Kirim Report Admin Keuangan (PDF & XLSX, dibuat di browser lalu dikirim
// sebagai base64) ke daftar tujuan email yang dipilih. payload:
// { dateStr, recipients: [email,...], pdfBase64, pdfFileName, xlsxBase64, xlsxFileName }
function sendFinanceReportEmail(payload, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_READ);
    if (!payload) return { success: false, message: "Data laporan tidak ditemukan." };
    let recipients = (payload.recipients || []).map(e => (e || "").toString().trim()).filter(Boolean);
    if (!recipients.length) return { success: false, message: "Pilih minimal 1 tujuan email." };
    if (!payload.pdfBase64 && !payload.xlsxBase64) return { success: false, message: "File laporan (PDF/XLSX) tidak ditemukan." };

    let attachments = [];
    if (payload.pdfBase64) {
      attachments.push(Utilities.newBlob(Utilities.base64Decode(payload.pdfBase64), "application/pdf", payload.pdfFileName || "Laporan_Keuangan.pdf"));
    }
    if (payload.xlsxBase64) {
      attachments.push(Utilities.newBlob(Utilities.base64Decode(payload.xlsxBase64),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", payload.xlsxFileName || "Laporan_Keuangan.xlsx"));
    }

    let settings = getSystemSettings();
    let companyName = settings.COMPANY_NAME || CONFIG.COMPANY_NAME;
    let subject = "Laporan Harian Admin Keuangan - " + companyName + " (" + (payload.dateStr || "") + ")";
    let body = "Yth. Bapak/Ibu,\n\n" +
      "Berikut kami lampirkan Laporan Harian Admin Keuangan " + companyName + " untuk tanggal " + (payload.dateStr || "") + ".\n\n" +
      "Terlampir file dalam format " + (payload.pdfBase64 && payload.xlsxBase64 ? "PDF dan Excel (XLSX)" : (payload.pdfBase64 ? "PDF" : "Excel (XLSX)")) + ".\n\n" +
      "Email ini dikirim otomatis oleh sistem, mohon tidak membalas ke alamat ini.\n\n" +
      "Terima kasih.";

    MailApp.sendEmail({ to: recipients.join(","), subject: subject, body: body, attachments: attachments, name: companyName });
    logActivity_(user.username, "SEND_FINANCE_REPORT_EMAIL", "Admin Keuangan",
      "Kirim laporan tanggal " + payload.dateStr + " ke " + recipients.join(", "), "");
    return { success: true, message: "Laporan berhasil dikirim ke " + recipients.length + " tujuan email." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function findFinanceBalanceRow_(sheet, dateStr) {
  let data = sheet.getDataRange().getValues();
  let headers = data[0];
  let dateCol = headers.indexOf("Date");
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizeDateStr_(data[i][dateCol]) === dateStr) return { rowIndex: i + 1, headers: headers, row: data[i] };
  }
  return null;
}

// Cari baris tanggal TERAKHIR sebelum dateStr yang sudah pernah diisi -
// dipakai untuk menyarankan Saldo Awal (= Saldo Akhir Aktual hari
// sebelumnya) supaya Admin Keuangan tidak perlu mengetik ulang tiap hari.
function findLatestFinanceBalanceBefore_(sheet, dateStr) {
  let data = sheet.getDataRange().getValues();
  let headers = data[0];
  let dateCol = headers.indexOf("Date");
  let best = null, bestDateStr = "";
  for (let i = 1; i < data.length; i++) {
    let d = normalizeDateStr_(data[i][dateCol]);
    if (d && d < dateStr && d > bestDateStr) { best = { rowIndex: i + 1, headers: headers, row: data[i] }; bestDateStr = d; }
  }
  return best;
}

function rowToObj_(headers, row) {
  let obj = {};
  headers.forEach((h, idx) => { obj[h] = row[idx]; });
  return obj;
}

// Ambil Saldo Akhir Brangkas/Rekening yang SUDAH DIRESOLVE (final) dari satu
// baris Finance_Balance: pakai snapshot SaldoAkhirBrangkas/SaldoAkhirRekening
// kalau sudah ada (baris yang disimpan dengan skema baru), atau fallback ke
// SaldoAkhirBrangkasAktual/SaldoAkhirRekeningAktual untuk baris lama (skema
// sebelum kolom snapshot ini ditambahkan, di mana Aktual adalah satu-satunya
// sumber - kalau Aktual juga kosong berarti data lama itu memang belum diisi).
function resolvedSaldoAkhir_(p) {
  let hasSnapshotBrangkas = p.SaldoAkhirBrangkas !== "" && p.SaldoAkhirBrangkas !== undefined && p.SaldoAkhirBrangkas !== null;
  let hasSnapshotRekening = p.SaldoAkhirRekening !== "" && p.SaldoAkhirRekening !== undefined && p.SaldoAkhirRekening !== null;
  return {
    brangkas: hasSnapshotBrangkas ? Number(p.SaldoAkhirBrangkas) : (Number(p.SaldoAkhirBrangkasAktual) || 0),
    rekening: hasSnapshotRekening ? Number(p.SaldoAkhirRekening) : (Number(p.SaldoAkhirRekeningAktual) || 0)
  };
}

// Hitung total & net Mutasi Saldo (pindah dana internal Brangkas <->
// Rekening) untuk SATU tanggal - dipakai computeFinanceCashFlow_() supaya
// Mutasi Saldo otomatis mempengaruhi Saldo Akhir Brangkas/Rekening hari itu
// (net positif = bucket itu bertambah bersih hari ini, net negatif =
// berkurang). netBrangkas + netRekening akan selalu 0 karena mutasi cuma
// memindahkan dana antara 2 bucket ini, bukan menciptakan/menghilangkan uang.
function computeFinanceMutationNet_(ss, dateStr) {
  let rows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE_MUTATIONS))
    .filter(m => normalizeDateStr_(m.Date) === dateStr);
  let masukBrangkas = 0, keluarBrangkas = 0, masukRekening = 0, keluarRekening = 0;
  rows.forEach(m => {
    let amt = Number(m.Amount) || 0;
    if (m.ToBucket === "Brangkas") masukBrangkas += amt;
    if (m.FromBucket === "Brangkas") keluarBrangkas += amt;
    if (m.ToBucket === "Rekening") masukRekening += amt;
    if (m.FromBucket === "Rekening") keluarRekening += amt;
  });
  return {
    masukBrangkas: masukBrangkas, keluarBrangkas: keluarBrangkas,
    masukRekening: masukRekening, keluarRekening: keluarRekening,
    netBrangkas: masukBrangkas - keluarBrangkas, netRekening: masukRekening - keluarRekening
  };
}

// Hitung arus kas masuk (tunai/non-tunai dari 3 outlet kasir) & total
// pengeluaran per Bucket (Omset/Brangkas/RekeningPT) untuk SATU tanggal.
// Dipakai bersama oleh getFinanceBalance() - supaya form Saldo & Kas bisa
// menampilkan Saldo Akhir Brangkas/Rekening (ESTIMASI) secara real-time
// begitu Saldo Awal diketik/diubah - dan getFinanceDailyReport().
function computeFinanceCashFlow_(ss, dateStr, tz) {
  function paySum_(unit) {
    let cfg = unitCfg_(unit);
    let rows = getSheetDataAsObjects_(ss.getSheetByName(cfg.paySheet))
      .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
    let b = { cash: 0, nonCash: 0 };
    rows.forEach(p => {
      let amt = Number(p.Amount) || 0;
      let method = (p.PaymentMethod || "").toLowerCase();
      if (method.includes("cash") || method.includes("tunai")) b.cash += amt; else b.nonCash += amt;
    });
    return b;
  }
  let payE = paySum_("Entrance"), payR = paySum_("Ride"), payS = paySum_("Resto");
  let totalCash = payE.cash + payR.cash + payS.cash;
  let totalNonCash = payE.nonCash + payR.nonCash + payS.nonCash;

  let expenses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES)).filter(r => normalizeDateStr_(r.Date) === dateStr);
  function expTotal_(cat) { return expenses.filter(e => e.Category === cat).reduce((t, e) => t + (Number(e.Amount) || 0), 0); }
  let totalExpOmset = expTotal_(FINANCE_EXPENSE_CATEGORY.OMSET);
  let totalExpBrangkas = expTotal_(FINANCE_EXPENSE_CATEGORY.BRANGKAS);
  let totalExpRekening = expTotal_(FINANCE_EXPENSE_CATEGORY.REKENING);
  // Asumsi: pengeluaran dari omset harian dibayar tunai dari laci kasir,
  // sehingga mengurangi porsi Cash (bukan porsi non-tunai).
  let cashBersih = totalCash - totalExpOmset;
  let mutasi = computeFinanceMutationNet_(ss, dateStr);
  return {
    totalCash: totalCash, totalNonCash: totalNonCash, totalExpOmset: totalExpOmset,
    totalExpBrangkas: totalExpBrangkas, totalExpRekening: totalExpRekening, cashBersih: cashBersih,
    mutasiNetBrangkas: mutasi.netBrangkas, mutasiNetRekening: mutasi.netRekening,
    mutasiMasukBrangkas: mutasi.masukBrangkas, mutasiKeluarBrangkas: mutasi.keluarBrangkas,
    mutasiMasukRekening: mutasi.masukRekening, mutasiKeluarRekening: mutasi.keluarRekening
  };
}

// Ambil (atau sarankan) data Saldo untuk satu tanggal. Kalau tanggal itu
// belum pernah diisi, Saldo Awal disarankan otomatis dari Saldo Akhir
// Aktual tanggal terakhir yang sudah diisi sebelumnya (0 kalau belum
// pernah ada data sama sekali). Juga menyertakan komponen arus kas
// (FlowCashBersih/FlowNonCash/FlowExpBrangkas/FlowExpRekening) hari itu
// supaya form web bisa menghitung & menampilkan Saldo Akhir Brangkas/
// Rekening (ESTIMASI) secara langsung begitu Saldo Awal diubah, tanpa
// perlu menyimpan dulu.
// Kunci Saldo Awal (default TERKUNCI) - kalau Yes/kosong, Admin Keuangan
// tidak boleh mengetik manual Saldo Awal Brangkas/Rekening; nilainya HARUS
// otomatis ikut Saldo Akhir (snapshot) tanggal terakhir yang sudah diisi.
// Hanya Owner yang bisa membuka kuncinya lewat submenu Profil Aplikasi &
// Fitur (lihat saveSystemSettings, key LOCK_MANUAL_SALDO_AWAL).
function isSaldoAwalLocked_(settings) {
  return (settings.LOCK_MANUAL_SALDO_AWAL || "Yes") !== "No";
}

function getFinanceBalance(dateStr, token) {
  try {
    checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    const tz = Session.getScriptTimeZone();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE);
    let settings = getSystemSettings();
    let locked = isSaldoAwalLocked_(settings);
    let flow = computeFinanceCashFlow_(ss, dateStr, tz);
    let found = findFinanceBalanceRow_(sheet, dateStr);
    if (found) {
      let obj = rowToObj_(found.headers, found.row);
      obj.isSaved = true;
      obj.HutangJSON = obj.HutangJSON || "[]";
      obj.PiutangJSON = obj.PiutangJSON || "[]";
      obj.FlowCashBersih = flow.cashBersih; obj.FlowNonCash = flow.totalNonCash;
      obj.FlowExpBrangkas = flow.totalExpBrangkas; obj.FlowExpRekening = flow.totalExpRekening;
      obj.FlowMutasiNetBrangkas = flow.mutasiNetBrangkas; obj.FlowMutasiNetRekening = flow.mutasiNetRekening;
      obj.LockManualSaldoAwal = locked;
      return { success: true, data: obj };
    }
    let prev = findLatestFinanceBalanceBefore_(sheet, dateStr);
    let suggestion = {
      Date: dateStr, isSaved: false,
      SaldoAwalBrangkas: 0, SaldoAwalRekening: 0, SaldoTHR: 0,
      SaldoAkhirBrangkasAktual: "", SaldoAkhirRekeningAktual: "",
      CatatanTambahan: "", HutangJSON: "[]", PiutangJSON: "[]"
    };
    if (prev) {
      let p = rowToObj_(prev.headers, prev.row);
      let resolved = resolvedSaldoAkhir_(p);
      suggestion.SaldoAwalBrangkas = resolved.brangkas;
      suggestion.SaldoAwalRekening = resolved.rekening;
      suggestion.SaldoTHR = Number(p.SaldoTHR) || 0;
    }
    suggestion.FlowCashBersih = flow.cashBersih; suggestion.FlowNonCash = flow.totalNonCash;
    suggestion.FlowExpBrangkas = flow.totalExpBrangkas; suggestion.FlowExpRekening = flow.totalExpRekening;
    suggestion.FlowMutasiNetBrangkas = flow.mutasiNetBrangkas; suggestion.FlowMutasiNetRekening = flow.mutasiNetRekening;
    suggestion.LockManualSaldoAwal = locked;
    suggestion.HasPrevBalance = !!prev;
    return { success: true, data: suggestion };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Ditulis berdasarkan NAMA kolom (bukan posisi/index tetap), supaya tetap
// benar walau urutan kolom fisik di sheet Finance_Balance berbeda dari
// FINANCE_BALANCE_HEADERS (mis. sheet lama yang kolom barunya ditambahkan
// belakangan oleh ensureFinanceBalanceSchema_). Ini yang memperbaiki bug
// Saldo Akhir tidak masuk ke database walau Saldo Awal tersimpan.
function saveFinanceBalance(data, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const tz = Session.getScriptTimeZone();
    ensureFinanceBalanceSchema_(ss);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE);
    if (!data.Date) return { success: false, message: "Tanggal wajib diisi." };
    let now = new Date();

    let lastCol = sheet.getLastColumn();
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let found = findFinanceBalanceRow_(sheet, data.Date);
    let existingRow = found ? found.row.slice() : new Array(headers.length).fill("");

    // Otorisasi Saldo Awal: kalau TERKUNCI (default), abaikan angka Saldo
    // Awal manual yang dikirim dari form dan PAKSA pakai Saldo Akhir
    // (snapshot) hari terakhir yang sudah diisi sebelum tanggal ini - Admin
    // Keuangan tidak bisa mengubahnya sendiri tanpa Owner membuka kunci di
    // Profil Aplikasi & Fitur. Kalau belum ada data hari sebelumnya sama
    // sekali (baris pertama sejak awal pemakaian), pertahankan nilai yang
    // sudah tersimpan (kalau edit) atau 0 (kalau baris baru).
    let settings = getSystemSettings();
    let saldoAwalBrangkas, saldoAwalRekening;
    if (isSaldoAwalLocked_(settings)) {
      let prevRow = findLatestFinanceBalanceBefore_(sheet, data.Date);
      if (prevRow) {
        let resolvedPrev = resolvedSaldoAkhir_(rowToObj_(prevRow.headers, prevRow.row));
        saldoAwalBrangkas = resolvedPrev.brangkas;
        saldoAwalRekening = resolvedPrev.rekening;
      } else {
        let awalIdx = headers.indexOf("SaldoAwalBrangkas"), awalIdxR = headers.indexOf("SaldoAwalRekening");
        saldoAwalBrangkas = found ? (Number(existingRow[awalIdx]) || 0) : 0;
        saldoAwalRekening = found ? (Number(existingRow[awalIdxR]) || 0) : 0;
      }
    } else {
      saldoAwalBrangkas = Number(data.SaldoAwalBrangkas) || 0;
      saldoAwalRekening = Number(data.SaldoAwalRekening) || 0;
    }
    let aktualBrangkas = (data.SaldoAkhirBrangkasAktual === "" || data.SaldoAkhirBrangkasAktual === undefined || data.SaldoAkhirBrangkasAktual === null) ? "" : Number(data.SaldoAkhirBrangkasAktual);
    let aktualRekening = (data.SaldoAkhirRekeningAktual === "" || data.SaldoAkhirRekeningAktual === undefined || data.SaldoAkhirRekeningAktual === null) ? "" : Number(data.SaldoAkhirRekeningAktual);

    // Selalu hitung & SIMPAN Saldo Akhir yang SUDAH DIRESOLVE (Aktual kalau
    // diisi, kalau tidak pakai hasil Hitungan otomatis kas masuk-keluar hari
    // ini) ke kolom snapshot SaldoAkhirBrangkas/SaldoAkhirRekening. Snapshot
    // inilah yang jadi acuan Saldo Awal besok (lihat resolvedSaldoAkhir_()),
    // BUKAN kolom Aktual saja - supaya rantai saldo antar hari tidak putus
    // hanya karena Aktual dibiarkan kosong (pakai Estimasi otomatis).
    let flow = computeFinanceCashFlow_(ss, data.Date, tz);
    let saldoAkhirBrangkasHitung = saldoAwalBrangkas + flow.cashBersih - flow.totalExpBrangkas + flow.mutasiNetBrangkas;
    let saldoAkhirRekeningHitung = saldoAwalRekening + flow.totalNonCash - flow.totalExpRekening + flow.mutasiNetRekening;
    let saldoAkhirBrangkas = aktualBrangkas === "" ? saldoAkhirBrangkasHitung : aktualBrangkas;
    let saldoAkhirRekening = aktualRekening === "" ? saldoAkhirRekeningHitung : aktualRekening;

    let values = {
      BalanceID: found ? found.row[found.headers.indexOf("BalanceID")] : genId_("FIN"),
      Date: data.Date,
      SaldoAwalBrangkas: saldoAwalBrangkas,
      SaldoAwalRekening: saldoAwalRekening,
      SaldoTHR: Number(data.SaldoTHR) || 0,
      SaldoAkhirBrangkasAktual: aktualBrangkas,
      SaldoAkhirRekeningAktual: aktualRekening,
      CatatanTambahan: data.CatatanTambahan || "",
      HutangJSON: data.HutangJSON || "[]",
      PiutangJSON: data.PiutangJSON || "[]",
      CreatedBy: user.fullName,
      UpdatedAt: now,
      SaldoAkhirBrangkas: saldoAkhirBrangkas,
      SaldoAkhirRekening: saldoAkhirRekening
    };

    let rowValues = headers.map((h, idx) => (h in values) ? values[h] : (existingRow[idx] !== undefined ? existingRow[idx] : ""));

    if (found) {
      sheet.getRange(found.rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowValues.length).setValues([rowValues]);
    }
    logActivity_(user.username, "SAVE_FINANCE_BALANCE", "Admin Keuangan", "Simpan saldo tanggal " + data.Date, "");
    return { success: true, message: "Data saldo berhasil disimpan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Hapus 1 baris data Saldo & Kas (dipakai tombol Hapus di tabel "Riwayat
// Saldo per Tanggal"). Dicari berdasarkan BalanceID supaya aman walau ada
// beberapa baris dengan tanggal yang sama (data lama/duplikat).
function deleteFinanceBalance(balanceId, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE);
    let rows = sheet.getDataRange().getValues();
    let headers = rows[0];
    let idCol = headers.indexOf("BalanceID");
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idCol] === balanceId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_FINANCE_BALANCE", "Admin Keuangan", "Hapus data saldo " + balanceId, "");
        return { success: true, message: "Data saldo berhasil dihapus." };
      }
    }
    return { success: false, message: "Data saldo tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// MUTASI SALDO (pindah dana internal Brangkas <-> Rekening PT)
// -----------------------------------------------------------------------
// Contoh pemakaian: Owner menyetorkan sebagian uang tunai di Brangkas ke
// Rekening PT (FromBucket=Brangkas, ToBucket=Rekening), atau menarik tunai
// dari Rekening ke Brangkas (sebaliknya). Setiap mutasi tercatat di sheet
// Finance_Balance_Mutations dan LANGSUNG mempengaruhi Saldo Akhir
// Brangkas & Rekening (snapshot) di tanggal ybs - lihat
// computeFinanceMutationNet_()/computeFinanceCashFlow_() dan
// recalcFinanceBalanceSaldoAkhirSnapshot_() di bawah.
// =====================================================================
const FINANCE_BUCKETS_ = ["Brangkas", "Rekening"];

// Setelah Finance_Balance_Mutations berubah (tambah/hapus mutasi) untuk
// suatu tanggal, baris Finance_Balance tanggal itu (kalau SUDAH pernah
// disimpan) perlu dihitung ulang kolom snapshot SaldoAkhirBrangkas/
// SaldoAkhirRekening-nya supaya konsisten dgn mutasi terbaru - TANPA
// mengubah Saldo Awal/Aktual/kolom lain yg sudah tersimpan. Kalau baris
// utk tanggal itu belum ada sama sekali, tidak perlu apa-apa (nanti
// otomatis benar begitu saveFinanceBalance/getFinanceBalance dipanggil
// utk tanggal tsb, karena keduanya sudah memanggil computeFinanceCashFlow_
// yang otomatis menyertakan mutasi terbaru).
function recalcFinanceBalanceSaldoAkhirSnapshot_(ss, dateStr) {
  const tz = Session.getScriptTimeZone();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE);
  let found = findFinanceBalanceRow_(sheet, dateStr);
  if (!found) return;
  let headers = found.headers;
  let obj = rowToObj_(headers, found.row);
  let flow = computeFinanceCashFlow_(ss, dateStr, tz);
  let saldoAwalBrangkas = Number(obj.SaldoAwalBrangkas) || 0;
  let saldoAwalRekening = Number(obj.SaldoAwalRekening) || 0;
  let aktualBrangkas = (obj.SaldoAkhirBrangkasAktual === "" || obj.SaldoAkhirBrangkasAktual === undefined || obj.SaldoAkhirBrangkasAktual === null) ? "" : Number(obj.SaldoAkhirBrangkasAktual);
  let aktualRekening = (obj.SaldoAkhirRekeningAktual === "" || obj.SaldoAkhirRekeningAktual === undefined || obj.SaldoAkhirRekeningAktual === null) ? "" : Number(obj.SaldoAkhirRekeningAktual);
  let hitungBrangkas = saldoAwalBrangkas + flow.cashBersih - flow.totalExpBrangkas + flow.mutasiNetBrangkas;
  let hitungRekening = saldoAwalRekening + flow.totalNonCash - flow.totalExpRekening + flow.mutasiNetRekening;
  let saldoAkhirBrangkas = aktualBrangkas === "" ? hitungBrangkas : aktualBrangkas;
  let saldoAkhirRekening = aktualRekening === "" ? hitungRekening : aktualRekening;
  let idxB = headers.indexOf("SaldoAkhirBrangkas"), idxR = headers.indexOf("SaldoAkhirRekening");
  if (idxB !== -1) sheet.getRange(found.rowIndex, idxB + 1).setValue(saldoAkhirBrangkas);
  if (idxR !== -1) sheet.getRange(found.rowIndex, idxR + 1).setValue(saldoAkhirRekening);
}

// Catat mutasi baru & langsung update snapshot Saldo Akhir tanggal ybs
// (kalau tanggal itu sudah punya baris Finance_Balance tersimpan).
function addFinanceMutation(data, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    if (!data || !data.Date) return { success: false, message: "Tanggal wajib diisi." };
    let fromBucket = data.FromBucket, toBucket = data.ToBucket;
    if (FINANCE_BUCKETS_.indexOf(fromBucket) === -1 || FINANCE_BUCKETS_.indexOf(toBucket) === -1) {
      return { success: false, message: "Bucket asal/tujuan tidak valid (harus Brangkas atau Rekening)." };
    }
    if (fromBucket === toBucket) return { success: false, message: "Bucket asal dan tujuan tidak boleh sama." };
    let amount = Number(data.Amount);
    if (!amount || amount <= 0) return { success: false, message: "Jumlah mutasi harus lebih dari 0." };

    const ss = getDatabase();
    ensureSheet_(ss, CONFIG.SHEETS.FINANCE_BALANCE_MUTATIONS, FINANCE_BALANCE_MUTATIONS_HEADERS);
    let now = new Date();
    let mutationId = genId_("FMU");
    ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE_MUTATIONS).appendRow([
      mutationId, data.Date, fromBucket, toBucket, amount, data.Notes || "", user.fullName, now
    ]);

    recalcFinanceBalanceSaldoAkhirSnapshot_(ss, data.Date);

    logActivity_(user.username, "ADD_FINANCE_MUTATION", "Admin Keuangan",
      "Mutasi saldo tanggal " + data.Date + ": " + fromBucket + " -> " + toBucket + " sebesar " + amount, "");
    return { success: true, message: "Mutasi saldo berhasil dicatat." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Hapus 1 baris Mutasi Saldo & sesuaikan kembali snapshot Saldo Akhir
// tanggal ybs (efek mutasi tsb otomatis "dibatalkan" dari perhitungan).
function deleteFinanceMutation(mutationId, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE_MUTATIONS);
    let rows = sheet.getDataRange().getValues();
    let headers = rows[0];
    let idCol = headers.indexOf("MutationID");
    let dateCol = headers.indexOf("Date");
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idCol] === mutationId) {
        let dateStr = normalizeDateStr_(rows[i][dateCol]);
        sheet.deleteRow(i + 1);
        recalcFinanceBalanceSaldoAkhirSnapshot_(ss, dateStr);
        logActivity_(user.username, "DELETE_FINANCE_MUTATION", "Admin Keuangan", "Hapus mutasi saldo " + mutationId, "");
        return { success: true, message: "Mutasi saldo berhasil dihapus." };
      }
    }
    return { success: false, message: "Data mutasi tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Riwayat Mutasi Saldo, opsional difilter rentang tanggal (startDate/
// endDate format "yyyy-MM-dd", INKLUSIF - kalau kosong berarti semua).
// Diurutkan terbaru dulu, dibatasi 500 baris supaya tetap ringan.
function getFinanceMutationHistory(startDate, endDate, token) {
  try {
    checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    let rows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE_MUTATIONS));
    if (startDate) rows = rows.filter(m => normalizeDateStr_(m.Date) >= startDate);
    if (endDate) rows = rows.filter(m => normalizeDateStr_(m.Date) <= endDate);
    rows.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
    let totalMatching = rows.length;
    rows = rows.slice(0, 500);
    return { success: true, history: rows, totalMatching: totalMatching };
  } catch (error) {
    return { success: false, message: error.message };
  }
}


// Tabel tracking seluruh histori penginputan Saldo per tanggal (dipakai
// tabel "Riwayat Saldo" di web app, submenu Saldo & Kas), diurutkan dari
// tanggal terbaru.
function getFinanceBalanceHistory(token) {
  try {
    checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    let data = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE));
    data.forEach(r => { r.Date = normalizeDateStr_(r.Date); });
    data.sort((a, b) => b.Date < a.Date ? -1 : (b.Date > a.Date ? 1 : 0));
    return { success: true, data: data };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getFinanceExpenses(dateStr, token) {
  try {
    checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    let data = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES))
      .filter(r => normalizeDateStr_(r.Date) === dateStr);
    return { success: true, data: data };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Dipakai Report bulanan/rentang & Dashboard - mengambil semua pengeluaran
// dalam rentang tanggal [startDate, endDate] (format yyyy-MM-dd, inklusif).
function getFinanceExpensesRange(startDate, endDate, token) {
  try {
    checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    let data = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES))
      .filter(r => { let d = normalizeDateStr_(r.Date); return d && d >= startDate && d <= endDate; });
    return { success: true, data: data };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Pengeluaran Harian (detail): Tanggal, Supplier, Belanja/Pengeluaran
// (Description), Harga (UnitPrice), Qty, Sat (Unit), Total Harga (Amount),
// Status Pembayaran (teks bebas dari Master Status Pembayaran - menentukan
// Bucket/Category: Omset/Brangkas/RekeningPT/BelumLunas), Pos Biaya
// (COACode/COAName dari Master Kode Akun), Pos Penggunaan (dari Master
// Pos Penggunaan/Kegiatan).
function saveFinanceExpense(data, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    if (!data.Date || !data.Description) return { success: false, message: "Tanggal dan Belanja/Pengeluaran wajib diisi." };
    if (!data.StatusPembayaran) return { success: false, message: "Status Pembayaran wajib dipilih." };

    const ss = getDatabase();

    // Tentukan Bucket (Category) dari Status Pembayaran yang dipilih, lewat Master Status Pembayaran.
    let statusRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_PAYMENT_STATUS));
    let statusRow = statusRows.find(s => s.StatusName === data.StatusPembayaran);
    let bucket = statusRow ? statusRow.Bucket : data.Bucket;
    if ([FINANCE_EXPENSE_CATEGORY.OMSET, FINANCE_EXPENSE_CATEGORY.BRANGKAS, FINANCE_EXPENSE_CATEGORY.REKENING, FINANCE_EXPENSE_CATEGORY.BELUM_LUNAS].indexOf(bucket) === -1) {
      return { success: false, message: "Status Pembayaran tidak valid/tidak dikenali. Cek Master Keuangan." };
    }

    // Nama Akun (COAName) dari Kode Akun (Pos Biaya) yang dipilih. Dibandingkan
    // via String() di kedua sisi karena Google Sheets otomatis mengubah
    // COACode yang berupa angka (mis. "5101") jadi tipe Number saat disimpan -
    // perbandingan === tanpa String() akan selalu gagal (Number !== String),
    // itulah penyebab COACode tersimpan tapi COAName selalu kosong.
    let coaName = "";
    if (data.COACode) {
      let coaRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_COA));
      let coaRow = coaRows.find(c => String(c.COACode) === String(data.COACode));
      coaName = coaRow ? coaRow.COAName : "";
    }

    let unitPrice = Number(data.UnitPrice) || 0;
    let qty = Number(data.Qty) || 1;
    let amount = data.Amount !== undefined && data.Amount !== "" ? Number(data.Amount) || 0 : unitPrice * qty;

    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES);
    let now = new Date();
    let headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    function colIdx_(name) { return headerRow.indexOf(name) + 1; }

    let valuesByHeader = {
      Date: data.Date, Category: bucket, Description: data.Description, Amount: amount,
      CreatedBy: user.fullName, Timestamp: now, Supplier: data.Supplier || "",
      UnitPrice: unitPrice, Qty: qty, Unit: data.Unit || "-", StatusPembayaran: data.StatusPembayaran,
      COACode: data.COACode || "", COAName: coaName, PosPenggunaan: data.PosPenggunaan || ""
    };

    if (data.isEdit && data.ExpenseID) {
      let rows = sheet.getDataRange().getValues();
      let idCol = headerRow.indexOf("ExpenseID");
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][idCol] === data.ExpenseID) {
          headerRow.forEach((h, idx) => {
            if (h === "ExpenseID") return;
            if (valuesByHeader.hasOwnProperty(h)) sheet.getRange(i + 1, idx + 1).setValue(valuesByHeader[h]);
          });
          logActivity_(user.username, "EDIT_FINANCE_EXPENSE", "Pengeluaran Harian", "Edit pengeluaran " + data.ExpenseID, "");
          return { success: true, message: "Pengeluaran berhasil diperbarui." };
        }
      }
      return { success: false, message: "Data pengeluaran tidak ditemukan." };
    }

    let newRow = headerRow.map(h => h === "ExpenseID" ? genId_("EXP") : (valuesByHeader.hasOwnProperty(h) ? valuesByHeader[h] : ""));
    sheet.appendRow(newRow);
    logActivity_(user.username, "CREATE_FINANCE_EXPENSE", "Pengeluaran Harian", "Tambah pengeluaran " + data.Description + " (" + bucket + ")", "");
    return { success: true, message: "Pengeluaran berhasil dicatat." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteFinanceExpense(expenseId, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES);
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === expenseId) {
        sheet.deleteRow(i + 1);
        logActivity_(user.username, "DELETE_FINANCE_EXPENSE", "Admin Keuangan", "Hapus pengeluaran " + expenseId, "");
        return { success: true, message: "Pengeluaran berhasil dihapus." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Import massal Pengeluaran Harian dari file Excel (lihat template unduhan
// di web app). Setiap baris divalidasi & disimpan lewat jalur yang sama
// persis dengan saveFinanceExpense (lookup Bucket dari Status Pembayaran,
// COAName dari Kode Akun) supaya konsisten & bebas dari salah ketik Bucket.
function bulkImportFinanceExpenses(rows, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_WRITE);
    const ss = getDatabase();
    let statusRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_PAYMENT_STATUS));
    let coaRows = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_COA));
    const sheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES);
    let headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let now = new Date();
    let inserted = 0, errors = [];

    rows.forEach((r, idx) => {
      try {
        if (!r.date || !r.description) { errors.push("Baris " + (idx + 2) + ": Tanggal / Belanja-Pengeluaran kosong, dilewati."); return; }
        if (!r.statusPembayaran) { errors.push("Baris " + (idx + 2) + ": Status Pembayaran kosong, dilewati."); return; }
        let statusRow = statusRows.find(s => s.StatusName === r.statusPembayaran);
        if (!statusRow) { errors.push("Baris " + (idx + 2) + ": Status Pembayaran '" + r.statusPembayaran + "' tidak ditemukan di Master Keuangan, dilewati."); return; }
        let coaName = "";
        if (r.coaCode) {
          let coaRow = coaRows.find(c => String(c.COACode) === String(r.coaCode));
          coaName = coaRow ? coaRow.COAName : "";
        }
        let unitPrice = Number(r.unitPrice) || 0;
        let qty = Number(r.qty) || 1;
        let amount = r.amount ? Number(r.amount) || 0 : unitPrice * qty;
        let valuesByHeader = {
          Date: r.date, Category: statusRow.Bucket, Description: r.description, Amount: amount,
          CreatedBy: user.fullName, Timestamp: now, Supplier: r.supplier || "",
          UnitPrice: unitPrice, Qty: qty, Unit: r.unit || "-", StatusPembayaran: r.statusPembayaran,
          COACode: r.coaCode || "", COAName: coaName, PosPenggunaan: r.posPenggunaan || ""
        };
        let newRow = headerRow.map(h => h === "ExpenseID" ? genId_("EXP") : (valuesByHeader.hasOwnProperty(h) ? valuesByHeader[h] : ""));
        sheet.appendRow(newRow);
        inserted++;
      } catch (rowErr) {
        errors.push("Baris " + (idx + 2) + ": " + rowErr.message);
      }
    });

    logActivity_(user.username, "IMPORT_FINANCE_EXPENSES", "Pengeluaran Harian", "Import " + inserted + " baris pengeluaran", "");
    return { success: true, inserted: inserted, errors: errors };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Rakit produk x rekap penjualan per outlet untuk satu tanggal - dipakai
// supaya report menampilkan SEMUA produk aktif di outlet itu (termasuk
// yang terjual 0), persis format laporan WA yang dikirim tiap hari.
function buildUnitProductBreakdown_(ss, unit, dateStr, tz) {
  const cfg = unitCfg_(unit);
  let products = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.PRODUCTS)).filter(p => {
    if (p.Status !== "Active") return false;
    let units = productUnits_(p.ApplicableUnit);
    return units.indexOf(unit) !== -1 || units.indexOf("General") !== -1;
  });
  let sales = getSheetDataAsObjects_(ss.getSheetByName(cfg.salesSheet))
    .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);

  let agg = {}; // key: ProductCode|Variant -> {qty, total}
  sales.forEach(r => {
    let key = r.ProductCode + "|" + (r.Variant || "-");
    if (!agg[key]) agg[key] = { qty: 0, total: 0 };
    agg[key].qty += Number(r.Qty) || 0;
    agg[key].total += Number(r.TotalPrice) || 0;
  });

  let lines = products.map(p => {
    let key = p.ProductCode + "|" + (p.Variant || "-");
    let a = agg[key] || { qty: 0, total: 0 };
    return { productName: p.ProductName, variant: p.Variant && p.Variant !== "-" ? p.Variant : "", unitPrice: Number(p.SellPrice) || 0, qty: a.qty, total: a.total };
  });

  // Produk yang lakunya lewat kode barang di luar Master Produk aktif saat
  // ini (mis. sudah dinonaktifkan) tetap ditampilkan supaya angkanya tidak hilang.
  let knownKeys = {};
  products.forEach(p => { knownKeys[p.ProductCode + "|" + (p.Variant || "-")] = true; });
  Object.keys(agg).forEach(key => {
    if (!knownKeys[key]) {
      let parts = key.split("|");
      lines.push({ productName: parts[0], variant: parts[1] !== "-" ? parts[1] : "", unitPrice: agg[key].qty ? Math.round(agg[key].total / agg[key].qty) : 0, qty: agg[key].qty, total: agg[key].total });
    }
  });

  let total = lines.reduce((t, l) => t + l.total, 0);
  let totalQty = lines.reduce((t, l) => t + (Number(l.qty) || 0), 0);
  return { lines: lines, total: total, totalQty: totalQty };
}

// Rekap Pendapatan per Kategori Produk untuk SEMUA outlet (Kasir Loket
// Masuk, Kasir Loket Wahana, Kasir Resto digabung jadi satu, per Kategori,
// Qty, Total) untuk SATU tanggal - ditampilkan di Report Admin Keuangan
// tepat di bawah tabel "Pendapatan Kasir Resto", lengkap dengan grafik.
// Diambil langsung dari kolom Category yang sudah ada di baris penjualan
// tiap outlet (Entrance_Sales/Ride_Sales/Resto_Sales) - tidak perlu join
// ulang ke Master Produk.
function buildCategoryBreakdownAllOutlets_(ss, dateStr, tz) {
  let agg = {};
  ["Entrance", "Ride", "Resto"].forEach(unit => {
    let cfg = unitCfg_(unit);
    let sales = getSheetDataAsObjects_(ss.getSheetByName(cfg.salesSheet))
      .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
    sales.forEach(r => {
      let cat = (r.Category || "").toString().trim() || "Tanpa Kategori";
      if (!agg[cat]) agg[cat] = { qty: 0, total: 0 };
      agg[cat].qty += Number(r.Qty) || 0;
      agg[cat].total += Number(r.TotalPrice) || 0;
    });
  });

  let lines = Object.keys(agg).map(cat => ({ category: cat, qty: agg[cat].qty, total: agg[cat].total }))
    .sort((a, b) => b.total - a.total);
  let total = lines.reduce((t, l) => t + l.total, 0);
  let totalQty = lines.reduce((t, l) => t + (Number(l.qty) || 0), 0);
  return { lines: lines, total: total, totalQty: totalQty };
}

// Rekap Total Penerimaan per Kasir (semua outlet, 1 baris per Nama Kasir +
// Posisi/outlet) untuk tabel "Total Penerimaan dari Kasir" di Report Admin
// Keuangan - ditampilkan setelah tabel Pendapatan Kasir Resto.
function buildKasirPenerimaanBreakdown_(ss, dateStr, tz) {
  let byKey = {};
  ["Entrance", "Ride", "Resto"].forEach(unit => {
    let cfg = unitCfg_(unit);
    let rows = getSheetDataAsObjects_(ss.getSheetByName(cfg.paySheet))
      .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
    rows.forEach(p => {
      let cashierName = p.Cashier || "-";
      let key = cashierName + "|" + cfg.label;
      if (!byKey[key]) byKey[key] = { cashierName: cashierName, posisi: cfg.label, tunai: 0, nonTunai: 0 };
      let amt = Number(p.Amount) || 0;
      let method = (p.PaymentMethod || "").toLowerCase();
      if (method.includes("cash") || method.includes("tunai")) byKey[key].tunai += amt; else byKey[key].nonTunai += amt;
    });
  });
  let order = { "Kasir Loket Masuk": 1, "Kasir Loket Wahana": 2, "Kasir Resto": 3 };
  let lines = Object.keys(byKey).map(k => {
    let l = byKey[k];
    l.jumlah = l.tunai + l.nonTunai;
    return l;
  }).sort((a, b) => (order[a.posisi] || 9) - (order[b.posisi] || 9) || a.cashierName.localeCompare(b.cashierName));
  let totalTunai = lines.reduce((t, l) => t + l.tunai, 0);
  let totalNonTunai = lines.reduce((t, l) => t + l.nonTunai, 0);
  return { lines: lines, totalTunai: totalTunai, totalNonTunai: totalNonTunai, totalJumlah: totalTunai + totalNonTunai };
}

// Report harian lengkap Admin Keuangan: gabungan otomatis dari data Kasir
// (Entrance + Ride + Resto) + input manual (Saldo, Pengeluaran, Hutang/
// Piutang). Dipakai untuk tampilan preview, export PDF/XLSX, dan disalin
// sebagai teks format WA.
function getFinanceDailyReport(dateStr, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    const tz = Session.getScriptTimeZone();
    let settings = getSystemSettings();

    let entrance = buildUnitProductBreakdown_(ss, "Entrance", dateStr, tz);
    let ride = buildUnitProductBreakdown_(ss, "Ride", dateStr, tz);
    let resto = buildUnitProductBreakdown_(ss, "Resto", dateStr, tz);
    let categoryBreakdown = buildCategoryBreakdownAllOutlets_(ss, dateStr, tz);

    function paySum_(unit) {
      let cfg = unitCfg_(unit);
      let rows = getSheetDataAsObjects_(ss.getSheetByName(cfg.paySheet))
        .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
      let b = { cash: 0, nonCash: 0 };
      rows.forEach(p => {
        let amt = Number(p.Amount) || 0;
        let method = (p.PaymentMethod || "").toLowerCase();
        if (method.includes("cash") || method.includes("tunai")) b.cash += amt; else b.nonCash += amt;
      });
      return b;
    }
    let payE = paySum_("Entrance"), payR = paySum_("Ride"), payS = paySum_("Resto");
    let totalCash = payE.cash + payR.cash + payS.cash;
    let totalNonCash = payE.nonCash + payR.nonCash + payS.nonCash;
    // Total Pendapatan Kotor = Total yang DITERIMA dari Kasir (Loket Masuk,
    // Loket Wahana, Resto), Tunai + Non Tunai (QRIS/Transfer/Debit/Kredit) -
    // bukan lagi dari total penjualan per produk, supaya konsisten dengan
    // tabel "Total Penerimaan dari Kasir".
    let totalPendapatanKotor = totalCash + totalNonCash;

    let expenses = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES)).filter(r => normalizeDateStr_(r.Date) === dateStr);
    function expByCat_(cat) { return expenses.filter(e => e.Category === cat); }
    let expOmset = expByCat_(FINANCE_EXPENSE_CATEGORY.OMSET);
    let expBrangkas = expByCat_(FINANCE_EXPENSE_CATEGORY.BRANGKAS);
    let expRekening = expByCat_(FINANCE_EXPENSE_CATEGORY.REKENING);
    let expBelumLunas = expByCat_(FINANCE_EXPENSE_CATEGORY.BELUM_LUNAS);
    let totalExpOmset = expOmset.reduce((t, e) => t + (Number(e.Amount) || 0), 0);
    let totalExpBrangkas = expBrangkas.reduce((t, e) => t + (Number(e.Amount) || 0), 0);
    let totalExpRekening = expRekening.reduce((t, e) => t + (Number(e.Amount) || 0), 0);
    let totalExpBelumLunas = expBelumLunas.reduce((t, e) => t + (Number(e.Amount) || 0), 0);
    // Catatan: Pengeluaran dari Omset Harian TIDAK lagi ditampilkan sebagai
    // bagian terpisah di Report (sudah dihapus dari tampilan/PDF/XLSX/WA),
    // sehingga tidak diikutkan lagi di Total Pengeluaran Seluruhnya. Nilainya
    // tetap dihitung (totalExpOmset) karena masih dipakai untuk menghitung
    // Saldo Akhir Brangkas (kas fisik di laci setelah dipotong pengeluaran
    // yang dibayar langsung dari omset harian).
    let totalExpSeluruhnya = totalExpBrangkas + totalExpRekening;

    // Total Pendapatan Bersih = sama dengan Total Pendapatan Kotor (tidak
    // lagi dikurangi Pengeluaran dari Omset Harian).
    let totalPendapatanBersih = totalPendapatanKotor;
    // cashBersih tetap dipakai khusus untuk estimasi Saldo Akhir Brangkas
    // (kas fisik) - asumsi pengeluaran dari omset harian dibayar tunai dari
    // laci kasir, sehingga mengurangi porsi Cash yang disetor ke brangkas.
    let cashBersih = totalCash - totalExpOmset;
    // Saldo Bersih = Pendapatan Bersih dikurangi Total Pengeluaran
    // Seluruhnya (Brangkas + Rekening) pada hari itu.
    let saldoBersih = totalPendapatanBersih - totalExpSeluruhnya;

    const balSheet = ss.getSheetByName(CONFIG.SHEETS.FINANCE_BALANCE);
    let balFound = findFinanceBalanceRow_(balSheet, dateStr);
    let bal = balFound ? rowToObj_(balFound.headers, balFound.row) : null;
    if (!bal) {
      let prev = findLatestFinanceBalanceBefore_(balSheet, dateStr);
      bal = prev ? rowToObj_(prev.headers, prev.row) : {};
      if (prev) {
        let resolvedPrev = resolvedSaldoAkhir_(bal);
        bal.SaldoAwalBrangkas = resolvedPrev.brangkas;
        bal.SaldoAwalRekening = resolvedPrev.rekening;
      } else {
        bal.SaldoAwalBrangkas = 0; bal.SaldoAwalRekening = 0;
      }
      bal.SaldoTHR = prev ? (Number(bal.SaldoTHR) || 0) : 0;
      bal.SaldoAkhirBrangkasAktual = ""; bal.SaldoAkhirRekeningAktual = "";
      bal.SaldoAkhirBrangkas = ""; bal.SaldoAkhirRekening = "";
      bal.CatatanTambahan = ""; bal.HutangJSON = "[]"; bal.PiutangJSON = "[]";
    }
    let saldoAwalBrangkas = Number(bal.SaldoAwalBrangkas) || 0;
    let saldoAwalRekening = Number(bal.SaldoAwalRekening) || 0;
    let saldoTHR = Number(bal.SaldoTHR) || 0;
    let saldoAwalTotal = saldoAwalBrangkas + saldoAwalRekening;

    // Saldo Akhir HITUNGAN (formula kas masuk-keluar hari ini) dipakai
    // sebagai ESTIMASI selama tanggal ini BELUM disimpan. Begitu tanggal
    // ini SUDAH disimpan (balFound), report memakai snapshot FINAL yang
    // sudah diresolve & dikunci saat Simpan Saldo (SaldoAkhirBrangkas/
    // SaldoAkhirRekening) - bukan dihitung ulang live - supaya angka
    // laporan yang sudah pernah dicetak/dikirim tidak diam-diam berubah
    // kalau ada transaksi tanggal ini yang diedit belakangan.
    let saldoAkhirBrangkasHitung = saldoAwalBrangkas + cashBersih - totalExpBrangkas;
    let saldoAkhirRekeningHitung = saldoAwalRekening + totalNonCash - totalExpRekening;
    let saldoAkhirBrangkas, saldoAkhirRekening;
    if (balFound) {
      let resolvedNow = resolvedSaldoAkhir_(bal);
      saldoAkhirBrangkas = resolvedNow.brangkas;
      saldoAkhirRekening = resolvedNow.rekening;
    } else {
      saldoAkhirBrangkas = saldoAkhirBrangkasHitung;
      saldoAkhirRekening = saldoAkhirRekeningHitung;
    }
    let saldoAkhirTotal = saldoAkhirBrangkas + saldoAkhirRekening;

    let hutang = []; let piutang = [];
    try { hutang = JSON.parse(bal.HutangJSON || "[]"); } catch (e) {}
    try { piutang = JSON.parse(bal.PiutangJSON || "[]"); } catch (e) {}

    // Rekap Omset & Pengeluaran bulan berjalan (sampai dengan tanggal
    // laporan), sesuai bagian NOTE di format laporan WA.
    let monthPrefix = dateStr.slice(0, 7); // yyyy-MM
    function monthSalesTotal_(unit) {
      let cfg = unitCfg_(unit);
      return getSheetDataAsObjects_(ss.getSheetByName(cfg.salesSheet)).reduce((t, r) => {
        if (r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd").slice(0, 7) === monthPrefix &&
          Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") <= dateStr) t += Number(r.TotalPrice) || 0;
        return t;
      }, 0);
    }
    let omsetBulanIni = monthSalesTotal_("Entrance") + monthSalesTotal_("Ride") + monthSalesTotal_("Resto");
    let pengeluaranBulanIni = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.FINANCE_EXPENSES))
      .filter(e => {
        let d = normalizeDateStr_(e.Date);
        return d && d.slice(0, 7) === monthPrefix && d <= dateStr;
      })
      .reduce((t, e) => t + (Number(e.Amount) || 0), 0);

    let kasirPenerimaan = buildKasirPenerimaanBreakdown_(ss, dateStr, tz);

    // PERBAIKAN: Subtotal Qty per outlet (Loket Masuk/Loket Wahana/Resto) +
    // Total Qty Rekapan Pendapatan Per Kategori (SEMUA outlet, harus sama
    // dengan jumlah 3 subtotal Qty di atas karena sumber datanya sama) -
    // ditampilkan di Report Admin Keuangan (layar, PDF, XLSX, format WA).
    let totalQtyEntrance = entrance.totalQty || 0;
    let totalQtyRide = ride.totalQty || 0;
    let totalQtyResto = resto.totalQty || 0;
    let totalQtyKategori = categoryBreakdown.totalQty || 0;

    // Data ringkas untuk grafik-grafik di Report Admin Keuangan (ditampilkan
    // di atas, sebelum bagian Saldo Awal).
    let chartData = {
      outletLabels: [UNIT_MAP.Entrance.label, UNIT_MAP.Ride.label, UNIT_MAP.Resto.label],
      outletValues: [entrance.total, ride.total, resto.total],
      paymentLabels: ["Tunai", "Non Tunai"],
      paymentValues: [totalCash, totalNonCash],
      expenseLabels: ["Brangkas", "Rekening PT"],
      expenseValues: [totalExpBrangkas, totalExpRekening],
      categoryLabels: categoryBreakdown.lines.map(l => l.category),
      categoryValues: categoryBreakdown.lines.map(l => l.total)
    };

    return {
      success: true, dateStr: dateStr,
      companyName: settings.COMPANY_NAME || CONFIG.COMPANY_NAME, address: settings.ADDRESS || "-",
      logoUrl: settings.LOGO_URL || "", cashierName: user.fullName, cashierPosition: "Admin Keuangan",
      printedAt: new Date().toISOString(),
      saldoAwalBrangkas: saldoAwalBrangkas, saldoAwalRekening: saldoAwalRekening, saldoAwalTotal: saldoAwalTotal, saldoTHR: saldoTHR,
      entrance: entrance, ride: ride, resto: resto, categoryBreakdown: categoryBreakdown, totalPendapatanKotor: totalPendapatanKotor,
      expOmset: expOmset, expBrangkas: expBrangkas, expRekening: expRekening, expBelumLunas: expBelumLunas,
      totalExpOmset: totalExpOmset, totalExpBrangkas: totalExpBrangkas, totalExpRekening: totalExpRekening,
      totalExpBelumLunas: totalExpBelumLunas, totalExpSeluruhnya: totalExpSeluruhnya,
      totalPendapatanBersih: totalPendapatanBersih, totalCash: totalCash, totalNonCash: totalNonCash, cashBersih: cashBersih,
      saldoBersih: saldoBersih,
      saldoAkhirBrangkasHitung: saldoAkhirBrangkasHitung, saldoAkhirRekeningHitung: saldoAkhirRekeningHitung,
      saldoAkhirBrangkas: saldoAkhirBrangkas, saldoAkhirRekening: saldoAkhirRekening, saldoAkhirTotal: saldoAkhirTotal,
      isBalanceSaved: !!balFound, hutang: hutang, piutang: piutang, catatanTambahan: bal.CatatanTambahan || "",
      omsetBulanIni: omsetBulanIni, pengeluaranBulanIni: pengeluaranBulanIni,
      kasirPenerimaan: kasirPenerimaan, chartData: chartData,
      totalQtyEntrance: totalQtyEntrance, totalQtyRide: totalQtyRide, totalQtyResto: totalQtyResto,
      totalQtyKategori: totalQtyKategori
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 11C. ADMIN KEUANGAN - REPORT SEMUA KASIR (Loket Masuk + Loket Wahana +
// Resto digabung dalam SATU file, lengkap dengan rekapan Total Penerimaan
// Seluruh Kasir per metode pembayaran & Total Penjualan per outlet).
// -----------------------------------------------------------------------
// Dipisah dari getFinanceDailyReport() karena tujuannya berbeda: laporan
// ini FOKUS ke rekapan penjualan & penerimaan tiap kasir per metode
// pembayaran (Tunai/QRIS/Transfer/Debit/Kredit) - bukan Saldo Brangkas/
// Rekening/Pengeluaran seperti Report Harian Admin Keuangan.
// "Balance": Total Penjualan Seluruh Kasir (dari data Penjualan/Sales)
// HARUS sama dengan Total Penerimaan Seluruh Kasir (dari data Penerimaan/
// Payments) kalau semua transaksi sudah tercatat lengkap & konsisten -
// selisih (kalau ada) dikirim ke client supaya kelihatan kalau ada data
// yang belum balance.
function buildUnitPaymentMethodBreakdown_(ss, unit, dateStr, tz) {
  const cfg = unitCfg_(unit);
  let rows = getSheetDataAsObjects_(ss.getSheetByName(cfg.paySheet))
    .filter(r => r.Date && Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd") === dateStr);
  let b = { cash: 0, qris: 0, transfer: 0, debit: 0, kredit: 0 };
  rows.forEach(p => {
    let amt = Number(p.Amount) || 0;
    let method = (p.PaymentMethod || "").toLowerCase();
    if (method.includes("debit")) b.debit += amt;
    else if (method.includes("kredit") || method.includes("credit")) b.kredit += amt;
    else if (method.includes("cash") || method.includes("tunai")) b.cash += amt;
    else if (method.includes("qris")) b.qris += amt;
    else if (method.includes("transfer")) b.transfer += amt;
  });
  b.total = b.cash + b.qris + b.transfer + b.debit + b.kredit;
  return b;
}

function getAllKasirReportData(dateStr, token) {
  try {
    let user = checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    const tz = Session.getScriptTimeZone();
    let settings = getSystemSettings();

    let entrance = buildUnitProductBreakdown_(ss, "Entrance", dateStr, tz);
    let ride = buildUnitProductBreakdown_(ss, "Ride", dateStr, tz);
    let resto = buildUnitProductBreakdown_(ss, "Resto", dateStr, tz);

    let payEntrance = buildUnitPaymentMethodBreakdown_(ss, "Entrance", dateStr, tz);
    let payRide = buildUnitPaymentMethodBreakdown_(ss, "Ride", dateStr, tz);
    let payResto = buildUnitPaymentMethodBreakdown_(ss, "Resto", dateStr, tz);

    let totalCash = payEntrance.cash + payRide.cash + payResto.cash;
    let totalQris = payEntrance.qris + payRide.qris + payResto.qris;
    let totalTransfer = payEntrance.transfer + payRide.transfer + payResto.transfer;
    let totalDebit = payEntrance.debit + payRide.debit + payResto.debit;
    let totalKredit = payEntrance.kredit + payRide.kredit + payResto.kredit;
    let totalPenerimaanSeluruhKasir = totalCash + totalQris + totalTransfer + totalDebit + totalKredit;
    let totalPenjualanSeluruhKasir = entrance.total + ride.total + resto.total;
    // PERBAIKAN: Subtotal Qty per kasir (Loket Masuk/Loket Wahana/Resto) +
    // Total Qty seluruh kasir untuk Report Semua Kasir.
    let totalQtyEntrance = entrance.totalQty || 0;
    let totalQtyRide = ride.totalQty || 0;
    let totalQtyResto = resto.totalQty || 0;
    let totalQtySeluruhKasir = totalQtyEntrance + totalQtyRide + totalQtyResto;

    return {
      success: true, dateStr: dateStr,
      companyName: settings.COMPANY_NAME || CONFIG.COMPANY_NAME, address: settings.ADDRESS || "-",
      logoUrl: settings.LOGO_URL || "",
      cashierName: user.fullName, cashierPosition: "Admin Keuangan - Report Semua Kasir",
      printedAt: new Date().toISOString(),
      entrance: entrance, ride: ride, resto: resto,
      payment: { Entrance: payEntrance, Ride: payRide, Resto: payResto },
      totals: {
        totalCash: totalCash, totalQris: totalQris, totalTransfer: totalTransfer, totalDebit: totalDebit, totalKredit: totalKredit,
        totalPenerimaanSeluruhKasir: totalPenerimaanSeluruhKasir,
        totalPenjualanEntrance: entrance.total, totalPenjualanRide: ride.total, totalPenjualanResto: resto.total,
        totalPenjualanSeluruhKasir: totalPenjualanSeluruhKasir,
        selisih: totalPenjualanSeluruhKasir - totalPenerimaanSeluruhKasir,
        totalQtyEntrance: totalQtyEntrance, totalQtyRide: totalQtyRide, totalQtyResto: totalQtyResto,
        totalQtySeluruhKasir: totalQtySeluruhKasir
      }
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// -----------------------------------------------------------------------
// WIDGET DASHBOARD "BALANCE REPORT SEMUA KASIR"
// -----------------------------------------------------------------------
// Mendeteksi tanggal & kasir (outlet) mana saja yang laporannya TIDAK
// BALANCE antara Total Penjualan (data Penjualan/Sales per outlet) dengan
// Total Penerimaan Kasir (data Penerimaan/Payments per outlet - baik yang
// otomatis tercatat per-item transaksi kalau SHOW_PAYMENT_METHOD_IN_SALES
// aktif, MAUPUN yang diinput manual lewat menu "Penerimaan Uang" per hari).
// Konsepnya sama persis dengan "selisih" di getAllKasirReportData() (Report
// Semua Kasir per 1 tanggal), tapi di sini dijalankan untuk SATU RENTANG
// tanggal sekaligus dan per-outlet, supaya Owner bisa langsung lihat di
// Dashboard tanggal & outlet mana yang perlu ditelusuri - tidak perlu buka
// Report Semua Kasir satu-satu per tanggal.
//
// PENTING soal performa: sheet Sales/Payment tiap outlet HANYA dibaca 1x
// (bukan 1x per tanggal per outlet) lalu dikelompokkan per tanggal di
// memori JS - supaya rentang panjang (mis. "Bulan Ini"/"Tahun Ini") tetap
// ringan, sama seperti pola yang sudah dipakai getDashboardData().
function getBalanceCheckReport(token, startDate, endDate) {
  try {
    checkPermission_(token, FINANCE_ROLES_READ);
    const ss = getDatabase();
    const tz = Session.getScriptTimeZone();
    let todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    let rangeStart = startDate || todayStr;
    let rangeEnd = endDate || todayStr;
    if (rangeStart > rangeEnd) { let t = rangeStart; rangeStart = rangeEnd; rangeEnd = t; }

    const UNITS_ = ["Entrance", "Ride", "Resto"];
    let mismatches = [];

    UNITS_.forEach(function(unit) {
      const cfg = unitCfg_(unit);

      // Total Penjualan per tanggal, dari sheet Sales outlet ini.
      let salesRows = getSheetDataAsObjects_(ss.getSheetByName(cfg.salesSheet));
      let salesByDate = {};
      salesRows.forEach(function(r) {
        if (!r.Date) return;
        let d = Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd");
        if (d < rangeStart || d > rangeEnd) return;
        salesByDate[d] = (salesByDate[d] || 0) + (Number(r.TotalPrice) || 0);
      });

      // Total Penerimaan Kasir per tanggal, dari sheet Payment outlet ini
      // (gabungan Tunai/QRIS/Transfer/Debit/Kredit apa pun sumbernya).
      let payRows = getSheetDataAsObjects_(ss.getSheetByName(cfg.paySheet));
      let payByDate = {};
      payRows.forEach(function(r) {
        if (!r.Date) return;
        let d = Utilities.formatDate(new Date(r.Date), tz, "yyyy-MM-dd");
        if (d < rangeStart || d > rangeEnd) return;
        payByDate[d] = (payByDate[d] || 0) + (Number(r.Amount) || 0);
      });

      // Gabungan semua tanggal yang punya data di salah satu sisi (Penjualan
      // ATAU Penerimaan) - tanggal tanpa transaksi sama sekali di kedua sisi
      // tidak perlu dicek (tidak relevan, otomatis balance/kosong).
      let allDates = {};
      Object.keys(salesByDate).forEach(function(d) { allDates[d] = true; });
      Object.keys(payByDate).forEach(function(d) { allDates[d] = true; });

      Object.keys(allDates).sort().forEach(function(d) {
        let totalSales = salesByDate[d] || 0;
        let totalPenerimaan = payByDate[d] || 0;
        // Pembulatan ke Rupiah terdekat supaya sisa desimal hasil pembagian
        // (mis. dari harga fleksibel) tidak dianggap "tidak balance" palsu.
        let selisih = Math.round(totalSales) - Math.round(totalPenerimaan);
        if (selisih !== 0) {
          mismatches.push({
            date: d, unit: unit, unitLabel: cfg.label,
            totalSales: totalSales, totalPenerimaan: totalPenerimaan, selisih: selisih
          });
        }
      });
    });

    // Terbaru dulu, lalu per outlet - supaya masalah paling baru langsung
    // kelihatan di paling atas widget.
    mismatches.sort(function(a, b) {
      if (a.date !== b.date) return a.date < b.date ? 1 : -1;
      return a.unitLabel < b.unitLabel ? -1 : (a.unitLabel > b.unitLabel ? 1 : 0);
    });

    return {
      success: true, rangeStart: rangeStart, rangeEnd: rangeEnd,
      mismatches: mismatches, mismatchCount: mismatches.length
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 12. USERS - MANAJEMEN USER (password tidak pernah dikirim ke client)
// =====================================================================
function getUsers(token) {
  try {
    // Manajemen User khusus Owner (sesuai kebijakan akses internal).
    checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    let users = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.USERS));
    users.forEach(u => { delete u.PasswordHash; delete u.PasswordPlain; });
    return { success: true, data: users };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveUser(userData, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);

    if (userData.isEdit) {
      if (userData.password && user.role !== ROLE.OWNER) {
        return { success: false, message: "Hanya Owner yang dapat mengubah password user." };
      }
      let data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === userData.userId) {
          sheet.getRange(i + 1, 2).setValue(userData.username);
          sheet.getRange(i + 1, 3).setValue(userData.fullName);
          sheet.getRange(i + 1, 4).setValue(userData.role);
          sheet.getRange(i + 1, 5).setValue(userData.role === ROLE.KASIR ? userData.assignedUnit : "All");
          sheet.getRange(i + 1, 7).setValue(userData.status);
          if (userData.password) {
            sheet.getRange(i + 1, 6).setValue(hashPassword_(userData.password));
            sheet.getRange(i + 1, 9).setValue(userData.password);
          }
          logActivity_(user.username, "EDIT_USER", "Manajemen User", "Edit user " + userData.username, "");
          return { success: true, message: "User berhasil diperbarui." };
        }
      }
      return { success: false, message: "User tidak ditemukan." };
    } else {
      let newId = genId_("U");
      let plainPwd = userData.password || "default123";
      let hashed = hashPassword_(plainPwd);
      sheet.appendRow([newId, userData.username, userData.fullName, userData.role,
        userData.role === ROLE.KASIR ? userData.assignedUnit : "All", hashed, "Active", new Date(), plainPwd]);
      logActivity_(user.username, "CREATE_USER", "Manajemen User", "Tambah user " + userData.username, "");
      return { success: true, message: "User berhasil ditambahkan." };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteUser(userId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        if (data[i][1] === "owner") return { success: false, message: "Akun Super Owner utama tidak dapat dihapus." };
        if (data[i][0] === user.userId) return { success: false, message: "Anda tidak dapat menghapus akun sendiri." };
        sheet.deleteRow(i + 1);
        return { success: true, message: "User berhasil dihapus." };
      }
    }
    return { success: false, message: "User tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Khusus Owner: lihat password login (plain text) milik user manapun.
// Password TIDAK PERNAH ikut terkirim lewat getUsers() - hanya bisa diambil
// satu-per-satu lewat fungsi ini, dan setiap pemanggilan dicatat di Activity
// Log supaya ada jejak audit siapa yang melihat password siapa dan kapan.
function getUserPassword(userId, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    let data = sheet.getDataRange().getValues();
    let headers = data[0];
    let idCol = headers.indexOf("UserID");
    let userCol = headers.indexOf("Username");
    let plainCol = headers.indexOf("PasswordPlain");
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === userId) {
        let plainPwd = plainCol !== -1 ? data[i][plainCol] : "";
        logActivity_(user.username, "VIEW_PASSWORD", "Manajemen User", "Lihat password user " + data[i][userCol], "");
        return {
          success: true,
          username: data[i][userCol],
          password: plainPwd ? plainPwd : "(Password lama - belum tersimpan, silakan set ulang password user ini)"
        };
      }
    }
    return { success: false, message: "User tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 13. SETTINGS - PENGATURAN UMUM (Owner only) & LOGO
// =====================================================================
function getSystemSettings() {
  const ss = getDatabase();
  let data = getSheetDataAsObjects_(ss.getSheetByName(CONFIG.SHEETS.SETTINGS));
  let settings = {};
  data.forEach(row => { settings[row.Key] = row.Value; });
  return settings;
}

function saveSystemSettings(settingsObj, token) {
  try {
    if (token) checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
    let data = sheet.getDataRange().getValues();

    for (let key in settingsObj) {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(settingsObj[key]); found = true; break; }
      }
      if (!found) sheet.appendRow([key, settingsObj[key], "Custom Setting"]);
    }
    return { success: true, message: "Pengaturan berhasil disimpan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 13b. PENGATURAN MENU & IKON (kustomisasi nama & ikon menu/submenu sidebar)
// -----------------------------------------------------------------------
// Struktur default seluruh menu & submenu yang bisa dikustomisasi Owner
// lewat Pengaturan > Menu & Ikon. `icon` memakai nama class Font Awesome
// (mis. "fa-chart-pie", TANPA prefix "fa-solid") - HANYA item yang tampil
// dengan ikon di sidebar (menu utama/grup) yang punya properti `icon`;
// item submenu (di dalam grup, menjorok ke dalam) hanya punya `label`
// karena memang tidak ditampilkan dengan ikon di UI saat ini.
// Override disimpan sebagai SATU baris JSON di sheet Settings (key
// MENU_CONFIG_JSON) lewat getSystemSettings()/saveSystemSettings() yang
// sudah ada, supaya tidak perlu sheet/skema baru.
const DEFAULT_MENU_CONFIG_ = {
  dashboard:             { label: "Dashboard",                      icon: "fa-chart-pie",    group: "Menu Utama" },
  unit_Entrance:         { label: "Kasir Loket Masuk",               icon: "fa-ticket",       group: "Menu Utama" },
  unit_Ride:             { label: "Kasir Loket Wahana",              icon: "fa-ferris-wheel", group: "Menu Utama" },
  unit_Resto:            { label: "Kasir Resto",                    icon: "fa-utensils",     group: "Menu Utama" },
  finance:               { label: "Admin Keuangan",                 icon: "fa-coins",        group: "Menu Utama" },
  warehouse:             { label: "Warehouse",                       icon: "fa-warehouse",    group: "Menu Utama" },
  master:                { label: "Master Produk",                  icon: "fa-database",     group: "Menu Utama" },
  settings:              { label: "Pengaturan",                     icon: "fa-gear",         group: "Menu Utama" },

  sub_sales:                   { label: "Laporan Penjualan",             group: "Submenu Outlet Kasir (berlaku utk ke-3 outlet)" },
  sub_payment:                  { label: "Penerimaan Uang",               group: "Submenu Outlet Kasir (berlaku utk ke-3 outlet)" },
  sub_stock:                    { label: "Manajemen Stok",                group: "Submenu Outlet Kasir (berlaku utk ke-3 outlet)" },
  sub_physicalstock:            { label: "Cek Stok Fisik Barang",         group: "Submenu Outlet Kasir (berlaku utk ke-3 outlet)" },
  sub_stockcard:                { label: "Kartu Stok",                    group: "Submenu Outlet Kasir (berlaku utk ke-3 outlet)" },
  sub_report:                   { label: "Report",                       group: "Submenu Outlet Kasir (berlaku utk ke-3 outlet)" },

  sub_finance_balance:          { label: "Saldo & Kas",                   group: "Submenu Admin Keuangan" },
  sub_finance_expenses:         { label: "Pengeluaran Harian",            group: "Submenu Admin Keuangan" },
  sub_finance_report:           { label: "Report",                       group: "Submenu Admin Keuangan" },
  sub_finance_report_allkasir:  { label: "Report Semua Kasir",            group: "Submenu Admin Keuangan" },
  sub_finance_master:           { label: "Master Keuangan",               group: "Submenu Admin Keuangan" },

  sub_warehouse_summary:        { label: "Ringkasan Semua Gudang",        group: "Submenu Warehouse" },
  sub_warehouse_stockcard:      { label: "Kartu Stok Gudang",             group: "Submenu Warehouse" },

  sub_catpay:                   { label: "Kategori & Metode Pembayaran",  group: "Submenu Pengaturan" },
  sub_users:                    { label: "Manajemen User",                group: "Submenu Pengaturan" },
  sub_settings:                 { label: "Profil Aplikasi & Fitur",       group: "Submenu Pengaturan" },
  sub_menuconfig:               { label: "Menu & Ikon",                   group: "Submenu Pengaturan" }
};

// Dibaca oleh SEMUA role (dipakai untuk render sidebar tiap user login),
// karenanya hanya memvalidasi sesi (allowedRoles = null), tidak dibatasi
// Owner saja. Mengembalikan hasil GABUNGAN antara default & override yang
// tersimpan - kalau field label/icon override kosong/tidak ada, otomatis
// jatuh kembali (fallback) ke nilai default.
function getMenuConfig(token) {
  try {
    checkPermission_(token, null);
    let settings = getSystemSettings();
    let overrides = {};
    if (settings.MENU_CONFIG_JSON) {
      try { overrides = JSON.parse(settings.MENU_CONFIG_JSON); } catch (e) { overrides = {}; }
    }
    let merged = {};
    for (let key in DEFAULT_MENU_CONFIG_) {
      let d = DEFAULT_MENU_CONFIG_[key];
      let o = overrides[key] || {};
      merged[key] = {
        label: (o.label && String(o.label).trim() !== "") ? String(o.label).trim() : d.label,
        icon: d.icon ? ((o.icon && String(o.icon).trim() !== "") ? String(o.icon).trim() : d.icon) : undefined,
        group: d.group
      };
    }
    return { success: true, config: merged };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// overrides: { [key]: { label, icon } , ... } - hanya key yang dikenal di
// DEFAULT_MENU_CONFIG_ yang disimpan (mencegah data sampah), dan hanya
// item yang memang punya slot `icon` di default yang boleh menyimpan
// override icon (submenu tidak dikirim ke sidebar dengan ikon).
function saveMenuConfig(overrides, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    let clean = {};
    for (let key in DEFAULT_MENU_CONFIG_) {
      if (!overrides || !overrides[key]) continue;
      let d = DEFAULT_MENU_CONFIG_[key];
      let o = overrides[key];
      let entry = {};
      if (o.label !== undefined && String(o.label).trim() !== "") entry.label = String(o.label).trim();
      if (d.icon && o.icon !== undefined && String(o.icon).trim() !== "") {
        entry.icon = String(o.icon).trim().replace(/^fa-solid\s+/, "").replace(/^fa[srb]?\s+/, "");
      }
      if (Object.keys(entry).length) clean[key] = entry;
    }
    saveSystemSettings({ MENU_CONFIG_JSON: JSON.stringify(clean) }, token);
    logActivity_(user.username, "UPDATE_MENU_CONFIG", "Pengaturan", "Perbarui nama/ikon menu & submenu sidebar", "");
    return { success: true, message: "Pengaturan menu & ikon berhasil disimpan." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function resetMenuConfig(token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    saveSystemSettings({ MENU_CONFIG_JSON: "" }, token);
    logActivity_(user.username, "RESET_MENU_CONFIG", "Pengaturan", "Kembalikan nama & ikon menu ke default", "");
    return { success: true, message: "Menu & ikon dikembalikan ke pengaturan default." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Logo disimpan sebagai data URL (base64) langsung di Settings, BUKAN link
// Google Drive. Ini menghindari masalah umum "logo tidak terbaca/tidak
// tampil" akibat pembatasan sharing/hotlink Google Drive untuk gambar -
// data URL selalu tampil instan di browser manapun tanpa perlu izin apapun.
function uploadLogo(base64Data, mimeType, fileName, token) {
  try {
    let user = checkPermission_(token, [ROLE.OWNER]);
    let allowedTypes = ["image/png", "image/jpeg", "image/jpg"];
    if (allowedTypes.indexOf(mimeType) === -1) {
      return { success: false, message: "Format logo harus PNG atau JPG." };
    }
    // Batas aman supaya muat disimpan sebagai 1 sel di Google Sheets (limit ~50.000 karakter/sel)
    if (base64Data.length > 60000) {
      return { success: false, message: "Ukuran file logo terlalu besar. Gunakan gambar PNG/JPG di bawah ±40KB." };
    }

    let dataUrl = "data:" + mimeType + ";base64," + base64Data;
    saveSystemSettings({ LOGO_URL: dataUrl }, null);
    logActivity_(user.username, "UPLOAD_LOGO", "Pengaturan", "Ganti logo aplikasi", "");
    return { success: true, url: dataUrl };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getDatabaseInfo(token) {
  try {
    checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    let counts = {};
    [CONFIG.SHEETS.ENTRANCE_SALES, CONFIG.SHEETS.RIDE_SALES, CONFIG.SHEETS.RESTO_SALES,
     CONFIG.SHEETS.ENTRANCE_PAYMENT, CONFIG.SHEETS.RIDE_PAYMENT, CONFIG.SHEETS.RESTO_PAYMENT,
     CONFIG.SHEETS.PRODUCTS].forEach(name => {
      let sh = ss.getSheetByName(name);
      counts[name] = sh ? Math.max(0, sh.getLastRow() - 1) : "N/A";
    });
    let executingAs = "-";
    try { executingAs = Session.getEffectiveUser().getEmail() || "-"; } catch (e) {}
    return { success: true, url: ss.getUrl(), name: ss.getName(), id: ss.getId(), counts: counts, executingAs: executingAs };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// UTILITAS DARURAT: kalau web app ternyata membaca/menulis ke spreadsheet
// yang BEDA dari yang dicek manual oleh Owner (gejala umum: "data masuk ke
// spreadsheet tapi tidak muncul di web"), Owner bisa menyambungkan paksa
// aplikasi ke spreadsheet database yang benar lewat sini - cukup tempel
// link atau ID spreadsheet-nya. Tidak mengubah/menghapus data apapun.
function setDatabaseId(idOrUrl, token) {
  try {
    checkPermission_(token, [ROLE.OWNER]);
    if (!idOrUrl) return { success: false, message: "ID atau link spreadsheet tidak boleh kosong." };
    let id = String(idOrUrl).trim();
    let match = id.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) id = match[1];

    let ss;
    try { ss = SpreadsheetApp.openById(id); } catch (e) {
      return { success: false, message: "Spreadsheet tidak ditemukan atau tidak bisa diakses. Pastikan ID/link benar dan Anda punya akses." };
    }
    // Validasi ringan: pastikan ini memang database aplikasi ini (ada sheet Products)
    if (!ss.getSheetByName(CONFIG.SHEETS.PRODUCTS)) {
      return { success: false, message: "Spreadsheet ini sepertinya bukan database aplikasi (sheet 'Products' tidak ditemukan)." };
    }
    PropertiesService.getScriptProperties().setProperty(CONFIG.DB_PROPERTY_KEY, ss.getId());
    logActivity_("system", "SWITCH_DATABASE", "Pengaturan", "Database disambungkan ulang ke: " + ss.getUrl(), "");
    return { success: true, message: "Berhasil disambungkan ke database: " + ss.getName(), url: ss.getUrl(), name: ss.getName() };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function triggerManualBackup(token) {
  try {
    checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    let backupFolder = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDERS.BACKUP).next();
    let copyFile = ss.getDriveFile().makeCopy("Backup_" + CONFIG.APP_NAME + "_" +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss"), backupFolder);
    ss.getSheetByName(CONFIG.SHEETS.BACKUP_LOG).appendRow([genId_("BAK"), new Date(), copyFile.getName(), "SUCCESS", copyFile.getUrl()]);
    return { success: true, message: "Backup berhasil dibuat.", url: copyFile.getUrl() };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 14. DIAGNOSTIC & PERBAIKAN DATABASE (dipanggil dari menu Pengaturan)
// =====================================================================
// Sheet referensi/master yang aman untuk DEDUP OTOMATIS kalau ditemukan
// baris dengan ID yang sama persis (data acuan, bukan transaksi keuangan/
// penjualan - jadi menghapus baris duplikat di sini tidak berisiko
// menghilangkan riwayat transaksi). Baris duplikat yang PALING BAWAH yang
// dihapus, baris pertama (paling atas) yang dipertahankan.
const DEDUPABLE_MASTER_SHEETS_ = [
  { sheet: CONFIG.SHEETS.PRODUCTS, idCol: "ProductCode", label: "Produk" },
  { sheet: CONFIG.SHEETS.CATEGORIES, idCol: "CategoryID", label: "Kategori" },
  { sheet: CONFIG.SHEETS.PAYMENT_METHODS, idCol: "MethodID", label: "Metode Pembayaran" },
  { sheet: CONFIG.SHEETS.PRODUCT_TYPES, idCol: "TypeID", label: "Tipe Produk" },
  { sheet: CONFIG.SHEETS.APPLICABLE_UNITS, idCol: "UnitID", label: "Unit Berlaku (Applicable Unit)" },
  { sheet: CONFIG.SHEETS.UNITS, idCol: "UnitID", label: "Satuan" },
  { sheet: CONFIG.SHEETS.WAREHOUSES, idCol: "WarehouseID", label: "Gudang" },
  { sheet: CONFIG.SHEETS.ROLES, idCol: "RoleID", label: "Roles" },
  { sheet: CONFIG.SHEETS.FINANCE_COA, idCol: "COACode", label: "Kode Akun (COA)" },
  { sheet: CONFIG.SHEETS.FINANCE_PAYMENT_STATUS, idCol: "StatusID", label: "Status Pembayaran" },
  { sheet: CONFIG.SHEETS.FINANCE_POS_PENGGUNAAN, idCol: "PosID", label: "Pos Penggunaan" }
];

// Sheet yang DICEK duplikat ID-nya tapi TIDAK dihapus otomatis (data
// transaksional/akun - kalau memang dobel, harus ditinjau manual dulu
// supaya tidak salah hapus baris yang justru valid).
const DEDUP_REVIEW_ONLY_SHEETS_ = [
  { sheet: CONFIG.SHEETS.USERS, idCol: "Username", label: "User/Akun Login" },
  { sheet: CONFIG.SHEETS.ENTRANCE_SALES, idCol: "SaleID", label: "Penjualan Loket Masuk" },
  { sheet: CONFIG.SHEETS.RIDE_SALES, idCol: "SaleID", label: "Penjualan Loket Wahana" },
  { sheet: CONFIG.SHEETS.RESTO_SALES, idCol: "SaleID", label: "Penjualan Resto" }
];

// Kolom yang HARUS berisi angka polos, per sheet. Dipakai untuk mendeteksi
// kasus nyata yang pernah terjadi: kolom "Quantity" di Warehouse_Stock
// tidak sengaja ikut ter-format sebagai Tanggal (biasanya gara-gara
// fill-handle/format painter menular dari kolom tanggal di sebelahnya).
// Akibatnya Apps Script membaca isi kolom itu sebagai objek Date, bukan
// angka - stok jadi salah baca dan beberapa halaman bisa gagal total
// dengan pesan "Invalid time value". Format TIDAK mengubah nilai asli yang
// tersimpan di sel, jadi memperbaikinya cukup dengan mengembalikan format
// kolom ke angka biasa - nilainya otomatis terbaca benar lagi sesudahnya.
const NUMERIC_COLUMNS_BY_SHEET_ = {
  [CONFIG.SHEETS.WAREHOUSE_STOCK]: ["Quantity", "MinStock"],
  [CONFIG.SHEETS.PRODUCTS]: ["BuyPrice", "SellPrice"],
  [CONFIG.SHEETS.STOCK_OPNAME]: ["StokAwal", "StokFisik", "Selisih", "SellPrice", "NilaiSelisih"],
  [CONFIG.SHEETS.STOCK_MUTATIONS]: ["Qty"],
  [CONFIG.SHEETS.KOLAM_STOCK]: ["Qty"],
  [CONFIG.SHEETS.TEKNISI_STOCK]: ["Qty"],
  [CONFIG.SHEETS.CLEANING_STOCK]: ["Qty"],
  [CONFIG.SHEETS.DAPUR_STOCK]: ["Qty"],
  [CONFIG.SHEETS.STOCK_CARD]: ["QtyIn", "QtyOut", "BalanceBefore", "BalanceAfter"],
  [CONFIG.SHEETS.ENTRANCE_SALES]: ["Qty", "TotalSerial", "StockStart", "StockIn", "StockOut", "StockEnd", "UnitPrice", "TotalPrice"],
  [CONFIG.SHEETS.RIDE_SALES]: ["Qty", "TotalSerial", "StockStart", "StockIn", "StockOut", "StockEnd", "UnitPrice", "TotalPrice"],
  [CONFIG.SHEETS.RESTO_SALES]: ["Qty", "StockStart", "StockIn", "StockOut", "StockEnd", "UnitPrice", "TotalPrice"],
  [CONFIG.SHEETS.ENTRANCE_PAYMENT]: ["Amount"],
  [CONFIG.SHEETS.RIDE_PAYMENT]: ["Amount"],
  [CONFIG.SHEETS.RESTO_PAYMENT]: ["Amount"],
  [CONFIG.SHEETS.FINANCE_BALANCE]: ["SaldoAwalBrangkas", "SaldoAwalRekening", "SaldoTHR", "SaldoAkhirBrangkasAktual", "SaldoAkhirRekeningAktual", "SaldoAkhirBrangkas", "SaldoAkhirRekening"],
  [CONFIG.SHEETS.FINANCE_EXPENSES]: ["Amount", "UnitPrice", "Qty"]
};

// Format Tanggal/Waktu bawaan Google Sheets biasanya mengandung salah satu
// pola berikut - dipakai untuk mendeteksi kolom angka yang "kena tular"
// format tanggal.
function looksLikeDateFormat_(fmt) {
  if (!fmt) return false;
  return /[ymdYMD]/.test(fmt) && (fmt.indexOf("h") !== -1 || fmt.indexOf("m/d") !== -1 || fmt.indexOf("d/m") !== -1 || fmt.indexOf("yyyy") !== -1 || fmt.indexOf("yy") !== -1);
}

function getSheetHeaders_(sheet) {
  let lastCol = sheet.getLastColumn();
  return lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim()) : [];
}

// Cari kolom duplikat (berdasarkan idCol) di satu sheet. Mengembalikan
// array of { value, rowNumbers } untuk tiap nilai ID yang muncul >1 kali.
// CATATAN PERFORMA: versi ini masih dipertahankan untuk pemakaian lain,
// tapi scan diagnostic (scanDatabaseIssues_) TIDAK memakai fungsi ini lagi
// karena akan memicu getRange()/getValues() ulang per sheet. Diagnostic
// memakai findDuplicateRowsFromValues_() di bawah yang bekerja dari data
// yang sudah terlanjur dibaca sekali lewat buildSheetScanCache_().
function findDuplicateRows_(sheet, idCol) {
  const headers = getSheetHeaders_(sheet);
  const idIdx = headers.indexOf(idCol);
  if (idIdx === -1 || sheet.getLastRow() < 3) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const map = {};
  data.forEach((row, i) => {
    const val = String(row[idIdx]).trim();
    if (!val) return;
    if (!map[val]) map[val] = [];
    map[val].push(i + 2); // nomor baris asli di sheet
  });
  return Object.keys(map).filter(v => map[v].length > 1).map(v => ({ value: v, rowNumbers: map[v] }));
}

// Versi findDuplicateRows_ yang bekerja dari array `values` yang SUDAH
// dibaca sebelumnya (mis. lewat range.getDataRange().getValues() satu kali
// per sheet), supaya tidak menambah panggilan getRange()/getValues() baru
// ke Google Sheets. `values[0]` dianggap baris header (baris 1 di sheet).
function findDuplicateRowsFromValues_(values, headers, idCol) {
  const idIdx = headers.indexOf(idCol);
  if (idIdx === -1 || values.length < 3) return [];
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const val = String(values[i][idIdx]).trim();
    if (!val) continue;
    if (!map[val]) map[val] = [];
    map[val].push(i + 1); // nomor baris asli di sheet (values index 0 = baris 1)
  }
  return Object.keys(map).filter(v => map[v].length > 1).map(v => ({ value: v, rowNumbers: map[v] }));
}

// =====================================================================
// OPTIMASI PERFORMA DIAGNOSTIC (v2 - lihat juga catatan di bawah tentang
// kenapa v1 masih bisa macet/timeout pada sheet transaksi yang sudah
// sangat besar, mis. setelah bertahun-tahun operasional):
//
// - Nilai (values) tetap dibaca 1x per sheet lewat getDataRange(), TAPI
//   HANYA untuk sheet yang jumlah barisnya masih di bawah
//   SCAN_DEEP_CHECK_MAX_ROWS_. Sheet yang lebih besar dari itu (biasanya
//   sheet transaksi kasir yang datanya menumpuk terus tiap hari) HANYA
//   dibaca baris headernya saja - duplikat & format TIDAK dicek otomatis
//   untuk sheet itu (dilaporkan sebagai info, bukan dianggap error),
//   supaya Diagnostic tidak pernah macet/timeout gara-gara 1 sheet raksasa.
// - getNumberFormats() JAUH lebih berat/lambat daripada getValues() untuk
//   jumlah sel yang sama (ini penyebab paling umum loading "menggantung"
//   puluhan menit bahkan sampai kena limit waktu eksekusi Apps Script).
//   Karena itu, format HANYA dicek pada NUMERIC_FORMAT_SAMPLE_ROWS_ baris
//   TERBARU (bukan seluruh riwayat), cukup untuk mendeteksi masalah karena
//   kasus nyatanya selalu mempengaruhi satu kolom penuh sekaligus (bukan
//   baris acak di tengah data lama). Kalau ditemukan & diperbaiki lewat
//   "Perbaiki Database", perbaikan tetap diterapkan ke SEMUA baris, bukan
//   cuma yang di-sample - lihat repairDatabaseAuto().
const SCAN_DEEP_CHECK_MAX_ROWS_ = 15000;
const NUMERIC_FORMAT_SAMPLE_ROWS_ = 3000;

function buildSheetScanCache_(ss) {
  const cache = {};
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const dataRows = Math.max(0, lastRow - 1);
    const entry = {
      sheet: sheet, values: [], headers: [], numberFormats: null,
      numberFormatsStartRow: 2, rowCount: dataRows, tooLargeForDeepCheck: false
    };
    if (lastRow > 0 && lastCol > 0) {
      if (dataRows > SCAN_DEEP_CHECK_MAX_ROWS_) {
        // Sheet terlalu besar untuk dibaca penuh dengan aman - cukup ambil
        // header (1 baris, murah) supaya inventory & cek kolom tetap jalan.
        entry.headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
        entry.tooLargeForDeepCheck = true;
      } else {
        const range = sheet.getDataRange();
        entry.values = range.getValues();
        entry.headers = entry.values.length ? entry.values[0].map(h => String(h).trim()) : [];
        if (NUMERIC_COLUMNS_BY_SHEET_[name] && dataRows > 0) {
          const rowsToCheck = Math.min(dataRows, NUMERIC_FORMAT_SAMPLE_ROWS_);
          const startRow = lastRow - rowsToCheck + 1; // ambil N baris TERBARU
          entry.numberFormats = sheet.getRange(startRow, 1, rowsToCheck, lastCol).getNumberFormats();
          entry.numberFormatsStartRow = startRow;
        }
      }
    }
    cache[name] = entry;
  });
  return cache;
}

// Scan lengkap kondisi database TANPA mengubah apapun (read-only). Dipakai
// baik untuk menampilkan hasil diagnostic maupun sebagai langkah pertama
// repairDatabaseAuto() sebelum benar-benar memperbaiki.
// Daftar LENGKAP semua sheet yang seharusnya ada (sumbernya SAMA PERSIS
// dengan skema yang dipakai untuk membuat/memperbaiki database -
// getFullSchemaMap_(), yang keynya diambil dari CONFIG.SHEETS di backend)
// berikut status OK/Missing & kelengkapan headernya. Ditampilkan penuh di
// Diagnostic Database supaya Owner bisa lihat SEMUA sheet database, bukan
// cuma yang bermasalah.
// `cache` opsional: kalau dilewatkan (hasil buildSheetScanCache_() yang
// sudah dibangun sekali di caller), fungsi ini TIDAK melakukan panggilan
// API apapun - murni membaca dari memori. Ini dipakai supaya
// runDatabaseDiagnostic()/repairDatabaseAuto() bisa berbagi 1 cache untuk
// getSheetInventory_() + scanDatabaseIssues_() sekaligus dalam 1 request.
function getSheetInventory_(ss, cache) {
  cache = cache || buildSheetScanCache_(ss);
  const schema = getFullSchemaMap_();
  return Object.keys(schema).map(sheetName => {
    const entry = cache[sheetName];
    if (!entry) {
      return { name: sheetName, status: "Missing", headerStatus: "-", rowCount: 0 };
    }
    const headers = entry.headers;
    const expected = schema[sheetName];
    let headerStatus = "OK";
    if (headers.length === 0) headerStatus = "Tidak ada header";
    else {
      const missing = expected.filter(h => headers.indexOf(h) === -1);
      if (missing.length) headerStatus = missing.length + " kolom kurang";
    }
    return { name: sheetName, status: "OK", headerStatus: headerStatus, rowCount: entry.rowCount };
  });
}

// `cache` opsional: kalau dilewatkan, dipakai apa adanya (dibangun sekali
// oleh caller lewat buildSheetScanCache_()) supaya seluruh scan di bawah
// ini TIDAK memicu getRange()/getValues()/getNumberFormats() berulang ke
// Google Sheets - hanya operasi di memori. Kalau tidak dilewatkan, cache
// dibangun sendiri (dipakai kalau fungsi ini dipanggil berdiri sendiri).
function scanDatabaseIssues_(ss, cache) {
  cache = cache || buildSheetScanCache_(ss);
  const issues = [];
  const schema = getFullSchemaMap_();

  // 1) Cek file spreadsheet dobel di Drive dengan nama database yang sama -
  //    penyebab umum "data sudah diinput tapi tidak muncul di web" karena
  //    aplikasi & user membuka spreadsheet yang berbeda.
  try {
    const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
    const found = [];
    while (files.hasNext()) {
      const f = files.next();
      found.push({ id: f.getId(), url: f.getUrl(), lastUpdated: f.getLastUpdated(), connected: f.getId() === ss.getId() });
    }
    if (found.length > 1) {
      issues.push({
        id: "dup_file",
        category: "duplicate_file",
        severity: "high",
        autoFixable: false,
        sheet: null,
        title: "Ditemukan " + found.length + " spreadsheet database dengan nama yang sama",
        description: "Ada beberapa file Spreadsheet bernama \"" + CONFIG.SPREADSHEET_NAME + "\" di Google Drive. Aplikasi saat ini tersambung ke salah satu di antaranya - kalau input data dilakukan di file yang lain, data tidak akan muncul di aplikasi. Tinjau daftar file berikut dan sambungkan ulang lewat menu \"Sambungkan ke Spreadsheet Ini\" kalau perlu.",
        detail: found
      });
    }
  } catch (e) {}

  // 2) Sheet yang seharusnya ada tapi tidak ditemukan (tabel/sheet kurang)
  const existingNames = Object.keys(cache);
  Object.keys(schema).forEach(sheetName => {
    if (existingNames.indexOf(sheetName) === -1) {
      issues.push({
        id: "missing_sheet:" + sheetName,
        category: "missing_sheet",
        severity: "high",
        autoFixable: true,
        sheet: sheetName,
        title: "Sheet \"" + sheetName + "\" tidak ditemukan",
        description: "Sheet ini seharusnya ada di database tapi hilang/belum pernah dibuat. Modul yang menggunakan sheet ini kemungkinan error.",
        detail: { headers: schema[sheetName] }
      });
    }
  });

  // 3) Untuk sheet yang ADA: cek header kosong / kolom yang hilang
  existingNames.forEach(sheetName => {
    if (!schema[sheetName]) return; // sheet di luar skema dikenal, lewati
    const expected = schema[sheetName];
    const headers = cache[sheetName].headers;

    if (headers.length === 0) {
      issues.push({
        id: "no_header:" + sheetName,
        category: "no_header",
        severity: "high",
        autoFixable: true,
        sheet: sheetName,
        title: "Sheet \"" + sheetName + "\" tidak punya baris judul kolom",
        description: "Baris pertama (header) sheet ini kosong, akan ditulis ulang sesuai skema standar.",
        detail: { headers: expected }
      });
    } else {
      const missing = expected.filter(h => headers.indexOf(h) === -1);
      if (missing.length) {
        issues.push({
          id: "missing_header:" + sheetName,
          category: "missing_header",
          severity: "medium",
          autoFixable: true,
          sheet: sheetName,
          title: "Sheet \"" + sheetName + "\" kekurangan " + missing.length + " kolom",
          description: "Kolom berikut belum ada di sheet ini dan akan ditambahkan di akhir tanpa mengubah data yang sudah ada: " + missing.join(", "),
          detail: { missingColumns: missing }
        });
      }
    }
  });

  // 4) Sheet tak dikenal (kemungkinan hasil duplikasi tidak sengaja, mis.
  //    "Products (1)", "Copy of Products", dsb.)
  existingNames.forEach(sheetName => {
    if (schema[sheetName]) return;
    const looksLikeCopy = /\(\d+\)$|copy of|salinan/i.test(sheetName);
    if (looksLikeCopy) {
      issues.push({
        id: "stray_sheet:" + sheetName,
        category: "stray_sheet",
        severity: "low",
        autoFixable: false,
        sheet: sheetName,
        title: "Sheet asing \"" + sheetName + "\" terdeteksi",
        description: "Nama sheet ini terlihat seperti hasil duplikasi tidak sengaja (copy-paste sheet). Periksa manual isinya sebelum dihapus - aplikasi tidak menggunakan sheet ini.",
        detail: null
      });
    }
  });

  // 5) Baris duplikat pada data master (aman didedup otomatis)
  DEDUPABLE_MASTER_SHEETS_.forEach(cfg => {
    const entry = cache[cfg.sheet];
    if (!entry || entry.tooLargeForDeepCheck) return;
    const dups = findDuplicateRowsFromValues_(entry.values, entry.headers, cfg.idCol);
    if (dups.length) {
      const totalExtraRows = dups.reduce((sum, d) => sum + (d.rowNumbers.length - 1), 0);
      issues.push({
        id: "dup_row:" + cfg.sheet,
        category: "duplicate_row",
        severity: "medium",
        autoFixable: true,
        sheet: cfg.sheet,
        title: "Data \"" + cfg.label + "\" punya " + dups.length + " " + cfg.idCol + " dobel",
        description: "Baris pertama untuk tiap " + cfg.idCol + " akan dipertahankan, " + totalExtraRows + " baris duplikat lainnya akan dihapus.",
        detail: { idCol: cfg.idCol, duplicates: dups }
      });
    }
  });

  // 6) Baris duplikat pada data transaksional/akun (hanya dilaporkan) -
  //    sheet yang terlalu besar (lihat tooLargeForDeepCheck) dilewati di
  //    sini dan dilaporkan lewat issue "sheet_too_large_skipped" di bawah.
  DEDUP_REVIEW_ONLY_SHEETS_.forEach(cfg => {
    const entry = cache[cfg.sheet];
    if (!entry || entry.tooLargeForDeepCheck) return;
    const dups = findDuplicateRowsFromValues_(entry.values, entry.headers, cfg.idCol);
    if (dups.length) {
      issues.push({
        id: "dup_row_review:" + cfg.sheet,
        category: "duplicate_row_review",
        severity: "high",
        autoFixable: false,
        sheet: cfg.sheet,
        title: "Data \"" + cfg.label + "\" punya " + dups.length + " " + cfg.idCol + " dobel (perlu tinjauan manual)",
        description: "Ditemukan " + cfg.idCol + " yang sama di beberapa baris. Karena ini data transaksi/akun, baris TIDAK dihapus otomatis - periksa nomor baris berikut langsung di spreadsheet sebelum memutuskan mana yang dihapus.",
        detail: { idCol: cfg.idCol, duplicates: dups }
      });
    }
  });

  // 6b) Info: sheet yang dilewati dari cek duplikat/format karena datanya
  //     sudah sangat besar (lihat SCAN_DEEP_CHECK_MAX_ROWS_). Ini BUKAN
  //     error - hanya transparansi ke Owner supaya tahu bagian mana yang
  //     otomatis tidak dicek dalam, dan jumlah barisnya tetap valid/akurat.
  Object.keys(cache).forEach(sheetName => {
    const entry = cache[sheetName];
    if (!entry.tooLargeForDeepCheck) return;
    issues.push({
      id: "size_skipped:" + sheetName,
      category: "sheet_too_large_skipped",
      severity: "low",
      autoFixable: false,
      sheet: sheetName,
      title: "Sheet \"" + sheetName + "\" (" + entry.rowCount + " baris) dilewati dari cek mendalam",
      description: "Sheet ini punya lebih dari " + SCAN_DEEP_CHECK_MAX_ROWS_ + " baris data, jadi demi menghindari Diagnostic macet/timeout, cek duplikat & format angka TIDAK dijalankan otomatis untuk sheet ini (header & jumlah kolom tetap dicek seperti biasa). Kalau perlu, periksa manual langsung di spreadsheet.",
      detail: { rowCount: entry.rowCount, threshold: SCAN_DEEP_CHECK_MAX_ROWS_ }
    });
  });

  // 7) Data master 8 gudang wajib (Utama/Loket Masuk/Loket Wahana/Resto/
  //    Kolam Renang/Teknisi/Dapur/Cleaning Service) - sumbernya SAMA PERSIS
  //    dengan yang dipakai migrasi otomatis ensureWarehousesSeed_(), jadi
  //    tidak ada daftar ganda yang bisa saling beda.
  try {
    const whEntry = cache[CONFIG.SHEETS.WAREHOUSES];
    if (whEntry) {
      const whData = whEntry.values;
      const idSet = {};
      for (let i = 1; i < whData.length; i++) idSet[whData[i][0]] = true;
      const missingWarehouses = CANONICAL_WAREHOUSES_.filter(w => !idSet[w[0]]);
      if (missingWarehouses.length) {
        issues.push({
          id: "missing_warehouse_seed",
          category: "missing_warehouse_seed",
          severity: "high",
          autoFixable: true,
          sheet: CONFIG.SHEETS.WAREHOUSES,
          title: missingWarehouses.length + " data gudang wajib belum ada di sheet Warehouses",
          description: "Gudang berikut seharusnya selalu ada (dipakai modul Warehouse & alur stok outlet) tapi belum tercatat: " + missingWarehouses.map(w => w[1] + " (" + w[0] + ")").join(", ") + ".",
          detail: { missing: missingWarehouses }
        });
      }
    }
  } catch (e) {}

  // 8) Kolom angka yang tidak sengaja ke-format sebagai Tanggal/Waktu -
  //    penyebab nyata dari error "Invalid time value" & stok/angka yang
  //    salah baca. Aman diperbaiki otomatis karena hanya mengubah FORMAT
  //    tampilan kolom, nilai asli yang tersimpan di sel tidak disentuh.
  Object.keys(NUMERIC_COLUMNS_BY_SHEET_).forEach(sheetName => {
    const entry = cache[sheetName];
    if (!entry || entry.tooLargeForDeepCheck || entry.values.length < 2 || !entry.numberFormats) return;
    const headers = entry.headers;
    const sampled = entry.numberFormats.length < entry.rowCount;
    NUMERIC_COLUMNS_BY_SHEET_[sheetName].forEach(colName => {
      const colIdx = headers.indexOf(colName);
      if (colIdx === -1) return;
      // formats sudah ada di memori (dibaca sekali, hanya untuk N baris
      // TERBARU kalau sheet-nya besar - lihat NUMERIC_FORMAT_SAMPLE_ROWS_).
      const badRowCount = entry.numberFormats.filter(r => looksLikeDateFormat_(r[colIdx])).length;
      if (badRowCount > 0) {
        const sampleNote = sampled
          ? " (dicek dari " + entry.numberFormats.length + " baris terbaru saja dari total " + entry.rowCount + " baris, karena sheet ini besar - kalau diperbaiki, SEMUA baris tetap ikut dirapikan)"
          : "";
        issues.push({
          id: "wrong_format:" + sheetName + ":" + colName,
          category: "wrong_number_format",
          severity: "high",
          autoFixable: true,
          sheet: sheetName,
          title: "Kolom \"" + colName + "\" di sheet \"" + sheetName + "\" ter-format sebagai Tanggal (" + badRowCount + " sel" + (sampled ? " dari sample" : "") + ")",
          description: "Kolom ini seharusnya berisi angka biasa tapi " + badRowCount + " selnya ter-format sebagai Tanggal/Waktu, sehingga nilainya salah terbaca oleh aplikasi (bisa menyebabkan error \"Invalid time value\" atau angka yang tampil salah). Memperbaikinya hanya mengubah format tampilan kembali ke angka - nilai yang sudah tersimpan TIDAK berubah." + sampleNote,
          detail: { column: colName, colIndex1: colIdx + 1, badRowCount: badRowCount }
        });
      }
    });
  });

  return issues;
}

function runDatabaseDiagnostic(token) {
  try {
    checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    // Baca semua sheet SEKALI (lihat buildSheetScanCache_) dan pakai cache
    // yang sama untuk scan issue + inventory, supaya tidak dobel baca.
    const cache = buildSheetScanCache_(ss);
    const issues = scanDatabaseIssues_(ss, cache);
    const inventory = getSheetInventory_(ss, cache);
    const autoFixableCount = issues.filter(i => i.autoFixable).length;
    return {
      success: true,
      checkedAt: new Date(),
      dbName: ss.getName(),
      dbUrl: ss.getUrl(),
      issues: issues,
      inventory: inventory,
      summary: {
        total: issues.length,
        autoFixable: autoFixableCount,
        needsReview: issues.length - autoFixableCount
      }
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Perbaiki semua masalah yang aman diperbaiki otomatis: buat sheet yang
// hilang, tulis ulang header yang kosong, tambahkan kolom yang kurang, dan
// hapus baris duplikat HANYA di sheet data master (lihat
// DEDUPABLE_MASTER_SHEETS_). Sheet asing & duplikat data transaksi/akun/
// file spreadsheet TIDAK disentuh - dilaporkan lagi di remainingIssues
// supaya Owner tahu apa yang masih perlu ditinjau manual.
function repairDatabaseAuto(token) {
  try {
    checkPermission_(token, [ROLE.OWNER]);
    const ss = getDatabase();
    const schema = getFullSchemaMap_();
    // Scan awal pakai 1 cache bersama (baca tiap sheet cukup sekali).
    const cache = buildSheetScanCache_(ss);
    const issues = scanDatabaseIssues_(ss, cache);
    const fixed = [];

    issues.filter(i => i.autoFixable).forEach(issue => {
      try {
        if (issue.category === "missing_sheet") {
          const headers = schema[issue.sheet];
          const sheet = ss.insertSheet(issue.sheet);
          sheet.getRange(1, 1, 1, headers.length).setValues([headers])
            .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
          sheet.setFrozenRows(1);
          fixed.push("Sheet \"" + issue.sheet + "\" dibuat ulang.");
        } else if (issue.category === "no_header") {
          const headers = schema[issue.sheet];
          const sheet = ss.getSheetByName(issue.sheet);
          sheet.getRange(1, 1, 1, headers.length).setValues([headers])
            .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
          sheet.setFrozenRows(1);
          fixed.push("Header sheet \"" + issue.sheet + "\" ditulis ulang.");
        } else if (issue.category === "missing_header") {
          const sheet = ss.getSheetByName(issue.sheet);
          const currentHeaders = getSheetHeaders_(sheet);
          const missing = issue.detail.missingColumns;
          sheet.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing])
            .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
          fixed.push("Sheet \"" + issue.sheet + "\": kolom " + missing.join(", ") + " ditambahkan.");
        } else if (issue.category === "duplicate_row") {
          const sheet = ss.getSheetByName(issue.sheet);
          // Kumpulkan semua nomor baris duplikat (kecuali baris pertama tiap
          // grup) dari SEMUA grup, urutkan dari bawah ke atas supaya nomor
          // baris yang belum dihapus tidak bergeser saat proses hapus.
          let rowsToDelete = [];
          issue.detail.duplicates.forEach(d => {
            rowsToDelete = rowsToDelete.concat(d.rowNumbers.slice(1));
          });
          rowsToDelete.sort((a, b) => b - a);
          rowsToDelete.forEach(r => sheet.deleteRow(r));
          fixed.push("Sheet \"" + issue.sheet + "\": " + rowsToDelete.length + " baris " + issue.detail.idCol + " duplikat dihapus (baris pertama tiap ID dipertahankan).");
        } else if (issue.category === "wrong_number_format") {
          const sheet = ss.getSheetByName(issue.sheet);
          const numRows = sheet.getLastRow() - 1;
          sheet.getRange(2, issue.detail.colIndex1, numRows, 1).setNumberFormat("#,##0");
          fixed.push("Sheet \"" + issue.sheet + "\": format kolom \"" + issue.detail.column + "\" (" + issue.detail.badRowCount + " sel) dikembalikan ke angka biasa.");
        } else if (issue.category === "missing_warehouse_seed") {
          ensureWarehousesSeed_(ss);
          fixed.push("Sheet \"Warehouses\": " + issue.detail.missing.map(w => w[1]).join(", ") + " ditambahkan.");
        }
      } catch (e) {
        fixed.push("GAGAL memperbaiki \"" + issue.title + "\": " + e.message);
      }
    });

    logActivity_("system", "DB_AUTO_REPAIR", "Pengaturan", fixed.length + " perbaikan otomatis dijalankan: " + fixed.join(" | "), "");

    // Scan ulang supaya Owner tahu status TERKINI setelah perbaikan. Cache
    // HARUS dibangun ulang di sini (bukan pakai `cache` dari atas) karena
    // datanya sudah berubah akibat proses perbaikan barusan (baris
    // dihapus/header ditambah/format diubah) - tapi tetap 1 cache dipakai
    // bersama untuk kedua fungsi supaya tidak dobel baca.
    const freshCache = buildSheetScanCache_(ss);
    const remainingIssues = scanDatabaseIssues_(ss, freshCache);
    const inventory = getSheetInventory_(ss, freshCache);
    return {
      success: true,
      fixedCount: fixed.length,
      fixedDetails: fixed,
      remainingIssues: remainingIssues,
      inventory: inventory
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =====================================================================
// 15. DIAGNOSTIK MANUAL - jalankan manual dari editor kalau perlu troubleshoot
// =====================================================================
function debugCheckUsersSheet() {
  const ss = getDatabase();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  Logger.log("Spreadsheet ID: " + ss.getId());
  Logger.log("Spreadsheet URL: " + ss.getUrl());
  Logger.log("Header Users: " + JSON.stringify(data[0]));
  for (let i = 1; i < data.length; i++) {
    Logger.log("Row " + i + " -> Username: [" + data[i][1] + "] | Role: " + data[i][3] + " | Status: " + data[i][data[i].length - 2]);
  }
}