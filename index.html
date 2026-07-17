// Menampilkan halaman web utama Senn Store
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Senn Store - Order Otomatis')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Ambil data dari Google Sheets
function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

// Fitur Login Customer & Admin
function loginUser(username, password) {
  if (username === 'admin' && password === 'admin123') {
    return { success: true, role: 'admin', name: 'Administrator', username: 'admin' };
  }
  
  const customers = getSheetData('Customer');
  const user = customers.find(c => String(c.username).trim() === String(username).trim() && String(c.password).trim() === String(password).trim());
  if (user) {
    return { success: true, role: 'customer', name: user.nama, username: user.username, no_hp: user.no_hp };
  }
  return { success: false, message: 'Username atau Password salah!' };
}

// Fitur Tambah Produk (Admin)
function addProduct(product) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produk');
  const nextId = sheet.getLastRow();
  sheet.appendRow([nextId, product.nama, product.harga, product.deskripsi, product.gambar_url, product.stok]);
  return { success: true };
}

// Fitur Hapus Produk (Admin)
function deleteProduct(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produk');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// Fitur Checkout Otomatis, Kurangi Stok, dan Kirim Notifikasi Email
function checkoutOrder(username, cartItems, totalHarga) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pesananSheet = ss.getSheetByName('Pesanan');
  const produkSheet = ss.getSheetByName('Produk');
  
  const orderId = 'SENN-' + new Date().getTime();
  const tanggal = new Date();
  const detail = JSON.stringify(cartItems);
  
  // 1. Catat ke sheet Pesanan
  pesananSheet.appendRow([orderId, tanggal, username, detail, totalHarga, 'Pending']);
  
  // 2. Potong Stok di sheet Produk otomatis
  const produkData = produkSheet.getDataRange().getValues();
  cartItems.forEach(item => {
    for (let i = 1; i < produkData.length; i++) {
      if (produkData[i][0] == item.id) {
        const stokLama = Number(produkData[i][5]);
        const stokBaru = Math.max(0, stokLama - item.qty);
        produkSheet.getRange(i + 1, 6).setValue(stokBaru); // Update kolom stok (kolom F / 6)
      }
    }
  });

  // 3. Kirim Email Notifikasi Otomatis ke Email Admin
  try {
    const adminEmail = Session.getEffectiveUser().getEmail(); // Otomatis terkirim ke email pemilik sheet/GAS ini
    const tanggalFormat = Utilities.formatDate(tanggal, "GMT+7", "dd-MM-yyyy HH:mm");
    
    let detailEmailTeks = "";
    cartItems.forEach(item => {
      detailEmailTeks += `• ${item.nama} (${item.qty}x) - Rp ${(item.harga * item.qty).toLocaleString('id-ID')}\n`;
    });

    const subjek = `🚨 [SENN STORE] ORDER BARU - ${orderId}`;
    const isiEmail = `Halo Owner Senn Store,\n\n` +
                     `Ada pesanan baru masuk dari website!\n\n` +
                     `ID Pesanan : ${orderId}\n` +
                     `Waktu      : ${tanggalFormat}\n` +
                     `Pelanggan  : ${username}\n` +
                     `-----------------------------------\n` +
                     `Daftar Belanja:\n${detailEmailTeks}\n` +
                     `-----------------------------------\n` +
                     `Total Bayar: Rp ${Number(totalHarga).toLocaleString('id-ID')}\n\n` +
                     `Sistem telah memotong stok produk secara otomatis di Google Sheets Anda. Silakan hubungi customer jika belum melakukan pembayaran via WhatsApp.`;
                     
    MailApp.sendEmail(adminEmail, subjek, isiEmail);
  } catch (error) {
    Logger.log("Gagal mengirim email: " + error.toString());
  }
  
  return { success: true, orderId: orderId };
}

// Membuat struktur database awal agar tidak perlu repot isi manual
function buatStrukturSheetOtomatis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const struktur = [
    {
      namaSheet: "Produk",
      kolom: ["id", "nama", "harga", "deskripsi", "gambar_url", "stok"],
      contohData: [
        [1, "Jersey Senn Store Spec", 150000, "Jersey olahraga premium senn store bahan dry-fit nyaman dipakai sehari-hari.", "https://images.unsplash.com/photo-1541167760496-1628856ab772?q=80&w=500", 99],
        [2, "Senn Hoodie Black", 299000, "Hoodie tebal katun premium dengan sablon awet anti pecah.", "https://images.unsplash.com/photo-1600891964599-f61ba0e24092?q=80&w=500", 50]
      ]
    },
    {
      namaSheet: "Customer",
      kolom: ["id", "username", "password", "nama", "no_hp"],
      contohData: [
        [1, "user1", "password123", "Budi Santoso", "08123456789"],
        [2, "user2", "pass456", "Siti Aminah", "08571234567"]
      ]
    },
    {
      namaSheet: "Pesanan",
      kolom: ["id_pesanan", "tanggal", "username", "detail_produk", "total_harga", "status"],
      contohData: []
    }
  ];

  struktur.forEach(item => {
    let sheet = ss.getSheetByName(item.namaSheet);
    if (!sheet) {
      sheet = ss.insertSheet(item.namaSheet);
    } else {
      sheet.clear();
    }
    sheet.getRange(1, 1, 1, item.kolom.length).setValues([item.kolom]);
    sheet.getRange(1, 1, 1, item.kolom.length)
         .setBackground("#0d6efd")
         .setFontColor("#FFFFFF")
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
         
    if (item.contohData.length > 0) {
      sheet.getRange(2, 1, item.contohData.length, item.kolom.length).setValues(item.contohData);
    }
    sheet.autoResizeColumns(1, item.kolom.length);
  });
  
  const sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1 && ss.getSheets().length > 3) {
    ss.deleteSheet(sheet1);
  }
}

// Fungsi Baru: Akan kita pasang untuk memeriksa sheet otomatis setiap 1 menit
function cekDanKirimEmailPesanan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pesanan');
  const lastRow = sheet.getLastRow();
  
  // Jika sheet kosong atau hanya ada header, hentikan proses
  if (lastRow <= 1) return;
  
  // 1. TULIS EMAIL GMAIL ANDA DI SINI
  const emailToko = "senhora311@gmail.com"; 
  
  // 2. Ambil data baris paling terakhir (pesanan paling baru)
  const orderId = sheet.getRange(lastRow, 1).getValue();
  const tanggal = sheet.getRange(lastRow, 2).getValue();
  const username = sheet.getRange(lastRow, 3).getValue();
  const detailProdukRaw = sheet.getRange(lastRow, 4).getValue();
  const totalHarga = sheet.getRange(lastRow, 5).getValue();
  const status = sheet.getRange(lastRow, 6).getValue();
  
  // Validasi: Cek apakah ID order ini sudah pernah dikirim emailnya atau belum
  const properti = PropertiesService.getScriptProperties();
  const lastSentId = properti.getProperty('LAST_SENT_ORDER_ID');
  if (lastSentId === String(orderId)) {
    return; // Jika ID sama dengan pesanan sebelumnya, hentikan (agar tidak spam email ganda)
  }
  
  // Format Tanggal
  const tanggalFormat = Utilities.formatDate(new Date(tanggal), "GMT+7", "dd-MM-yyyy HH:mm");
  
  // Urai Detail Produk Belanjaan
  let detailTeks = "";
  try {
    const items = JSON.parse(detailProdukRaw);
    items.forEach(item => {
      detailTeks += `• ${item.nama} (${item.qty}x) - Rp ${(item.harga * item.qty).toLocaleString('id-ID')}\n`;
    });
  } catch(err) {
    detailTeks = detailProdukRaw;
  }

  // Format Konten Isi Email Notifikasi
  const subjek = `🚨 [SENN STORE] MASUK ORDERAN BARU - ${orderId}`;
  const isiEmail = `Halo Owner Senn Store,\n\n` +
                   `Sistem mendeteksi ada transaksi baru yang masuk ke Google Sheets!\n\n` +
                   `ID Pesanan : ${orderId}\n` +
                   `Waktu      : ${tanggalFormat}\n` +
                   `Pelanggan  : ${username}\n` +
                   `-----------------------------------\n` +
                   `Daftar Belanja:\n${detailTeks}\n` +
                   `-----------------------------------\n` +
                   `Total Bayar: Rp ${Number(totalHarga).toLocaleString('id-ID')}\n` +
                   `Status Awal: ${status}\n\n` +
                   `Silakan periksa WhatsApp Admin Anda untuk melihat bukti transfer/QRIS yang dikirimkan oleh pelanggan.`;

  // Kirim email dari server Google Anda
  GmailApp.sendEmail(emailToko, subjek, isiEmail);
  
  // Kunci ID ini agar tidak terkirim dua kali pada menit berikutnya
  properti.setProperty('LAST_SENT_ORDER_ID', String(orderId));
}


// FUNGSI UNTUK MENGIRIM PESAN WA OTOMATIS VIA FONNTE
function kirimAkunViaFonnte(orderId, noHp, namaPelanggan, detailAkun) {
  // ⚠️ GANTI DENGAN TOKEN FONNTE ANDA SENDIRI
  const fonnteToken = "BDUjyYLQTYmcoDxPJX7Z"; 
  
  const pesan = `Halo *${namaPelanggan}*,\n\n` +
                `Terima kasih telah berbelanja di *Senn Store* dan melakukan pembayaran. Otorisasi pesanan Anda dengan *ID: ${orderId}* telah sukses diverifikasi! ✅\n\n` +
                `Berikut adalah detail akun produk digital Anda:\n` +
                `----------------------------------------\n` +
                `${detailAkun}\n` +
                `----------------------------------------\n\n` +
                `Jika ada kendala, silakan hubungi Pusat Bantuan kami kembali ya. Selamat menikmati! 🙏`;

  const url = "https://api.fonnte.com/send";
  const payload = {
    "target": String(noHp),
    "message": pesan
  };
  
  const options = {
    "method": "post",
    "headers": {
      "Authorization": fonnteToken
    },
    "payload": payload,
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const resText = response.getContentText();
    Logger.log("Respon Fonnte: " + resText);
    
    // Jika sukses terkirim, ubah status di Google Sheets jadi 'Selesai'
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pesanan');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == orderId) {
        sheet.getRange(i + 1, 6).setValue('Selesai'); // Kolom F (Status) jadi Selesai
        sheet.getRange(i + 1, 7).setValue(detailAkun); // Kolom G (Keterangan Akun) diisi data akun
        break;
      }
    }
    return { success: true, message: "Akun berhasil dikirim langsung ke WA pelanggan!" };
  } catch(e) {
    return { success: false, message: "Gagal kirim WA Gateway: " + e.toString() };
  }
}

// Tambahan fungsi untuk Admin mengambil data list Pesanan
function getDaftarPesanan() {
  return getSheetData('Pesanan');
}
