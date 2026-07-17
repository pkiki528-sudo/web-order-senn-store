function doGet(e) {
  const action = e.parameter.action;
  let data;
  
  if (action === 'getPesanan') {
    data = getSheetData('Pesanan');
  } else {
    data = getSheetData('Produk');
  }
  
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*');
}

function doPost(e) {
  let requestData;
  try {
    requestData = JSON.parse(e.postData.contents);
  } catch(err) {
    requestData = e.parameter;
  }
  
  const action = requestData.action;
  let hasil = { success: false, message: 'Aksi tidak dikenali' };
  
  if (action === 'checkout') {
    hasil = checkoutOrder(requestData.username, requestData.cartItems, requestData.totalHarga);
  } else if (action === 'kirimAkun') {
    hasil = kirimAkunViaFonnte(requestData.orderId, requestData.noHp, requestData.namaPelanggan, requestData.detailAkun);
  }
  
  return ContentService.createTextOutput(JSON.stringify(hasil))
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*');
}

function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => { obj[header] = row[index]; });
    return obj;
  });
}

function checkoutOrder(username, cartItems, totalHarga) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pesananSheet = ss.getSheetByName('Pesanan');
  const produkSheet = ss.getSheetByName('Produk');
  
  const orderId = 'SENN-' + new Date().getTime();
  const tanggal = new Date();
  pesananSheet.appendRow([orderId, tanggal, username, JSON.stringify(cartItems), totalHarga, 'Pending']);
  
  try {
    const produkData = produkSheet.getDataRange().getValues();
    cartItems.forEach(item => {
      for (let i = 1; i < produkData.length; i++) {
        if (produkData[i][0] == item.id) {
          produkSheet.getRange(i + 1, 6).setValue(Math.max(0, Number(produkData[i][5]) - item.qty));
        }
      }
    });
  } catch(e){}

  return { success: true, orderId: orderId };
}

function kirimAkunViaFonnte(orderId, noHp, namaPelanggan, detailAkun) {
  const fonnteToken = "TOKEN_FONNTE_KAMU_DISINI"; // Isi dengan token fonnte kamu
  const pesan = `Halo *${namaPelanggan}*,\n\nTerima kasih telah berbelanja di *Senn Store*. Detail akun produk digital Anda:\n----------------------------------------\n${detailAkun}\n----------------------------------------\n Selamat menikmati! 🙏`;
  
  const options = {
    "method": "post",
    "headers": { "Authorization": fonnteToken },
    "payload": { "target": String(noHp), "message": pesan },
    "muteHttpExceptions": true
  };
  
  try {
    UrlFetchApp.fetch("https://api.fonnte.com/send", options);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pesanan');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == orderId) {
        sheet.getRange(i + 1, 6).setValue('Selesai');
        sheet.getRange(i + 1, 7).setValue(detailAkun);
        break;
      }
    }
    return { success: true, message: "Akun sukses dikirim langsung ke WA pelanggan!" };
  } catch(e) {
    return { success: false, message: "Gagal kirim: " + e.toString() };
  }
}
