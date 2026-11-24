const SHEET_NAME = 'Orders';

// ฟังก์ชันตั้งค่าหัวตารางอัตโนมัติ (กดรันครั้งเดียวตอนติดตั้ง)
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  const headers = [
    'Order_ID', 'Timestamp', 'Customer_Name', 'Phone', 
    'Address', 'SKU1_Qty', 'SKU2_Qty', 'Total_Price', 
    'Payment_Method', 'Status'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#065f46').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

// รับข้อมูลคำสั่งซื้อจากหน้า Landing Page
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }

    const data = JSON.parse(e.postData.contents);
    const orderId = 'ORD-' + Math.floor(100000 + Math.random() * 900000);
    const timestamp = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });

    sheet.appendRow([
      orderId,
      timestamp,
      data.name,
      "'" + data.phone,
      data.address,
      data.sku1Qty,
      data.sku2Qty,
      data.totalPrice,
      data.paymentMethod,
      'รอจัดเตรียม (Pending)'
    ]);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', orderId: orderId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
