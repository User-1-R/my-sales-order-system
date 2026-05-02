// ១. មុខងារទទួលទិន្នន័យពី GitHub (POST Method)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = saveSalesOrder(data); 
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "orderId": result.orderId }))
           .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": err.toString() }))
           .setMimeType(ContentService.MimeType.JSON);
  }
}

// ២. មុខងារបញ្ជូនទិន្នន័យទៅ GitHub (GET Method)
function doGet(e) {
  // ប្រសិនបើហៅមកដើម្បីយកបញ្ជីផលិតផល
  if (e.parameter.action === 'getProducts') {
    const products = getProducts();
    return ContentService.createTextOutput(JSON.stringify(products))
           .setMimeType(ContentService.MimeType.JSON);
  }
  
  // បើហៅមកធម្មតា ឱ្យបង្ហាញទំព័រ HTML
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('Sales Order System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- មុខងារ getProducts, saveSalesOrder និង generateSalesPDF របស់អ្នករក្សានៅដដែលខាងក្រោមនេះ ---

function getProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName("Products");
  const sSheet = ss.getSheetByName("Stock");
  const pData = pSheet.getDataRange().getValues();
  const sData = sSheet.getDataRange().getValues();
  const products = [];
  let stockMap = {};
  for (let j = 1; j < sData.length; j++) { stockMap[sData[j][0].toString()] = sData[j][2]; }
  for (let i = 1; i < pData.length; i++) {
    if (pData[i][0]) { 
      let code = pData[i][0].toString();
      products.push({ code: code, name: pData[i][1], price: pData[i][2] || 0, image: pData[i][4] || "", stock: stockMap[code] || 0 });
    }
  }
  return products;
}

function saveSalesOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders") || ss.insertSheet("Orders");
  const stockSheet = ss.getSheetByName("Stock");
  const sData = stockSheet.getDataRange().getValues();
  const orderId = "IN-" + Utilities.formatDate(new Date(), "GMT+7", "ddMM-") + Math.floor(1000 + Math.random() * 9000);
  const incentivePercent = parseFloat(data.discountPerc) || 0; 

  data.items.reverse().forEach(item => {
    const subTotal = item.price * item.qty;
    const vat = subTotal * 0.10;
    const totalWithVat = subTotal + vat; 
    const incentiveAmount = totalWithVat * (incentivePercent / 100);
    const finalTotal = totalWithVat - incentiveAmount;
    orderSheet.insertRowBefore(2);
    const rowValues = [new Date(), orderId, data.customer.manager, data.customer.name, data.customer.phone, data.customer.address, item.code, item.name, item.qty, item.price, subTotal, incentivePercent + "%", totalWithVat, incentiveAmount, finalTotal];
    orderSheet.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);
    for (let j = 1; j < sData.length; j++) {
      if (sData[j][0].toString() === item.code.toString()) {
        const currentStock = sData[j][2] || 0;
        stockSheet.getRange(j + 1, 3).setValue(currentStock - item.qty); 
        break;
      }
    }
  });
  return { orderId: orderId };
}

function generateSalesPDF(data) {
  const subTotalNum = parseFloat(data.subtotal.replace(/[$,]/g, '')) || 0;
  const vatAmount = subTotalNum * 0.10;
  const discPerc = parseFloat(data.discountPerc) || 0;
  const totalBeforeDisc = subTotalNum + vatAmount;
  const discountCash = totalBeforeDisc * (discPerc / 100);
  const grandTotal = totalBeforeDisc - discountCash;
  let html = `<html><body style="font-family:sans-serif;"><h2>Order ID: ${data.orderId}</h2><p>Customer: ${data.customer.name}</p><h3>Total: $${grandTotal.toFixed(2)}</h3></body></html>`;
  const blob = Utilities.newBlob(html, "text/html", "invoice.html");
  return Utilities.base64Encode(blob.getAs("application/pdf").getBytes());
}