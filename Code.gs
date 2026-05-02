function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Sales Order System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName("Products");
  const sSheet = ss.getSheetByName("Stock");
  const pData = pSheet.getDataRange().getValues();
  const sData = sSheet.getDataRange().getValues();
  const products = [];
  
  let stockMap = {};
  for (let j = 1; j < sData.length; j++) {
    stockMap[sData[j][0].toString()] = sData[j][2]; 
  }
  
  for (let i = 1; i < pData.length; i++) {
    if (pData[i][0]) { 
      let code = pData[i][0].toString();
      products.push({
        code: code,
        name: pData[i][1],
        price: pData[i][2] || 0,
        image: pData[i][4] || "",
        stock: stockMap[code] || 0
      });
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
  
  // យកភាគរយ Incentive ពី Form ដែល User បញ្ចូល (Column L)
  const incentivePercent = parseFloat(data.discountPerc) || 0; 

  // បញ្ច្រាស List ទំនិញ ដើម្បីរក្សាលំដាប់ពេល Insert ទៅជួរទី ២
  data.items.reverse().forEach(item => {
    const subTotal = item.price * item.qty; // Column K: Total Price
    const vat = subTotal * 0.10;
    const totalWithVat = subTotal + vat; // Column M: Total+VAT 10%
    
    // គណនាប្រាក់ Incentive ($) តាមភាគរយជាក់ស្តែង (Column N)
    const incentiveAmount = totalWithVat * (incentivePercent / 100);
    
    // រូបមន្តចុងក្រោយ៖ Grand Total (Column O)
    const finalTotal = totalWithVat - incentiveAmount;

    orderSheet.insertRowBefore(2);
    
    const rowValues = [
      new Date(),           // A: Date
      orderId,              // B: Order ID
      data.customer.manager,// C: Manager
      data.customer.name,    // D: Customer Name
      data.customer.phone,   // E: Phone
      data.customer.address, // F: Address
      item.code,            // G: Product Code
      item.name,            // H: Name Product
      item.qty,             // I: QTY
      item.price,           // J: Price
      subTotal,             // K: Total Price
      incentivePercent + "%", // L: Incentive (ភាគរយដែលបញ្ចូលពី Form)
      totalWithVat,         // M: Total+VAT 10%
      incentiveAmount,      // N: Incentive% (ទឹកលុយដែលបានគណនាដក)
      finalTotal            // O: Grand Total
    ];

    orderSheet.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);

    // Update Stock
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

  let html = `
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Kantumruy+Pro&display=swap');
        body { font-family: 'Kantumruy Pro', sans-serif; padding: 20px; color: #333; }
        .header { display: flex; justify-content: space-between; border-bottom: 2px solid #1877f2; padding-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background: #f8f9fa; border-bottom: 2px solid #dee2e6; padding: 10px; text-align: left; }
        td { border-bottom: 1px solid #ddd; padding: 10px; }
        .total-box { text-align: right; margin-top: 20px; }
        .grand-total { font-size: 18px; color: #1877f2; font-weight: bold; border-top: 2px solid #1877f2; padding-top: 5px; }
      </style>
    </head>
    <body>
      <div class="header">
        <div><h2 style="color:#1877f2; margin:0;">Sales Order</h2><p>Order ID: ${data.orderId}</p></div>
        <div style="text-align:right">កាលបរិច្ឆេទ: ${new Date().toLocaleDateString('km-KH')}</div>
      </div>
      <div style="margin: 20px 0;">
        <strong>ព័ត៌មានអតិថិជន:</strong><br>
        ឈ្មោះ: ${data.customer.name} | ទូរស័ព្ទ: ${data.customer.phone}<br>
        អ្នកលក់: ${data.customer.manager}
      </div>
      <table>
        <thead><tr><th>ទំនិញ</th><th>ចំនួន</th><th>តម្លៃ</th><th>សរុប</th></tr></thead>
        <tbody>
          ${data.items.map(item => `
            <tr>
              <td><strong>${item.code}</strong> - ${item.name}</td>
              <td>${item.qty}</td>
              <td>$${item.price.toFixed(2)}</td>
              <td>$${(item.price * item.qty).toFixed(2)}</td>
            </tr>`).join('')}
        </tbody>
      </table>
      <div class="total-box">
        <p>Subtotal: $${subTotalNum.toFixed(2)}</p>
        <p>VAT (10%): $${vatAmount.toFixed(2)}</p>
        ${discPerc > 0 ? `<p style="color:orange;">Discount (${discPerc}%): -$${discountCash.toFixed(2)}</p>` : ''}
        <div class="grand-total">សរុបរួម: $${grandTotal.toFixed(2)}</div>
      </div>
    </body>
    </html>`;
  
  const blob = Utilities.newBlob(html, "text/html", "invoice.html");
  return Utilities.base64Encode(blob.getAs("application/pdf").getBytes());
}
