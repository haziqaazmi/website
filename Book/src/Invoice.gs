function generateInvoiceFromSelectedRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Form Responses 1');

  const activeRow = sheet.getActiveRange().getRow();

  if (activeRow < 2) {
    SpreadsheetApp.getUi().alert("Please select a guest row first.");
    return;
  }

  const data = sheet.getRange(activeRow, 1, 1, 18).getValues()[0];
  const pricingDb = getPricingData();

  const guestEmail = data[1];
  const guestName = data[2];
  const checkIn = new Date(data[4]);
  const checkOut = new Date(data[5]);
  const ref = data[9];
  const totalAmount = Number(data[10]);
  const rawRoomNos = data[16].toString();

  const typesArr = data[7].toString()
    .split(',')
    .map(s => s.trim());

  const numsArr = rawRoomNos
    .split(',')
    .map(s => s.trim());

  let invoiceRows = [];
  let tempDate = new Date(checkIn);
  let totalEstimatedPrice = 0;

  while (tempDate < checkOut) {
    let currentTime = tempDate.getTime();

    typesArr.forEach(roomType => {
      let priceInfo = pricingDb[roomType];

      if (priceInfo) {
        totalEstimatedPrice += calculateNightlyRate(
          priceInfo,
          currentTime
        );
      }
    });

    tempDate.setDate(tempDate.getDate() + 1);
  }

  const copy = DriveApp
    .getFileById(TEMPLATE_ID)
    .makeCopy(
      `Invoice_${ref}_${guestName}`,
      DriveApp.getRootFolder()
    );

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  const guestPhone = data[3];
  const bookstatus = data[13];

  body.replaceText("{{Name}}", guestName);
  body.replaceText("{{Ref}}", ref);
  body.replaceText("{{Phone}}", guestPhone);
  body.replaceText("{{Status}}", bookstatus);
  body.replaceText(
    "{{Date}}",
    Utilities.formatDate(
      new Date(),
      "GMT+8",
      "dd/MM/yyyy"
    )
  );
  body.replaceText("{{RoomNo}}", rawRoomNos);

  const itemTable = body.getTables()[1];

  itemTable.setBorderColor('#FFFFFF');

  const templateRow = itemTable.getRow(1);

  invoiceRows.forEach((item, index) => {
    let newRow =
      (index === 0)
        ? templateRow
        : itemTable.appendTableRow(templateRow.copy());

    newRow.getCell(0).setText(item.date);
    newRow.getCell(1).setText(item.desc);
    newRow.getCell(2).setText(item.rate);
  });

  let totalRow = itemTable.appendTableRow();

  totalRow.appendTableCell("");

  totalRow
    .appendTableCell("GRAND TOTAL")
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.RIGHT
    )
    .setBold(true);

  totalRow
    .appendTableCell(
      "RM " + totalAmount.toFixed(2)
    )
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.RIGHT
    )
    .setBold(true);

  try {
    let headerRow = itemTable.getRow(0);

    for (let i = 0; i < 3; i++) {
      headerRow.getCell(i).setAttributes({
        [DocumentApp.Attribute.BORDER_BOTTOM_COLOR]: '#000000',
        [DocumentApp.Attribute.BORDER_BOTTOM_WIDTH]: 1
      });
    }

    for (let i = 1; i < 3; i++) {
      totalRow.getCell(i).setAttributes({
        [DocumentApp.Attribute.BORDER_TOP_COLOR]: '#000000',
        [DocumentApp.Attribute.BORDER_TOP_WIDTH]: 1
      });
    }
  } catch (e) {
    // Fallback
  }

  doc.saveAndClose();

  const pdf = copy.getAs(MimeType.PDF);

  MailApp.sendEmail({
    to: guestEmail,
    subject: `Receipt - ${ref}`,
    htmlBody:
      `Thank you ${guestName}. We received your payment. ` +
      `Please find attached receipt for your reference.`,
    attachments: [pdf]
  });

  let cleanPhone = guestPhone
    .toString()
    .replace(/[^0-9]/g, '');

  if (!cleanPhone.startsWith('6')) {
    cleanPhone = '6' + cleanPhone;
  }

  sendTelegramAlert(
    `✅ <b>Manual Invoice Sent</b>\n` +
    `👤 Guest: ${guestName}\n` +
    `📱 Phone: <a href="https://wa.me/${cleanPhone}">WhatsApp Guest</a>\n` +
    `💰 Total: <b>RM ${totalAmount.toFixed(2)}</b>`
  );

  copy.setTrashed(true);
}


/**
 * Generate official PDF invoice
 */
function generateSelectedRowInvoice(
  sheet,
  statusCell,
  rowNum,
  adminSheet
) {
  const TEMPLATE_ID =
    "1J3AT_PQMBsCP5me1C92tvQ4G7qcBWRKJ_MIbnzh_OcE";

  const FOLDER_ID =
    "1_fBWd1d9EX2PAVBpc3AxckBtjkbgOj_B";

  if (!rowNum || isNaN(rowNum) || rowNum <= 1) {
    statusCell.setValue("❌ Error: Row # in B5");
    adminSheet.getRange("B4").setValue(false);
    return;
  }

  try {
    statusCell.setValue("⏳ Building PDF...");

    const pricingRules = getPricingData();

    const data =
      sheet.getRange(rowNum, 1, 1, 24).getValues()[0];

    const guestName = data[2];
    const guestPhone = data[3];
    const checkIn = new Date(data[4]);
    const checkOut = new Date(data[5]);
    const roomTypesString = data[7].toString();
    const ref = data[9];
    const rawRoomNos = data[16];

    const roomNumbersList = rawRoomNos
      .toString()
      .split(',')
      .map(s => s.trim());

    const totalPaidToDate =
      Number(data[17]) || 0;

    const addonHours =
      Number(data[22]) || 0;

    const addonTax =
      Number(data[23]) || 0;

    let invoiceRows = [];
    let tempDate = new Date(checkIn);
    let correctedGrandTotal = 0;

    // PRICING ENGINE
    while (tempDate < checkOut) {
      tempDate.setHours(12, 0, 0, 0);

      let formattedDate =
        Utilities.formatDate(
          tempDate,
          "GMT+8",
          "dd MMM yyyy"
        );

      let currentTime =
        tempDate.getTime();

      roomNumbersList.forEach((num, index) => {
        let rTypes =
          roomTypesString
            .split(',')
            .map(s => s.trim());

        let rType =
          rTypes[index] || rTypes[0];

        let rule = pricingRules[rType];

        let nightPrice =
          calculateNightlyRate(
            rule,
            currentTime
          );

        correctedGrandTotal += Number(nightPrice);

        invoiceRows.push({
          date: formattedDate,
          desc: `${rType.toUpperCase()} (${num})`,
          rate: Number(nightPrice).toFixed(2)
        });
      });

      tempDate.setDate(
        tempDate.getDate() + 1
      );
    }

    // ADD-ON HOURS
    if (addonHours > 0) {
      const addonTotal =
        addonHours * 20;

      correctedGrandTotal += addonTotal;

      invoiceRows.push({
        date: "-",
        desc:
          `ADD-ON HOUR (${addonHours} HRS x RM20)`,
        rate:
          addonTotal.toFixed(2)
      });
    }

    // BOOKING.COM TAX
    if (addonTax > 0) {
      correctedGrandTotal += addonTax;

      invoiceRows.push({
        date: "-",
        desc: "BOOKING.COM TAX & CHARGES",
        rate: addonTax.toFixed(2)
      });
    }

    const finalBalanceDue =
      Math.max(
        0,
        Number(
          (
            correctedGrandTotal -
            totalPaidToDate
          ).toFixed(2)
        )
      );

    const paymentStatusLabel =
      finalBalanceDue <= 0
        ? "FULLY PAID"
        : "PARTIAL (DEPOSIT)";

    // DOCUMENT
    const copy =
      DriveApp
        .getFileById(TEMPLATE_ID)
        .makeCopy(
          `Invoice_${ref}_${guestName}`,
          DriveApp.getRootFolder()
        );

    const doc =
      DocumentApp.openById(copy.getId());

    const body =
      doc.getBody();

    body.replaceText(
      "{{Name}}",
      guestName
    );

    body.replaceText(
      "{{Ref}}",
      ref
    );

    body.replaceText(
      "{{Phone}}",
      guestPhone
    );

    body.replaceText(
      "{{Status}}",
      paymentStatusLabel
    );

    body.replaceText(
      "{{Checkin}}",
      Utilities.formatDate(
        checkIn,
        "GMT+8",
        "dd MMM yyyy"
      )
    );

    body.replaceText(
      "{{Checkout}}",
      Utilities.formatDate(
        checkOut,
        "GMT+8",
        "dd MMM yyyy"
      )
    );

    body.replaceText(
      "{{Date}}",
      Utilities.formatDate(
        new Date(),
        "GMT+8",
        "dd/MM/yyyy"
      )
    );

    body.replaceText(
      "{{RoomNo}}",
      rawRoomNos
    );

    const itemTable =
      body.getTables()[1];

    while (
      itemTable.getNumRows() > 1
    ) {
      itemTable.removeRow(1);
    }

    invoiceRows.forEach(item => {
      let row =
        itemTable.appendTableRow();

      row.appendTableCell(
        item.date
      );

      row.appendTableCell(
        item.desc
      );

      row
        .appendTableCell(
          "RM " + item.rate
        )
        .getChild(0)
        .asParagraph()
        .setAlignment(
          DocumentApp.HorizontalAlignment.RIGHT
        );
    });

    addSummaryRowToDoc(
      itemTable,
      "GRAND TOTAL",
      "RM " +
        correctedGrandTotal.toFixed(2),
      true
    );

    addSummaryRowToDoc(
      itemTable,
      "TOTAL PAID",
      "RM " +
        totalPaidToDate.toFixed(2),
      false
    );

    addSummaryRowToDoc(
      itemTable,
      "BALANCE DUE",
      "RM " +
        finalBalanceDue.toFixed(2),
      true
    );

    itemTable.setBorderWidth(0);

    doc.saveAndClose();

    const pdfBlob =
      copy.getAs(MimeType.PDF);

    const pdfFile =
      DriveApp
        .getFolderById(FOLDER_ID)
        .createFile(pdfBlob);

    copy.setTrashed(true);

    // COLUMN U
    sheet
      .getRange(rowNum, 21)
      .setValue(pdfFile.getUrl());

    // COLUMN V - WHATSAPP
    let cleanPhone =
      guestPhone
        .toString()
        .replace(/[^0-9]/g, '');

    cleanPhone =
      cleanPhone.startsWith('0')
        ? '60' +
          cleanPhone.substring(1)
        : (
            cleanPhone.startsWith('60')
              ? cleanPhone
              : '60' + cleanPhone
          );

    const waMessage =
      encodeURIComponent(
        `Hi ${guestName}, this is D'Anjung Cottage. ` +
        `Here is your invoice for booking ${ref}.`
      );

    const waUrl =
      `https://wa.me/${cleanPhone}?text=${waMessage}`;

    sheet
      .getRange(rowNum, 22)
      .setFormula(
        `=HYPERLINK("${waUrl}", "📲 SEND WA")`
      );

    statusCell.setValue(
      "✅ PDF Ready!"
    );

  } catch (e) {
    statusCell.setValue(
      "❌ Error: " + e.toString()
    );
  } finally {
    adminSheet
      .getRange("B4")
      .setValue(false);
  }
}


function addSummaryRowToDoc(
  table,
  label,
  value,
  isBold
) {
  let row =
    table.appendTableRow();

  row.appendTableCell("");

  row
    .appendTableCell(label)
    .getChild(0)
    .asParagraph()
    .setBold(isBold)
    .setAlignment(
      DocumentApp.HorizontalAlignment.RIGHT
    );

  row
    .appendTableCell(value)
    .getChild(0)
    .asParagraph()
    .setBold(isBold)
    .setAlignment(
      DocumentApp.HorizontalAlignment.RIGHT
    );
}
