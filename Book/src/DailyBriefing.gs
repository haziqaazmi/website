function sendDailyBriefing() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Form Responses 1'
    );

  const data =
    sheet.getDataRange().getValues();

  const today = new Date();

  today.setHours(0, 0, 0, 0);

  const todayStr =
    Utilities.formatDate(
      today,
      "GMT+8",
      "yyyy-MM-dd"
    );

  const allRooms = [
    "1","2","3","4","5",
    "6","7","8","9","10",
    "11","12","13","14","15"
  ];

  let occupiedRooms = [];
  let checkIns = [];
  let checkOuts = [];
  let stayOvers = [];

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const status = row[13];

    if (
      status &&
      status.includes("Canceled")
    ) {
      continue;
    }

    const name = row[2];
    const phone = row[3];

    const ciDate =
      row[4] instanceof Date
        ? Utilities.formatDate(
            row[4],
            "GMT+8",
            "MM-dd-yyyy"
          )
        : row[4];

    const coDate =
      row[5] instanceof Date
        ? Utilities.formatDate(
            row[5],
            "GMT+8",
            "MM-dd-yyyy"
          )
        : row[5];

    const roomNosStr =
      row[16]
        ? row[16].toString()
        : "TBD";

    const roomList =
      roomNosStr
        .split(',')
        .map(r => r.trim());

    const roomType = row[7];

    let cleanPhone =
      phone
        .toString()
        .replace(/[^0-9]/g, '');

    if (
      cleanPhone.startsWith('0')
    ) {
      cleanPhone =
        '60' +
        cleanPhone.substring(1);
    }

    const getWaLink = msg =>
      `https://wa.me/${cleanPhone}?text=${encodeURIComponent(msg)}`;

    const checkInObj =
      new Date(row[4]);

    const checkOutObj =
      new Date(row[5]);

    // ARRIVING TODAY
    if (ciDate === todayStr) {

      const msg =
        `Salam dan Hi ${name}. Selamat datang ke D’Anjung Cottage Paka😊\n\n` +
        `*Berikut maklumat self check-in:* \n\n` +
        `🏠 Room Type: (${roomType})\n` +
        `🔐 Keybox: Room (${roomNosStr})\n` +
        `🔢 Passcode: xxx \n\n` +
        `🕒 Check-in: Selepas 3:00 PM\n` +
        `🕛 Check-out: Sebelum 12:00 PM\n\n` +
        `*Cas late check-out atau early check-in: Rm20/jam*\n\n` +
        `Sila tonton video sebelum tiba untuk memudahkan proses check-in.\n\n` +
        `*Peraturan penginapan (Mohon Kerjasama Semua🙏🏼):*\n\n` +
        `🚭 Dilarang merokok dalam bilik (RM100 denda akan dikenakan)\n` +
        `🚫 Dilarang bawa durian masuk bilik\n` +
        `🏊 Pool time: 7AM – 10PM\n` +
        `🚯 Dilarang membuang sampah merata2\n` +
        `🙊 Dilarang membuat bising\n` +
        `👶🏻 Pampers bayi perlu dibuang didalam tong sampah luar bilik\n` +
        `💦 Elakkan memasuki bilik dalam keadaan basah kuyup\n\n` +
        `*****************************\n\n` +
        `*Info penggunaan Pantry Dan Laundry room*\n\n` +
        `☕Kami ada menyediakan kopi,teh,susu,gula dan water dispenser.\n` +
        `👕Penggunaan mesin basuh,sabun dan pengering adalah percuma\n\n` +
        `*****************************\n\n` +
        `Jika perlukan apa-apa bantuan sepanjang penginapan, boleh hubungi kami pada bila-bila masa.\n\n` +
        `Terima kasih,\n` +
        `Enjoy your stay! 😊`;

      checkIns.push(
        `➡️ <b>${name}</b> ` +
        `(R: ${roomNosStr}) - ` +
        `<a href="${getWaLink(msg)}">Send Check-in Info</a>`
      );

    }

    // DEPARTING TODAY
    else if (coDate === todayStr) {

      const msg =
        `Hi ${name}, Terima kasih kerana memilih kami 😊\n` +
        `Cara self check-out seperti dalam video yang dilampirkan ya.\n\n` +
        `Sebelum keluar, mohon:\n` +
        `• Tutup semua peralatan elektrik terutamanya aircond\n` +
        `• Tutup lampu & paip air\n` +
        `• Buang sisa makanan & pampers ke tong sampah di luar bilik\n\n` +
        `Jika berpuas hati dengan penginapan anda, mohon jasa baik beri rating ⭐⭐⭐⭐⭐ di Google Review:\n\n` +
        `https://tinyurl.com/Danjung-review\n\n` +
        `Terima kasih atas kerjasama & semoga berjumpa lagi 🌟`;

      checkOuts.push(
        `⬅️ <b>${name}</b> ` +
        `(R: ${roomNosStr}) - ` +
        `<a href="${getWaLink(msg)}">Send Check-out Info</a>`
      );

    }

    // STAYING OVER
    else if (
      today > checkInObj &&
      today < checkOutObj
    ) {

      occupiedRooms =
        occupiedRooms.concat(
          roomList
        );

      const msg =
        `Hi ${name}, would you like housekeeping for your room (${roomNosStr}) today? 🧹`;

      stayOvers.push(
        `🏡 <b>${name}</b> ` +
        `(R: ${roomNosStr}) - ` +
        `<a href="${getWaLink(msg)}">Housekeeping?</a>`
      );
    }
  }

  const uniqueOccupied =
    [...new Set(occupiedRooms)];

  const availableRooms =
    allRooms.filter(
      room =>
        !uniqueOccupied.includes(room)
    );

  let msg =
    `📅 <b>BRIEFING: ${Utilities.formatDate(today, "GMT+8", "dd MMM")}</b>\n\n`;

  msg +=
    `🛬 <b>CHECK-INS (${checkIns.length})</b>\n` +
    `${checkIns.length > 0 ? checkIns.join('\n') : "<i>None</i>"}\n\n`;

  msg +=
    `🛫 <b>CHECK-OUTS (${checkOuts.length})</b>\n` +
    `${checkOuts.length > 0 ? checkOuts.join('\n') : "<i>None</i>"}\n\n`;

  msg +=
    `🛌 <b>STAY-OVERS (${stayOvers.length})</b>\n` +
    `${stayOvers.length > 0 ? stayOvers.join('\n') : "<i>None</i>"}\n\n`;

  msg +=
    `✨ <b>AVAILABLE:</b> ` +
    `<code>${availableRooms.join(', ')}</code>`;

  sendTelegramAlert(msg);
}
