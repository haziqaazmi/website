function cleanupExpiredBookings() {
  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      'Form Responses 1'
    );

  const data =
    sheet.getDataRange().getValues();

  const now =
    new Date().getTime();

  const hour =
    60 * 60 * 1000;

  const fortySevenHours =
    47 * hour;

  const fortyEightHours =
    48 * hour;

  for (let i = 1; i < data.length; i++) {

    const timestamp =
      new Date(data[i][0]).getTime();

    if (isNaN(timestamp)) continue;

    const guestName = data[i][2];
    const guestEmail = data[i][1];
    const status = data[i][8];
    const ref = data[i][9];
    const paymentStatus = data[i][13];

    const timeElapsed =
      now - timestamp;

    // PAID / DEPOSIT
    if (
      paymentStatus === "Paid (BCOM)" ||
      paymentStatus === "Fully Paid" ||
      paymentStatus === "Partial (Deposit)"
    ) {
      if (
        status !== "Confirmed" &&
        status !== "Cancelled"
      ) {
        sheet
          .getRange(i + 1, 9)
          .setValue("Confirmed");
      }

      continue;
    }

    if (status !== "Cancelled") {

      // AUTO CANCEL AFTER 48 HOURS
      if (
        timeElapsed >= fortyEightHours
      ) {

        sheet
          .getRange(i + 1, 9)
          .setValue("Cancelled");

        try {

          const calendarMap = {
            "1": "b8a24f392804fbed91419ba4f45f5c9bf32bb0f42c6a7bd6f97d66c3a5d8f711@group.calendar.google.com",
            "2": "625ea9611e1c00708a04d7a360a41037d773fef416b92c15b3446c244ddea170@group.calendar.google.com",
            "3": "51e4fc458f175d7be8dc773b8ee0cfd99cb748de56b11f5e35d2cb0117625bc5@group.calendar.google.com",
            "4": "52f1e794dee3c260960d38e1eaa17d7ad86061d9e1c45b6d20f7d5e005791f08@group.calendar.google.com",
            "5": "5f994d449bbed887a61ae4838ea679dde064d6fe4a8b2fbe53a42a24a7403d@group.calendar.google.com",
            "6": "9905a281ea9f4fbe8525fbd00473b33827f403e2836b182964055e5e04de92cc@group.calendar.google.com",
            "7": "45818db6387e8af625ebda2f8cec9ee009f6e2b789058e4ce36b0607210fa3ad@group.calendar.google.com",
            "8": "f9a04818ff2bc5c71a7b34f2406007940ed6d9c8c0dbb80760d3aeb6b95c19e3@group.calendar.google.com",
            "9": "cc8e1ed0a8916a96fca075962dee4fe8452f8b4785bbce324cc6fb544f5ba6f9@group.calendar.google.com",
            "10": "015049b3a50ab1329ad142b7759623d088976942aefde22fcba18e69942d152b@group.calendar.google.com",
            "11": "af2f60eebc116990a1dad57fc1750591b2c9179df4ea5a29a5097ed715fe2115@group.calendar.google.com",
            "12": "b1c350d921e8071a5bbba5edae289211d0e8d072e849953c3eeffa62a485150d@group.calendar.google.com",
            "13": "b5d8a6b6a816e0b378de6fe77fb846ca0005db1af324dd863f3136b7dd36c40f@group.calendar.google.com",
            "14": "75fcc6ef9275d37af6e177944fa3ebaf1242c371da1e75f4d5c91c3110860a3c@group.calendar.google.com",
            "15": "46a2015378e21df371cae842aeb39cbea390973789344b11916622053c9e8b14@group.calendar.google.com"
          };

          const checkIn =
            new Date(data[i][4]);

          const checkOut =
            new Date(data[i][5]);

          const assignedRooms =
            data[i][16]
              .toString()
              .split(',')
              .map(r => r.trim())
              .filter(String);

          assignedRooms.forEach(roomNum => {

            const calId =
              calendarMap[roomNum];

            if (!calId) return;

            const cal =
              CalendarApp.getCalendarById(
                calId
              );

            if (!cal) return;

            const events =
              cal.getEvents(
                checkIn,
                checkOut
              );

            events.forEach(ev => {

              if (
                ev.getTitle().includes(ref) ||
                ev.getTitle().includes(guestName)
              ) {
                ev.deleteEvent();
              }

            });

          });

        } catch (e) {
          console.log(
            "Calendar removal error: " +
            e.message
          );
        }

        sendTelegramAlert(
          `🔴 <b>AUTO-CANCELLED</b>\n` +
          `Guest: ${guestName}\n` +
          `Ref: ${ref}\n` +
          `Rooms Cleared: ${data[i][16]}`
        );

        try {
          MailApp.sendEmail(
            guestEmail,
            "Booking Cancelled - D'Anjung Cottage",
            `Hi ${guestName}, your booking (${ref}) has been cancelled as we did not receive payment within the 48-hour window.`
          );
        } catch (e) {
          console.log(
            "Email failed for: " +
            guestEmail
          );
        }
      }

      // REMINDER 47-48 HOURS
      else if (
        timeElapsed >= fortySevenHours &&
        data[i][15] !== "Reminder Sent"
      ) {

        sendTelegramAlert(
          `⚠️ <b>1-HOUR EXPIRY</b>\n` +
          `Guest: ${guestName}\n` +
          `Ref: ${ref}\n` +
          `Last chance to pay before cancellation!`
        );

        try {
          MailApp.sendEmail(
            guestEmail,
            "Your Booking Expires in 1 Hour",
            `Hi ${guestName}, your booking ${ref} will be automatically cancelled in 1 hour. Please complete your payment now.`
          );
        } catch (e) {
          console.log(
            "Email failed for: " +
            guestEmail
          );
        }

        sheet
          .getRange(i + 1, 16)
          .setValue("Reminder Sent");
      }

      // ACTIVE
      else {
        if (status !== "Active") {
          sheet
            .getRange(i + 1, 9)
            .setValue("Active");
        }
      }
    }
  }
}
