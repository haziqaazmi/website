function pullBookingToSheets() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Form Responses 1'
    );

  const syncSheet =
    ss.getSheetByName(
      'SyncLinks'
    );

  if (!syncSheet) return;

  const syncData =
    syncSheet.getDataRange().getValues();

  const pricingDb =
    getPricingData();

  const lastRow =
    sheet.getLastRow();

  const existingRefs =
    lastRow > 1
      ? sheet
          .getRange(1, 10, lastRow, 1)
          .getValues()
          .flat()
          .map(r =>
            String(r)
              .trim()
              .toLowerCase()
          )
      : [];

  for (
    let i = 1;
    i < syncData.length;
    i++
  ) {

    const roomId =
      syncData[i][0].toString();

    const icalUrl =
      syncData[i][1];

    const roomType =
      syncData[i][2];

    if (
      !icalUrl ||
      !icalUrl.startsWith("http")
    ) {
      continue;
    }

    try {

      const response =
        UrlFetchApp.fetch(
          icalUrl +
          "?nocache=" +
          new Date().getTime()
        )
        .getContentText();

      const events =
        parseICal(response);

      events.forEach(event => {

        const ref =
          event.uid
            ? String(event.uid)
                .trim()
                .toLowerCase()
            : (
                "BCOM-" +
                roomId +
                "-" +
                event.start.getTime()
              ).toLowerCase();

        if (
          existingRefs.indexOf(ref) === -1
        ) {

          let totalEstimatedPrice = 0;

          let tempDate =
            new Date(event.start);

          const checkOut =
            new Date(event.end);

          const nights =
            Math.max(
              1,
              Math.round(
                (
                  checkOut -
                  tempDate
                ) /
                (
                  1000 *
                  60 *
                  60 *
                  24
                )
              )
            );

          while (
            tempDate < checkOut
          ) {

            let priceInfo =
              pricingDb[roomType];

            if (priceInfo) {

              let currentTime =
                tempDate.getTime();

              totalEstimatedPrice +=
                calculateNightlyRate(
                  priceInfo,
                  currentTime
                );
            }

            tempDate.setDate(
              tempDate.getDate() + 1
            );
          }

          sheet.appendRow([
            new Date(),
            "bcom@manual.com",
            event.name ||
              "Booking.com Guest",
            "N/A",
            event.start,
            event.end,
            nights,
            roomType,
            "Confirmed",
            ref,
            totalEstimatedPrice,
            "",
            "",
            "Paid (BCOM)",
            "Booking.com",
            "",
            roomId
          ]);

          existingRefs.push(ref);

          SpreadsheetApp.flush();

          sendTelegramAlert(
            `🏨 <b>New B.com Import</b>\n` +
            `👤 Guest: ${event.name || "Guest"}\n` +
            `🚪 Room: ${roomId} (${roomType})\n` +
            `📅 ${Utilities.formatDate(event.start, "GMT+8", "dd MMM")} - ${Utilities.formatDate(event.end, "GMT+8", "dd MMM")}\n` +
            `💰 Est. Price: <b>RM ${totalEstimatedPrice.toFixed(2)}</b>`
          );

          Logger.log(
            "✅ Added: " + ref
          );
        }
      });

    } catch (e) {

      Logger.log(
        `Error Room ${roomId}: ${e.message}`
      );
    }
  }
}
