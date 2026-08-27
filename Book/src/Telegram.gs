function sendTelegramAlert(message) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID,
      text: message,
      parse_mode: 'HTML',
      disable_web_page_preview: true
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (!result.ok) {
      Logger.log("Telegram API Error: " + response.getContentText());
    }
  } catch (e) {
    Logger.log("Telegram Connection Error: " + e.toString());
  }
}

function testMyBot() {
  const message =
    "🚀 <b>D'Anjung Cottage System Online!</b>\n\n" +
    "Your Telegram bot is now successfully linked to your Google Sheet. " +
    "You will receive booking and payment alerts here.";

  sendTelegramAlert(message);
}
