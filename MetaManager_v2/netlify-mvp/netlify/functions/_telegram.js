const TELEGRAM_BASE = "https://api.telegram.org";

function getTelegramConfig() {
  const token = process.env.TELEGRAM_BOT_TOKEN || "";
  const chatId = process.env.TELEGRAM_CHAT_ID || "";
  return {
    token,
    chatId,
    configured: Boolean(token && chatId)
  };
}

async function sendTelegramMessage({ token, chatId, text }) {
  const url = `${TELEGRAM_BASE}/bot${token}/sendMessage`;
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chat_id: chatId,
      text
    })
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Telegram sendMessage failed: ${response.status} ${body}`);
  }
}

async function sendTelegramDocument({ token, chatId, filename, textContent, caption }) {
  const url = `${TELEGRAM_BASE}/bot${token}/sendDocument`;
  const form = new FormData();
  form.append("chat_id", String(chatId));
  if (caption) {
    form.append("caption", caption);
  }
  const blob = new Blob([textContent], { type: "text/plain;charset=utf-8" });
  form.append("document", blob, filename);

  const response = await fetch(url, {
    method: "POST",
    body: form
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Telegram sendDocument failed: ${response.status} ${body}`);
  }
}

module.exports = {
  getTelegramConfig,
  sendTelegramMessage,
  sendTelegramDocument
};
