exports.handler = async (event) => {
  if (event.httpMethod !== "GET") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const backendUrl = process.env.BACKEND_URL || "";
    if (!backendUrl) {
      return {
        statusCode: 400,
        body: JSON.stringify({
          error:
            "Нужен Python backend. Задайте переменную BACKEND_URL в Netlify " +
            "и CHECKO_API_KEY на backend."
        })
      };
    }

    const qs = event.queryStringParameters || {};
    const inn = (qs.inn || "").trim();
    if (!inn) {
      return { statusCode: 400, body: JSON.stringify({ error: "Параметр inn обязателен" }) };
    }

    const base = backendUrl.replace(/\/+$/, "");
    const url = `${base}/lookup/company?inn=${encodeURIComponent(inn)}`;
    const response = await fetch(url, { method: "GET" });
    const data = await response.json();
    if (!response.ok) {
      return { statusCode: response.status, body: JSON.stringify(data) };
    }

    return { statusCode: 200, body: JSON.stringify(data) };
  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message || "Internal error" })
    };
  }
};
