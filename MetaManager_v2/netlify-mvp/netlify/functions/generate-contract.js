exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const payload = JSON.parse(event.body || "{}");
    if (!payload.contractNumber || !payload.customerShortname) {
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "contractNumber и customerShortname обязательны" })
      };
    }

    if (payload.includeWorkAddress && !payload.workAddress) {
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "Укажите адрес проведения работ" })
      };
    }

    const backendUrl = process.env.BACKEND_URL || "";
    if (!backendUrl) {
      return {
        statusCode: 400,
        body: JSON.stringify({
          error:
            "Для отправки договора в формате ПК-версии нужен Python backend. " +
            "Задайте переменную BACKEND_URL в Netlify."
        })
      };
    }

    const response = await fetch(`${backendUrl.replace(/\/+$/, "")}/generate/contract`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!response.ok) {
      return { statusCode: response.status, body: JSON.stringify(data) };
    }

    return {
      statusCode: 200,
      body: JSON.stringify(data)
    };
  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message || "Internal error" })
    };
  }
};
