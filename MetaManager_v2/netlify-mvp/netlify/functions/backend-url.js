exports.handler = async () => {
  const backendUrl = (process.env.BACKEND_URL || "").trim();
  return {
    statusCode: 200,
    body: JSON.stringify({
      ok: true,
      backendUrl
    })
  };
};
