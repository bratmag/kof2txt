exports.handler = async function (event) {
  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      body: "Use POST"
    };
  }

  try {
    const data = JSON.parse(event.body);
    const url = data.url;
    const token = data.token;

    const response = await fetch(url, {
      method: data.method || "GET",
      headers: {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json"
      }
    });

    const text = await response.text();

    return {
      statusCode: response.status,
      body: text
    };

  } catch (err) {
    return {
      statusCode: 500,
      body: "Proxy error: " + err.toString()
    };
  }
};
