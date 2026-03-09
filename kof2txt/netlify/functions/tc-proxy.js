exports.handler = async function (event) {
  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      body: "Use POST"
    };
  }

  try {
    const data = JSON.parse(event.body || "{}");
    const url = data.url;
    const token = data.token;
    const method = data.method || "GET";
    const body = data.body;

    if (!url) {
      return {
        statusCode: 400,
        body: JSON.stringify({
          ok: false,
          error: "Missing url"
        })
      };
    }

    if (!token) {
      return {
        statusCode: 400,
        body: JSON.stringify({
          ok: false,
          error: "Missing token"
        })
      };
    }

    const headers = {
      Authorization: "Bearer " + token,
      Accept: "application/json"
    };

    if (body) {
      headers["Content-Type"] = "application/json";
    }

    console.log("Proxy request starting");
    console.log("Method:", method);
    console.log("URL:", url);

    const response = await fetch(url, {
      method,
      headers,
      body: body ? JSON.stringify(body) : undefined
    });

    const text = await response.text();

    console.log("Proxy response status:", response.status);
    console.log("Proxy response preview:", text.slice(0, 500));

    return {
      statusCode: response.status,
      headers: {
        "Content-Type": response.headers.get("content-type") || "text/plain"
      },
      body: text
    };
  } catch (err) {
    console.error("Proxy fetch error:", err);
    return {
      statusCode: 500,
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        ok: false,
        error: String(err),
        message: err && err.message ? err.message : null,
        cause: err && err.cause ? String(err.cause) : null,
        stack: err && err.stack ? err.stack : null
      })
    };
  }
};
