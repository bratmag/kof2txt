// kof2txt/netlify/functions/tc-proxy.js
// Netlify Function (CommonJS) - proxy for Trimble Connect Core API (unngår CORS i browser)

exports.handler = async (event) => {
  try {
    if (event.httpMethod !== "POST") {
      return {
        statusCode: 405,
        headers: { "content-type": "text/plain; charset=utf-8" },
        body: "Use POST",
      };
    }

    const payload = JSON.parse(event.body || "{}");
    const { url, method = "GET", token, headers = {}, body } = payload || {};

    if (!url || typeof url !== "string") {
      return { statusCode: 400, body: "Missing url" };
    }
    if (!token || typeof token !== "string") {
      return { statusCode: 400, body: "Missing token" };
    }

    // Allowlist hosts
    const allowedHosts = new Set([
      "app.connect.trimble.com",
      "app21.connect.trimble.com",
      "app31.connect.trimble.com",
      "realtime-service.az.quadri.trimble.com",
      "graphql.az.quadri.trimble.com",
    ]);

    const u = new URL(url);
    if (!allowedHosts.has(u.host)) {
      return { statusCode: 403, body: `Host not allowed: ${u.host}` };
    }

    const h = {
      ...headers,
      Authorization: `Bearer ${token}`,
      Accept: headers.Accept || "application/json",
    };

    let fetchBody;
    if (body !== undefined && body !== null) {
      if (typeof body === "string") {
        fetchBody = body;
      } else {
        fetchBody = JSON.stringify(body);
        if (!h["Content-Type"] && !h["content-type"]) {
          h["Content-Type"] = "application/json";
        }
      }
    }

    const res = await fetch(url, { method, headers: h, body: fetchBody });
    const text = await res.text();

    return {
      statusCode: res.status,
      headers: {
        "content-type": res.headers.get("content-type") || "text/plain; charset=utf-8",
      },
      body: text,
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: { "content-type": "text/plain; charset=utf-8" },
      body: `Proxy error: ${String(e)}`,
    };
  }
};