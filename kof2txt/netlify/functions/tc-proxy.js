// netlify/functions/tc-proxy.js
// En enkel proxy for Trimble Connect Core API for å unngå CORS fra browser.
// Viktig: Vi lagrer ikke token, vi videresender bare.

export default async (req, context) => {
  try {
    if (req.method !== "POST") {
      return new Response("Use POST", { status: 405 });
    }

    const payload = await req.json();
    const { url, method = "GET", token, headers = {}, body } = payload || {};

    if (!url || typeof url !== "string") {
      return new Response("Missing url", { status: 400 });
    }
    if (!token || typeof token !== "string") {
      return new Response("Missing token", { status: 400 });
    }

    // Allowlist: bare tillat kall til Trimble Connect hosts + presigned download URL-er
    const allowedHosts = new Set([
      "app.connect.trimble.com",
      "app21.connect.trimble.com",
      "app31.connect.trimble.com",
    ]);

    const u = new URL(url);
    if (!allowedHosts.has(u.host)) {
      return new Response(`Host not allowed: ${u.host}`, { status: 403 });
    }

    const h = new Headers(headers);
    h.set("Authorization", `Bearer ${token}`);
    if (!h.has("Accept")) h.set("Accept", "application/json");

    // Hvis vi sender body, sørg for content-type
    let fetchBody = undefined;
    if (body !== undefined && body !== null) {
      // body kan være string eller object
      if (typeof body === "string") {
        fetchBody = body;
      } else {
        fetchBody = JSON.stringify(body);
        if (!h.has("Content-Type")) h.set("Content-Type", "application/json");
      }
    }

    const res = await fetch(url, {
      method,
      headers: h,
      body: fetchBody,
    });

    const ct = res.headers.get("content-type") || "";
    const isJson = ct.includes("application/json");

    const outHeaders = {
      "content-type": isJson ? "application/json" : "text/plain; charset=utf-8",
    };

    const text = await res.text();

    return new Response(text, {
      status: res.status,
      headers: outHeaders,
    });
  } catch (e) {
    return new Response(`Proxy error: ${String(e)}`, { status: 500 });
  }
};