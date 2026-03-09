exports.handler = async function (event) {
  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      body: "Use POST"
    };
  }

  try {
    const data = JSON.parse(event.body || "{}");
    const { action, token, projectId, projectLocation } = data || {};

    if (!token) {
      return {
        statusCode: 400,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ok: false, error: "Missing token" })
      };
    }

    if (action !== "debugProjectAccess") {
      return {
        statusCode: 400,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ok: false, error: "Unknown action" })
      };
    }

    // Midlertidig diagnose:
    // 1) bekreft at functionen kan gjøre et eksternt kall
    // 2) returner nok info til å se hva som feiler
    const testUrl = "https://developer.trimble.com/docs/connect/concepts/";

    const testResponse = await fetch(testUrl, {
      method: "GET",
      headers: {
        Accept: "text/html"
      }
    });

    const testText = await testResponse.text();

    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        ok: true,
        message: "Netlify outbound fetch virker",
        projectId,
        projectLocation,
        fetchTest: {
          url: testUrl,
          status: testResponse.status,
          ok: testResponse.ok,
          preview: testText.slice(0, 200)
        },
        note: "Neste steg er å erstatte hardkodet prosjekt-URL med region discovery + riktig regionalt Connect-endepunkt."
      })
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        ok: false,
        error: String(err),
        message: err?.message || null,
        cause: err?.cause ? String(err.cause) : null,
        stack: err?.stack || null
      })
    };
  }
};
