// PEER / Esti Heller — Anthropic API proxy
// המפתח נשמר אך ורק כ-Environment Variable ב-Vercel — לעולם לא בקוד ולא בדפדפן.

export default async function handler(req, res) {
  // CORS — allow all origins (API key is server-side only, so this is safe)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { model, messages, system, tools, tool_choice, max_tokens } = req.body || {};

  if (!Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: "Invalid or empty messages array" });
  }
  if (messages.length > 40) {
    return res.status(400).json({ error: "Conversation too long" });
  }

  const payload = {
    model: model || "claude-sonnet-4-6",
    max_tokens: max_tokens || 4096,
    messages: messages
  };
  if (system) payload.system = system;
  if (tools) payload.tools = tools;
  if (tool_choice) payload.tool_choice = tool_choice;

  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify(payload)
    });

    const data = await r.json();
    return res.status(r.status).json(data);
  } catch (err) {
    console.error("Proxy error:", err);
    return res.status(500).json({ error: "Proxy internal error" });
  }
}
