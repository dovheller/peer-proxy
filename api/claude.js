// PEER / Esti Heller — Anthropic API proxy
// משרת גם את פיקסל (claude-haiku-4-5) וגם את הכלי של Esti (claude-sonnet-4-5 + tools).
// המפתח נשמר אך ורק כ-Environment Variable ב-Vercel — לעולם לא בקוד ולא בדפדפן.

export default async function handler(req, res) {
  // --- CORS ---
  // להתחלה אפשר '*' לבדיקות. לפני שימוש קבוע — צמצם לדומיינים שלך בלבד.
  const ALLOWED_ORIGINS = [
    "https://dovheller.github.io",
    "http://localhost:3000",
    "http://127.0.0.1:5500", // Live Server נפוץ
    "null"                    // פתיחת קובץ HTML מקומי (file://)
  ];
  const origin = req.headers.origin || "";
  const allowOrigin = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];

  res.setHeader("Access-Control-Allow-Origin", allowOrigin);
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  // --- קריאת הפרמטרים מהבקשה ---
  const { model, messages, system, tools, tool_choice, max_tokens } = req.body || {};

  // --- ולידציה ---
  if (!Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: "Invalid or empty messages array" });
  }
  if (messages.length > 40) {
    return res.status(400).json({ error: "Conversation too long" });
  }

  // --- בניית ה-payload ל-Anthropic ---
  const payload = {
    model: model || "claude-sonnet-4-5",   // פיקסל שולח haiku; ברירת מחדל sonnet
    max_tokens: max_tokens || 1024,
    messages: messages
  };
  if (system) payload.system = system;
  if (tools) payload.tools = tools;             // Esti משתמש בכלי update_scene
  if (tool_choice) payload.tool_choice = tool_choice;

  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": process.env.ANTHROPIC_API_KEY,   // ← מ-Environment Variable בלבד
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
