import { createHmac, timingSafeEqual } from "node:crypto";

function sign(value, secret) {
  return createHmac("sha256", secret).update(value).digest("base64url");
}
function issueAdminToken(secret) {
  const payload = Buffer.from(JSON.stringify({ role: "admin", exp: Date.now() + 4 * 60 * 60 * 1000 })).toString("base64url");
  return payload + "." + sign(payload, secret);
}
function validAdminToken(token, secret) {
  if (!token || typeof token !== "string") return false;
  const [payload, signature] = token.split(".");
  if (!payload || !signature) return false;
  const expected = sign(payload, secret);
  const a = Buffer.from(signature), b = Buffer.from(expected);
  if (a.length !== b.length || !timingSafeEqual(a, b)) return false;
  try {
    const claims = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
    return claims.role === "admin" && Number(claims.exp) > Date.now();
  } catch { return false; }
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ success: false, message: "Method not allowed" });
  }
  const gasUrl = process.env.GAS_WEB_APP_URL;
  const secret = process.env.CBT_API_SECRET;
  if (!gasUrl || !secret) {
    return res.status(500).json({ success: false, message: "Environment API belum diatur di Vercel." });
  }
  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : (req.body || {});
    const action = String(body.action || "");
    const adminActions = new Set(["getExamResultsForAdmin", "saveSingleQuestion"]);
    if (adminActions.has(action) && !validAdminToken(body.adminToken, secret)) {
      return res.status(401).json({ success: false, message: "Sesi admin tidak sah atau kedaluwarsa. Silakan login kembali." });
    }
    const upstream = await fetch(gasUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action, data: body.data || {}, secret })
    });
    const raw = await upstream.text();
    let parsed;
    try { parsed = JSON.parse(raw); }
    catch { throw new Error("Respons GAS bukan JSON. Periksa URL deployment /exec dan izin akses Web App."); }
    if (action === "verifyAdminLogin" && parsed && parsed.success) {
      parsed.adminToken = issueAdminToken(secret);
    }
    return res.status(200).json(parsed);
  } catch (err) {
    return res.status(502).json({ success: false, message: "Gagal menghubungi Google Apps Script: " + err.message });
  }
}
