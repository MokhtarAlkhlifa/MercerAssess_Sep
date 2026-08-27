// Handles attendance check-ins for QR-code-based session sign-in. Uses the
// same Redis list pattern as survey-response.js — atomic rpush/lrange so
// concurrent check-ins from many people scanning at once don't clobber
// each other.
import { Redis } from "@upstash/redis";

const kv = Redis.fromEnv();

export default async function handler(req, res) {
  const { sessionId } = req.query;
  if (!sessionId || typeof sessionId !== "string") {
    return res.status(400).json({ error: "Missing 'sessionId' query parameter" });
  }
  const key = `attendance-${sessionId}`;

  try {
    if (req.method === "GET") {
      const raw = await kv.lrange(key, 0, -1);
      const checkins = (raw || []).map((r) => (typeof r === "string" ? JSON.parse(r) : r));
      return res.status(200).json({ checkins });
    }

    if (req.method === "POST") {
      const { name, employeeId, cohort, dept } = req.body || {};
      if (!name || typeof name !== "string" || !name.trim()) {
        return res.status(400).json({ error: "Body must include a non-empty 'name'" });
      }
      const entry = {
        name: name.trim(),
        employeeId: (employeeId || "").trim(),
        cohort: (cohort || "").trim(),
        dept: (dept || "").trim(),
        checkedInAt: new Date().toISOString(),
      };
      await kv.rpush(key, JSON.stringify(entry));
      return res.status(200).json({ ok: true });
    }

    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.error("KV error:", err);
    return res.status(500).json({ error: "Storage backend error. Is a Redis database connected via Vercel Marketplace (Upstash)?" });
  }
}
