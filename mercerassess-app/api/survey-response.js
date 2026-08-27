// Handles real participant responses to admin-built surveys (the ones shared
// via QR code / link). Uses Redis's native list operations (rpush/lrange)
// rather than a single read-modify-write blob, so concurrent submissions
// from multiple people at once don't clobber each other.
import { Redis } from "@upstash/redis";

const kv = Redis.fromEnv();

export default async function handler(req, res) {
  const { surveyId } = req.query;
  if (!surveyId || typeof surveyId !== "string") {
    return res.status(400).json({ error: "Missing 'surveyId' query parameter" });
  }
  const key = `survey-responses-${surveyId}`;

  try {
    if (req.method === "GET") {
      const raw = await kv.lrange(key, 0, -1);
      const responses = (raw || []).map((r) => (typeof r === "string" ? JSON.parse(r) : r));
      return res.status(200).json({ responses });
    }

    if (req.method === "POST") {
      const { answers, respondentEmployeeId, respondentName } = req.body || {};
      if (!answers || typeof answers !== "object") {
        return res.status(400).json({ error: "Body must be { answers: { [questionId]: value } }" });
      }
      const entry = {
        answers,
        submittedAt: new Date().toISOString(),
        // Present only when submitted by a logged-in participant from within
        // the app (see the Surveys section); the anonymous QR/link flow
        // never sends these, keeping that path fully anonymous.
        respondentEmployeeId: respondentEmployeeId || null,
        respondentName: respondentName || null,
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
