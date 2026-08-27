// A single generic key-value endpoint backed by Upstash Redis (via Vercel's
// Marketplace Redis integration), so every device that opens the deployed
// site reads and writes the same shared data instead of each browser's own
// local storage.
//
// Requires a Redis database connected via Vercel Marketplace (Vercel
// dashboard → Storage tab → Marketplace Database Providers → Upstash →
// install/connect to this project). That automatically injects the
// KV_REST_API_URL / KV_REST_API_TOKEN environment variables Redis.fromEnv()
// reads below — nothing to configure manually.
import { Redis } from "@upstash/redis";

const kv = Redis.fromEnv();

export default async function handler(req, res) {
  const { key } = req.query;
  if (!key || typeof key !== "string") {
    return res.status(400).json({ error: "Missing 'key' query parameter" });
  }

  try {
    if (req.method === "GET") {
      const value = await kv.get(key);
      if (value === null || value === undefined) {
        return res.status(404).json({ error: "not found" });
      }
      // Redis.fromEnv() may auto-parse JSON-looking strings back into objects on
      // read. App.jsx always expects a plain string it can JSON.parse itself,
      // so normalize back to a string here regardless of what KV handed back.
      const stringValue = typeof value === "string" ? value : JSON.stringify(value);
      return res.status(200).json({ key, value: stringValue });
    }

    if (req.method === "POST") {
      const { value } = req.body || {};
      if (typeof value !== "string") {
        return res.status(400).json({ error: "Body must be { value: <string> }" });
      }
      await kv.set(key, value);
      return res.status(200).json({ key, value });
    }

    if (req.method === "DELETE") {
      await kv.del(key);
      return res.status(200).json({ key, deleted: true });
    }

    res.setHeader("Allow", "GET, POST, DELETE");
    return res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.error("KV error:", err);
    return res.status(500).json({ error: "Storage backend error. Is a Redis database connected via Vercel Marketplace (Upstash)?" });
  }
}
