// Grades a participant's written reflection against weighted criteria using
// Google's Gemini API. Requires GEMINI_API_KEY to be set in this project's
// Vercel environment variables (Settings → Environment Variables) — this
// route reads it from process.env and never exposes it to the browser.
export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "GEMINI_API_KEY is not set in this project's environment variables." });
  }

  const { reflection, criteria, assessmentName } = req.body || {};
  if (!reflection || typeof reflection !== "string" || !reflection.trim()) {
    return res.status(400).json({ error: "Body must include non-empty 'reflection' text." });
  }
  if (!Array.isArray(criteria) || criteria.length === 0) {
    return res.status(400).json({ error: "Body must include a 'criteria' array of { name, weight }." });
  }

  const criteriaList = criteria.map((c) => `- ${c.name} (weight: ${c.weight}%)`).join("\n");

  const prompt = `You are grading a participant's written reflection for a professional development program${assessmentName ? ` called "${assessmentName}"` : ""}.

Grade the reflection below against each of the following criteria, scoring each from 0-100 based only on the actual content of the reflection:
${criteriaList}

For each criterion, give an integer score from 0-100 and 1-2 sentences of specific, constructive feedback that references the actual content of the reflection (not generic praise or generic criticism).

Reflection:
"""
${reflection.trim()}
"""

Respond with ONLY valid JSON, no markdown formatting or code fences, in exactly this shape:
{
  "criteria": [
    { "name": "<criterion name exactly as given above>", "score": <integer 0-100>, "feedback": "<1-2 sentences>" }
  ]
}`;

  try {
    // "gemini-flash-latest" is Google's auto-updating alias for their
    // current stable Flash model — pointing at a specific version string
    // (e.g. "gemini-2.5-flash") risks it being retired within months, since
    // Google has been cycling model generations quickly. This alias gets
    // hot-swapped to the current model automatically, with 2 weeks' notice
    // before any change, so this shouldn't need updating again.
    const model = "gemini-flash-latest";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
    const geminiRes = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey,
      },
      body: JSON.stringify({
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.3,
          responseMimeType: "application/json",
        },
      }),
    });

    const data = await geminiRes.json();
    if (!geminiRes.ok) {
      const message = (data && data.error && data.error.message) || "Gemini API error";
      return res.status(geminiRes.status).json({ error: message });
    }

    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) {
      return res.status(500).json({ error: "Gemini returned no content — the reflection may have been blocked by safety filters, or the response was empty." });
    }

    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      // Fallback in case the model wrapped the JSON in code fences despite
      // being asked not to.
      const cleaned = text.replace(/```json|```/g, "").trim();
      try {
        parsed = JSON.parse(cleaned);
      } catch (e2) {
        return res.status(500).json({ error: "Gemini's response wasn't valid JSON. Try again." });
      }
    }

    if (!parsed || !Array.isArray(parsed.criteria)) {
      return res.status(500).json({ error: "Unexpected response format from Gemini." });
    }

    // Clamp scores defensively in case the model returns something out of range.
    const cleanCriteria = parsed.criteria.map((c) => ({
      name: String(c.name || ""),
      score: Math.max(0, Math.min(100, Math.round(Number(c.score) || 0))),
      feedback: String(c.feedback || ""),
    }));

    return res.status(200).json({ criteria: cleanCriteria });
  } catch (err) {
    console.error("Gemini error:", err);
    return res.status(500).json({ error: "Failed to grade reflection." });
  }
}
