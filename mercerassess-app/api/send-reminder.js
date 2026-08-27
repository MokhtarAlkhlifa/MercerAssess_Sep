// Sends real meeting-reminder emails via Resend. Requires RESEND_API_KEY to
// be set in the Vercel project's environment variables (Settings →
// Environment Variables) — this route reads it from process.env and never
// exposes it to the browser.
//
// If you haven't verified your own sending domain in Resend yet, set
// RESEND_FROM to something like "MercerAssess <onboarding@resend.dev>" (their
// shared test address) so sending still works during setup; switch it to
// your own verified domain once that's ready.
export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "RESEND_API_KEY is not set in this project's environment variables." });
  }

  const { recipients, subject, message } = req.body || {};
  if (!Array.isArray(recipients) || recipients.length === 0) {
    return res.status(400).json({ error: "Body must include a non-empty 'recipients' array of email addresses." });
  }
  if (!subject || !message) {
    return res.status(400).json({ error: "Body must include 'subject' and 'message'." });
  }

  const from = process.env.RESEND_FROM || "MercerAssess <onboarding@resend.dev>";
  const html = String(message).replace(/\n/g, "<br />");

  try {
    // Resend allows up to 50 "to" addresses per call; batch just in case a
    // large audience is selected.
    const batches = [];
    for (let i = 0; i < recipients.length; i += 50) batches.push(recipients.slice(i, i + 50));

    const results = [];
    for (const batch of batches) {
      const resendRes = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          from,
          to: batch,
          subject,
          html,
        }),
      });
      const data = await resendRes.json();
      if (!resendRes.ok) {
        return res.status(resendRes.status).json({ error: data.message || "Resend API error", details: data });
      }
      results.push(data);
    }

    return res.status(200).json({ ok: true, sent: recipients.length, results });
  } catch (err) {
    console.error("Resend error:", err);
    return res.status(500).json({ error: "Failed to send reminder emails." });
  }
}
