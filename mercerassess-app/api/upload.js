// Uploads a resource file (Word, PDF, PowerPoint, etc.) to Vercel Blob
// storage and returns a public URL. Used by the Resources feature so admins
// can upload a file directly instead of needing to host it elsewhere and
// paste in a link.
//
// Requires Vercel Blob storage connected to this project (Vercel dashboard →
// Storage → Marketplace Database Providers → Blob → connect to this
// project). That automatically injects the BLOB_READ_WRITE_TOKEN
// environment variable this file needs — nothing to configure here.
import { put } from "@vercel/blob";

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "25mb",
    },
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { fileName, fileData, mimeType } = req.body || {};
  if (!fileData || typeof fileData !== "string" || !fileName) {
    return res.status(400).json({ error: "Missing file data." });
  }

  try {
    const buffer = Buffer.from(fileData, "base64");
    const safeName = fileName.replace(/[^a-zA-Z0-9._-]+/g, "_");
    const blob = await put(`resources/${Date.now()}-${safeName}`, buffer, {
      access: "public",
      contentType: mimeType || "application/octet-stream",
    });
    return res.status(200).json({ url: blob.url, fileName });
  } catch (err) {
    console.error("File upload error:", err);
    return res.status(500).json({ error: "Failed to upload file. Is Blob storage connected to this project?" });
  }
}
