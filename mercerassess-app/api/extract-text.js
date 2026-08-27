// Extracts plain text from an uploaded .docx or .pdf reflection file,
// server-side (Node has no trouble with PDF/DOCX parsing the way a browser
// does — no worker configuration, no client-side library weight). The
// client sends the file as base64; this never touches any AI model itself,
// it only turns a file into text that then gets sent to /api/grade-reflection.
import mammoth from "mammoth";
import pdfParse from "pdf-parse";

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "10mb",
    },
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { fileName, fileData, mimeType } = req.body || {};
  if (!fileData || typeof fileData !== "string") {
    return res.status(400).json({ error: "Missing file data." });
  }

  const lowerName = (fileName || "").toLowerCase();
  const isPdf = mimeType === "application/pdf" || lowerName.endsWith(".pdf");
  const isDocx =
    mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    lowerName.endsWith(".docx");

  if (!isPdf && !isDocx) {
    return res.status(400).json({ error: "Unsupported file type. Please upload a .docx or .pdf file." });
  }

  try {
    const buffer = Buffer.from(fileData, "base64");
    let text = "";

    if (isPdf) {
      const result = await pdfParse(buffer);
      text = result.text;
    } else {
      const result = await mammoth.extractRawText({ buffer });
      text = result.value;
    }

    text = (text || "").trim();
    if (!text) {
      return res.status(400).json({ error: "Couldn't find any text in this file — it may be empty, scanned/image-based, or corrupted." });
    }

    return res.status(200).json({ text });
  } catch (err) {
    console.error("Text extraction error:", err);
    return res.status(500).json({ error: "Failed to extract text from this file. Try a different file, or paste the text directly instead." });
  }
}
