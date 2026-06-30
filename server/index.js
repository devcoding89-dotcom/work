
const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const OpenAI = require("openai");
require("dotenv").config();

const { tailorDocxWithGPTExtraction } = require("./docxGpt");

const app = express();
const PORT = process.env.PORT || 4000;

if (!process.env.OPENAI_API_KEY) {
  console.error("Missing OPENAI_API_KEY");
  process.exit(1);
}

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

const upload = multer({
  dest: path.join(__dirname, "uploads"),
  limits: { fileSize: 10 * 1024 * 1024 },
});

app.use((req, res, next) => {
  const allowedOrigins = [
    "http://localhost:3000",
    "https://resume-tailor.vercel.app",
    "https://*.vercel.app"
  ];
  const origin = req.headers.origin;
  if (allowedOrigins.includes(origin) || origin?.includes("vercel.app")) {
    res.header("Access-Control-Allow-Origin", origin || "*");
  }
  res.header("Access-Control-Allow-Headers", "Content-Type");
  res.header("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  if (req.method === "OPTIONS") return res.sendStatus(200);
  next();
});

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.get("/health", (_req, res) => res.json({ status: "ok" }));

function safeFilename(name) {
  return (name || "modified").replace(/[^\w\- .]/g, "_");
}

app.post("/api/modify-resume", upload.single("resume"), async (req, res) => {
  let uploadedPath = null;

  try {
    if (!req.file) return res.status(400).send("No file uploaded.");
    uploadedPath = req.file.path;

    const jdText = (req.body.jd || "").toString();
    const companyName = ((req.body.company || "") + "").trim();

    const originalBuffer = fs.readFileSync(uploadedPath);
    const zip = await JSZip.loadAsync(originalBuffer);

    const docFile = zip.file("word/document.xml");
    if (!docFile) {
      const profileName = path.basename(req.file.originalname, path.extname(req.file.originalname));
      const fileName = safeFilename(`${profileName} - ${companyName || "modified"}.docx`);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
      return res.send(originalBuffer);
    }

    const documentXml = await docFile.async("string");

    // Main: gpt-4.1 extraction + rewrite
    const outputBuffer = await tailorDocxWithGPTExtraction({
      zip,
      documentXml,
      jdText,
      openai,
    });

    const profileName = path.basename(req.file.originalname, path.extname(req.file.originalname));
    const fileName = safeFilename(`${profileName} - ${companyName || "modified"}.docx`);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    res.send(outputBuffer || originalBuffer);
  } catch (e) {
    console.error("Error tailoring DOCX:", e);
    res.status(500).send(`Error tailoring DOCX: ${e.message || e}`);
  } finally {
    if (uploadedPath) fs.unlink(uploadedPath, () => {});
  }
});

app.listen(PORT, () => console.log("Server running on port", PORT));

module.exports = app;
