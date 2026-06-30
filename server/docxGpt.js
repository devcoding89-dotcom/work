
// server/docxGpt.js
const { paragraphVisibleText, rewriteParagraphTextKeepingStructure, rewriteParagraphTextKeepingStructureWithBreaks, rewriteParagraphTextKeepingStructureWithBreaksAndBolding } = require("./docxSafe");
const { sanitizeBullets } = require("./utils/sanitizeBullets");

// Paragraph regex for DOCX document.xml
const PARA_REGEX = /<w:p[\s\S]*?<\/w:p>/g;

// Prefix helpers (we preserve original bullet/label prefix)
function detectPrefix(originalText) {
  const t = (originalText || "").trimStart();

  // Skills category labels: "Backend: "
  const cat = t.match(/^([A-Za-z &/]+:\s+)(.*)$/);
  if (cat) return { prefix: cat[1], core: (cat[2] || "").trim() };

  // Bullet-like prefixes (o / • / - / * / numbered) should NOT be preserved.
// The DOCX paragraph style controls bullets; preserving markers causes double bullets.
  const bullet = t.match(/^((?:•|o|-|–|\*|\d+[.)])\s+)(.*)$/);
  if (bullet) return { prefix: "", core: (bullet[2] || "").trim() };

  return { prefix: "", core: t.trim() };
}


// If a skills paragraph contains multiple category labels on one line, format it with line breaks.
// Uses common labels but also supports generic "Word: " category patterns.
function formatSkillsMultiline(text) {
  const t = (text || "").toString().replace(/\s+/g, " ").trim();
  if (!t) return t;

  // If it already has obvious separators like " Back End:" etc on one line, normalize to newlines.
  const labels = [
    "Back End:", "Backend:", "Front End:", "Frontend:", "Cloud & DevOps:", "Cloud:", "DevOps:",
    "Data & Testing:", "Testing:", "Data:", "Other:", "Languages:", "Frameworks:", "Tools:"
  ];

  let found = 0;
  for (const lab of labels) {
    const re = new RegExp("\\\\b" + lab.replace(/[.*+?^${}()|[\]\\\\]/g, "\\\\$&"), "i");
    if (re.test(t)) found += 1;
  }

  // Generic "Something: " categories
  const genericCats = (t.match(/[A-Za-z][A-Za-z &/]{1,30}:\s+/g) || []);
  const hasManyCats = genericCats.length >= 2;

  if (found < 2 && !hasManyCats) return t;

  // Insert newline before each category label after the first.
  // Do it by splitting on category tokens, then rejoining with \n.
  // First, create a tokenized version using a regex that matches "Category: "
  const parts = t.split(/(?=[A-Za-z][A-Za-z &/]{1,30}:\s+)/g).map((s) => s.trim()).filter(Boolean);
  if (parts.length <= 1) return t;
  return parts.join("\n");
}
function isCompanyOverviewLine(text) {
  const l = (text || "").trim().toLowerCase();
  return (
    l.startsWith("company summary:") ||
    l.startsWith("company overview:") ||
    l.startsWith("company description:") ||
    l.startsWith("company summary") ||
    l.startsWith("company overview") ||
    l.startsWith("company description")
  );
}

function extractJson(raw) {
  if (!raw) return null;
  let s = raw.trim();

  // strip fenced code blocks
  if (s.startsWith("```")) {
    s = s.replace(/^```[^\n]*\n/, "").replace(/```$/, "").trim();
  }

  // slice around the outermost JSON object/array
  const firstCurly = s.indexOf("{");
  const firstBracket = s.indexOf("[");
  let start = -1;
  if (firstCurly !== -1 && firstBracket !== -1) start = Math.min(firstCurly, firstBracket);
  else start = firstCurly !== -1 ? firstCurly : firstBracket;

  if (start > 0) s = s.slice(start);

  // try parse
  try {
    return JSON.parse(s);
  } catch {
    // last resort: locate last closing bracket/brace
    const lastCurly = s.lastIndexOf("}");
    const lastBracket2 = s.lastIndexOf("]");
    const end = Math.max(lastCurly, lastBracket2);
    if (end !== -1) {
      const ss = s.slice(0, end + 1);
      try { return JSON.parse(ss); } catch { return null; }
    }
    return null;
  }
}

/**
 * GPT pass 1+2 in one call:
 * - Given paragraph list w/ indices and JD
 * - Identify Experience + Skills sections
 * - Produce edits: [{ index, newText }]
 *
 * newText MUST NOT include job titles, dates, headings, or company overview lines.
 * It must keep the same bullet count/shape by only editing existing lines.
 */
async function getEditsFromGPT(openai, paragraphs, jdText) {
  // Keep prompt size under control: include up to 220 paragraphs, truncate each line
  const maxParas = 220;
  const view = paragraphs.slice(0, maxParas).map((p, i) => {
    const t = (p.text || "").replace(/\s+/g, " ").trim();
    const clipped = t.length > 240 ? t.slice(0, 240) + "…" : t;
    return `${i}: ${clipped}`;
  }).join("\n");

  const prompt = `
You are an expert technical resume editor.

Task:
1) Identify ONLY the paragraphs that are:
   - responsibility bullet lines inside the Experience section (including fake bullets like "o " or "-" or "•")
   - skills lines inside the Skills section (including "Category: value" lines)
2) Rewrite ONLY those lines to better match the job description.

Hard rules:
- Do NOT change Objective/Summary/Profile.
- Do NOT change job titles, company names, date/location lines, headings, education, projects, certifications.
- Do NOT change company overview/description lines (Company Summary/Overview/Description), even if they look like bullets.
- Do NOT add or remove lines. Only rewrite existing ones.
- Keep content truthful: no invented experience.
- For Experience responsibility lines: return ONLY the sentence text. Do NOT include any bullet characters (•, -, –, o, *), numbering (1., 1)), or list formatting.
- Output MUST be valid JSON only (no markdown).

Output schema:
{
  "edits": [
    { "index": <paragraph_index_integer>, "text": "<rewritten line text without changing its bullet marker/label>" }
  ]
}

Notes:
- The "index" refers to the paragraph number shown in the paragraph list.
- The "text" should be the rewritten line's content INCLUDING any category label like "Backend: " if it is part of the same line.
- Do NOT include company overview lines in edits.

JOB DESCRIPTION:
${jdText}

PARAGRAPHS:
${view}

JSON ONLY:
`;

  const response = await openai.responses.create({
    model: "gpt-4.1",
    input: prompt,
  });

  const raw =
    response.output_text ||
    (response.output &&
      response.output[0] &&
      response.output[0].content &&
      response.output[0].content[0] &&
      response.output[0].content[0].text) ||
    "";

  const parsed = extractJson(raw);
  if (!parsed || !parsed.edits || !Array.isArray(parsed.edits)) return null;

  // Normalize edits
  const cleaned = [];
  const seen = new Set();
  for (const e of parsed.edits) {
    if (!e || typeof e.index !== "number" || typeof e.text !== "string") continue;
    const idx = e.index;
    if (idx < 0 || idx >= paragraphs.length) continue;
    if (seen.has(idx)) continue;
    const text = e.text.trim();
    if (!text) continue;

    // extra protection: never allow overview lines
    if (isCompanyOverviewLine(paragraphs[idx].text) || isCompanyOverviewLine(text)) continue;

    cleaned.push({ index: idx, text });
    seen.add(idx);
  }

  return cleaned;
}

async function tailorDocxWithGPTExtraction({ zip, documentXml, jdText, openai }) {
  // Build paragraph list (xml + visible text)
  const paraXmls = documentXml.match(PARA_REGEX) || [];
  const paragraphs = paraXmls.map((px) => ({ xml: px, text: paragraphVisibleText(px) }));


  // Build a bold-term dictionary from the Skills section (used to bold important skill words inside EXPERIENCE bullets only).
  function extractBoldTermsFromSkills(paragraphs) {
    const texts = paragraphs.map(p => (p.text || "").replace(/\s+/g, " ").trim());
    const skillsIdx = texts.findIndex(t => /^skills\b/i.test(t));
    if (skillsIdx === -1) return [];
    const terms = [];
    for (let i = skillsIdx + 1; i < texts.length; i++) {
      const t = texts[i];
      if (!t) break;
      // Stop if we hit a likely new section header
      if (/^(education|projects|certifications|professional experience|experience)\b/i.test(t)) break;
      // Split on commas; also split category lines "X: a, b"
      const cleaned = t.replace(/\([^)]*\)/g, " ");
      const parts = cleaned.split(/[,•]/g).map(s => s.trim()).filter(Boolean);
      for (const p of parts) {
        const core = p.replace(/^([A-Za-z &/]{1,30}:)\s*/,"").trim();
        if (core.length >= 2 && core.length <= 40) terms.push(core);
      }
    }
    // Normalize and de-dupe
    return [...new Set(terms)]
      .map(s => s.replace(/\s+/g," ").trim())
      .filter(s => s.length >= 2)
      .slice(0, 250);
  }

  const boldTerms = extractBoldTermsFromSkills(paragraphs);


  // Ask GPT for which paragraph indices to rewrite + new text
  const edits = await getEditsFromGPT(openai, paragraphs, jdText);
  if (!edits || edits.length === 0) {
    // nothing to change / GPT failed -> return original zip unchanged
    return null;
  }

  // Create map for fast lookup
  const editMap = new Map();
  for (const e of edits) editMap.set(e.index, e.text);

  // Apply edits safely:
  // - preserve original bullet/label prefix by using original paragraph's prefix
  // - for Skills category lines: preserve "Category: " prefix from original if present
  let i = -1;
  const newXml = documentXml.replace(PARA_REGEX, (pXml) => {
    i += 1;
    let newText = sanitizeBullets(editMap.get(i));
    if (!newText) return pXml;

    const originalText = paragraphs[i].text;
    const o = detectPrefix(originalText);
    const n = detectPrefix(newText);

    // If original has a prefix, keep it and rewrite only core. If it's a category label, keep it.
    // If GPT already included the label, prefer GPT's label only if it matches the original label (case-insensitive).
    let combined = newText.trim();

    if (o.prefix) {
      if (o.prefix.match(/:\s+$/)) {
        // category label
        const origLabel = o.prefix.trim().toLowerCase();
        const gptLabel = (n.prefix || "").trim().toLowerCase();
        const core = (n.prefix ? n.core : newText).trim();
        if (gptLabel && gptLabel === origLabel) combined = `${o.prefix}${core}`.trim();
        else combined = `${o.prefix}${core}`.trim();
      } else {
        // bullet marker
        const core = (n.prefix ? n.core : newText).trim();
        combined = `${o.prefix}${core}`.trim();
      }
    }

    
// Preserve Skills formatting: if the original skills line had multiple categories in one paragraph,
// convert to multi-line using <w:br/> so Word renders line breaks.
let finalText = combined;
if (o.prefix && o.prefix.match(/:\s+$/)) {
  // category-style skills line
  finalText = formatSkillsMultiline(finalText);
} else if (formatSkillsMultiline(originalText) !== originalText) {
  // original had multiple categories on one line; keep it multi-line
  finalText = formatSkillsMultiline(finalText);
}

    // Bold important skill words inside EXPERIENCE bullet lines only (skills section itself is not bolded).
    const isBulletLine = !!(o.prefix && !o.prefix.match(/:\s+$/));
    if (isBulletLine && boldTerms && boldTerms.length) {
      return rewriteParagraphTextKeepingStructureWithBreaksAndBolding(pXml, finalText, boldTerms);
    }
    return rewriteParagraphTextKeepingStructureWithBreaks(pXml, finalText);

  });

  zip.file("word/document.xml", newXml);
  return await zip.generateAsync({ type: "nodebuffer" });
}

module.exports = { tailorDocxWithGPTExtraction };
