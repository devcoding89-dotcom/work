
// server/docxSafe.js

function escapeXml(str) {
  return (str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function paragraphVisibleText(paraXml) {
  return paraXml
    .replace(/<w:tab\/>/g, " ")
    .replace(/<w:br\/>/g, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Safe rewrite:
 * - preserve XML structure
 * - only update the text inside <w:t> nodes
 * - first <w:t> gets full text; remaining cleared
 */
function rewriteParagraphTextKeepingStructure(paraXml, newText) {
  const tRegex = /<w:t\b[^>]*>[\s\S]*?<\/w:t>/g;
  const matches = [...paraXml.matchAll(tRegex)];
  if (!matches.length) return paraXml;

  let out = "";
  let last = 0;

  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    const full = m[0];
    const start = m.index;
    const end = start + full.length;

    out += paraXml.slice(last, start);

    const openEnd = full.indexOf(">") + 1;
    const closeStart = full.lastIndexOf("</w:t>");
    const openTag = full.slice(0, openEnd);
    const closeTag = full.slice(closeStart);

    const inner = i === 0 ? escapeXml(newText) : "";
    out += openTag + inner + closeTag;

    last = end;
  }

  out += paraXml.slice(last);
  return out;
}


/**
 * Rewrite paragraph text while preserving paragraph properties, with support for '\n' line breaks AND bolding specific terms.
 * This rebuilds the paragraph's run content using the first run's <w:rPr> as a base, and adds <w:b/> for matched segments.
 * Safety: If the paragraph contains complex elements (hyperlinks/fields), caller should fall back to rewriteParagraphTextKeepingStructureWithBreaks.
 */
function rewriteParagraphTextKeepingStructureWithBreaksAndBolding(paraXml, newText, boldTerms) {
  const text = (newText || "").toString();
  const terms = (boldTerms || []).filter(Boolean);
  if (!terms.length) return rewriteParagraphTextKeepingStructureWithBreaks(paraXml, text);

  // Skip complex paragraphs
  if (/<w:hyperlink\b|<w:fldSimple\b|<w:instrText\b|<w:smartTag\b/i.test(paraXml)) {
    return rewriteParagraphTextKeepingStructureWithBreaks(paraXml, text);
  }

  // Pull paragraph properties (keep as-is)
  const pOpen = paraXml.match(/^<w:p\b[^>]*>/);
  const pClose = "</w:p>";
  const pPr = (paraXml.match(/<w:pPr[\s\S]*?<\/w:pPr>/) || [""])[0];

  // Base run properties from first run, if any
  const firstRPr = (paraXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/) || [""])[0];

  const escape = escapeXml;

  // Build a regex that matches any term (prefer longer terms first)
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const sorted = [...new Set(terms)].sort((a,b)=>b.length-a.length);
  const re = new RegExp(sorted.map(esc).join("|"), "gi");

  // Split by lines, then create segments with optional <w:br/>
  const lines = text.split("\n");
  const runs = [];
  const makeRun = (t, bold) => {
    const needsPreserve = /^\s|\s$/.test(t);
    const tAttrs = needsPreserve ? ' xml:space="preserve"' : "";
    const rPrBold = bold
      ? (firstRPr
          ? (/<w:b\b/i.test(firstRPr) ? firstRPr : firstRPr.replace(/<\/w:rPr>\s*$/i, "<w:b/></w:rPr>"))
          : "<w:rPr><w:b/></w:rPr>")
      : firstRPr;

    return `<w:r>${rPrBold}<w:t${tAttrs}>${escape(t)}</w:t></w:r>`;
  };

  for (let li=0; li<lines.length; li++) {
    const line = lines[li];

    let last = 0;
    for (const m of line.matchAll(re)) {
      const start = m.index ?? 0;
      const end = start + m[0].length;
      if (start > last) runs.push(makeRun(line.slice(last, start), false));
      runs.push(makeRun(line.slice(start, end), true));
      last = end;
    }
    if (last < line.length) runs.push(makeRun(line.slice(last), false));

    // line break between lines
    if (li !== lines.length - 1) {
      runs.push(`<w:r>${firstRPr}<w:br/></w:r>`);
    }
  }

  const newInner = `${pPr}${runs.join("")}`;
  const openTag = pOpen ? pOpen[0] : "<w:p>";
  return `${openTag}${newInner}${pClose}`;
}

module.exports = {
  escapeXml,
  paragraphVisibleText,
  rewriteParagraphTextKeepingStructure,
};


/**
 * Like rewriteParagraphTextKeepingStructure, but supports visible line breaks.
 * If newText contains '\n', it inserts <w:br/> between lines by reusing the first <w:t> open/close tags.
 * Remaining original <w:t> nodes are cleared.
 */
function rewriteParagraphTextKeepingStructureWithBreaks(paraXml, newText) {
  const text = (newText || "").toString();
  if (!text.includes("\n")) {
    return rewriteParagraphTextKeepingStructure(paraXml, text);
  }

  const tRegex = /<w:t\b[^>]*>[\s\S]*?<\/w:t>/g;
  const matches = [...paraXml.matchAll(tRegex)];
  if (!matches.length) return paraXml;

  const first = matches[0][0];
  const openEnd = first.indexOf(">") + 1;
  const closeStart = first.lastIndexOf("</w:t>");
  const openTag = first.slice(0, openEnd);
  const closeTag = first.slice(closeStart);

  const lines = text.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  const rebuilt = lines
    .map((line) => openTag + escapeXml(line) + closeTag)
    .join("<w:br/>");

  // Replace first <w:t> with rebuilt, clear any subsequent <w:t>
  let out = "";
  let last = 0;
  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    const full = m[0];
    const start = m.index;
    const end = start + full.length;

    out += paraXml.slice(last, start);

    if (i === 0) {
      out += rebuilt;
    } else {
      // keep same tags but empty text
      const oe = full.indexOf(">") + 1;
      const cs = full.lastIndexOf("</w:t>");
      const ot = full.slice(0, oe);
      const ct = full.slice(cs);
      out += ot + "" + ct;
    }
    last = end;
  }
  out += paraXml.slice(last);
  return out;
}


/**
 * Rewrite paragraph text while preserving paragraph properties, with support for '\n' line breaks AND bolding specific terms.
 * This rebuilds the paragraph's run content using the first run's <w:rPr> as a base, and adds <w:b/> for matched segments.
 * Safety: If the paragraph contains complex elements (hyperlinks/fields), caller should fall back to rewriteParagraphTextKeepingStructureWithBreaks.
 */
function rewriteParagraphTextKeepingStructureWithBreaksAndBolding(paraXml, newText, boldTerms) {
  const text = (newText || "").toString();
  const terms = (boldTerms || []).filter(Boolean);
  if (!terms.length) return rewriteParagraphTextKeepingStructureWithBreaks(paraXml, text);

  // Skip complex paragraphs
  if (/<w:hyperlink\b|<w:fldSimple\b|<w:instrText\b|<w:smartTag\b/i.test(paraXml)) {
    return rewriteParagraphTextKeepingStructureWithBreaks(paraXml, text);
  }

  // Pull paragraph properties (keep as-is)
  const pOpen = paraXml.match(/^<w:p\b[^>]*>/);
  const pClose = "</w:p>";
  const pPr = (paraXml.match(/<w:pPr[\s\S]*?<\/w:pPr>/) || [""])[0];

  // Base run properties from first run, if any
  const firstRPr = (paraXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/) || [""])[0];

  const escape = escapeXml;

  // Build a regex that matches any term (prefer longer terms first)
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const sorted = [...new Set(terms)].sort((a,b)=>b.length-a.length);
  const re = new RegExp(sorted.map(esc).join("|"), "gi");

  // Split by lines, then create segments with optional <w:br/>
  const lines = text.split("\n");
  const runs = [];
  const makeRun = (t, bold) => {
    const needsPreserve = /^\s|\s$/.test(t);
    const tAttrs = needsPreserve ? ' xml:space="preserve"' : "";
    const rPrBold = bold
      ? (firstRPr
          ? (/<w:b\b/i.test(firstRPr) ? firstRPr : firstRPr.replace(/<\/w:rPr>\s*$/i, "<w:b/></w:rPr>"))
          : "<w:rPr><w:b/></w:rPr>")
      : firstRPr;

    return `<w:r>${rPrBold}<w:t${tAttrs}>${escape(t)}</w:t></w:r>`;
  };

  for (let li=0; li<lines.length; li++) {
    const line = lines[li];

    let last = 0;
    for (const m of line.matchAll(re)) {
      const start = m.index ?? 0;
      const end = start + m[0].length;
      if (start > last) runs.push(makeRun(line.slice(last, start), false));
      runs.push(makeRun(line.slice(start, end), true));
      last = end;
    }
    if (last < line.length) runs.push(makeRun(line.slice(last), false));

    // line break between lines
    if (li !== lines.length - 1) {
      runs.push(`<w:r>${firstRPr}<w:br/></w:r>`);
    }
  }

  const newInner = `${pPr}${runs.join("")}`;
  const openTag = pOpen ? pOpen[0] : "<w:p>";
  return `${openTag}${newInner}${pClose}`;
}

module.exports.rewriteParagraphTextKeepingStructureWithBreaks = rewriteParagraphTextKeepingStructureWithBreaks;
