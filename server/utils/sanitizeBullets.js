/**
 * Removes leading bullet characters from text.
 * Ensures DOCX paragraph styles control bullets.
 */
function sanitizeBullets(text) {
  if (!text) return text;
  return text.replace(/^\s*(•|-|–|\*|\d+[.)])\s*/g, '');
}

module.exports = { sanitizeBullets };
