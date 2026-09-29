/** Approximate layout estimates, ported from bruce-md2word (src/core/layout.ts). */

const WIDE_CHAR = /[\p{Script=Han}\p{Script=Hangul}\p{Script=Hiragana}\p{Script=Katakana}　-￯]/u;

/** Approximate text width in 12pt half-width characters, not browser pixels. */
function textColumns(text) {
  let columns = 0;
  for (const ch of text) columns += WIDE_CHAR.test(ch) ? 2 : 1;
  return columns;
}

/**
 * Column widths in twips that always sum to the available table width.
 * Short identifier/status columns keep a floor; long prose gets a bounded larger share.
 */
function columnWidths(rows, count, total) {
  const weights = Array.from({ length: count }, (_, i) => {
    // reduce, not spread: very long tables exceed the engine's argument limit.
    const longest = rows.reduce((max, row) => Math.max(max, textColumns((row.children[i] && row.children[i].textContent) || '')), 1);
    return Math.max(6, Math.min(40, Math.sqrt(longest) * 3));
  });
  const sum = weights.reduce((a, b) => a + b, 0);
  const floor = Math.min(900, Math.floor(total / count / 2));
  const widths = weights.map(w => Math.floor(floor + (total - floor * count) * w / sum));
  for (let i = 0, remaining = total - widths.reduce((a, b) => a + b, 0); i < remaining; i++) widths[i % count]++;
  return widths;
}

function estimatedLines(text, widthTwips) {
  const columns = Math.max(1, widthTwips / 120);
  return text.split('\n').reduce((lines, line) => lines + Math.max(1, Math.ceil(textColumns(line) / columns)), 0);
}

module.exports = { textColumns, columnWidths, estimatedLines };
