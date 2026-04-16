import katex from "katex";

/**
 * KaTeX math pre-processor for Markdown preview.
 *
 * Strategy: extract $…$ / $$…$$ blocks before marked.parse() runs,
 * replace them with safe placeholders, then restore rendered HTML after.
 * This avoids conflicts with marked's emphasis / code parsing.
 */

const PH = "\x00MATH";
const PH_END = "\x00";

interface MathBlock {
  placeholder: string;
  html: string;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderLatex(latex: string, displayMode: boolean): string {
  try {
    return katex.renderToString(latex, {
      displayMode,
      throwOnError: false,
    });
  } catch {
    return `<code class="msp-math-error">${escapeHtml(latex)}</code>`;
  }
}

/**
 * Replace all math delimiters with placeholder tokens.
 * Returns the stripped text and the list of rendered blocks.
 *
 * Order: block ($$…$$) first so nested $ inside $$ is not mis-detected.
 */
export function extractMath(
  markdown: string,
): { text: string; blocks: MathBlock[] } {
  const blocks: MathBlock[] = [];
  let counter = 0;

  // 1. Block math: $$...$$ (single-line or multi-line)
  let text = markdown.replace(/\$\$([\s\S]+?)\$\$/g, (_m, latex: string) => {
    const html = renderLatex(latex.trim(), true);
    const ph = `${PH}${counter}${PH_END}`;
    blocks.push({ placeholder: ph, html });
    counter++;
    return `\n${ph}\n`; // ensure block-level placement
  });

  // 2. Inline math: $...$ (no linebreaks allowed inside)
  text = text.replace(/\$([^\$\n]+?)\$/g, (_m, latex: string) => {
    const html = renderLatex(latex.trim(), false);
    const ph = `${PH}${counter}${PH_END}`;
    blocks.push({ placeholder: ph, html });
    counter++;
    return ph;
  });

  return { text, blocks };
}

/**
 * Restore rendered math HTML into the marked output.
 */
export function restoreMath(html: string, blocks: MathBlock[]): string {
  let result = html;
  // marked may wrap the placeholder in <p> — that's fine for inline.
  // For block equations wrapped in <p>, we unwrap after replacement.
  for (const block of blocks) {
    result = result.replace(block.placeholder, block.html);
  }
  // Unwrap block equations that ended up inside <p> tags
  result = result.replace(
    /<p>\s*(<span class="katex-display[\s\S]*?<\/span>)\s*<\/p>/g,
    "$1",
  );
  return result;
}
