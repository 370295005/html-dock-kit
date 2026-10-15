import juice from 'juice/client';
import { load } from 'cheerio';
import { html as htmlNamespace, parse, serialize } from 'parse5';
import { adapter } from 'parse5-htmlparser2-tree-adapter';

interface RuleRange { start: number; open: number; end: number }

// Expose only structural characters. Quotes, comments and CSS escapes cannot
// accidentally delimit a selector list or a block (e.g. content: "};").
function* structural(css: string): Generator<readonly [number, string]> {
  let quote: string | undefined;
  for (let index = 0; index < css.length; index += 1) {
    const character = css[index]!;
    if (character === '\\') { index += 1; continue; }
    if (quote) {
      if (character === quote) quote = undefined;
      continue;
    }
    if (character === '"' || character === "'") { quote = character; continue; }
    if (character === '/' && css[index + 1] === '*') {
      const close = css.indexOf('*/', index + 2);
      if (close < 0) return;
      index = close + 1;
      continue;
    }
    yield [index, character];
  }
}

function topLevelRules(css: string): RuleRange[] {
  const rules: RuleRange[] = [];
  let start: number | undefined;
  let open: number | undefined;
  let parentheses = 0;
  let brackets = 0;
  let braces = 0;
  for (const [index, character] of structural(css)) {
    if (start === undefined) {
      if (/\s/.test(character)) continue;
      start = index;
    }
    if (character === '(') parentheses += 1;
    else if (character === ')') parentheses = Math.max(0, parentheses - 1);
    else if (character === '[') brackets += 1;
    else if (character === ']') brackets = Math.max(0, brackets - 1);
    if (parentheses !== 0 || brackets !== 0) continue;
    if (character === '{') {
      if (open === undefined) open = index;
      braces += 1;
    } else if (character === '}' && open !== undefined) {
      braces -= 1;
      if (braces === 0) {
        rules.push({ start, open, end: index + 1 });
        start = undefined;
        open = undefined;
      }
    } else if (character === ';' && open === undefined) {
      // At-statements such as @import/@charset remain untouched.
      start = undefined;
    }
  }
  return rules;
}

function selectors(header: string): string[] {
  const result: string[] = [];
  let start = 0;
  let parentheses = 0;
  let brackets = 0;
  for (const [index, character] of structural(header)) {
    if (character === '(') parentheses += 1;
    else if (character === ')') parentheses -= 1;
    else if (character === '[') brackets += 1;
    else if (character === ']') brackets -= 1;
    else if (character === ',' && parentheses === 0 && brackets === 0) {
      result.push(header.slice(start, index).trim());
      start = index + 1;
    }
  }
  result.push(header.slice(start).trim());
  return result;
}

function mayPrune(selector: string): boolean {
  // Word relies on simple global stylesheet rules for its default paragraph
  // styles and inherited fonts, even after body styles have been inlined.
  if (/^(?:[A-Za-z][\w-]*|\*)$/.test(selector)) return false;
  // Retain pseudo rules. Some matched pseudo selectors describe state or
  // generated content and therefore cannot be represented by inline styles.
  for (const [, character] of structural(selector)) if (character === ':') return false;
  return true;
}

function staticCss(css: string): string {
  if (/\/\*\s*juice\s+(?:ignore|start\s+ignore|end\s+ignore)\b/.test(css)) return '';
  return topLevelRules(css)
    .filter(rule => !css.slice(rule.start, rule.open).trimStart().startsWith('@'))
    .map(rule => css.slice(rule.start, rule.end)).join('\n');
}

function pruneStaticCss(css: string, inlined: ReadonlySet<string>): string {
  // Respect Juice's explicit author opt-out directives. Keeping this stylesheet
  // is safer than confusing identical selectors across ignored/applied rules.
  if (/\/\*\s*juice\s+(?:ignore|start\s+ignore|end\s+ignore)\b/.test(css)) return css;
  const chunks: string[] = [];
  let cursor = 0;
  for (const rule of topLevelRules(css)) {
    const header = css.slice(rule.start, rule.open);
    if (header.trimStart().startsWith('@')) continue;
    const original = selectors(header);
    const remaining = original.filter(selector => !inlined.has(selector) || !mayPrune(selector));
    if (remaining.length === original.length) continue;
    chunks.push(css.slice(cursor, rule.start));
    if (remaining.length > 0) {
      chunks.push(remaining.join(', '), css.slice(rule.open, rule.end));
    }
    cursor = rule.end;
  }
  chunks.push(css.slice(cursor));
  return chunks.join('');
}

/** Inline static CSS, removing redundant complex rules that confuse Word. */
export function inlineStyles(html: string): string {
  // Build the scoped Cheerio DOM with the same HTML5 parser used for MHTML.
  // Passing a DOM to load avoids reparsing RCDATA/inert text as markup.
  const document = parse(html, { treeAdapter: adapter });
  const $ = load(document, { xml: { xmlMode: false, decodeEntities: true } }, false);
  const blocks: Array<{ element: ReturnType<typeof $>; css: string }> = [];
  $('style').each((_index, node) => {
    const element = $(node);
    const type = element.attr('type')?.trim().toLowerCase();
    const media = element.attr('media')?.trim().toLowerCase();
    let templateContent = false;
    for (let parent = node.parent; parent; parent = parent.parent) {
      if ('name' in parent && parent.name === 'template') templateContent = true;
    }
    if (node.namespace !== htmlNamespace.NS.HTML || element.attr('data-embed') !== undefined || templateContent ||
      (type && type !== 'text/css') || (media && media !== 'all')) return;
    blocks.push({ element, css: element.text() });
  });
  const inlined = juice.inlineDocument($, blocks.map(block => staticCss(block.css)).join('\n'), {
    removeInlinedSelectors: true,
    preserveImportant: true,
    applyWidthAttributes: false,
    applyHeightAttributes: false,
    applyAttributesTableElements: false,
    inlinePseudoElements: false,
    insertPreservedExtraCss: false,
    resolveCSSVariables: true,
  });
  if (inlined) {
    for (const block of blocks) {
      const remaining = pruneStaticCss(block.css, inlined);
      if (remaining.trim()) block.element.text(remaining);
      else block.element.remove();
    }
  }
  return serialize(document, { treeAdapter: adapter });
}
