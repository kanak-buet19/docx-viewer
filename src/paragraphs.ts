// Paragraph formatting from Word: alignment, spacing before/after, line spacing, indents,
// font size, and bold/italic set by the paragraph style.
//
// mammoth keeps only the structure (headings, lists, bold/italic set on a run). For each
// paragraph the effective properties are resolved here (direct w:pPr, then the paragraph
// style and the styles it is based on, then the document defaults), turned into CSS, and a
// placeholder run carrying the index of that CSS is put first in the paragraph. After mammoth
// the placeholder becomes the style attribute of the <p>/<hN>/<li> it starts.

import { W_NS, XNode, XDoc, parseXml, placeholderRun } from './prepare';

const RE_PARA = /<(p|h[1-6]|li)(\s[^>]*)?>(\d+)(<\/\1>)?/g;
const RE_LEFT = /\d+/g;

interface Style { basedOn: string | null; pPr: XNode | null; rPr: XNode | null }

export interface ParagraphData {
    css: string[];                   // CSS of each distinct paragraph format, by index
    font: string | null;             // document font (docDefaults / default paragraph style)
}

/** Put a format placeholder first in every paragraph of the parsed document.xml. */
export function markParagraphs(doc: XDoc, stylesXml: string | undefined): ParagraphData {
    const styles = new Map<string, Style>();
    let defaultPStyle: string | null = null;
    let defPPr: XNode | null = null, defRPr: XNode | null = null;
    if (stylesXml) {
        const sdoc = parseXml(stylesXml);
        const dd = child(sdoc.documentElement, 'docDefaults');
        defPPr = child(child(dd, 'pPrDefault'), 'pPr');
        defRPr = child(child(dd, 'rPrDefault'), 'rPr');
        for (const s of kids(sdoc.documentElement).filter(k => isW(k, 'style'))) {
            const id = attr(s, 'styleId');
            styles.set(id, { basedOn: attr(child(s, 'basedOn'), 'val') || null, pPr: child(s, 'pPr'), rPr: child(s, 'rPr') });
            if (attr(s, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(attr(s, 'default'))) {
                defaultPStyle = id;
            }
        }
    }
    const chain = (id: string | null): Style[] => {             // style, basedOn, ... (no loops)
        const out: Style[] = [];
        for (let i = 0; id && i < 20; i++) {
            const s = styles.get(id);
            if (!s || out.includes(s)) {
                break;
            }
            out.push(s);
            id = s.basedOn;
        }
        return out;
    };

    const fontOf = (r: XNode | null) => attr(child(r, 'rFonts'), 'ascii') || null;
    const baseChain = chain(defaultPStyle);
    const font = [...baseChain.map(s => fontOf(s.rPr)), fontOf(defRPr)].find(f => f) ?? null;

    const css: string[] = [];
    const index = new Map<string, number>();
    for (const p of descendants(doc.documentElement, 'p')) {
        const pPr = child(p, 'pPr');
        const sc = chain(attr(child(pPr, 'pStyle'), 'val') || defaultPStyle);
        const pPrs = [pPr, ...sc.map(s => s.pPr), defPPr];
        const rPrs = [...sc.map(s => s.rPr), defRPr];
        const pv = (el: string, a: string) => first(pPrs, n => attrOrNull(child(n, el), a));
        const rFlag = (el: string) => {                            // bold / italic from the style
            const v = first(rPrs, n => { const e = child(n, el); return e ? (attr(e, 'val') || '1') : null; });
            return v !== null && !['0', 'false', 'off'].includes(v);
        };

        const rules: string[] = [];
        const jc = pv('jc', 'val');
        const align = ({ center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' } as { [k: string]: string })[jc ?? ''];
        rules.push('text-align:' + (align ?? 'left'));
        rules.push('margin-top:' + pt(pv('spacing', 'before'), 0));
        rules.push('margin-bottom:' + pt(pv('spacing', 'after'), 0));
        const line = num(pv('spacing', 'line'));
        const rule = pv('spacing', 'lineRule') ?? 'auto';
        if (line) {
            rules.push('line-height:' + (rule === 'auto' ? (line / 240 * 1.15).toFixed(3) : line / 20 + 'pt'));
        } else {
            rules.push('line-height:1.15');
        }
        const left = num(pv('ind', 'left') ?? pv('ind', 'start')), right = num(pv('ind', 'right') ?? pv('ind', 'end'));
        const firstLine = num(pv('ind', 'firstLine')), hanging = num(pv('ind', 'hanging'));
        if (left) { rules.push('padding-left:' + left / 20 + 'pt'); }
        if (right) { rules.push('padding-right:' + right / 20 + 'pt'); }
        if (hanging) { rules.push('text-indent:' + (-hanging / 20) + 'pt'); }
        else if (firstLine) { rules.push('text-indent:' + firstLine / 20 + 'pt'); }
        const size = dominantSize(p, rPrs);
        if (size) { rules.push('font-size:' + size / 2 + 'pt'); }
        if (rFlag('b')) { rules.push('font-weight:bold'); }
        if (rFlag('i')) { rules.push('font-style:italic'); }

        const key = rules.join(';');
        let i = index.get(key);
        if (i === undefined) {
            i = css.length;
            css.push(key);
            index.set(key, i);
        }
        let firstContent: XNode | null = null;                      // after w:pPr
        for (let k = 0; k < p.childNodes.length; k++) {
            const c = p.childNodes[k];
            if (!(c.nodeType === 1 && isW(c, 'pPr'))) {
                firstContent = c;
                break;
            }
        }
        p.insertBefore(placeholderRun(doc, '' + i + ''), firstContent);
    }
    return { css, font };
}

/** Placeholders -> style attributes; an empty paragraph keeps one line, as in Word. */
export function insertParagraphStyles(html: string, d: ParagraphData): string {
    return html.replace(RE_PARA, (_m, tag: string, attrs: string | undefined, i: string, close: string | undefined) => {
        const css = d.css[Number(i)];
        const open = `<${tag}${attrs ?? ''}${css ? ` style="${css}"` : ''}>`;
        return close ? open + '<br>' + close : open;
    }).replace(RE_LEFT, '');
}

/** Font size (half-points) carrying most of the paragraph's text; run size, else style. */
function dominantSize(p: XNode, styleRPrs: (XNode | null)[]): number | null {
    const styleSz = num(first(styleRPrs, n => attrOrNull(child(n, 'sz'), 'val')));
    const count = new Map<number, number>();
    for (const r of descendants(p, 'r')) {
        const len = descendants(r, 't').reduce((a, t) => a + (t.textContent ?? '').length, 0);
        if (len === 0) {
            continue;
        }
        const sz = num(attrOrNull(child(child(r, 'rPr'), 'sz'), 'val')) || styleSz || 20;
        count.set(sz, (count.get(sz) ?? 0) + len);
    }
    let best: number | null = styleSz || null, n = 0;
    count.forEach((c, sz) => { if (c > n) { n = c; best = sz; } });
    return best;
}

// ---------------------------------------------------------------- helpers

function first<T>(list: (XNode | null)[], f: (n: XNode) => T | null): T | null {
    for (const n of list) {
        if (n) {
            const v = f(n);
            if (v !== null && v !== undefined) {
                return v;
            }
        }
    }
    return null;
}

function pt(twips: string | null, dflt: number): string {
    const v = num(twips);
    return (v === null ? dflt : v / 20) + 'pt';
}

function num(s: string | null | undefined): number | null {
    if (s === null || s === undefined || s === '') {
        return null;
    }
    const v = Number(s);
    return isFinite(v) ? v : null;
}

function isW(n: XNode, name: string): boolean {
    return n.nodeType === 1 && n.namespaceURI === W_NS && n.localName === name;
}

function kids(n: XNode | null): XNode[] {
    const out: XNode[] = [];
    if (n) {
        for (let i = 0; i < n.childNodes.length; i++) {
            if (n.childNodes[i].nodeType === 1) {
                out.push(n.childNodes[i]);
            }
        }
    }
    return out;
}

function child(n: XNode | null, name: string): XNode | null {
    return kids(n).find(k => isW(k, name)) ?? null;
}

function attrOrNull(n: XNode | null, name: string): string | null {
    if (!n) {
        return null;
    }
    if (n.hasAttributeNS(W_NS, name)) {
        return n.getAttributeNS(W_NS, name);
    }
    return n.hasAttribute('w:' + name) ? n.getAttribute('w:' + name) : null;
}

function attr(n: XNode | null, name: string): string {
    return attrOrNull(n, name) ?? '';
}

function descendants(root: XNode, name: string): XNode[] {
    const out: XNode[] = [];
    const walk = (n: XNode) => {
        for (let i = 0; i < n.childNodes.length; i++) {
            const c = n.childNodes[i];
            if (c.nodeType === 1) {
                if (isW(c, name)) {
                    out.push(c);
                }
                walk(c);
            }
        }
    };
    walk(root);
    return out;
}
