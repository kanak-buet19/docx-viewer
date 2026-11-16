// Paragraph and run formatting from Word: alignment, spacing before/after (added, not
// collapsed, as in Word), line spacing, indents, font size, bold/italic of the paragraph style;
// per run: colour, highlight, shading, size, font, underline, capitals, hidden text and the
// effects of character styles.
//
// mammoth keeps only the structure (headings, lists, bold/italic set on a run). For each
// paragraph the effective properties are resolved here (direct w:pPr, then the paragraph
// style and the styles it is based on, then the document defaults), turned into CSS, and a
// placeholder run carrying the index of that CSS is put first in the paragraph. After mammoth
// the placeholder becomes the style attribute of the <p>/<hN>/<li> it starts. Runs that differ
// from their paragraph get start/end placeholders, which become <span style="...">...</span>.
//
// Line height: Word's "single" is the font's own line height, the same as CSS "normal"; the
// viewer measures it for the document font and sets --lh (multiples are calc(var(--lh) * m)).

import { W_NS, XNode, XDoc, parseXml, placeholderRun } from './prepare';

const RE_PARA = /<(p|h[1-6]|li)(\s[^>]*)?>(\d+)(<\/\1>)?/g;
const RE_LEFT = /\uE00A\d+\uE00B/g;
const RE_RUN = /\uE00C(\d+)\uE00D|\uE00E/g;

const HIGHLIGHT: { [k: string]: string } = {
    yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff',
    red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
    darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
    lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

interface Style { basedOn: string | null; pPr: XNode | null; rPr: XNode | null }

export interface ParagraphData {
    css: string[];                   // CSS of each distinct paragraph format, by index
    runCss: string[];                // CSS of each distinct run format, by index
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
    const runCss: string[] = [];
    const runIndex = new Map<string, number>();
    const paras = descendants(doc.documentElement, 'p');
    const styleIdOf = (p: XNode) => attr(child(child(p, 'pPr'), 'pStyle'), 'val') || defaultPStyle;
    const sameStyleSibling = (p: XNode, dir: 'previousSibling' | 'nextSibling') => {
        let s = p[dir];
        while (s && s.nodeType !== 1) {
            s = s[dir];
        }
        return !!s && isW(s, 'p') && styleIdOf(s) === styleIdOf(p);
    };
    for (const p of paras) {
        const pPr = child(p, 'pPr');
        const sc = chain(styleIdOf(p));
        const pPrs = [pPr, ...sc.map(s => s.pPr), defPPr];
        const rPrs = [...sc.map(s => s.rPr), defRPr];
        const pv = (el: string, a: string) => first(pPrs, n => attrOrNull(child(n, el), a));
        const rFlag = (el: string) => {                            // bold / italic from the style
            const v = first(rPrs, n => { const e = child(n, el); return e ? (attr(e, 'val') || '1') : null; });
            return v !== null && !['0', 'false', 'off'].includes(v);
        };

        const rules: string[] = ['margin:0'];
        const jc = pv('jc', 'val');
        const align = ({ center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' } as { [k: string]: string })[jc ?? ''];
        rules.push('text-align:' + (align ?? 'left'));
        // spacing as padding: Word adds space after + space before, CSS margins would collapse;
        // none between paragraphs of the same style when the style says so (contextualSpacing)
        const ctx = first(pPrs, n => { const e = child(n, 'contextualSpacing'); return e ? (attr(e, 'val') || '1') : null; });
        const contextual = ctx !== null && !['0', 'false', 'off'].includes(ctx);
        const before = contextual && sameStyleSibling(p, 'previousSibling') ? '0' : pv('spacing', 'before');
        const after = contextual && sameStyleSibling(p, 'nextSibling') ? '0' : pv('spacing', 'after');
        rules.push('padding-top:' + pt(before, 0));
        rules.push('padding-bottom:' + pt(after, 0));
        const line = num(pv('spacing', 'line'));
        const rule = pv('spacing', 'lineRule') ?? 'auto';
        if (line && rule === 'exact') {
            rules.push('line-height:' + line / 20 + 'pt');
        } else if (line && rule === 'atLeast') {
            rules.push('line-height:max(' + line / 20 + 'pt, calc(var(--lh, 1.15) * 1em))');
        } else {
            rules.push('line-height:calc(var(--lh, 1.15) * ' + ((line ?? 240) / 240).toFixed(3) + ')');
        }
        const left = num(pv('ind', 'left') ?? pv('ind', 'start')), right = num(pv('ind', 'right') ?? pv('ind', 'end'));
        const firstLine = num(pv('ind', 'firstLine')), hanging = num(pv('ind', 'hanging'));
        if (left) { rules.push('padding-left:' + left / 20 + 'pt'); }
        if (right) { rules.push('padding-right:' + right / 20 + 'pt'); }
        if (hanging) { rules.push('text-indent:' + (-hanging / 20) + 'pt'); }
        else if (firstLine) { rules.push('text-indent:' + firstLine / 20 + 'pt'); }
        const size = dominantSize(p, rPrs);
        if (size) { rules.push('font-size:' + size / 2 + 'pt'); }
        const pBold = rFlag('b'), pItalic = rFlag('i');
        if (pBold) { rules.push('font-weight:bold'); }
        if (pItalic) { rules.push('font-style:italic'); }

        // runs: direct formatting and character style, where they differ from the paragraph
        for (const r of descendants(p, 'r')) {
            const rPr = child(r, 'rPr');
            if (!rPr) {
                continue;
            }
            const cc = chain(attr(child(rPr, 'rStyle'), 'val') || null);
            const rr = [rPr, ...cc.map(s => s.rPr)];
            const rv = (el: string, a: string) => first(rr, n => attrOrNull(child(n, el), a));
            const on = (el: string) => {
                const v = first(rr, n => { const e = child(n, el); return e ? (attr(e, 'val') || '1') : null; });
                return v === null ? null : !['0', 'false', 'off', 'none'].includes(v);
            };
            const rs: string[] = [];
            const color = rv('color', 'val');
            if (color && color !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(color) && color !== '000000') {
                rs.push('color:#' + color);
            }
            const hl = rv('highlight', 'val');
            const fill = rv('shd', 'fill');
            if (hl && HIGHLIGHT[hl]) {
                rs.push('background-color:' + HIGHLIGHT[hl]);
            } else if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(fill) && fill.toUpperCase() !== 'FFFFFF') {
                rs.push('background-color:#' + fill);
            }
            const sz = num(rv('sz', 'val'));
            if (sz && sz !== size) {
                rs.push('font-size:' + sz / 2 + 'pt');
            }
            const rf = rv('rFonts', 'ascii');
            if (rf && rf !== font) {
                rs.push(`font-family:'${rf.replace(/['"<>]/g, '')}'`);
            }
            const b = on('b'), it = on('i');
            if (b !== null && b !== pBold) { rs.push('font-weight:' + (b ? 'bold' : 'normal')); }
            if (it !== null && it !== pItalic) { rs.push('font-style:' + (it ? 'italic' : 'normal')); }
            const u = rv('u', 'val');
            if (u && u !== 'none') { rs.push('text-decoration:underline'); }
            if (on('caps')) { rs.push('text-transform:uppercase'); }
            if (on('smallCaps')) { rs.push('font-variant:small-caps'); }
            if (on('vanish')) { rs.push('display:none'); }
            if (rs.length === 0) {
                continue;
            }
            const rk = rs.join(';');
            let ri = runIndex.get(rk);
            if (ri === undefined) {
                ri = runCss.length;
                runCss.push(rk);
                runIndex.set(rk, ri);
            }
            const t0 = doc.createElementNS(W_NS, 'w:t');
            t0.appendChild(doc.createTextNode('\uE00C' + ri + '\uE00D'));
            const t1 = doc.createElementNS(W_NS, 'w:t');
            t1.appendChild(doc.createTextNode('\uE00E'));
            r.insertBefore(t0, rPr.nextSibling);
            r.appendChild(t1);
        }

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
    return { css, runCss, font };
}

/** Placeholders -> style attributes; an empty paragraph keeps one line, as in Word. */
export function insertParagraphStyles(html: string, d: ParagraphData): string {
    return html.replace(RE_PARA, (_m, tag: string, attrs: string | undefined, i: string, close: string | undefined) => {
        const css = d.css[Number(i)];
        const open = `<${tag}${attrs ?? ''}${css ? ` style="${css}"` : ''}>`;
        return close ? open + '<br>' + close : open;
    }).replace(RE_LEFT, '').replace(RE_RUN, (_m, i: string | undefined) =>
        i === undefined ? '</span>' : `<span style="${d.runCss[Number(i)] ?? ''}">`);
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
