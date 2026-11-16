// Word equations (OMML, Office Math Markup Language) -> MathML.
//
// mammoth drops <m:oMath> elements, so equations vanish from the rendered document.
// Before mammoth runs, every equation in word/document.xml is converted to MathML and
// replaced by a placeholder text run; after mammoth, the placeholders in the HTML are
// replaced by the MathML. The webview (Chromium) renders MathML natively, so no extra
// script or font is needed.

import { M_NS, W_NS, XNode, XDoc, placeholderRun } from './prepare';

// Private-use characters: never in real text, passed through by mammoth unescaped.
const PH_OPEN = '\uE000MATH';
const PH_CLOSE = '\uE001';
const PH_RE = /\uE000MATH(\d+)\uE001/g;

/** Replace every equation in the parsed document.xml by a placeholder run; returns the
 *  MathML of each, in placeholder order. */
export function replaceEquations(doc: XDoc): string[] {
    // Outermost equations: <m:oMathPara> (display) or <m:oMath> not inside one (inline).
    const found: XNode[] = [];
    const walk = (n: XNode) => {
        for (let i = 0; i < n.childNodes.length; i++) {
            const c = n.childNodes[i];
            if (c.nodeType !== 1) {
                continue;
            }
            if (c.namespaceURI === M_NS && (c.localName === 'oMathPara' || c.localName === 'oMath')) {
                found.push(c);
            } else {
                walk(c);
            }
        }
    };
    walk(doc.documentElement);

    const equations: string[] = [];
    for (const eq of found) {
        let mathml: string;
        try {
            mathml = eq.localName === 'oMathPara'
                ? kids(eq).filter(k => isM(k, 'oMath')).map(m => mathElement(m, true)).join('')
                : mathElement(eq, false);
            if (!/<(mi|mn|mo|mtext)\b/.test(mathml)) {
                mathml = '';                         // only tracked deletions: nothing left to show
            }
        } catch (e) {
            console.error('omml: equation not converted:', e);
            mathml = '<span style="opacity:0.6">[equation]</span>';
        }
        const run = placeholderRun(doc, PH_OPEN + equations.length + PH_CLOSE);
        equations.push(mathml);

        // A run must sit inside a paragraph; wrap it where the equation is block-level.
        const parent = eq.parentNode as XNode;
        let replacement = run;
        if (parent.namespaceURI === W_NS && ['body', 'tc', 'txbxContent'].includes(parent.localName ?? '')) {
            replacement = doc.createElementNS(W_NS, 'w:p');
            replacement.appendChild(run);
        }
        parent.insertBefore(replacement, eq);
        parent.removeChild(eq);
    }
    return equations;
}

/** Put the MathML back in place of the placeholders in mammoth's HTML. */
export function insertEquations(html: string, equations: string[]): string {
    if (equations.length === 0) {
        return html;
    }
    return html.replace(PH_RE, (m, i) => equations[Number(i)] ?? m);
}

// ---------------------------------------------------------------- converter

function mathElement(oMath: XNode, display: boolean): string {
    const attr = display ? ' display="block"' : '';
    return `<math xmlns="http://www.w3.org/1998/Math/MathML"${attr}>${row(oMath)}</math>`;
}

function isM(n: XNode, name: string): boolean {
    return n.nodeType === 1 && n.namespaceURI === M_NS && n.localName === name;
}

function kids(n: XNode): XNode[] {
    const out: XNode[] = [];
    for (let i = 0; i < n.childNodes.length; i++) {
        if (n.childNodes[i].nodeType === 1) {
            out.push(n.childNodes[i]);
        }
    }
    return out;
}

function child(n: XNode | null, name: string): XNode | null {
    return n ? kids(n).find(k => isM(k, name)) ?? null : null;
}

/** m:val of a property element; undefined if the element is absent. */
function prop(pr: XNode | null, name: string): string | undefined {
    const p = child(pr, name);
    if (!p) {
        return undefined;
    }
    if (p.hasAttributeNS(M_NS, 'val')) {
        return p.getAttributeNS(M_NS, 'val') ?? '';
    }
    return p.hasAttribute('m:val') ? p.getAttribute('m:val') ?? '' : '';
}

/** On/off property: present without a value, or with 1/on/true, means on. */
function flag(pr: XNode | null, name: string): boolean {
    const v = prop(pr, name);
    return v !== undefined && (v === '' || v === '1' || v === 'on' || v === 'true');
}

function esc(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** All content children of an element as one <mrow>; property elements (*Pr) skipped. */
function row(n: XNode | null): string {
    if (!n) {
        return '<mrow></mrow>';
    }
    return '<mrow>' + kids(n).map(conv).join('') + '</mrow>';
}

function arg(n: XNode, name: string): string {
    return row(child(n, name));
}

function op(ch: string, attrs = ''): string {
    return `<mo${attrs}>${esc(ch)}</mo>`;
}

function conv(n: XNode): string {
    const name = n.localName ?? '';
    if (name.endsWith('Pr')) {
        return '';                                   // rPr, fPr, ctrlPr, ... (formatting only)
    }
    if (n.namespaceURI === W_NS) {
        if (name === 'r') {                          // plain Word run inside an equation
            const text = kids(n).filter(k => k.localName === 't').map(k => k.textContent ?? '').join('');
            return text ? `<mtext>${esc(text)}</mtext>` : '';
        }
        if (name === 'ins' || name === 'moveTo') {   // tracked insertion: keep (as mammoth does)
            return kids(n).map(conv).join('');
        }
        return '';                                   // tracked deletion (del, moveFrom), other
    }
    if (n.namespaceURI !== M_NS) {
        return '';
    }
    const pr = child(n, name + 'Pr');
    switch (name) {
        case 'r':
            return run(n);
        case 'oMath':
        case 'e': case 'num': case 'den': case 'sub': case 'sup':
        case 'deg': case 'lim': case 'fName':
            return row(n);
        case 'f': {
            const type = prop(pr, 'type');
            const num = arg(n, 'num'), den = arg(n, 'den');
            if (type === 'lin') {
                return `<mrow>${num}<mo>/</mo>${den}</mrow>`;
            }
            if (type === 'noBar') {
                return `<mfrac linethickness="0">${num}${den}</mfrac>`;
            }
            return `<mfrac>${num}${den}</mfrac>`;
        }
        case 'sSup':
            return `<msup>${arg(n, 'e')}${arg(n, 'sup')}</msup>`;
        case 'sSub':
            return `<msub>${arg(n, 'e')}${arg(n, 'sub')}</msub>`;
        case 'sSubSup':
            return `<msubsup>${arg(n, 'e')}${arg(n, 'sub')}${arg(n, 'sup')}</msubsup>`;
        case 'sPre':
            return `<mmultiscripts>${arg(n, 'e')}<mprescripts/>${arg(n, 'sub')}${arg(n, 'sup')}</mmultiscripts>`;
        case 'rad': {
            const deg = child(n, 'deg');
            if (flag(pr, 'degHide') || !deg || kids(deg).every(k => (k.localName ?? '').endsWith('Pr'))) {
                return `<msqrt>${arg(n, 'e')}</msqrt>`;
            }
            return `<mroot>${arg(n, 'e')}${row(deg)}</mroot>`;
        }
        case 'nary': {
            const chr = prop(pr, 'chr') || '∫';
            const integral = '∫∬∭∮∯∰∱∲∳'.includes(chr);
            const loc = prop(pr, 'limLoc') ?? (integral ? 'subSup' : 'undOvr');
            const hideSub = flag(pr, 'subHide'), hideSup = flag(pr, 'supHide');
            const o = op(chr, ' largeop="true"');
            const sub = arg(n, 'sub'), sup = arg(n, 'sup');
            let base: string;
            if (hideSub && hideSup) {
                base = o;
            } else if (loc === 'undOvr') {
                base = hideSub ? `<mover>${o}${sup}</mover>`
                    : hideSup ? `<munder>${o}${sub}</munder>`
                    : `<munderover>${o}${sub}${sup}</munderover>`;
            } else {
                base = hideSub ? `<msup>${o}${sup}</msup>`
                    : hideSup ? `<msub>${o}${sub}</msub>`
                    : `<msubsup>${o}${sub}${sup}</msubsup>`;
            }
            return `<mrow>${base}${arg(n, 'e')}</mrow>`;
        }
        case 'd': {
            const beg = prop(pr, 'begChr') ?? '(';
            const end = prop(pr, 'endChr') ?? ')';
            const sep = prop(pr, 'sepChr') ?? '|';
            const fence = ' fence="true" stretchy="true"';
            const parts = kids(n).filter(k => isM(k, 'e')).map(row);
            return '<mrow>' + (beg ? op(beg, fence) : '')
                + parts.join(sep ? op(sep, ' separator="true"') : '')
                + (end ? op(end, fence) : '') + '</mrow>';
        }
        case 'func':
            return `<mrow>${arg(n, 'fName')}<mo>⁡</mo>${arg(n, 'e')}</mrow>`;
        case 'acc': {
            const chr = ACCENTS[prop(pr, 'chr') ?? '̂'] ?? prop(pr, 'chr') ?? 'ˆ';
            return `<mover accent="true">${arg(n, 'e')}${op(chr, ' stretchy="true"')}</mover>`;
        }
        case 'bar':
            return prop(pr, 'pos') === 'top'
                ? `<mover accent="true">${arg(n, 'e')}${op('¯', ' stretchy="true"')}</mover>`
                : `<munder accentunder="true">${arg(n, 'e')}${op('_', ' stretchy="true"')}</munder>`;
        case 'groupChr': {
            const chr = prop(pr, 'chr') || '⏟';
            return prop(pr, 'pos') === 'top'
                ? `<mover>${arg(n, 'e')}${op(chr, ' stretchy="true"')}</mover>`
                : `<munder>${arg(n, 'e')}${op(chr, ' stretchy="true"')}</munder>`;
        }
        case 'limLow':
            return `<munder>${arg(n, 'e')}${arg(n, 'lim')}</munder>`;
        case 'limUpp':
            return `<mover>${arg(n, 'e')}${arg(n, 'lim')}</mover>`;
        case 'borderBox':
            return `<menclose notation="box">${arg(n, 'e')}</menclose>`;
        case 'box':
            return arg(n, 'e');
        case 'phant': {                              // show is on by default
            const show = prop(pr, 'show');
            return show === '0' || show === 'off' || show === 'false'
                ? `<mphantom>${arg(n, 'e')}</mphantom>` : arg(n, 'e');
        }
        case 'eqArr':                                // equation array: one row per m:e
            return '<mtable displaystyle="true">'
                + kids(n).filter(k => isM(k, 'e')).map(alignedRow).join('')
                + '</mtable>';
        case 'm':
            return '<mrow><mtable>' + kids(n).filter(k => isM(k, 'mr')).map(r =>
                '<mtr>' + kids(r).filter(k => isM(k, 'e')).map(e => `<mtd>${row(e)}</mtd>`).join('') + '</mtr>'
            ).join('') + '</mtable></mrow>';
        default:
            return row(n);                           // unknown wrapper: keep its content
    }
}

/** Row of an equation array. Word marks alignment points with '&' in the text: cells split
 *  there, aligned right / left alternately (so "a &= b" lines up at '='). */
function alignedRow(e: XNode): string {
    const inner = kids(e).map(conv).join('');
    const cells = inner.split(/<mo[^>]*>&amp;<\/mo>/);
    const balanced = cells.every(c => (c.match(/<mrow>/g) ?? []).length === (c.match(/<\/mrow>/g) ?? []).length);
    if (cells.length === 1 || !balanced) {           // no '&', or '&' nested deeper: drop it
        return `<mtr><mtd style="text-align:left"><mrow>${cells.join('')}</mrow></mtd></mtr>`;
    }
    return '<mtr>' + cells.map((c, i) =>
        `<mtd style="text-align:${i % 2 === 0 ? 'right' : 'left'};padding:0.1em 0">${c}</mtd>`).join('') + '</mtr>';
}

// Combining accents (as stored by Word) -> spacing characters MathML expects over a base.
const ACCENTS: { [c: string]: string } = {
    '̀': '`', '́': '´', '̂': 'ˆ', '̃': '˜',
    '̄': '¯', '̅': '¯', '̆': '˘', '̇': '˙',
    '̈': '¨', '̌': 'ˇ', '⃗': '→', '⃖': '←',
};

/** One equation run (m:r): text split into numbers, identifiers and operators. */
function run(r: XNode): string {
    const rPr = child(r, 'rPr');
    const text = runText(r);
    if (!text) {
        return '';
    }
    if (child(rPr, 'nor')) {
        return `<mtext>${esc(text)}</mtext>`;        // "normal text" inside an equation
    }
    const sty = prop(rPr, 'sty') ?? 'i';             // p plain, b bold, i italic, bi bold italic
    const scr = prop(rPr, 'scr');                    // double-struck, script, fraktur, ...
    const plain = sty === 'p' || sty === 'b';
    const bold = sty === 'b' || sty === 'bi';
    const style = bold ? ' style="font-weight:bold"' : '';

    let out = '';
    const re = /([0-9]+(?:[.,][0-9]+)*)|(\p{L}+)|(\s+)|(.)/gu;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
        if (m[1]) {
            out += `<mn${style}>${m[1]}</mn>`;
        } else if (m[2]) {
            let letters = m[2];
            if (scr) {
                letters = alphabet(letters, scr);
            }
            if (plain || scr) {                      // upright: one identifier per word
                out += `<mi mathvariant="normal"${style}>${esc(letters)}</mi>`;
            } else {                                 // italic: one identifier per letter
                for (const ch of letters) {
                    out += `<mi${style}>${esc(ch)}</mi>`;
                }
            }
        } else if (m[4]) {
            out += `<mo${style}>${esc(m[4])}</mo>`;
        }                                            // whitespace: MathML spaces itself
    }
    return out;
}

/** Text of a run: m:t children, also inside tracked insertions; tracked deletions dropped. */
function runText(n: XNode): string {
    let s = '';
    for (const k of kids(n)) {
        if (isM(k, 't')) {
            s += k.textContent ?? '';
        } else if (k.namespaceURI === W_NS && (k.localName === 'ins' || k.localName === 'moveTo')) {
            s += runText(k);
        }
    }
    return s;
}

// Letters of the special math alphabets (Chromium ignores mathvariant other than "normal",
// so the Unicode mathematical characters are used instead).
const ALPHA: { [scr: string]: { base: number; holes: { [c: string]: string } } } = {
    'double-struck': { base: 0x1D538, holes: { C: 'ℂ', H: 'ℍ', N: 'ℕ', P: 'ℙ', Q: 'ℚ', R: 'ℝ', Z: 'ℤ' } },
    'script': { base: 0x1D49C, holes: { B: 'ℬ', E: 'ℰ', F: 'ℱ', H: 'ℋ', I: 'ℐ', L: 'ℒ', M: 'ℳ', R: 'ℛ', e: 'ℯ', g: 'ℊ', o: 'ℴ' } },
    'fraktur': { base: 0x1D504, holes: { C: 'ℭ', H: 'ℌ', I: 'ℑ', R: 'ℜ', Z: 'ℨ' } },
};

function alphabet(s: string, scr: string): string {
    const a = ALPHA[scr];
    if (!a) {
        return s;
    }
    let out = '';
    for (const ch of s) {
        const c = ch.charCodeAt(0);
        if (a.holes[ch]) {
            out += a.holes[ch];
        } else if (c >= 65 && c <= 90) {
            out += String.fromCodePoint(a.base + c - 65);
        } else if (c >= 97 && c <= 122) {
            out += String.fromCodePoint(a.base + 26 + c - 97);
        } else {
            out += ch;
        }
    }
    return out;
}
