// Preparation of a .docx before mammoth converts it, and completion of the HTML after.
//
// mammoth drops what it does not understand (equations, comment anchors) and knows nothing
// about the page. word/document.xml is therefore parsed once here: equations and comment
// anchors are replaced by placeholder text runs, which mammoth passes through, and the
// placeholders in its HTML are then replaced by MathML / anchor spans (finishHtml). The page
// size and the comments themselves are read from the package as well.

import JSZip = require('jszip');
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { replaceEquations, insertEquations } from './omml';
import { markComments, insertCommentAnchors, CommentData } from './comments';
import { markChanges, insertChanges, ChangeData } from './trackchanges';
import { markParagraphs, insertParagraphStyles, ParagraphData } from './paragraphs';

export const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Minimal view of the xmldom node API used here (the project has no DOM lib types).
export interface XNode {
    nodeType: number;
    localName: string | null;
    namespaceURI: string | null;
    textContent: string | null;
    parentNode: XNode | null;
    nextSibling: XNode | null;
    previousSibling: XNode | null;
    childNodes: { length: number; [i: number]: XNode };
    getAttributeNS(ns: string, name: string): string | null;
    getAttribute(name: string): string | null;
    hasAttributeNS(ns: string, name: string): boolean;
    hasAttribute(name: string): boolean;
    insertBefore(n: XNode, ref: XNode | null): XNode;
    removeChild(n: XNode): XNode;
    appendChild(n: XNode): XNode;
    replaceChild(n: XNode, old: XNode): XNode;
}
export interface XDoc extends XNode {
    documentElement: XNode;
    createElementNS(ns: string, qname: string): XNode;
    createTextNode(text: string): XNode;
}

export function parseXml(xml: string): XDoc {
    return new DOMParser().parseFromString(xml, 'text/xml') as unknown as XDoc;
}

/** <w:r><w:t>text</w:t></w:r> */
export function placeholderRun(doc: XDoc, text: string): XNode {
    const run = doc.createElementNS(W_NS, 'w:r');
    const t = doc.createElementNS(W_NS, 'w:t');
    t.appendChild(doc.createTextNode(text));
    run.appendChild(t);
    return run;
}

/** Page size and margins in inches (last w:pgSz / w:pgMar = main section). */
export interface PageSize { w: number; h: number; top: number; right: number; bottom: number; left: number }

export interface Prepared {
    buffer: Buffer;                  // docx for mammoth (placeholders in place)
    equations: string[];             // MathML, by placeholder index
    comments: CommentData[];
    changes: ChangeData[];           // tracked changes, by placeholder index
    paragraphs: ParagraphData | null; // paragraph formats, by placeholder index
    page: PageSize | null;
}

export async function prepareDocx(buffer: Buffer): Promise<Prepared> {
    const zip = await JSZip.loadAsync(buffer);
    const part = zip.file('word/document.xml');
    if (!part) {
        return { buffer, equations: [], comments: [], changes: [], paragraphs: null, page: null };
    }
    const xml = await part.async('string');
    const page = pageSize(xml);
    const hasMath = xml.indexOf(M_NS) >= 0;
    const hasComments = xml.indexOf('commentRangeStart') >= 0 || xml.indexOf('commentReference') >= 0;
    const hasMarks = /<w:(ins|del|moveTo|moveFrom|lastRenderedPageBreak)\b|<w:br [^>]*w:type="page"/.test(xml);

    const doc = parseXml(xml);
    let equations: string[] = [];
    let comments: CommentData[] = [];
    let changes: ChangeData[] = [];
    try {
        equations = hasMath ? replaceEquations(doc) : [];   // first: changes inside equations are its own
    } catch (e) {
        console.error('prepare: equations skipped:', e);
    }
    try {
        changes = hasMarks ? markChanges(doc) : [];
    } catch (e) {
        console.error('prepare: tracked changes skipped:', e);
    }
    try {
        comments = hasComments ? await markComments(doc, zip) : [];
    } catch (e) {
        console.error('prepare: comments skipped:', e);
    }
    let paragraphs: ParagraphData | null = null;
    try {                                                   // last: its marker must come first
        paragraphs = markParagraphs(doc, await zip.file('word/styles.xml')?.async('string'));
    } catch (e) {
        console.error('prepare: paragraph formats skipped:', e);
    }
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc as any));
    const out = await zip.generateAsync({ type: 'nodebuffer' });
    return { buffer: out, equations, comments, changes, paragraphs, page };
}

/** mammoth's HTML -> final HTML: MathML and comment anchors in place, the comments as JSON
 *  for the viewer script, and the page size as CSS; wrapped in one .docx-page, which the
 *  viewer script splits into pages. */
export function finishHtml(html: string, p: Prepared): string {
    if (p.paragraphs) {
        html = insertParagraphStyles(html, p.paragraphs);   // first: markers sit right after the tag
    }
    html = insertChanges(insertCommentAnchors(insertEquations(html, p.equations)), p.changes);
    let out = '';
    if (p.comments.length) {
        out += '<script type="application/json" id="docx-comments">'
            + JSON.stringify(p.comments).replace(/</g, '\\u003c') + '</script>';
    }
    if (!p.page) {
        return out + html;
    }
    return out + pageCss(p.page, p.paragraphs?.font ?? null) + '<div class="docx-page">' + html + '</div>';
}

function pageSize(xml: string): PageSize {
    const inches = (tag: string, attr: string, dflt: number): number => {
        const all = [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*\\bw:${attr}="(\\d+)"`, 'g'))];
        return all.length ? Number(all[all.length - 1][1]) / 1440 : dflt;   // twips -> in
    };
    // default: A4 with 1 in margins
    return {
        w: inches('pgSz', 'w', 8.27), h: inches('pgSz', 'h', 11.69),
        top: inches('pgMar', 'top', 1), right: inches('pgMar', 'right', 1),
        bottom: inches('pgMar', 'bottom', 1), left: inches('pgMar', 'left', 1),
    };
}

/** Fixed page width (lines wrap as on the printed page, not with the window), page boxes,
 *  and the comment cards in a column to the right of the pages. */
function pageCss(g: PageSize, font: string | null): string {
    const cm = 3.4;                  // comment column, in
    const family = font ? `font-family: "${font.replace(/["<>]/g, '')}", serif;` : '';
    return `<style>
    .docx-page { ${family} font-size: 10pt; }
    .docx-page h1, .docx-page h2, .docx-page h3, .docx-page h4, .docx-page h5, .docx-page h6 { color: inherit; }
    #document.docx-document { width: ${g.w}in; min-width: ${g.w}in; max-width: none; padding: 0;
        background: transparent !important; box-shadow: none; border-radius: 0; }
    #document.docx-document.has-comments { width: ${g.w + cm}in; min-width: ${g.w + cm}in; }
    .docx-page { position: relative; box-sizing: border-box; width: ${g.w}in; min-height: ${g.h}in;
        padding: ${g.top}in ${g.right}in ${g.bottom}in ${g.left}in; margin: 0 0 16px 0;
        background: var(--document-bg); box-shadow: 0 0 12px var(--viewer-shadow); overflow-wrap: break-word; }
    .cm-card { position: absolute; left: calc(100% + 0.2in); width: ${cm - 0.3}in; box-sizing: border-box;
        padding: 6px 8px; font-size: 12px; line-height: 1.35; background: var(--document-bg); color: var(--viewer-fg);
        border: 1px solid rgba(128,128,128,0.35); border-left: 3px solid #e8a317; border-radius: 4px;
        box-shadow: 0 1px 3px var(--viewer-shadow); cursor: pointer; }
    .cm-card.done { opacity: 0.55; }
    .cm-card.done:not(.active) p, .cm-card.done:not(.active) .cm-reply { display: none; }   /* folded */
    .cm-card.active { border-left-color: #d9480f; box-shadow: 0 0 0 2px rgba(217,72,15,0.45); opacity: 1; }
    .cm-card .cm-head { font-weight: 600; }
    .cm-card .cm-date { font-weight: 400; opacity: 0.65; margin-left: 6px; font-size: 11px; }
    .cm-card .cm-tag { font-weight: 400; font-size: 10px; margin-left: 6px; padding: 0 4px; border-radius: 3px;
        background: rgba(128,128,128,0.2); }
    .cm-card p { margin: 2px 0; }
    .cm-card .cm-reply { margin-top: 6px; padding-top: 6px; border-top: 1px solid rgba(128,128,128,0.25); }
    #document { counter-reset: docx-page; }
    .docx-page { counter-increment: docx-page; }
    .docx-page::after { content: counter(docx-page); position: absolute; left: 0; right: 0;
        bottom: ${Math.min(0.5, g.bottom / 2)}in; text-align: center; font-size: 11px; opacity: 0.55; }
    ins.tc { color: #1a7f37; text-decoration: underline; text-decoration-color: rgba(26,127,55,0.6); }
    del.tc { color: #c62828; text-decoration: line-through; }
    del.tc-img img { outline: 3px solid #c62828; opacity: 0.45; }
    ins.tc-img img { outline: 3px solid #1a7f37; }
    ins.tc-img, del.tc-img { text-decoration: none; }
    #document.tc-final del.tc { display: none; }
    #document.tc-final ins.tc { color: inherit; text-decoration: none; }
    ::highlight(docx-comment) { background-color: rgba(255, 196, 0, 0.30); }
    ::highlight(docx-comment-active) { background-color: rgba(255, 140, 0, 0.55); }
    </style>`;
}
