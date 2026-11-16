// Word comments (word/comments.xml, replies and resolved state from commentsExtended.xml).
//
// mammoth ignores comment anchors. Each w:commentRangeStart / w:commentRangeEnd /
// w:commentReference in document.xml is replaced by a placeholder text run; after mammoth
// the placeholders become empty <span class="cm-s|cm-e|cm-r" data-cid="id"> anchors, which
// the viewer script uses to highlight the commented text and place the comment cards.

import JSZip = require('jszip');
import { W_NS, XNode, XDoc, parseXml, placeholderRun } from './prepare';

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

const PH_RE = /([SER])(\d+)/g;

export interface CommentData {
    id: string;
    author: string;
    initials: string;
    date: string;                    // ISO, may be empty
    paras: string[];                 // plain text of each paragraph
    parent: string | null;           // id of the comment this one replies to
    done: boolean;                   // resolved
}

/** Replace the comment anchors in document.xml by placeholder runs and read the comments. */
export async function markComments(doc: XDoc, zip: JSZip): Promise<CommentData[]> {
    const anchors: XNode[] = [];
    const walk = (n: XNode) => {
        for (let i = 0; i < n.childNodes.length; i++) {
            const c = n.childNodes[i];
            if (c.nodeType !== 1) {
                continue;
            }
            if (c.namespaceURI === W_NS && ['commentRangeStart', 'commentRangeEnd', 'commentReference'].includes(c.localName ?? '')) {
                anchors.push(c);
            } else {
                walk(c);
            }
        }
    };
    walk(doc.documentElement);

    for (const a of anchors) {
        const id = attrW(a, 'id');
        const kind = a.localName === 'commentRangeStart' ? 'S' : a.localName === 'commentRangeEnd' ? 'E' : 'R';
        const text = '' + kind + id + '';
        const parent = a.parentNode as XNode;
        if (kind === 'R') {                          // inside a w:r: becomes a w:t of that run
            const t = doc.createElementNS(W_NS, 'w:t');
            t.appendChild(doc.createTextNode(text));
            parent.replaceChild(t, a);
            continue;
        }
        const run = placeholderRun(doc, text);
        if (inParagraph(a)) {
            // Not inside a tracked deletion (mammoth drops those): put it next to the deletion.
            let ref: XNode = a;
            for (let p = a.parentNode; p && !isW(p, 'p'); p = p.parentNode) {
                if (isW(p, 'del') || isW(p, 'moveFrom')) {
                    ref = p;
                }
            }
            if (ref === a) {
                parent.replaceChild(run, a);
            } else {
                const rp = ref.parentNode as XNode;
                rp.insertBefore(run, kind === 'S' ? ref : ref.nextSibling);
                parent.removeChild(a);
            }
        } else {
            // Between paragraphs (or table rows): start of the next / end of the previous paragraph.
            const p = kind === 'S' ? nextParagraph(a) : prevParagraph(a);
            parent.removeChild(a);
            if (p) {
                if (kind === 'S') {
                    const first = firstContent(p);
                    p.insertBefore(run, first);
                } else {
                    p.appendChild(run);
                }
            }
        }
    }
    return readComments(zip);
}

/** Placeholders in mammoth's HTML -> empty anchor spans. */
export function insertCommentAnchors(html: string): string {
    return html.replace(PH_RE, (_m, k, id) =>
        `<span class="cm-${k === 'S' ? 's' : k === 'E' ? 'e' : 'r'}" data-cid="${id}"></span>`);
}

async function readComments(zip: JSZip): Promise<CommentData[]> {
    const xml = await zip.file('word/comments.xml')?.async('string');
    if (!xml) {
        return [];
    }
    // Replies and resolved state: keyed by the w14:paraId of a comment's last paragraph.
    const ext = new Map<string, { parent: string | null; done: boolean }>();
    const exXml = await zip.file('word/commentsExtended.xml')?.async('string');
    if (exXml) {
        for (const e of elements(parseXml(exXml).documentElement, W15_NS, 'commentEx')) {
            ext.set(e.getAttributeNS(W15_NS, 'paraId') ?? '', {
                parent: e.getAttributeNS(W15_NS, 'paraIdParent') || null,
                done: e.getAttributeNS(W15_NS, 'done') === '1',
            });
        }
    }

    const list: (CommentData & { paraId: string })[] = [];
    for (const c of elements(parseXml(xml).documentElement, W_NS, 'comment')) {
        const paras = elements(c, W_NS, 'p');
        const texts = paras.map(p => elements(p, W_NS, 't').map(t => t.textContent ?? '').join(''))
            .filter((t, i, all) => t.trim() !== '' || all.length === 1);
        const last = paras[paras.length - 1];
        list.push({
            id: attrW(c, 'id'),
            author: attrW(c, 'author'),
            initials: attrW(c, 'initials'),
            date: attrW(c, 'date'),
            paras: texts,
            parent: null,
            done: false,
            paraId: last ? last.getAttributeNS(W14_NS, 'paraId') ?? '' : '',
        });
    }
    const idOfPara = new Map(list.map(c => [c.paraId, c.id]));
    return list.map(({ paraId, ...c }) => {
        const e = ext.get(paraId);
        return { ...c, done: e?.done ?? false, parent: e?.parent ? idOfPara.get(e.parent) ?? null : null };
    });
}

// ---------------------------------------------------------------- helpers

function isW(n: XNode, name: string): boolean {
    return n.nodeType === 1 && n.namespaceURI === W_NS && n.localName === name;
}

function attrW(n: XNode, name: string): string {
    return n.getAttributeNS(W_NS, name) ?? n.getAttribute('w:' + name) ?? '';
}

function elements(root: XNode, ns: string, name: string): XNode[] {
    const out: XNode[] = [];
    const walk = (n: XNode) => {
        for (let i = 0; i < n.childNodes.length; i++) {
            const c = n.childNodes[i];
            if (c.nodeType === 1) {
                if (c.namespaceURI === ns && c.localName === name) {
                    out.push(c);
                }
                walk(c);
            }
        }
    };
    walk(root);
    return out;
}

function inParagraph(n: XNode): boolean {
    for (let p = n.parentNode; p; p = p.parentNode) {
        if (isW(p, 'p')) {
            return true;
        }
        if (isW(p, 'body') || isW(p, 'tc') || isW(p, 'txbxContent')) {
            return false;
        }
    }
    return false;
}

/** First w:p at or below n (document order), or null. */
function firstParagraph(n: XNode): XNode | null {
    if (isW(n, 'p')) {
        return n;
    }
    for (let i = 0; i < n.childNodes.length; i++) {
        const p = n.childNodes[i].nodeType === 1 ? firstParagraph(n.childNodes[i]) : null;
        if (p) {
            return p;
        }
    }
    return null;
}

function lastParagraph(n: XNode): XNode | null {
    if (isW(n, 'p')) {
        return n;
    }
    for (let i = n.childNodes.length - 1; i >= 0; i--) {
        const p = n.childNodes[i].nodeType === 1 ? lastParagraph(n.childNodes[i]) : null;
        if (p) {
            return p;
        }
    }
    return null;
}

function nextParagraph(n: XNode): XNode | null {
    for (let a: XNode | null = n; a; a = a.parentNode) {
        for (let s = a.nextSibling; s; s = s.nextSibling) {
            const p = s.nodeType === 1 ? firstParagraph(s) : null;
            if (p) {
                return p;
            }
        }
    }
    return null;
}

function prevParagraph(n: XNode): XNode | null {
    for (let a: XNode | null = n; a; a = a.parentNode) {
        for (let s = a.previousSibling; s; s = s.previousSibling) {
            const p = s.nodeType === 1 ? lastParagraph(s) : null;
            if (p) {
                return p;
            }
        }
    }
    return null;
}

/** First child of a paragraph after its properties (w:pPr), or null. */
function firstContent(p: XNode): XNode | null {
    for (let i = 0; i < p.childNodes.length; i++) {
        const c = p.childNodes[i];
        if (!(c.nodeType === 1 && isW(c, 'pPr'))) {
            return c;
        }
    }
    return null;
}
