// Tracked changes (w:ins / w:del, w:moveTo / w:moveFrom) and page breaks.
//
// mammoth shows inserted text as plain text and drops deleted text. Before mammoth, each
// run-level change gets placeholder runs at its start and end, and deletions are turned into
// ordinary text (w:delText -> w:t, the w:del wrapper removed) so mammoth keeps them. After
// mammoth, the text between the placeholders is wrapped in <ins class="tc"> / <del class="tc">
// (text pieces only, so the HTML stays well formed); the viewer shows them as markup or hides
// the deletions ("Final").
//
// Page breaks: Word's own breaks from its last layout (w:lastRenderedPageBreak) and manual
// breaks (w:br w:type="page") become empty <span class="pb" data-k="r|h"> markers for the
// viewer's page splitting.

import { W_NS, XNode, XDoc, placeholderRun } from './prepare';

export interface ChangeData { type: 'ins' | 'del'; author: string; date: string }

// innermost first: no other change may start inside (a deletion inside an insertion)
const RE_CHANGE = /\uE004S(\d+)\uE005((?:(?!\uE004S)[\s\S])*?)\uE004E\1\uE005/g;
const RE_LEFT = /[SE]\d+/g;
const RE_BREAK = /([rh])/g;

/** Mark tracked changes and page breaks in the parsed document.xml. */
export function markChanges(doc: XDoc): ChangeData[] {
    const changes: ChangeData[] = [];
    const found: XNode[] = [];
    const breaks: XNode[] = [];
    const walk = (n: XNode) => {
        for (let i = 0; i < n.childNodes.length; i++) {
            const c = n.childNodes[i];
            if (c.nodeType !== 1) {
                continue;
            }
            if (c.namespaceURI === W_NS) {
                const name = c.localName ?? '';
                if (['ins', 'del', 'moveTo', 'moveFrom'].includes(name) && hasRun(c)) {
                    found.push(c);                   // run-level change (not a paragraph-mark one)
                } else if (name === 'lastRenderedPageBreak'
                    || (name === 'br' && (c.getAttributeNS(W_NS, 'type') ?? c.getAttribute('w:type')) === 'page')) {
                    breaks.push(c);
                }
            }
            walk(c);
        }
    };
    walk(doc.documentElement);

    for (const b of breaks) {                        // inside a w:r: becomes a w:t of that run
        const t = doc.createElementNS(W_NS, 'w:t');
        t.appendChild(doc.createTextNode('' + (b.localName === 'br' ? 'h' : 'r') + ''));
        (b.parentNode as XNode).replaceChild(t, b);
    }

    for (const c of found) {
        const del = c.localName === 'del' || c.localName === 'moveFrom';
        const n = changes.length;
        changes.push({
            type: del ? 'del' : 'ins',
            author: c.getAttributeNS(W_NS, 'author') ?? '',
            date: c.getAttributeNS(W_NS, 'date') ?? '',
        });
        c.insertBefore(placeholderRun(doc, 'S' + n + ''), c.childNodes.length ? c.childNodes[0] : null);
        c.appendChild(placeholderRun(doc, 'E' + n + ''));
        if (del) {
            for (const t of descendants(c, 'delText')) {          // deleted text -> normal text
                const nt = doc.createElementNS(W_NS, 'w:t');
                nt.appendChild(doc.createTextNode(t.textContent ?? ''));
                (t.parentNode as XNode).replaceChild(nt, t);
            }
            const parent = c.parentNode as XNode;                 // unwrap: mammoth drops w:del
            while (c.childNodes.length) {
                parent.insertBefore(c.childNodes[0], c);
            }
            parent.removeChild(c);
        }
    }
    return changes;
}

/** Placeholders in mammoth's HTML -> <ins>/<del> around the text pieces, page-break markers. */
export function insertChanges(html: string, changes: ChangeData[]): string {
    const wrap = (_m: string, id: string, inner: string) => {
        const c = changes[Number(id)];
        if (!c) {
            return inner;
        }
        const title = esc(c.author + (c.date ? ', ' + c.date.slice(0, 10) : ''));
        // wrap only text between tags, and not text inside <math> (own markup) or inside a
        // change already wrapped (a deletion inside an insertion stays a deletion)
        let skip = 0;
        return inner.replace(/(<[^>]*>)|([^<]+)/g, (_t, tag: string, text: string) => {
            if (tag) {
                if (/^<img\b/.test(tag) && skip === 0) {     // images: marked (and hidden) too
                    return `<${c.type} class="tc tc-img" title="${title}">${tag}</${c.type}>`;
                }
                if (/^<math\b|^<(ins|del) class="tc/.test(tag)) { skip++; }
                if (/^<\/math>|^<\/(ins|del)>/.test(tag) && skip > 0) { skip--; }
                return tag;
            }
            return skip > 0 ? text : `<${c.type} class="tc" title="${title}">${text}</${c.type}>`;
        });
    };
    for (let prev = ''; prev !== html;) {            // nested changes: repeat, inner to outer
        prev = html;
        html = html.replace(RE_CHANGE, wrap);
    }
    return html.replace(RE_LEFT, '').replace(RE_BREAK, (_m, k) => `<span class="pb" data-k="${k}"></span>`);
}

function hasRun(n: XNode): boolean {
    for (let i = 0; i < n.childNodes.length; i++) {
        const c = n.childNodes[i];
        if (c.nodeType === 1 && c.namespaceURI === W_NS && ['r', 'hyperlink', 'smartTag', 'fldSimple'].includes(c.localName ?? '')) {
            return true;
        }
    }
    return false;
}

function descendants(root: XNode, name: string): XNode[] {
    const out: XNode[] = [];
    const walk = (n: XNode) => {
        for (let i = 0; i < n.childNodes.length; i++) {
            const c = n.childNodes[i];
            if (c.nodeType === 1) {
                if (c.namespaceURI === W_NS && c.localName === name) {
                    out.push(c);
                }
                walk(c);
            }
        }
    };
    walk(root);
    return out;
}

function esc(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}
