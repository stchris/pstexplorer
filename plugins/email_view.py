"""
Datasette plugin: email_view

Formats the `messages` table as a readable email viewer.

After exporting an SQLite db with `pstexplorer export my.pst` run:

uvx datasette serve my.db --plugins-dir plugins/

Run the playwright-driven integration test with:

uv run --with datasette --with pytest --with pytest-playwright pytest

Table view:
  - body_text column is styled with pre-wrap so line-breaks are preserved
  - Noisy/binary columns are hidden via CSS

Row view (single message):
  - JavaScript builds an email-style header card (Subject, From, To, CC, Date)
  - Full body text shown in a comfortable reading pane
  - Raw datasette table is hidden
"""

import json
from datasette import hookimpl


_CSS = """
/* ── shared ───────────────────────────────────────────────── */

/* Preserve line-breaks in the raw body_text cell */
td.col-body_text {
    white-space: pre-wrap;
    word-break: break-word;
    font-family: ui-sans-serif, system-ui, sans-serif;
    font-size: 13px;
    line-height: 1.55;
    vertical-align: top;
    min-width: 340px;
    max-width: 560px;
}

th.col-sender, td.col-sender,
th.col-subject, td.col-subject {
    min-width: 140px;
    max-width: 220px;
    word-break: break-word;
    vertical-align: top;
}

/* ── table view: hide noisy/binary columns ────────────────── */

body.table th.col-body_html,  body.table td.col-body_html,
body.table th.col-body_rtf,   body.table td.col-body_rtf,
body.table th.col-folder_id,  body.table td.col-folder_id,
body.table th.col-message_class, body.table td.col-message_class,
body.table th.col-delivery_time, body.table td.col-delivery_time {
    display: none;
}

/* ── table view: cap body preview height ─────────────────── */

body.table td.col-body_text {
    max-height: 260px;
    overflow-y: auto;
    display: block;
}

/* ── row view: email layout ──────────────────────────────── */

.email-viewer {
    margin: 0 0 2rem 0;
    max-width: 820px;
}

.email-envelope {
    background: #f7f8fa;
    border: 1px solid #dde1e7;
    border-radius: 8px;
    padding: 18px 22px;
    margin-bottom: 18px;
    font-size: 14px;
    line-height: 1.5;
}

.email-subject-line {
    font-size: 20px;
    font-weight: 600;
    color: #1a1a1a;
    margin-bottom: 14px;
    word-break: break-word;
}

.email-field {
    display: flex;
    gap: 8px;
    margin-bottom: 4px;
}

.email-field-label {
    font-weight: 600;
    color: #6b7280;
    min-width: 48px;
    flex-shrink: 0;
}

.email-field-value {
    color: #1f2937;
    word-break: break-word;
}

.email-body-pane {
    background: #ffffff;
    border: 1px solid #dde1e7;
    border-radius: 8px;
    padding: 22px 26px;
    white-space: pre-wrap;
    word-break: break-word;
    font-family: ui-sans-serif, system-ui, sans-serif;
    font-size: 14px;
    line-height: 1.75;
    color: #1f2937;
    min-height: 200px;
}

/* ── row view: prev/next navigation ──────────────────────── */

.email-nav {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 14px;
    max-width: 820px;
}

.email-nav a {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 6px 14px;
    border: 1px solid #dde1e7;
    border-radius: 6px;
    background: #f7f8fa;
    color: #2563eb;
    text-decoration: none;
    font-size: 13px;
    font-weight: 500;
    transition: background 0.15s;
}

.email-nav a:hover {
    background: #e8ecf1;
}

.email-nav .disabled {
    color: #9ca3af;
    pointer-events: none;
    border-color: #e5e7eb;
    background: #fafafa;
}

.email-nav-spacer {
    flex: 1;
}
"""

_JS = """
(function () {
    var body = document.body;
    if (!body.classList.contains('table-messages')) return;

    /* Inject stylesheet */
    var style = document.createElement('style');
    style.textContent = CSS_PLACEHOLDER;
    document.head.appendChild(style);

    /* Only continue for the single-row view */
    if (!body.classList.contains('row')) return;

    function cellText(cls) {
        var td = document.querySelector('td.' + cls);
        if (!td) return '';
        return (td.textContent || td.innerText || '').trim();
    }

    function mkDiv(cls) {
        var d = document.createElement('div');
        d.className = cls;
        return d;
    }

    function mkField(label, value) {
        if (!value || value === '\u00a0') return null;
        var row = mkDiv('email-field');
        var lbl = mkDiv('email-field-label');
        lbl.textContent = label;
        var val = mkDiv('email-field-value');
        val.textContent = value;
        row.appendChild(lbl);
        row.appendChild(val);
        return row;
    }

    var subject  = cellText('col-subject');
    var sender   = cellText('col-sender');
    var toRecip  = cellText('col-to_recipients');
    var ccRecip  = cellText('col-cc_recipients');
    var date     = cellText('col-submit_time');
    var bodyText = cellText('col-body_text');

    /* Build envelope */
    var envelope = mkDiv('email-envelope');
    if (subject) {
        var subj = mkDiv('email-subject-line');
        subj.textContent = subject;
        envelope.appendChild(subj);
    }
    [['From', sender], ['To', toRecip], ['CC', ccRecip], ['Date', date]].forEach(function (pair) {
        var el = mkField(pair[0], pair[1]);
        if (el) envelope.appendChild(el);
    });

    /* Build body pane */
    var bodyPane = mkDiv('email-body-pane');
    bodyPane.textContent = bodyText;

    /* Build prev/next navigation */
    var nav = mkDiv('email-nav');
    var prevLink = document.createElement('a');
    prevLink.className = 'disabled';
    prevLink.textContent = '\u2190 Previous';
    prevLink.href = '#';
    var nextLink = document.createElement('a');
    nextLink.className = 'disabled';
    nextLink.textContent = 'Next \u2192';
    nextLink.href = '#';
    var spacer = mkDiv('email-nav-spacer');
    nav.appendChild(prevLink);
    nav.appendChild(spacer);
    nav.appendChild(nextLink);

    /* Resolve prev/next IDs via the JSON API */
    var currentId = cellText('col-id');
    var pathParts = window.location.pathname.split('/');
    var basePath = pathParts.slice(0, -1).join('/');
    var apiBase = basePath + '.json?_shape=array&_col=id&_sort=id';

    fetch(apiBase)
        .then(function (r) { return r.json(); })
        .then(function (rows) {
            var ids = rows.map(function (r) { return String(r.id); });
            var idx = ids.indexOf(currentId);
            if (idx > 0) {
                prevLink.href = basePath + '/' + ids[idx - 1];
                prevLink.className = '';
            }
            if (idx >= 0 && idx < ids.length - 1) {
                nextLink.href = basePath + '/' + ids[idx + 1];
                nextLink.className = '';
            }
        });

    /* Assemble viewer and swap out the datasette table */
    var viewer = mkDiv('email-viewer');
    viewer.appendChild(nav);
    viewer.appendChild(envelope);
    viewer.appendChild(bodyPane);

    var table = document.querySelector('table.rows-and-columns');
    if (table) {
        table.parentNode.insertBefore(viewer, table);
        table.style.display = 'none';
    }
}());
"""


@hookimpl
def extra_body_script(
    template, database, table, columns, view_name, request, datasette
):
    if table != "messages":
        return ""
    return _JS.replace("CSS_PLACEHOLDER", json.dumps(_CSS))
