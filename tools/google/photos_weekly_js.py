"""Page scripts of the weekly Photos album: what ``evaluate`` runs inside Google Photos.

Split from ``photos_weekly_album`` so the browser flow stays small; nothing here runs outside a page.
"""

# Scan the rendered part of a day-grouped grid (library timeline, collection or album). Day
# headers are dated by the first tile after them (their own text is relative: "Today",
# "Wednesday"). Elements get data-pwa-* marks so Playwright can click them with real input.
SCAN_JS = r"""() => {
  const HEAD = '[role=checkbox][aria-label^="Select all photos from "]';
  document.querySelectorAll('[data-pwa-head]').forEach(e => e.removeAttribute('data-pwa-head'));
  const nodes = [...document.querySelectorAll(HEAD + ', a[href*="/photo/"]')];
  const headers = [], tiles = [], matched = new Set();
  let pending = [];
  for (const n of nodes) {
    if (n.tagName !== 'A') { pending.push(n); continue; }
    const label = n.getAttribute('aria-label') || '';
    const id = (n.getAttribute('href').split('/photo/')[1] || '').split(/[?#/]/)[0];
    for (const h of pending) {
      h.setAttribute('data-pwa-head', id);
      headers.push({key: id, label, checked: h.getAttribute('aria-checked') === 'true'});
    }
    pending = [];
    let box = null;
    for (let p = n.parentElement, i = 0; p && i < 5 && !box; p = p.parentElement, i++) {
      box = p.querySelector('[role=checkbox]:not(' + HEAD + ')');
    }
    if (box && box.getAttribute('aria-label') !== label) box = null;
    if (box) { box.setAttribute('data-pwa-tile', id); matched.add(box); }
    tiles.push({id, label, checked: box ? box.getAttribute('aria-checked') === 'true' : null, box: !!box});
  }
  // Item checkboxes with no tile link beside them: an item the day header counts that tiles never show.
  const orphans = [...document.querySelectorAll('[role=checkbox]:not(' + HEAD + ')')]
    .filter(b => !matched.has(b)).map(b => b.getAttribute('aria-label') || '');
  return {headers, tiles, orphans, visible: document.visibilityState === 'visible'};
}"""

# How far the page has scrolled, summed over every scroller: the library grid scrolls inside a c-wiz,
# the screenshots collection inside a plain div and the album list may use the window.
SCROLL_POS_JS = r"""() => [...document.querySelectorAll('*')].reduce((sum, e) => sum + e.scrollTop, window.scrollY)"""

COUNTER_JS = r"""() => {
  for (const e of document.querySelectorAll('div, span')) {
    if (e.childElementCount) continue;
    const m = (e.textContent || '').trim().match(/^(\d+) selected$/);
    if (m && e.offsetParent !== null) return Number(m[1]);
  }
  return 0;
}"""

SCROLL_TOP_JS = r"""(px) => {
  const sc = [...document.querySelectorAll('c-wiz')].sort((a, b) => b.scrollHeight - a.scrollHeight)[0];
  if (sc) sc.scrollTop = px;
}"""

HEAD_NEAR_TOP_JS = r"""(key) => {
  const el = document.querySelector(`[data-pwa-head="${key}"]`);
  const sc = [...document.querySelectorAll('c-wiz')].sort((a, b) => b.scrollHeight - a.scrollHeight)[0];
  if (!el || !sc) return false;
  el.scrollIntoView({block: 'start'});
  sc.scrollTop = Math.max(0, sc.scrollTop - 150);
  return true;
}"""

# A link-shared album's card links to /share/<id>?key=... instead of /album/<id>. Albums others
# shared with the account are /share/ cards too; a title match there refuses creation (fails safe).
ALBUM_CARDS_JS = r"""() => [...document.querySelectorAll('a[href*="/album/"], a[href*="/share/"]')].map(a => ({
  href: a.href, lines: (a.innerText || '').split('\n').map(s => s.trim()).filter(Boolean)}))"""

# The visible dialogs, in DOM order (the Share dialog, then the "Create link to share" confirmation
# stacked over it). Buttons are marked data-pwa-share="<dialog>:<label>" for a real click. ``texts``
# is everywhere a link can show: the text, the inputs and the markup (an existing link is only in
# the markup).
SHARE_DIALOGS_JS = r"""() => {
  const shown = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  document.querySelectorAll('[data-pwa-share]').forEach(e => e.removeAttribute('data-pwa-share'));
  return [...document.querySelectorAll('[role=dialog]')].filter(shown).map((d, i) => ({
    heading: (d.innerText || '').trim().split('\n')[0],
    buttons: [...d.querySelectorAll('button, [role=button]')].filter(shown).map(b => {
      const label = (b.innerText || '').trim() || b.getAttribute('aria-label') || '';
      b.setAttribute('data-pwa-share', i + ':' + label);
      return label;
    }),
    texts: [d.innerText || '', ...[...d.querySelectorAll('input')].map(x => x.value), d.innerHTML]}));
}"""
