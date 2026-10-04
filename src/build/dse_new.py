#!/usr/bin/env python3
"""Parsers for the new DSE site (www.dse.com.bd) — packet 3A03.

Standard library only. Every function takes raw text (a JSON body or an
HTML page, as saved under tests/fixtures/dse_com_bd/) and returns plain
dicts or lists of dicts. Missing values come back as None, never 0.

Shapes handled:
  parse_latest   the /api/live/prices JSON (cols + rows) and the old
                 scroll-page marquee (code, LTP, change, percent)
  parse_archive  the /api/live/data-archive/day-end JSON and day-end
                 archive tables with a Date column
  parse_index    the /api/live/data-archive/market-summary JSON
                 (daily DSEX/DS30) and the /api/live/index-history JSON
                 (a dated point series)
  parse_companies  the /api/live/prices JSON (also the companies list with
                 category/sector) and the companies directory table
  parse_markets  the /api/live/market-statistics JSON (block trades), the
                 markets page (price-limit and marginable tables) and the
                 marginable-securities page
  parse_news     the announcements page notice list (server-rendered
                 title links; the page carries no dates, bodies or codes)
"""
import datetime
import html as _html
import json
import re
from html.parser import HTMLParser

# ── numbers and dates ────────────────────────────────────────────────────

_MONTHS = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def to_num(v):
    """Read one cell as a float. Commas come off, '-' / U+2212 / '' / None
    become None, and '5.5%' reads as 5.5. Never 0 for an empty cell."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '')
    if s.endswith('%'):
        s = s[:-1].strip()
    s = s.replace('\u2212', '-').strip()
    if s in ('', '-', '\u2014', '\u2013', '--', 'N/A', 'NA', '--'):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(s):
    """Read DD-MM-YYYY, DD.MM.YYYY, YYYY-MM-DD or '01 Oct 2026' (also a
    yearless '1 Sep') into an ISO date. Returns None when unparseable."""
    if s is None:
        return None
    t = ' '.join(str(s).split())
    if re.fullmatch(r'\d{4}-\d{1,2}-\d{1,2}', t):
        y, m, d = (int(x) for x in t.split('-'))
    else:
        m3 = re.fullmatch(r'(\d{1,2})\s+([A-Za-z]{3,9})(?:\s+(\d{4}))?', t)
        if m3:
            mon = _MONTHS.get(m3.group(2)[:3].lower())
            if mon is None:
                return None
            d, m = int(m3.group(1)), mon
            y = int(m3.group(3)) if m3.group(3) else None
            if y is None:
                return None
        else:
            m2 = re.fullmatch(r'(\d{1,2})[-.](\d{1,2})[-.](\d{4})', t)
            if not m2:
                return None
            d, m, y = (int(x) for x in m2.groups())
    try:
        return datetime.date(y, m, d).isoformat()
    except ValueError:
        return None


def _series_year(mon, day, today):
    """The year for a yearless '1 Sep' point in a recent series: this year
    when that day has come or passed, else last year."""
    if (mon, day) <= (today.month, today.day):
        return today.year
    return today.year - 1


def _clean_code(v):
    """Uppercase and strip a trading code; keep '&' and '()' as written."""
    if v is None:
        return None
    c = str(v).strip().upper()
    return c or None


# ── table and marquee markup ─────────────────────────────────────────────

class _TableFinder(HTMLParser):
    """Collects the top-level tables (rows of cell text). Only the outermost
    table is tracked, so a marquee's nested tables do not become phantom
    rows; their text either lands in the enclosing cell or is dropped."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self._depth = 0
        self._row = None
        self._cell = None

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self._depth += 1
            if self._depth == 1:
                self.tables.append([])
        elif self._depth == 1:
            if tag == 'tr' and self._row is None:
                self._row = []
            elif tag in ('td', 'th') and self._row is not None and self._cell is None:
                self._cell = []
        # start tags of inner tables: their text lands in self._cell

    def handle_endtag(self, tag):
        if tag == 'table':
            if self._depth == 1:
                self._depth = 0
            elif self._depth > 1:
                self._depth -= 1
        elif tag == 'tr' and self._row is not None:
            if self.tables and self._depth == 1 and any(c.strip() for c in self._row):
                self.tables[-1].append(self._row)
            self._row = None
            self._cell = None
        elif tag in ('td', 'th') and self._cell is not None:
            self._row.append(' '.join(''.join(self._cell).split()))
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _tables(text):
    p = _TableFinder()
    p.feed(text)
    return p.tables


def _find_table(tables, need):
    """The first table whose header row mentions every word in `need`
    (case-insensitive)."""
    for rows in tables:
        if not rows:
            continue
        head = ' '.join(rows[0]).lower()
        if all(w in head for w in need):
            return rows
    return None


def _marquee_rows(text):
    """Old scroll page marquee: <a class='abhead' href=
    "displayCompany.php?name=CODE"> CODE &nbsp; LTP &nbsp; [img] <br>
    CHANGE &nbsp; PERCENT% </a>. Only code, LTP and the change are on the
    marquee (the ab1 links of the data table are not matched)."""
    rows = []
    seen = set()
    for m in re.finditer(
            r"displayCompany\.php\?name=([^\"&]+)\"[^>]*abhead[^>]*>(.*?)</a>",
            text, re.S):
        code = _clean_code(m.group(1))
        if not code or code in seen:
            continue
        seen.add(code)
        inner = re.sub(r'<[^>]+>', ' ', m.group(2))
        inner = ' '.join(_html.unescape(inner).split())
        ltp = change = None
        for t in inner.split():
            if t.endswith('%') and change is None:
                change = to_num(t)
            elif ltp is None and re.fullmatch(r'-?\d[\d,]*(\.\d+)?', t):
                ltp = to_num(t)
        rows.append({
            'code': code, 'ltp': ltp, 'high': None, 'low': None,
            'open': None, 'close': ltp, 'ycp': None, 'change': change,
            'trades': None, 'value_mn': None, 'volume': None,
        })
    return rows


def _header_index(head, words, taken):
    for i, h in enumerate(head):
        if i in taken:
            continue
        if any(w in h.lower() for w in words):
            return i
    return None


def _rows_from_table(rows, mapping):
    """Rows of a table whose header names the columns in `mapping`
    (field name -> list of header words, first unused match wins)."""
    head = rows[0]
    idx = {}
    for field, words in mapping.items():
        i = _header_index(head, words, set(idx.values()))
        if i is not None:
            idx[field] = i
    if 'code' not in idx:
        return []
    out = []
    seen = set()
    for row in rows[1:]:
        if len(row) > len(head) * 2:
            continue  # marquee or ticker junk row, wider than the header
        code = _clean_code(row[idx['code']])
        if not code or code in seen:
            continue
        seen.add(code)
        d = {}
        for field, i in idx.items():
            d[field] = code if field == 'code' else (
                to_num(row[i]) if i < len(row) else None)
        out.append(d)
    return out


def _is_junk_row(code, ncols, nhead):
    """Ticker or marquee rows: no usable code, or a row much wider than the
    header (a marquee cell that flattened into the table)."""
    if not code:
        return True
    if nhead and ncols > nhead * 2:
        return True
    return False


# ── the parsers ──────────────────────────────────────────────────────────

def parse_latest(text):
    """[{code, ltp, high, low, open, close, ycp, change, trades, value_mn,
    volume}]. A close of 0 (the board while the market is open) falls back
    to the LTP."""
    try:
        data = json.loads(text)
    except ValueError:
        data = None
    if isinstance(data, dict) and 'cols' in data and 'rows' in data:
        cols = list(data['cols'])
        out = []
        seen = set()
        for row in data['rows']:
            if not isinstance(row, (list, tuple)) or len(row) != len(cols):
                continue  # junk row: not aligned with the header
            d = dict(zip(cols, row))
            code = _clean_code(d.get('code'))
            if not code or code in seen:
                continue
            seen.add(code)
            ltp = to_num(d.get('ltp'))
            close = to_num(d.get('close'))
            if close in (None, 0.0):
                close = ltp
            out.append({
                'code': code,
                'ltp': ltp,
                'high': to_num(d.get('high')),
                'low': to_num(d.get('low')),
                'open': to_num(d.get('open')),
                'close': close,
                'ycp': to_num(d.get('ycp')),
                'change': to_num(d.get('percent')),
                'trades': to_num(d.get('trades')),
                'value_mn': to_num(d.get('value')),
                'volume': to_num(d.get('volume')),
            })
        return out
    if isinstance(data, dict) and isinstance(data.get('rows'), list):
        # archive-style row objects: latest fields only
        out = []
        seen = set()
        for row in data['rows']:
            if not isinstance(row, dict):
                continue
            code = _clean_code(row.get('tradingCode') or row.get('code'))
            if not code or code in seen:
                continue
            seen.add(code)
            out.append({
                'code': code,
                'ltp': to_num(row.get('ltp')),
                'high': to_num(row.get('high')),
                'low': to_num(row.get('low')),
                'open': to_num(row.get('openp') if 'openp' in row else row.get('open')),
                'close': to_num(row.get('closep') if 'closep' in row else row.get('close')),
                'ycp': to_num(row.get('ycp')),
                'change': to_num(row.get('percent')),
                'trades': to_num(row.get('trade') if 'trade' in row else row.get('trades')),
                'value_mn': to_num(row.get('value')),
                'volume': to_num(row.get('volume')),
            })
        return out
    # Old scroll page: the data table wins over the marquee when both are
    # on the page.
    tables = _tables(text)
    tbl = _find_table(tables, ['trading', 'ltp']) or \
        _find_table(tables, ['code', 'ltp'])
    if tbl and len(tbl) >= 2:
        rows = _rows_from_table(tbl, {
            'code': ['trading code', 'code', 'symbol'],
            'ltp': ['ltp'],
            'high': ['high'],
            'low': ['low'],
            'open': ['open'],
            'close': ['closep', 'close'],
            'ycp': ['ycp', 'previous'],
            'change': ['change'],
            'trades': ['trade'],
            'value_mn': ['value'],
            'volume': ['volume', 'shares'],
        })
        if rows:
            for d in rows:
                if d['close'] in (None, 0.0):
                    d['close'] = d['ltp']
            return rows
    return _marquee_rows(text)


def parse_archive(text):
    """[{date, code, open, high, low, close, ycp, trades, value_mn, volume}].
    A (date, code) pair appears at most once."""
    try:
        data = json.loads(text)
    except ValueError:
        data = None
    if isinstance(data, dict) and isinstance(data.get('rows'), list):
        out = []
        seen = set()
        for row in data['rows']:
            if not isinstance(row, dict):
                continue
            code = _clean_code(row.get('tradingCode') or row.get('code'))
            date = parse_date(row.get('date'))
            if not code:
                continue
            if (date, code) in seen:
                continue
            seen.add((date, code))
            out.append({
                'date': date,
                'code': code,
                'open': to_num(row.get('openp') if 'openp' in row else row.get('open')),
                'high': to_num(row.get('high')),
                'low': to_num(row.get('low')),
                'close': to_num(row.get('closep') if 'closep' in row else row.get('close')),
                'ycp': to_num(row.get('ycp')),
                'trades': to_num(row.get('trade') if 'trade' in row else row.get('trades')),
                'value_mn': to_num(row.get('value')),
                'volume': to_num(row.get('volume')),
            })
        return out
    rows = _find_table(_tables(text), ['date'])
    if not rows or len(rows) < 2:
        return []
    head = [h.lower() for h in rows[0]]
    idx = {}
    for key, words in (('date', ['date']), ('code', ['code', 'symbol', 'name']),
                       ('open', ['open']), ('high', ['high']), ('low', ['low']),
                       ('close', ['close']), ('ycp', ['previous', 'prev', 'ycp']),
                       ('trades', ['trades', 'trade']),
                       ('value_mn', ['value']), ('volume', ['volume', 'shares'])):
        for i, h in enumerate(head):
            if any(w in h for w in words) and i not in idx.values():
                idx[key] = i
                break
    if 'code' not in idx or 'date' not in idx:
        return []
    out = []
    seen = set()
    for row in rows[1:]:
        if len(row) != len(head):
            continue
        code = _clean_code(row[idx['code']])
        if _is_junk_row(code, len(row), len(head)):
            continue
        date = parse_date(row[idx['date']])
        if (date, code) in seen:
            continue
        seen.add((date, code))
        out.append({
            'date': date,
            'code': code,
            'open': to_num(row[idx['open']]) if 'open' in idx else None,
            'high': to_num(row[idx['high']]) if 'high' in idx else None,
            'low': to_num(row[idx['low']]) if 'low' in idx else None,
            'close': to_num(row[idx['close']]) if 'close' in idx else None,
            'ycp': to_num(row[idx['ycp']]) if 'ycp' in idx else None,
            'trades': to_num(row[idx['trades']]) if 'trades' in idx else None,
            'value_mn': to_num(row[idx['value_mn']]) if 'value_mn' in idx else None,
            'volume': to_num(row[idx['volume']]) if 'volume' in idx else None,
        })
    return out


def parse_index(text, today=None):
    """[{date, dsex, ds30, dses}] in date order. market-summary rows give
    DSEX and DS30 (DSES is not in that endpoint); index-history points give
    one series — the file's key picks which — with the year of a yearless
    '1 Sep' label inferred from `today` (defaults to the run date)."""
    today = today or datetime.date.today()
    try:
        data = json.loads(text)
    except ValueError:
        data = None
    if not isinstance(data, dict):
        return []
    if 'rows' in data:
        out = []
        for row in data['rows']:
            if not isinstance(row, dict):
                continue
            date = parse_date(row.get('date'))
            out.append({
                'date': date,
                'dsex': to_num(row.get('dsex')),
                'ds30': to_num(row.get('ds30')),
                'dses': to_num(row.get('dses')),
            })
        out.sort(key=lambda r: r['date'] or '')
        return out
    points = data.get('points')
    if isinstance(points, list):
        which = str(data.get('key') or data.get('code') or 'DSEX').upper()
        field = which if which in ('DSEX', 'DS30', 'DSES') else 'DSEX'
        out = []
        for p in points:
            if not isinstance(p, dict):
                continue
            date = parse_date(p.get('t'))
            if date is None and isinstance(p.get('t'), str):
                m = re.fullmatch(r'(\d{1,2})\s+([A-Za-z]{3,9})', ' '.join(p['t'].split()))
                if m and _MONTHS.get(m.group(2)[:3].lower()):
                    mon = _MONTHS[m.group(2)[:3].lower()]
                    try:
                        date = datetime.date(
                            _series_year(mon, int(m.group(1)), today), mon,
                            int(m.group(1))).isoformat()
                    except ValueError:
                        date = None
            out.append({
                'date': date,
                'dsex': to_num(p.get('value')) if field == 'DSEX' else None,
                'ds30': to_num(p.get('value')) if field == 'DS30' else None,
                'dses': to_num(p.get('value')) if field == 'DSES' else None,
            })
        out.sort(key=lambda r: r['date'] or '')
        return out
    return []


def parse_companies(text):
    """[{code, name, category, sector}]. The live prices JSON carries all
    four per row; the companies directory table gives code, name and the
    board letter (category) with the sector after a ' · '."""
    try:
        data = json.loads(text)
    except ValueError:
        data = None
    if isinstance(data, dict) and 'cols' in data and 'rows' in data:
        cols = list(data['cols'])
        out = []
        seen = set()
        for row in data['rows']:
            if not isinstance(row, (list, tuple)) or len(row) != len(cols):
                continue
            d = dict(zip(cols, row))
            code = _clean_code(d.get('code'))
            if not code or code in seen:
                continue
            seen.add(code)
            out.append({
                'code': code,
                'name': d.get('name') if 'name' in d else None,
                'category': d.get('category') or None,
                'sector': d.get('sector') or None,
            })
        return out
    # Companies directory page: one anchor per row (class 'block flex-1
    # min-w-0' — the header marquee and nav links have other classes).
    # The cell text alone loses the badge boundary, so the anchor block is
    # read structurally: the badge span's title leads with the board
    # letter, the muted div holds 'Name · Sector'.
    out = []
    seen = set()
    for m in re.finditer(
            r'<a class="block flex-1 min-w-0"[^>]*href="/company/([^"]+)"'
            r'[^>]*>(.*?)</a>', text, re.S):
        code = _clean_code(_html.unescape(m.group(1)))
        if not code or code in seen:
            continue
        seen.add(code)
        block = m.group(2)
        badge = re.search(r'title="([A-Z]) — ', block)
        name = None
        sector = None
        nm = re.search(r'--text-muted\)">(.*?)</div>', block, re.S)
        if nm:
            seg = re.sub(r'<[^>]+>', ' ', nm.group(1))
            seg = ' '.join(_html.unescape(seg).split())
            sm = seg.find('·')
            if sm >= 0:
                name = seg[:sm].strip()
                sector = seg[sm + 1:].strip()
            else:
                name = seg
        out.append({
            'code': code,
            'name': name or None,
            'category': badge.group(1) if badge else None,
            'sector': sector or None,
        })
    return out


def parse_markets(text):
    """{limits: [{code, upper, lower}], block: [{code, price, qty, value}],
    marginable: [code]}. The market-statistics JSON carries block trades
    (price = maxPrice); the markets page carries the price-limit and
    marginable tables; the marginable page carries only the list."""
    out = {'limits': [], 'block': [], 'marginable': []}
    try:
        data = json.loads(text)
    except ValueError:
        data = None
    if isinstance(data, dict) and 'block' in data and isinstance(data['block'], dict):
        for row in data['block'].get('rows') or []:
            if not isinstance(row, dict):
                continue
            code = _clean_code(row.get('code'))
            if not code:
                continue
            out['block'].append({
                'code': code,
                'price': to_num(row.get('maxPrice') if row.get('maxPrice') is not None
                                 else row.get('minPrice')),
                'qty': to_num(row.get('quantity')),
                'value': to_num(row.get('valueMn')),
            })
        return out
    tables = _tables(text)
    lim = _find_table(tables, ['lower', 'upper'])
    if lim and len(lim) >= 2:
        for row in lim[1:]:
            if len(row) < 5:
                continue
            code = _clean_code(row[0])
            if not code:
                continue
            out['limits'].append({
                'code': code,
                'upper': to_num(row[4]),
                'lower': to_num(row[3]),
            })
    blk = _find_table(tables, ['instr', 'max', 'min']) or \
          _find_table(tables, ['code', 'max', 'min'])
    if blk and len(blk) >= 2:
        for row in blk[1:]:
            if len(row) < 6:
                continue
            code = _clean_code(row[0])
            if not code:
                continue
            out['block'].append({
                'code': code,
                'price': to_num(row[1]),
                'qty': to_num(row[4]),
                'value': to_num(row[5]),
            })
    mar = _find_table(tables, ['trading code', 'category']) or \
          _find_table(tables, ['ticker', 'category'])
    if mar and len(mar) >= 2:
        head = mar[0]
        ci = next((i for i, h in enumerate(head)
                   if 'code' in h.lower() or 'ticker' in h.lower()), 0)
        for row in mar[1:]:
            if len(row) <= ci:
                continue
            code = _clean_code(row[ci])
            if code and code not in out['marginable']:
                out['marginable'].append(code)
    return out


def parse_news(text):
    """[{date, code, title, body, url}] from the announcements page. The
    server-rendered list carries only a title and a link; date, code and
    body stay None rather than guessed."""
    items = []
    for m in re.finditer(
            r'<a href="([^"]+)"[^>]*class="[^"]*announcement-page-row-link[^"]*"'
            r'[^>]*>(.*?)</a>', text, re.S):
        title = ' '.join(re.sub(r'<[^>]+>', ' ', m.group(2)).split())
        title = _html.unescape(title)
        if not title:
            continue
        items.append({
            'date': None,
            'code': None,
            'title': title,
            'body': None,
            'url': _html.unescape(m.group(1)),
        })
    return items
