#!/usr/bin/env python3
"""
Finbodh daily refresh — self-contained.

Reads the published bundle's data.js, replaces every layer that changes daily,
writes it back. It needs nothing from the repository, so it runs in a fresh
session with only the published artifact to work from.

    python3 refresh.py data.js [news.json]        rewrite a published bundle
    python3 refresh.py --overlay ov.json [news.json]  build the daily layer alone
    python3 refresh.py --upload https://base [ov.json]  push the chunks to the site
    python3 refresh.py --publish DIR [side.json]    write the chunks as plain files

The overlay is the same work, emitted as a standalone document instead of a
rewritten 5 MB bundle. It is stored in the site's own database and fetched by
the page after first paint, so the live site is current every morning without a
redeploy. Because a calendar and the series indexed against it must come from
one run, the overlay replaces the whole price layer or none of it.

Sources, all plain HTTPS on GitHub so this runs unattended:

  niaz86/DSE  dse_companies.json   category, shareholding, NAV, EPS, P/E,
                                   52-week range, paid-up capital — every
                                   listed security, from dsebd.org's own pages
              dse_stocks.xlsx      one worksheet per trading code, a dated
                                   OHLCV row appended each session
              dse_market_summary.xlsx  DSEX and DS30 daily
  Suhried/dse_share  output.json   the day's movers with a seven-session trail

news.json, if given, is [{"title","url","source","published","summary"}, ...].

Nothing here estimates. Anything that fails to fetch is left exactly as it was,
and any indicator that needs more sessions than a ticker has stays absent.
"""
import json, sys, os, math, re, time, urllib.request, urllib.parse, datetime, tempfile

RAW = "https://raw.githubusercontent.com/"
DSE_JSON = RAW + "niaz86/DSE/master/dse_companies.json"
DSE_XLSX = RAW + "niaz86/DSE/master/dse_stocks.xlsx"
IDX_XLSX = RAW + "niaz86/DSE/master/dse_market_summary.xlsx"
MOV_URL = RAW + "Suhried/dse_share/main/docs/output.json"
UA = "Finbodh/0.2 (+https://finbodh.com)"
DSE_BASE = "https://www.dsebd.org/"
PREFIX = "window.__FB__ = "
GAP_PCT = 12.0
MIN_FOR = {"ma20": 20, "ma50": 50, "ma100": 100, "rsi14": 15, "macd": 35,
           "bb20": 20, "atr14": 15, "ret_1w": 6, "ret_1m": 22, "ret_3m": 64,
           "ret_6m": 127}


def get(url, timeout=300, binary=False):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        b = r.read()
    return b if binary else json.loads(b.decode("utf-8", "replace"))


def n_(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace(",", "").strip()
    if s in ("", "-", "n/a", "N/A", "None"): return None
    try: return float(s)
    except Exception: return None


def as_date(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.date().isoformat() if isinstance(v, datetime.datetime) else v.isoformat()
    s = str(v or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m: return m.group(0)
    for fmt in ("%b %d, %Y", "%d %b %Y", "%d-%b-%Y", "%Y/%m/%d"):
        try: return datetime.datetime.strptime(s, fmt).date().isoformat()
        except Exception: pass
    return None


# ── the exchange's day-end archive ────────────────────────────────────────
# The mirror lags a session or two; the exchange publishes every session in its
# day-end archive the same afternoon, so the archive is the first source for
# the most recent sessions and the mirror the fallback for history.
def dhaka_today():
    """Dhaka is UTC+6 and has no daylight saving."""
    return (datetime.datetime.utcnow() + datetime.timedelta(hours=6)).date().isoformat()


_MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
           "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}


def _as_date(v):
    """Port of deploy/daily.js asDate: ISO, day-month-year in the order the
    archive writes it, and day-first D-M-YYYY (recent pages write 14-09-2026)."""
    s = str(v or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m: return m.group(0)
    m = re.match(r"^(\d{1,2})[\s\-/]*([A-Za-z]{3,})[\s\-/,]*(\d{4})", s)
    if m and _MONTHS.get(m.group(2)[:3].lower()):
        return "%s-%02d-%02d" % (m.group(3), _MONTHS[m.group(2)[:3].lower()], int(m.group(1)))
    m = re.match(r"^([A-Za-z]{3,})[\s\-/]*(\d{1,2})[\s\-/,]*(\d{4})", s)
    if m and _MONTHS.get(m.group(1)[:3].lower()):
        return "%s-%02d-%02d" % (m.group(3), _MONTHS[m.group(1)[:3].lower()], int(m.group(2)))
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return "%s-%02d-%02d" % (m.group(3), int(m.group(2)), int(m.group(1)))
    return None


def _table_rows(html):
    out = []
    for tr in re.findall(r"<tr[^>]*>([\s\S]*?)</tr>", html, re.I):
        cells = []
        for c in re.findall(r"<t[dh][^>]*>([\s\S]*?)</t[dh]>", tr, re.I):
            s = re.sub(r"<[^>]*>", " ", c)
            s = re.sub(r"&nbsp;?", " ", s, flags=re.I).replace("&amp;", "&")
            cells.append(re.sub(r"\s+", " ", s).strip())
        if cells: out.append(cells)
    return out


def dayEndParse(html):
    """Port of deploy/daily.js dayEndParse. day_end_archive.php columns:
    #, DATE, TRADING CODE, LTP*, HIGH, LOW, OPENP*, CLOSEP*, YCP, TRADE,
    VALUE (mn), VOLUME. CLOSEP is the session's weighted average close — the
    figure the rest of the product is built on — and LTP only stands in if it
    is missing. The 2026 relaunch serves the same rows as a JSON document, so
    a response that is not table markup is read by its row fields, which carry
    the same columns.
    Returns {code: [(date, open, high, low, close, volume), ...]}."""
    out = {}
    if html.lstrip()[:1] in "{[":
        try:
            d = json.loads(html)
        except Exception:
            return out
        for r in (d.get("rows") or []):
            date = _as_date(r.get("date"))
            code = str(r.get("tradingCode") or "").strip().upper()
            if not date or not re.match(r"^[A-Z0-9.()\-]{2,20}$", code): continue
            close = n_(r.get("closep")) or n_(r.get("ltp"))
            if not close or close <= 0: continue
            out.setdefault(code, []).append(
                (date, n_(r.get("openp")), n_(r.get("high")), n_(r.get("low")),
                 close, n_(r.get("volume"))))
        return out
    for cells in _table_rows(html):
        # Every page carries a marquee of the whole board in one enormous row.
        if len(cells) < 11 or len(cells) > 20: continue
        d0 = _as_date(cells[0])
        d1 = _as_date(cells[1]) if len(cells) > 1 else None
        off = 0 if d0 else (1 if d1 else -1)
        if off < 0: continue                        # header, and the marquee
        d = d0 if off == 0 else d1
        code = (cells[off + 1] or "").strip().upper()
        if not d or not re.match(r"^[A-Z0-9.()\-]{2,20}$", code): continue
        close = n_(cells[off + 6]) or n_(cells[off + 2])
        if not close or close <= 0: continue
        out.setdefault(code, []).append(
            (d, n_(cells[off + 5]), n_(cells[off + 3]), n_(cells[off + 4]),
             close, n_(cells[off + 10])))
    return out


# The 2026 relaunch serves the archive as JSON from /api/live. The day-end
# endpoint (from=, to=, optional single instrument: the site's own page sends
# nothing else) answers at most 500 of the day's ~634 rows — the first in
# trading-code order, `total` the true count. The board endpoint
# /api/live/prices answers uncapped, and its `close` is the same
# weighted-average figure the day-end's `closep` carries (500/500 overlapping
# codes matched to the kopeck on 2026-09-25; see plan/reports/P10c.md). So the
# cheap complete path is: the board for the newest session (one uncapped
# request), then the day-end archive — whole market per earlier session the
# mirror lacks, one request at a time, one per second — with per-instrument
# requests only for the listed securities the 500-row cap dropped.
ARCHIVE_BUDGET_SEC = 12 * 60
ARCHIVE_MAX_REQ = 450


def _http_json(url, timeout=60, referer=None):
    headers = {"User-Agent": UA, "Accept": "*/*"}
    if referer: headers["Referer"] = referer
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def read_board():
    """The live board: every listed security for its latest session, uncapped.
    Returns ({code: (date, open, high, low, close, volume)}, session_date)."""
    doc = _http_json(DSE_BASE + "api/live/prices", timeout=60)
    cols = doc.get("cols") or []
    rows = doc.get("rows") or []
    if not cols or not rows or not all(k in cols for k in ("code", "close")):
        raise ValueError("board response has no usable columns")
    date = str(((doc.get("session") or {}).get("sessionDate") or "")).strip()
    out = {}
    for r in rows:
        d = dict(zip(cols, r))
        code = str(d.get("code") or "").strip().upper()
        if not re.match(r"^[A-Z0-9.()\-]{2,20}$", code): continue
        close = n_(d.get("close"))
        if not close or close <= 0: continue
        out[code] = (date, n_(d.get("open")), n_(d.get("high")), n_(d.get("low")),
                     close, n_(d.get("volume")))
    return out, date


def _day_end_request(log, date, inst=None):
    """One whole-market (or single-instrument) day-end fetch for one session.
    Returns ({code: row-tuple}, total) or (None, None) — the failure leaves
    whatever the mirror has exactly as it was; nothing is guessed."""
    inst_q = ("&inst=" + urllib.parse.quote(inst)) if inst else ""
    url = (DSE_BASE + "api/live/data-archive/day-end?from=" + date
           + "&to=" + date + inst_q)
    try:
        doc = _http_json(url, timeout=60, referer=DSE_BASE + "data-archives")
    except Exception as e:
        log.append("day-end %s FAILED, left as-is: %s" % (inst or date, e))
        return None, None
    rows = dayEndParse(json.dumps(doc))
    total = doc.get("total") if isinstance(doc, dict) else None
    return rows, total


def _market_summary(log, f, t):
    """The exchange's own session calendar for a window, one row per trading
    day it actually ran, with the DSEX and DS30 close of each. A day the
    window omits is not a trading day — the day-end endpoint re-dates rows for
    a non-trading day to the last session that ran, so a weekday guess is not
    safe. Returns [(date, dsex, ds30)] or None on failure."""
    try:
        doc = _http_json(DSE_BASE + "api/live/data-archive/market-summary?from="
                         + f + "&to=" + t, timeout=60,
                         referer=DSE_BASE + "data-archives")
    except Exception as e:
        log.append("market-summary %s..%s FAILED, weekday calendar instead: %s"
                   % (f, t, e))
        return None
    out = []
    for r in (doc.get("rows") or []):
        d = str(r.get("date") or "").strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d): continue
        out.append((d, n_(r.get("dsex")), n_(r.get("ds30"))))
    return out


def read_archive(log, since, tickers=None):
    """Every listed security for every session after `since` (the mirror's
    newest). The board covers the newest session in one uncapped request; each
    other session the mirror still lacks is a whole-market day-end request,
    one per second, and the few listed securities its 500-row cap dropped get
    per-instrument requests (still one per second, under 450 a night, a 12-
    minute stop). `tickers` is the mirror's set of listed codes, used to find
    the capped ones. The market-summary is the session calendar and carries
    the index closes the mirror's index file may lack. Returns
    ({code: {date: row-tuple}}, newest, [(date, dsex, ds30)]) or
    (None, None, None) to leave the mirror exactly as it was."""
    t0 = time.time()
    out, nreq, newest, summary = {}, 0, None, []
    tickers = set(tickers or ())

    def wait():
        if nreq and time.time() - t0 > ARCHIVE_BUDGET_SEC:
            raise TimeoutError("the 12-minute budget is reached")
        if nreq >= ARCHIVE_MAX_REQ:
            raise TimeoutError("the 450-request budget is reached")
        if nreq: time.sleep(1)          # one request per second, never more

    def fill(rows):
        nonlocal newest
        if not rows: return
        for code, sess in rows.items():
            for s in sess:
                if s[0] > since:
                    out.setdefault(code, {})[s[0]] = s
                    newest = s[0] if newest is None else max(newest, s[0])

    try:
        # The newest session: the board, uncapped, one request.
        board, bdate = read_board()
        nreq += 1
        if not bdate or not board:
            raise ValueError("the board gave no session or rows")
        # The board is live during trading hours (10:00-14:30 Dhaka). Its
        # session only counts once the day has closed; a run before 15:00 on
        # the board's own date would store a half-day as a session.
        dhaka_now = datetime.datetime.utcnow() + datetime.timedelta(hours=6)
        if bdate >= dhaka_now.date().isoformat() and dhaka_now.hour < 15:
            raise ValueError("the board is today's and the session has not "
                             "closed; run after 15:00 Dhaka")
        fill({c: [s] for c, s in board.items()})
        newest = bdate
        # The sessions the mirror still lacks, per the exchange's own
        # calendar: everything between `since` and the board's session.
        d_board = datetime.date.fromisoformat(bdate)
        d_since = datetime.date.fromisoformat(since)
        wait()
        window = _market_summary(log, d_since.isoformat(), d_board.isoformat())
        nreq += 1
        if window is not None:
            summary = [(d, a, b) for d, a, b in window if d > since]
            gaps = sorted(d for d, _a, _b in window
                          if since < d < bdate)
        else:
            # Without the exchange's calendar there is no safe guess: the
            # day-end endpoint re-dates a holiday to the last session, so a
            # guessed date can mint a session that never traded. The earlier
            # sessions stay with the mirror tonight; the board still counts.
            log.append("exchange: no session calendar, earlier sessions stay "
                       "with the mirror")
            gaps = []
    except Exception as e:
        log.append("exchange: board FAILED: %s" % e)
        return None, None, None
    for f in gaps:
        try:
            wait()
        except TimeoutError:
            log.append("exchange: backfill stopped before %s; that session "
                       "stays with the mirror" % f)
            break
        rows, total = _day_end_request(log, f)
        nreq += 1
        if rows is None:
            continue
        fill(rows)
        # The 500-row cap: the listed securities past it get their own
        # requests, one per second. A code the exchange has no day-end row for
        # is not traded that session; the request returns nothing and it stays
        # with the mirror.
        if isinstance(total, int) and total > len(rows):
            for code in sorted(tickers - set(rows)):
                try:
                    wait()
                except TimeoutError:
                    log.append("exchange: capped backfill stopped before %s; the "
                               "rest stay with the mirror" % code)
                    break
                per, _ = _day_end_request(log, f, inst=code)
                nreq += 1
                if per:
                    fill(per)
    if not out:
        log.append("exchange: no rows beyond the mirror; mirror left unchanged")
        return None, None, None
    log.append("exchange: path b, %d requests, %d seconds, newest %s"
               % (nreq, int(time.time() - t0), newest))
    return out, newest, summary


def merge_archive(series, arc, tickers=None):
    """Mirror rows are the base; a session the mirror lacks is appended from the
    exchange; where both have a session and the closes differ by more than 0.5%,
    the exchange wins (it is the primary source). `tickers` narrows to the
    securities the product tracks — the exchange's board also lists bonds and
    fund units, and those are not part of this product, so a session for a code
    the product does not track is not added (a new dash-only company page would
    be a regression)."""
    tickers = set(tickers) if tickers is not None else set(series)
    newer = corrected = omitted = 0
    newest = None
    for code, dates in sorted(arc.items()):
        if code not in tickers:
            omitted += len(dates)
            continue
        base = series.get(code)
        for d, s in sorted(dates.items()):
            (dd, o, h, l, c, v) = s
            newest = d if newest is None else max(newest, d)
            if base is None:
                series[code] = [{"d": d, "o": o, "h": h, "l": l, "c": c, "v": v}]
                newer += 1
                continue
            r = next((x for x in base if x["d"] == d), None)
            if r is None:
                base.append({"d": d, "o": o, "h": h, "l": l, "c": c, "v": v})
                newer += 1
            elif r.get("c") and c and abs(c - r["c"]) > 0.005 * r["c"]:
                r.update({"c": c, "h": h if h is not None else r.get("h"),
                          "l": l if l is not None else r.get("l"),
                          "v": v if v is not None else r.get("v")})
                corrected += 1
    for rows in series.values(): rows.sort(key=lambda x: x["d"])
    line = (f"archive: {newer} sessions newer than the mirror (newest {newest}), "
            f"{corrected} closes corrected")
    if omitted:
        line += f", {omitted} sessions for codes the product does not track omitted"
    return line


# ── indicators ────────────────────────────────────────────────────────────
def sma(xs, n):  return sum(xs[-n:]) / n if len(xs) >= n else None
def ema(xs, n):
    if len(xs) < n: return []
    k, e = 2.0 / (n + 1), sum(xs[:n]) / n
    out = [e]
    for x in xs[n:]:
        e = x * k + e * (1 - k); out.append(e)
    return out
def rsi(xs, n=14):
    if len(xs) < n + 1: return None
    g = [max(xs[i] - xs[i-1], 0.0) for i in range(1, len(xs))]
    l = [max(xs[i-1] - xs[i], 0.0) for i in range(1, len(xs))]
    ag, al = sum(g[:n]) / n, sum(l[:n]) / n
    for i in range(n, len(g)):
        ag = (ag * (n - 1) + g[i]) / n; al = (al * (n - 1) + l[i]) / n
    if al == 0: return 100.0 if ag > 0 else 50.0
    return 100.0 - 100.0 / (1.0 + ag / al)
def macd(xs):
    if len(xs) < 35: return None
    ef, es = ema(xs, 12), ema(xs, 26)
    ef = ef[len(ef) - len(es):]
    line = [a - b for a, b in zip(ef, es)]
    sl = ema(line, 9)
    if not sl: return None
    return {"hist": line[-1] - sl[-1],
            "prev": (line[-2] - sl[-2]) if len(sl) > 1 and len(line) > 1 else None}
def boll(xs, n=20, k=2.0):
    if len(xs) < n: return None
    w = xs[-n:]; m = sum(w) / n
    sd = math.sqrt(sum((x - m) ** 2 for x in w) / n)
    up, lo = m + k * sd, m - k * sd
    wd = up - lo
    return {"pct_b": ((xs[-1] - lo) / wd * 100.0) if wd > 0 else None,
            "bandwidth": (wd / m * 100.0) if m else None}
def atr(rows, n=14):
    if len(rows) < n + 1: return None
    trs = []
    for i in range(1, len(rows)):
        h = rows[i].get("h") or rows[i]["c"]; l = rows[i].get("l") or rows[i]["c"]
        pc = rows[i-1]["c"]
        trs.append(max(h - l, abs(h - pc), abs(l - pc)))
    a = sum(trs[:n]) / n
    for x in trs[n:]: a = (a * (n - 1) + x) / n
    return a
def ret_over(xs, n):
    if len(xs) < n + 1 or xs[-(n+1)] <= 0: return None
    return 100.0 * (xs[-1] / xs[-(n+1)] - 1.0)
def find_gaps(rows):
    out = []
    for i in range(1, len(rows)):
        a, b = rows[i-1]["c"], rows[i]["c"]
        if not a or not b: continue
        ch = 100.0 * (b / a - 1.0)
        if abs(ch) > GAP_PCT:
            out.append({"d": rows[i]["d"], "pct": round(ch, 2), "from": a, "to": b})
    return out


def technicals(rows, idx_ret):
    rows = [r for r in rows if r.get("c")]
    if len(rows) < 6: return None
    cs = [r["c"] for r in rows]; vs = [r.get("v") or 0 for r in rows]
    n, last = len(cs), cs[-1]
    g = lambda k, v: None if (MIN_FOR.get(k) and n < MIN_FOR[k]) else v
    m20, m50, m100 = g("ma20", sma(cs, 20)), g("ma50", sma(cs, 50)), g("ma100", sma(cs, 100))
    bb, mc, a = g("bb20", boll(cs)), g("macd", macd(cs)), g("atr14", atr(rows))
    hi, lo = max(cs), min(cs)
    vol20 = sum(vs[-20:]) / 20 if n >= 20 else None
    gaps = find_gaps(rows)
    clean = n - 1
    if gaps:
        lastg = max(i for i in range(1, n) if any(x["d"] == rows[i]["d"] for x in gaps))
        clean = n - 1 - lastg
    r3 = g("ret_3m", ret_over(cs, 63))
    out = {"n": n, "first": rows[0]["d"], "last": rows[-1]["d"],
           "ma20": m20, "ma50": m50, "ma100": m100, "rsi14": g("rsi14", rsi(cs)),
           "macd_hist": mc["hist"] if mc else None,
           "macd_cross": (None if not mc or mc["prev"] is None else
                          ("up" if mc["hist"] > 0 >= mc["prev"] else
                           "down" if mc["hist"] < 0 <= mc["prev"] else None)),
           "bb_pct": bb["pct_b"] if bb else None,
           "bb_width": bb["bandwidth"] if bb else None,
           "atr_pct": (100.0 * a / last) if (a and last) else None,
           "gaps": gaps[-4:] or None, "clean_sessions": clean,
           "ret_1w": g("ret_1w", ret_over(cs, 5)), "ret_1m": g("ret_1m", ret_over(cs, 21)),
           "ret_3m": r3, "ret_6m": g("ret_6m", ret_over(cs, 126)),
           "rel_3m": (r3 - idx_ret[63]) if (r3 is not None and idx_ret.get(63) is not None) else None,
           "range_hi": hi, "range_lo": lo,
           "from_hi": 100.0 * (last / hi - 1.0) if hi else None,
           "from_lo": 100.0 * (last / lo - 1.0) if lo else None,
           "vol_avg20": vol20,
           "vol_ratio": (vs[-1] / vol20) if (vol20 and vol20 > 0 and vs[-1]) else None,
           "traded_days": sum(1 for v in vs[-60:] if v and v > 0),
           "window_days": min(60, n),
           "above_ma20": (last > m20) if m20 else None,
           "above_ma50": (last > m50) if m50 else None,
           "ma20_over_ma50": (m20 > m50) if (m20 and m50) else None}
    return {k: (round(v, 3) if isinstance(v, float) else v)
            for k, v in out.items() if v is not None}


# ── the workbooks ─────────────────────────────────────────────────────────
def clean_sessions(series, index_dates=(), min_n=20):
    """Sessions the exchange never ran, out; a re-dated Thursday, back.
    A copy of src/build/sessions.py (this script needs nothing from the
    repository); tests/sessions.py holds the two to the same answers.

    The day-end archive answers a request for a day the exchange was shut
    with the last session again under the requested date, so a day on which
    not one security's close or volume differs from its own previous row did
    not trade. A Friday or Saturday the exchange's calendar does not list,
    carrying fresh rows, with the Thursday before it missing, is that
    Thursday. Returns (series, copies removed, {from: to})."""
    seen = {}
    for rows in series.values():
        for a, b in zip(rows, rows[1:]):
            n = seen.setdefault(b["d"], [0, 0])
            n[0] += 1
            n[1] += (a.get("c") == b.get("c") and (a.get("v") or 0) == (b.get("v") or 0))
    copies = {d for d, (n, s) in seen.items() if n >= min_n and s == n}
    have = {r["d"] for rows in series.values() for r in rows} - copies
    idx = set(index_dates or ())
    moved = {}
    for d in sorted(have):
        if d in idx:
            continue
        day = datetime.date.fromisoformat(d)
        wd = day.weekday()                                   # Fri 4, Sat 5
        if wd not in (4, 5):
            continue
        thu = (day - datetime.timedelta(days=wd - 3)).isoformat()
        if thu not in have and thu not in moved.values():
            moved[d] = thu
    out = {}
    for code, rows in series.items():
        rr = [dict(r, d=moved.get(r["d"], r["d"])) for r in rows if r["d"] not in copies]
        rr.sort(key=lambda r: r["d"])
        out[code] = rr
    return out, sorted(copies), moved


def read_prices(log):
    try:
        import openpyxl                                   # noqa
    except ImportError:
        os.system(sys.executable + " -m pip install openpyxl --break-system-packages -q")
        try:
            import openpyxl                               # noqa
        except ImportError:
            log.append("prices SKIPPED: openpyxl unavailable")
            return None, None
    import openpyxl
    tmp = tempfile.gettempdir()
    sp, ip = os.path.join(tmp, "s.xlsx"), os.path.join(tmp, "i.xlsx")
    try:
        open(sp, "wb").write(get(DSE_XLSX, binary=True))
        open(ip, "wb").write(get(IDX_XLSX, binary=True))
    except Exception as e:
        log.append(f"price workbooks FAILED, prices left unchanged: {e}")
        return None, None

    wb = openpyxl.load_workbook(sp, read_only=True, data_only=True)
    series = {}
    for sheet in wb.sheetnames:
        ws, head, rows = wb[sheet], None, []
        for r in ws.iter_rows(values_only=True):
            if head is None:
                head = [str(x or "").strip().upper() for x in r]; continue
            rec = dict(zip(head, r))
            d = as_date(rec.get("DATE"))
            c = n_(rec.get("CLOSEP*")) or n_(rec.get("CLOSEP")) or n_(rec.get("LTP*")) or n_(rec.get("LTP"))
            if not d or not c or c <= 0: continue
            rows.append({"d": d, "c": c, "h": n_(rec.get("HIGH")), "l": n_(rec.get("LOW")),
                         "v": n_(rec.get("VOLUME"))})
        if not rows: continue
        rows.sort(key=lambda x: x["d"])
        seen, ded = set(), []
        for r in reversed(rows):
            if r["d"] in seen: continue
            seen.add(r["d"]); ded.append(r)
        ded.reverse()
        series[sheet.strip().upper()] = ded
    wb.close()

    wb = openpyxl.load_workbook(ip, read_only=True, data_only=True)
    ws, head, idx = wb[wb.sheetnames[0]], None, []
    for r in ws.iter_rows(values_only=True):
        if head is None:
            head = [str(x or "").strip() for x in r]; continue
        rec = dict(zip(head, r))
        d, dx = as_date(rec.get("Date")), n_(rec.get("DSEX Index"))
        if d and dx: idx.append([d, round(dx, 2), n_(rec.get("DS30 Index"))])
    wb.close()
    idx.sort(key=lambda x: x[0])
    seen, ded = set(), []
    for r in reversed(idx):
        if r[0] in seen: continue
        seen.add(r[0]); ded.append(r)
    ded.reverse()
    log.append(f"prices: {len(series)} tickers, index {len(ded)} sessions")
    # The exchange's API is the first source for the sessions the mirror
    # still lacks; the mirror above stays the base and the fallback.
    try:
        since = max((r["d"] for rows in series.values() for r in rows), default=None)
        if since:
            arc, arc_newest, arc_idx = read_archive(log, since, tickers=list(series))
            if arc:
                log.append(merge_archive(series, arc, tickers=set(series)))
            if arc_idx:
                # The mirror's index file can sit a session ahead of its own
                # prices; the exchange's market-summary is the authoritative
                # index close, so a session it names wins over the mirror row,
                # and a session the mirror lacks is appended (never invented:
                # only rows the exchange actually sent).
                have = {r[0]: i for i, r in enumerate(ded)}
                added = 0
                for (d, dsex, ds30) in arc_idx:
                    if d in have:
                        if dsex is not None and abs(dsex - ded[have[d]][1]) > 0.005 * dsex:
                            ded[have[d]][1] = round(dsex, 2)
                        if ds30 is not None:
                            ded[have[d]][2] = ds30
                    else:
                        row = [d] + ([round(dsex, 2)] if dsex is not None else [None]) \
                              + ([ds30] if ds30 is not None else [None])
                        ded.append(row)
                        have[d] = len(ded) - 1
                        added += 1
                ded.sort(key=lambda x: x[0])
                log.append(f"index: {added} sessions appended from the exchange")
    except Exception as e:
        log.append(f"archive FAILED, mirror only: {e}")
    series, copies, moved = clean_sessions(series, [r[0] for r in ded])
    if copies or moved:
        log.append(f"sessions: {len(copies)} copies of a closed day removed, "
                   f"{len(moved)} re-dated to their Thursday")
    return series, ded


def main(path, news_path=None):
    raw = open(path, encoding="utf-8").read()
    body = raw[len(PREFIX):] if raw.startswith(PREFIX) else raw
    D = json.loads(body.rstrip().rstrip(";"))
    C = D["companies"]
    today = datetime.date.today().isoformat()
    log = []

    # ---- the exchange's company pages
    try:
        dse = get(DSE_JSON)
        touched = 0
        for tk, r in dse.items():
            c = C.get(tk)
            if not c: continue
            d = c.get("dse") or {}
            for key, src in (("nav_per_share", "nav_per_share"), ("eps_basic", "latest_eps_basic"),
                             ("trailing_pe", "trailing_pe"), ("week52_high", "week52_high"),
                             ("week52_low", "week52_low"), ("opening_price", "opening_price"),
                             ("market_cap_mn", "market_cap_mn"),
                             ("paid_up_capital_mn", "paid_up_capital_mn")):
                v = n_(r.get(src))
                if v is not None: d[key] = v
            for key, src in (("cash_dividend", "cash_dividend"), ("bonus_dividend", "bonus_stock_dividend")):
                if (r.get(src) or "").strip(): d[key] = r[src].strip()
            if r.get("market_category"):
                d["category"] = r["market_category"]; c["cat"] = r["market_category"]
            d["as_of"] = r.get("last_updated") or d.get("as_of")
            c["dse"] = d
            if d.get("week52_high") is not None: c["hi52"] = d["week52_high"]
            if d.get("week52_low") is not None: c["lo52"] = d["week52_low"]
            if d.get("market_cap_mn"): c["mcap"] = d["market_cap_mn"] * 1e6
            own = c.get("own") or {}
            sp, ins, pub = n_(r.get("sponsor_director")), n_(r.get("institute")), n_(r.get("public"))
            if sp is not None and pub is not None:
                own.update({"sponsor_director_pct": sp, "institution_pct": ins,
                            "public_pct": pub, "source_id": "dse",
                            "tradable_pct": 100.0 - sp - (own.get("government_pct") or 0.0),
                            "category": r.get("market_category") or own.get("category")})
                c["own"] = own
            touched += 1
        log.append(f"exchange pages: {touched} securities refreshed")
    except Exception as e:
        log.append(f"exchange pages FAILED, left unchanged: {e}")

    # ---- daily sessions, indicators, index
    series, idx = read_prices(log)
    if series:
        cal = sorted({r["d"] for rows in series.values() for r in rows})
        cali = {x: i for i, x in enumerate(cal)}
        idx_c = [r[1] for r in (idx or []) if r[1]]
        idx_ret = {n: ret_over(idx_c, n) for n in (5, 21, 63, 126)}
        D["cal"] = cal
        if idx: D["index"] = idx
        done = 0
        for tk, rows in series.items():
            c = C.get(tk)
            if not c: continue
            rr = [r for r in rows if r.get("c") and r["d"] in cali]
            if len(rr) >= 3:
                cc, vv, prev = [], [], 0
                for r in rr:
                    x = int(round(r["c"] * 100)); cc.append(x - prev); prev = x
                    vv.append(int(round((r.get("v") or 0) / 1000.0)))
                ks = [cali[r["d"]] for r in rr]
                px = {"i0": ks[0], "c": cc}
                if ks != list(range(ks[0], ks[0] + len(ks))): px["k"] = ks
                if any(vv): px["v"] = vv
                c["px"] = px
                c["price"] = rr[-1]["c"]
                c["qdate"] = rr[-1]["d"]
                c["psrc"] = "dse"
                if len(rr) > 1 and rr[-2]["c"]:
                    c["chg"] = round(100.0 * (rr[-1]["c"] / rr[-2]["c"] - 1.0), 3)
            t = technicals(rows, idx_ret)
            if t: c["tech"] = t
            done += 1
        log.append(f"indicators recomputed for {done} securities")

    # ---- movers
    try:
        m = get(MOV_URL)
        D.setdefault("live", {})["movers"] = {
            "gainers": m.get("gainers") or [], "losers": m.get("losers") or [],
            "as_of": m.get("generated_at_bd") or m.get("generated_at"), "source": MOV_URL}
        log.append(f"movers: {len(m.get('gainers') or [])} up, {len(m.get('losers') or [])} down")
    except Exception as e:
        log.append(f"movers FAILED, left unchanged: {e}")

    # ---- news
    if news_path:
        try:
            items = json.load(open(news_path, encoding="utf-8"))
            if isinstance(items, dict): items = items.get("items") or []
            items = [i for i in items if i.get("title") and i.get("url")]
            if items:
                D.setdefault("live", {})["news"] = {"items": items[:120], "fetched": today}
                log.append(f"news: {len(items)} items")
        except Exception as e:
            log.append(f"news FAILED, left unchanged: {e}")

    D["generated"] = today
    # data.js is loaded with <script src>, so the browser looks for </script>
    # inside it before any JSON parsing happens. json.dumps does not escape
    # angle brackets, and the news titles come from four newspapers verbatim.
    # One headline containing the closing tag ends the script element early,
    # window.__FB__ is never assigned, and the site is a blank page until the
    # next rebuild. Escaping them as \u003c keeps the JSON identical in value.
    body = json.dumps(D, ensure_ascii=False, separators=(",", ":"))
    body = body.replace("<", "\\u003c").replace(">", "\\u003e")
    out = PREFIX + body + ";"
    open(path, "w", encoding="utf-8").write(out)
    log.append(f"wrote {path} ({len(out)/1024:.0f} KB), generated {today}")
    print("\n".join("  " + l for l in log))
    failed = sum(1 for l in log if "FAILED" in l)
    return 1 if failed >= 3 else 0


# ── the financial press ───────────────────────────────────────────────────
FEEDS = [
    ("The Business Standard", "https://www.tbsnews.net/economy/stocks/rss.xml"),
    ("The Business Standard", "https://www.tbsnews.net/economy/rss.xml"),
    ("The Daily Star", "https://www.thedailystar.net/business/rss.xml"),
    ("Dhaka Tribune", "https://www.dhakatribune.com/feed/business"),
]


def _rss_date(s):
    """RFC 822 as the feeds write it, to ISO. Unparseable stays absent."""
    s = (s or "").strip()
    if not s: return None
    for fmt in ("%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S %Z",
                "%a, %d %b %Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d %H:%M:%S"):
        try:
            d = datetime.datetime.strptime(s, fmt)
            return d.isoformat()
        except Exception:
            pass
    return None


def parse_feed(xml_bytes, source):
    """RSS and Atom. Namespaces are stripped: these four feeds use four
    different ones for the same fields."""
    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return []
    for el in root.iter():
        if "}" in el.tag: el.tag = el.tag.split("}", 1)[1]
    out = []
    for it in list(root.iter("item")) + list(root.iter("entry")):
        g = lambda t: (it.findtext(t) or "").strip()
        title = g("title")
        link = g("link")
        if not link:
            a = it.find("link")
            if a is not None: link = (a.get("href") or "").strip()
        if not title or not link.startswith("http"): continue
        body = g("description") or g("summary") or g("content")
        body = re.sub(r"<[^>]+>", " ", body)
        body = re.sub(r"&[a-z]+;|&#\d+;", " ", body)
        body = re.sub(r"\s+", " ", body).strip()
        out.append({"title": re.sub(r"\s+", " ", title),
                    "url": link, "source": source,
                    "published": _rss_date(g("pubDate") or g("published") or g("updated")),
                    "summary": body[:240] or None})
    return out


def fetch_news(log):
    items, seen, failed = [], set(), 0
    for source, url in FEEDS:
        try:
            got = parse_feed(get(url, timeout=45, binary=True), source)
        except Exception as e:
            failed += 1
            log.append(f"news {url.split('/')[2]} unreachable: {e}")
            continue
        for it in got:
            if it["url"] in seen: continue
            seen.add(it["url"]); items.append(it)
    items.sort(key=lambda x: x.get("published") or "", reverse=True)
    if items:
        log.append(f"news: {len(items)} items from {len(FEEDS) - failed} of {len(FEEDS)} feeds")
    else:
        log.append("news FAILED on every feed, omitted")
    return items[:120]


# ── overlay mode ──────────────────────────────────────────────────────────
# Builds the layer that changes daily as a standalone document, so it can be
# stored once and served to the live site without redeploying the 5 MB bundle.
# It reads nothing from the repository and nothing from the published bundle:
# every field below comes from a source fetched in this run, which is why the
# price layer is replaced wholesale rather than appended to. A calendar and the
# series indexed against it must never come from two different runs.

OV_VERSION = 1


def build_overlay(news_path=None):
    log = []
    ov = {"v": OV_VERSION, "generated": datetime.date.today().isoformat(), "co": {}}
    co = ov["co"]

    def slot(tk):
        return co.setdefault(tk, {})

    # ---- the exchange's company pages
    try:
        dse = get(DSE_JSON)
        for tk, r in dse.items():
            tk = (tk or "").strip().upper()
            if not tk: continue
            d = {}
            for key, src in (("nav_per_share", "nav_per_share"), ("eps_basic", "latest_eps_basic"),
                             ("trailing_pe", "trailing_pe"), ("week52_high", "week52_high"),
                             ("week52_low", "week52_low"), ("opening_price", "opening_price"),
                             ("market_cap_mn", "market_cap_mn"),
                             ("paid_up_capital_mn", "paid_up_capital_mn")):
                v = n_(r.get(src))
                if v is not None: d[key] = v
            for key, src in (("cash_dividend", "cash_dividend"), ("bonus_dividend", "bonus_stock_dividend")):
                if (r.get(src) or "").strip(): d[key] = r[src].strip()
            if r.get("market_category"): d["category"] = r["market_category"]
            if r.get("last_updated"): d["as_of"] = r["last_updated"]
            e = slot(tk)
            if d: e["dse"] = d
            if r.get("market_category"): e["cat"] = r["market_category"]
            if d.get("week52_high") is not None: e["hi52"] = d["week52_high"]
            if d.get("week52_low") is not None: e["lo52"] = d["week52_low"]
            if d.get("market_cap_mn"): e["mcap"] = d["market_cap_mn"] * 1e6
            sp, ins, pub = n_(r.get("sponsor_director")), n_(r.get("institute")), n_(r.get("public"))
            if sp is not None and pub is not None:
                # tradable_pct is left to the client: it needs government_pct,
                # which this source does not carry. Nothing is assumed here.
                own = {"sponsor_director_pct": sp, "institution_pct": ins,
                       "public_pct": pub, "source_id": "dse",
                       "category": r.get("market_category")}
                # A null would overwrite a good value on merge. Absent stays absent.
                e["own"] = {k: v for k, v in own.items() if v is not None}
        log.append(f"exchange pages: {len(co)} securities")
    except Exception as e:
        log.append(f"exchange pages FAILED, omitted: {e}")

    # ---- daily sessions, indicators, index
    series, idx = read_prices(log)
    if series:
        cal = sorted({r["d"] for rows in series.values() for r in rows})
        cali = {x: i for i, x in enumerate(cal)}
        idx_c = [r[1] for r in (idx or []) if r[1]]
        idx_ret = {n: ret_over(idx_c, n) for n in (5, 21, 63, 126)}
        ov["cal"] = cal
        if idx: ov["index"] = idx
        done = 0
        for tk, rows in series.items():
            rr = [r for r in rows if r.get("c") and r["d"] in cali]
            e = slot(tk)
            if len(rr) >= 3:
                cc, vv, prev = [], [], 0
                for r in rr:
                    x = int(round(r["c"] * 100)); cc.append(x - prev); prev = x
                    vv.append(int(round((r.get("v") or 0) / 1000.0)))
                ks = [cali[r["d"]] for r in rr]
                px = {"i0": ks[0], "c": cc}
                if ks != list(range(ks[0], ks[0] + len(ks))): px["k"] = ks
                if any(vv): px["v"] = vv
                # The session's open, high and low, each minus its close, in
                # paisa, aligned with c. A row whose source carries no open,
                # high or low gets 0 in all three and its position in x; a
                # row whose open, high or low contradicts its own close gets
                # the same — a candle the source would not support is not
                # drawn, and nothing is invented to fill it. A series with no
                # known open, high or low at all carries none of the four
                # keys.
                oo, hh, ll, xmiss, any_k = [], [], [], [], False
                for j, r in enumerate(rr):
                    cp = int(round(r["c"] * 100))
                    try:
                        op = int(round(r["o"] * 100)) if r.get("o") else 0
                        hp = int(round(r["h"] * 100)) if r.get("h") else 0
                        lp = int(round(r["l"] * 100)) if r.get("l") else 0
                        if not (op > 0 and hp > 0 and lp > 0): raise ValueError
                        if not (lp <= min(op, cp) and max(op, cp) <= hp): raise ValueError
                    except (TypeError, ValueError):
                        xmiss.append(j); oo.append(0); hh.append(0); ll.append(0); continue
                    any_k = True
                    oo.append(op - cp); hh.append(hp - cp); ll.append(lp - cp)
                if any_k:
                    px["o"] = oo; px["h"] = hh; px["l"] = ll
                    if xmiss: px["x"] = xmiss
                e["px"] = px
                e["price"] = rr[-1]["c"]
                e["qdate"] = rr[-1]["d"]
                e["psrc"] = "dse"
                if len(rr) > 1 and rr[-2]["c"]:
                    e["chg"] = round(100.0 * (rr[-1]["c"] / rr[-2]["c"] - 1.0), 3)
            t = technicals(rows, idx_ret)
            if t: e["tech"] = t
            done += 1
        log.append(f"indicators: {done} securities over {len(cal)} sessions")
    else:
        log.append("prices absent: the overlay carries no price layer, "
                   "so the site keeps the series it shipped with")

    # ---- movers
    try:
        m = get(MOV_URL)
        ov.setdefault("live", {})["movers"] = {
            "gainers": m.get("gainers") or [], "losers": m.get("losers") or [],
            "as_of": m.get("generated_at_bd") or m.get("generated_at"), "source": MOV_URL}
        log.append(f"movers: {len(m.get('gainers') or [])} up, {len(m.get('losers') or [])} down")
    except Exception as e:
        log.append(f"movers FAILED, omitted: {e}")

    # ---- news: a file if one is handed over, otherwise the feeds themselves
    items = []
    if news_path and os.path.exists(news_path):
        try:
            items = json.load(open(news_path, encoding="utf-8"))
            if isinstance(items, dict): items = items.get("items") or []
            items = [i for i in items if i.get("title") and i.get("url")]
            log.append(f"news: {len(items)} items from {news_path}")
        except Exception as e:
            log.append(f"news file FAILED, falling back to the feeds: {e}")
    if not items:
        items = fetch_news(log)
    if items:
        ov.setdefault("live", {})["news"] = {"items": items[:120], "fetched": ov["generated"]}

    co.pop("", None)
    ov["log"] = log
    return ov


STALE_DAYS = 4


def check_freshness(ov):
    """A stale mirror must never be published as today. If the newest session
    in the fetched data is more than four calendar days before today
    (Asia/Dhaka), say why and stop before anything is uploaded. Returns the
    newest session date, or None when the guard has spoken."""
    cal = [d for d in (ov.get("cal") or []) if d]
    if not cal:
        print("freshness: no sessions in the fetched data; nothing to check")
        return None
    newest = max(cal)
    try:
        from zoneinfo import ZoneInfo
        today = datetime.datetime.now(ZoneInfo("Asia/Dhaka")).date()
    except Exception:
        today = datetime.date.today()
    age = (today - datetime.date.fromisoformat(newest)).days
    if age > STALE_DAYS:
        print(f"freshness: newest session {newest} is {age} calendar days before "
              f"{today.isoformat()} (Asia/Dhaka); refusing to publish a stale mirror")
        return None
    return newest


def write_overlay(out_path, news_path=None, chunk_kb=88):
    """Writes the overlay, and beside it the base64 chunks that go into D1."""
    import base64, gzip as _gz
    ov = build_overlay(news_path)
    newest = check_freshness(ov)
    if newest is None:
        print("  overlay NOT written: the data failed the freshness guard")
        return 2
    print(f"  freshness: newest session {newest}")
    body = json.dumps(ov, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    gz = _gz.compress(body, 9)
    b64 = base64.b64encode(gz).decode("ascii")
    step = chunk_kb * 1024
    chunks = [b64[i:i + step] for i in range(0, len(b64), step)]
    open(out_path, "w", encoding="utf-8").write(json.dumps(ov, ensure_ascii=False))
    meta = {"v": OV_VERSION, "generated": ov["generated"], "parts": len(chunks),
            "bytes_json": len(body), "bytes_gz": len(gz), "bytes_b64": len(b64),
            "securities": len(ov["co"]), "sessions": len(ov.get("cal") or []),
            "has_prices": bool(ov.get("cal"))}
    side = out_path + ".chunks.json"
    open(side, "w", encoding="utf-8").write(json.dumps({"meta": meta, "chunks": chunks}))
    for l in ov["log"]:
        print("  " + l)
    print(f"  overlay: {len(body)/1024:.0f} KB json, {len(gz)/1024:.0f} KB gzip, "
          f"{len(chunks)} chunks -> {side}")
    failed = sum(1 for l in ov["log"] if "FAILED" in l)
    return 1 if failed >= 2 else 0


def do_publish(out_dir, side_path=None):
    """Publish mode: write the day's chunks as plain files for a public repo.

    Replaces the key-gated push with a pull the Worker needs no secret for. A
    GitHub job in a public repository runs this, force-pushes the output to a
    data branch, and the Worker reads the files back on a cron. The side file
    is the one --overlay writes. Each chunk is written byte-for-byte, and the
    manifest beside it carries the generation id, the part count, each part's
    length and the overlay's meta. Validation happens before anything is
    written: a missing side file, or one with no meta, no chunks, or a chunk
    over 200000 characters, fails the run and leaves the directory empty."""
    side_path = side_path or "ov.json.chunks.json"
    if not os.path.exists(side_path):
        print(f"publish: {side_path} not found; run --overlay first")
        return 1
    try:
        side = json.load(open(side_path, encoding="utf-8"))
    except Exception as e:
        print(f"publish: cannot read {side_path}: {e}")
        return 1
    meta, chunks = side.get("meta") or {}, side.get("chunks") or []
    if not meta or not chunks:
        print(f"publish: {side_path} has no meta or chunks")
        return 1
    if any(len(c) > 200000 for c in chunks):
        print("publish: a chunk is over 200000 characters; nothing written")
        return 1
    gen = str(int(time.time()))
    os.makedirs(out_dir, exist_ok=True)
    manifest = {"v": 1, "gen": gen, "parts": len(chunks),
                "bytes": [len(c) for c in chunks], "meta": meta}
    open(os.path.join(out_dir, "manifest.json"), "w", encoding="utf-8").write(
        json.dumps(manifest, separators=(",", ":")))
    for i, chunk in enumerate(chunks):
        with open(os.path.join(out_dir, "part-%03d.txt" % i), "w", encoding="utf-8") as f:
            f.write(chunk)
    print(f"publish: {len(chunks)} parts as gen {gen} -> {out_dir}")
    return 0


def do_upload(base_url, side_path=None):
    """Upload mode: push the chunks a --overlay run wrote, then commit.
    Runs after --overlay; it reads REFRESH_KEY from the environment and fails
    the run (non-zero exit) on any error, so the workflow is red when the site
    is not current. The generation id is this moment in Unix seconds: ten
    digits for the next ~550 years, and a commit always replaces the one
    before it, which the Worker keeps for a rollback."""
    side_path = side_path or "/tmp/ov.json.chunks.json"
    if not os.path.exists(side_path):
        print(f"upload: {side_path} not found; run --overlay first")
        return 1
    key = os.environ.get("REFRESH_KEY", "")
    if not key:
        print("upload: REFRESH_KEY is not set in the environment; nothing sent")
        return 1
    try:
        side = json.load(open(side_path, encoding="utf-8"))
    except Exception as e:
        print(f"upload: cannot read {side_path}: {e}")
        return 1
    meta, chunks = side.get("meta") or {}, side.get("chunks") or []
    if not meta or not chunks:
        print(f"upload: {side_path} has no meta or chunks")
        return 1
    gen = str(int(time.time()))
    base = base_url.rstrip("/")
    hdr = {"x-refresh-key": key}

    def send(method, path, body, ctype=None, attempts=3, waits=(2, 4, 8)):
        last = None
        for i in range(attempts):
            try:
                h = dict(hdr)
                if ctype: h["content-type"] = ctype
                req = urllib.request.Request(base + path, data=body, headers=h, method=method)
                with urllib.request.urlopen(req, timeout=120) as r:
                    r.read()
                    return True
            except Exception as e:
                last = e
                if i < attempts - 1:
                    print(f"upload: attempt {i + 1} failed for {path}: {e}; "
                          f"retrying in {waits[i]}s")
                    time.sleep(waits[i])
        print(f"upload: {path} failed after {attempts} attempts: {last}")
        return False

    for i, chunk in enumerate(chunks):
        if not send("PUT", f"/api/daily/part?gen={gen}&seq={i}", chunk.encode("utf-8")):
            return 1
    if not send("POST", f"/api/daily/commit?gen={gen}&parts={len(chunks)}",
                json.dumps(meta).encode("utf-8"), ctype="application/json"):
        return 1
    print(f"upload: {len(chunks)} parts committed as gen {gen}")
    return 0


if __name__ == "__main__":
    a = sys.argv[1:]
    if a and a[0] == "--overlay":
        if len(a) < 2:
            print(__doc__); sys.exit(2)
        sys.exit(write_overlay(a[1], a[2] if len(a) > 2 else None))
    if a and a[0] == "--upload":
        if len(a) < 2:
            print(__doc__); sys.exit(2)
        sys.exit(do_upload(a[1], a[2] if len(a) > 2 else None))
    if a and a[0] == "--publish":
        if len(a) < 2:
            print(__doc__); sys.exit(2)
        sys.exit(do_publish(a[1], a[2] if len(a) > 2 else None))
    if not a:
        print(__doc__); sys.exit(2)
    sys.exit(main(a[0], a[1] if len(a) > 1 else None))
