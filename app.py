#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Muse 报告本地服务（macOS）：一个跑在本机的小 web 应用。

功能：
- 后台每 5 分钟自动从 163 邮箱（IMAP）拉取 Muse 发来的报告邮件，
  HTML 附件（完整版）与正文 HTML（简要）按日期存进 ~/选股报告/<日期>/；
- 有新报告时弹 macOS 通知提醒；
- 浏览器打开 http://127.0.0.1:8765 就是索引页：按日期列出全部报告，
  未读带红点、页面每分钟自动刷新，点开即看（本地打开，表格排序、
  整行点击跳经典版等交互完整）。
只用 Python 标准库；只监听 127.0.0.1，不对局域网开放。
"""
import email
import email.policy
import imaplib
import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from html.parser import HTMLParser
from urllib.parse import unquote, urlparse, parse_qs, quote, urljoin, urlencode
import urllib.request

LOCAL_TZ = timezone(timedelta(hours=8))  # 报告日期/时间一律按北京时间

APP_DIR = os.path.expanduser("~/Library/Application Support/MuseMailReports")
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
STATE_PATH = os.path.join(APP_DIR, "state.json")
LOG_PATH = os.path.expanduser("~/Library/Logs/muse-mailreports.log")
PORT = 8765

CFG = {}
LOCK = threading.Lock()


def log(msg):
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass
    print(line, flush=True)


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def reports_root():
    return os.path.expanduser(CFG.get("reports_dir", "~/选股报告"))


def manifest_path():
    return os.path.join(reports_root(), "manifest.json")


def load_manifest():
    return load_json(manifest_path(), [])


def save_manifest(m):
    save_json(manifest_path(), m)


def sanitize(name, fallback="report"):
    name = re.sub(r'[\\/:*?"<>|\s]+', "-", (name or "").strip())
    return name.strip("-.")[:80] or fallback


FETCH_STATE = {"running": False, "new": None, "ts": 0}


def trigger_fetch():
    if FETCH_STATE["running"]:
        return False
    FETCH_STATE["running"] = True

    def _run():
        try:
            FETCH_STATE["new"] = len(fetch_once(limit=5))
        except Exception:
            FETCH_STATE["new"] = -1
        FETCH_STATE["ts"] = time.time()
        FETCH_STATE["running"] = False

    threading.Thread(target=_run, daemon=True).start()
    return True


def notify(title, text):
    """macOS 通知（osascript 系统自带）。"""
    def esc(s):
        return s.replace("\\", "\\\\").replace('"', '\\"')
    try:
        subprocess.run(
            ["osascript", "-e",
             f'display notification "{esc(text)}" with title "{esc(title)}"'],
            timeout=10, capture_output=True)
    except Exception:
        pass


# ---------------- 抓取 ----------------

def send_imap_id(conn):
    """163 邮箱要求客户端报一次 ID，否则可能以 Unsafe Login 拒绝。"""
    try:
        imaplib.Commands["ID"] = ("AUTH", "SELECTED")
        conn._simple_command(
            "ID", '("name" "muse-mailreports" "version" "1.0" "vendor" "muse")')
        conn._untagged_response("OK", [None], "ID")
    except Exception:
        pass


def process_message(msg, done):
    """把一封邮件的 HTML 存进 <日期>/ 目录，返回 manifest 条目或 None。"""
    mid = (msg["Message-ID"] or "").strip()
    if mid and mid in done:
        return None
    subject = str(msg["Subject"] or "（无主题）")
    raw_dt = msg["Date"].datetime if msg["Date"] else None
    dt = raw_dt.astimezone(LOCAL_TZ) if raw_dt else datetime.now(LOCAL_TZ)
    date_s, time_s = f"{dt:%Y-%m-%d}", f"{dt:%H:%M}"
    day_dir = os.path.join(reports_root(), date_s)
    entry = {"id": mid or f"{date_s}-{time_s}-{subject}",
             "date": date_s, "time": time_s, "subject": subject,
             "full": None, "body": None, "read": False}
    for part in msg.walk():
        ctype = part.get_content_type()
        disp = str(part.get("Content-Disposition") or "")
        fname = part.get_filename()
        if fname and fname.lower().endswith(".html"):
            payload = part.get_payload(decode=True) or b""
            os.makedirs(day_dir, exist_ok=True)
            rel = f"{date_s}/{time_s.replace(':', '')}_{sanitize(fname[:-5])}.html"
            with open(os.path.join(reports_root(), rel), "wb") as fh:
                fh.write(payload)
            entry["full"] = rel
        elif ctype == "text/html" and "attachment" not in disp:
            payload = part.get_payload(decode=True) or b""
            if len(payload) < 200:
                continue
            os.makedirs(day_dir, exist_ok=True)
            rel = f"{date_s}/{time_s.replace(':', '')}_{sanitize(subject)}-正文.html"
            with open(os.path.join(reports_root(), rel), "wb") as fh:
                fh.write(payload)
            entry["body"] = rel
    if entry["full"] or entry["body"]:
        return entry
    return None



def _clipboard_text():
    try:
        r = subprocess.run(["pbpaste"], capture_output=True, timeout=5)
        return r.stdout.decode("utf-8", errors="replace")
    except Exception:
        return ""


_ART_JUNK = ("script", "style", "noscript", "iframe", "form", "button",
             "input", "select", "textarea", "svg", "canvas", "video",
             "audio", "header", "footer", "nav", "aside")



_VOID_TAGS = {"img", "br", "hr", "input", "meta", "link", "area",
              "base", "col", "embed", "source", "track", "wbr"}
_SKIP_TAGS = {"script", "style", "noscript", "iframe", "form", "button",
              "select", "textarea", "svg", "canvas", "video", "audio",
              "header", "footer", "nav", "aside"}
_JUNK_RE = re.compile(
    r"(nav|menu|header|footer|sidebar|breadcrumb|login|share|"
    r"recommend|related|rank|advert|banner|toolbar|pagination|"
    r"pagebar|tuijian|guanzhu)", re.I)
_TS_RE = re.compile(
    r"20\d\d[-/.]\d\d?[-/.]\d\d?(?:\s+\d\d?:\d\d)?"
    r"|\d\d?-\d\d?\s+\d\d?:\d\d")


class _Node:
    __slots__ = ("tag", "attrs", "children", "parent")

    def __init__(self, tag, attrs, parent):
        self.tag = tag
        self.attrs = attrs
        self.children = []
        self.parent = parent


class _Dom(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = _Node("root", {}, None)
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        n = _Node(tag, dict(attrs), self.stack[-1])
        self.stack[-1].children.append(n)
        if tag not in _VOID_TAGS:
            self.stack.append(n)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(
            _Node(tag, dict(attrs), self.stack[-1]))

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                break

    def handle_data(self, data):
        if data:
            self.stack[-1].children.append(data)


def _walk(node):
    yield node
    for c in node.children:
        if isinstance(c, _Node):
            yield from _walk(c)


def _text_of(node):
    parts = []
    for c in node.children:
        if isinstance(c, str):
            parts.append(c)
        else:
            parts.append(_text_of(c))
    return "".join(parts)


def _link_text_len(node):
    total = 0
    for n in _walk(node):
        if n.tag == "a":
            total += len(_text_of(n).strip())
    return total


def _cls(node):
    return node.attrs.get("class") or ""


def _tokens(node):
    return _cls(node).split()


def _serialize(node, base_url):
    out = []
    for c in node.children:
        if isinstance(c, str):
            out.append(c.replace("&", "&amp;").replace("<", "&lt;")
                       .replace(">", "&gt;"))
            continue
        if c.tag in _SKIP_TAGS:
            continue
        toks = _tokens(c)
        if "comment-data-left" in toks or "comment-data-button" in toks:
            continue
        ident = _cls(c) + " " + (c.attrs.get("id") or "")
        if _JUNK_RE.search(ident):
            t = _text_of(c).strip()
            if len(t) < 300 or _link_text_len(c) / max(len(t), 1) > 0.5:
                continue
        if c.tag == "img":
            srcv = (c.attrs.get("data-src")
                    or c.attrs.get("data-original")
                    or c.attrs.get("src") or "")
            if not srcv or srcv.startswith("data:"):
                continue
            if not re.match(r"https?:", srcv):
                srcv = urljoin(base_url, srcv)
            out.append(f'<img src="{_esc(srcv)}" '
                       f'alt="{_esc(c.attrs.get("alt") or "")}">')
            continue
        if c.tag == "br":
            out.append("<br>")
            continue
        if c.tag == "hr":
            out.append("<hr>")
            continue
        inner = _serialize(c, base_url)
        if c.tag == "a":
            href = c.attrs.get("href") or "#"
            if href.startswith("javascript:"):
                href = "#"
            elif not re.match(r"(https?:|#|mailto:)", href):
                href = urljoin(base_url, href)
            out.append(f'<a href="{_esc(href)}">{inner}</a>')
        elif c.tag in ("div", "p", "span", "section", "article", "ul",
                       "ol", "li", "table", "thead", "tbody", "tr",
                       "td", "th", "blockquote", "pre", "h1", "h2",
                       "h3", "h4", "h5", "h6", "strong", "b", "em",
                       "i", "u", "font", "small", "sub", "sup",
                       "code", "figure", "figcaption", "main", "dl",
                       "dt", "dd"):
            out.append(f"<{c.tag}>{inner}</{c.tag}>")
        else:
            out.append(inner)
    return "".join(out)


def _pick_body(root):
    for n in _walk(root):
        if n.attrs.get("id") == "js_content":
            return n
    for n in _walk(root):
        if "article-text" in _tokens(n):
            return n
    best, best_score = None, 0.0
    for n in _walk(root):
        if n.tag not in ("div", "section", "article", "main", "td"):
            continue
        t = _text_of(n).strip()
        tl = len(t)
        if tl < 200:
            continue
        lr = _link_text_len(n) / tl
        ps = sum(1 for d in _walk(n) if d.tag in ("p", "br", "h2", "h3"))
        ntags = sum(1 for _ in _walk(n))
        score = (tl * (1 - lr) + 60 * ps) * (tl / (tl + 25 * ntags))
        if score > best_score:
            best, best_score = n, score
    if best is not None:
        t = _text_of(best).strip()
        net = len(t) * (1 - _link_text_len(best) / max(len(t), 1))
        if net < 150:
            return None
        long_paras = 0
        for d in _walk(best):
            own = "".join(c for c in d.children
                          if isinstance(c, str)).strip()
            if len(own) >= 60:
                long_paras += 1
        if long_paras < 2:
            return None
    return best


def _render_comment(block, base_url):
    user, when = "", ""
    for d in _walk(block):
        if d.tag == "a" and ("user-name" in _tokens(d)
                             or "/blog/" in (d.attrs.get("href") or "")):
            user = _text_of(d).strip()
            if user:
                break
    m = _TS_RE.search(_text_of(block))
    if m:
        when = m.group(0)
    text_node = None
    for d in _walk(block):
        if "comment-data-text" in _tokens(d):
            text_node = d
            break
    if text_node is not None:
        inner = _serialize(text_node, base_url)
    else:
        inner = _serialize(block, base_url)
    meta = ""
    if user:
        meta += f"<b>{_esc(user)}</b>"
    if when:
        meta += f" · {_esc(when)}"
    return (f'<div class="cmt"><div class="cmt-meta">{meta}</div>'
            f'<div class="cmt-body">{inner}</div></div>')


def _pick_comments(root, body, base_url):
    blocks = [n for n in _walk(root) if "comment-data" in _tokens(n)]
    if len(blocks) >= 2:
        return [_render_comment(b, base_url) for b in blocks]
    body_ids = set()
    node = body
    while node is not None:
        body_ids.add(id(node))
        node = node.parent
    if body is not None:
        for d in _walk(body):
            body_ids.add(id(d))
    best, best_score = None, 0
    for n in _walk(root):
        if id(n) in body_ids:
            continue
        if n.tag not in ("div", "section", "ul", "article"):
            continue
        t = _text_of(n).strip()
        if len(t) < 100:
            continue
        ts = len(_TS_RE.findall(t))
        if ts < 3:
            continue
        ident = (_cls(n) + " " + (n.attrs.get("id") or "")).lower()
        score = ts * 120
        if re.search(r"comment|reply|huifu|pinglun", ident):
            score += 400
        if score > best_score:
            best, best_score = n, score
    if best is not None:
        return [_serialize(best, base_url)]
    return []


def _extract_article(html, base_url):
    title = ""
    m = re.search(r'<meta[^>]+property=["\']og:title["\'][^>]*'
                  r'content=["\']([^"\']+)', html, re.I)
    if not m:
        m = re.search(r"<title[^>]*>(.*?)</title>", html, re.S | re.I)
    if m:
        title = re.sub(r"<[^>]+>", "", m.group(1)).strip()
    dom = _Dom()
    try:
        dom.feed(html)
    except Exception:
        pass
    body = _pick_body(dom.root)
    if body is None:
        return title, "", ""
    frag = _serialize(body, base_url)
    cmts = _pick_comments(dom.root, body, base_url)
    if cmts:
        frag += ('<h2 class="cmth">💬 评论区</h2>' + "".join(cmts))
    text = re.sub(r"<[^>]+>", "", frag)
    return title, frag, text


def _article_page(title, frag, url):
    t = _esc(title)
    return f"""<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{t}</title><style>
body {{ margin:0; background:#eef1f5; color:#1c2733;
  font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif; }}
.wrap {{ max-width:860px; margin:0 auto; padding:26px 22px 60px;
  background:#fff; min-height:100vh;
  box-shadow:0 2px 18px rgba(31,58,95,.08); }}
h1.at {{ font-size:24px; line-height:1.4; margin:2px 0 8px; }}
.src {{ color:#7a8794; font-size:13px; margin-bottom:18px;
  border-bottom:1px solid #e5e9ee; padding-bottom:12px;
  word-break:break-all; }}
.src a {{ color:#2563eb; }}
.article {{ font-size:16.5px; line-height:1.85; word-break:break-word; }}
.article img {{ max-width:100%; height:auto; }}
.article table {{ border-collapse:collapse; max-width:100%; }}
.article td, .article th {{ border:1px solid #dfe4ea; padding:5px 8px; }}
.article a {{ color:#2563eb; }}
.article h1 {{ font-size:21px; }} .article h2 {{ font-size:19px; }}
.article h3 {{ font-size:17px; }}
.cmth {{ font-size:19px; margin:34px 0 6px; padding-top:18px;
  border-top:2px solid #e5e9ee; }}
.cmt {{ padding:12px 0; border-bottom:1px solid #eef1f5; }}
.cmt-meta {{ color:#7a8794; font-size:13.5px; margin-bottom:4px; }}
.cmt-meta b {{ color:#33475c; }}
.cmt-body {{ font-size:15.5px; line-height:1.75; }}
</style></head><body><div class="wrap">
<h1 class="at">{t}</h1>
<div class="src">来源：<a href="{_esc(url)}">{_esc(url)}</a> · 本地抓取存档</div>
<div class="article">{frag}</div>
</div></body></html>"""


def grab_article():
    clip = _clipboard_text()
    m = re.search(r"https?://[^\s\"'<>]+", clip or "")
    if not m:
        return {"ok": False,
                "msg": "剪贴板里没有链接——先在微信或浏览器里复制文章链接，再点「抓文章」"}
    url = m.group(0).rstrip(")\uFF09\u3011]\"'")
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                          "AppleWebKit/605.1.15 (KHTML, like Gecko) "
                          "Version/14.1.2 Safari/605.1.15",
            "Accept": "text/html,application/xhtml+xml"})
        with urllib.request.urlopen(req, timeout=25) as resp:
            raw = resp.read()
            charset = resp.headers.get_content_charset() or "utf-8"
            final_url = resp.geturl()
    except Exception as e:
        return {"ok": False, "msg": f"打开链接失败：{e}"}
    try:
        html = raw.decode(charset, errors="replace")
    except (LookupError, UnicodeDecodeError):
        html = raw.decode("utf-8", errors="replace")
    if "tgb.cn/bbs/" in final_url and "article-text" not in html:
        return {"ok": False,
                "msg": "这个淘股吧链接是列表版式，抓不到正文——打开文章本体后再复制链接试试"}
    title, frag, text = _extract_article(html, final_url)
    if len(text.strip()) < 200:
        return {"ok": False,
                "msg": "没抓到正文——这个来源可能不让抓（要登录或有验证），换个链接试试"}
    title = (title or "未命名文章")[:80]
    now = datetime.now(LOCAL_TZ)
    date_s, time_s = f"{now:%Y-%m-%d}", f"{now:%H:%M}"
    day_dir = os.path.join(reports_root(), date_s)
    os.makedirs(day_dir, exist_ok=True)
    stem = f"{time_s.replace(':', '')}_{sanitize(title)}"
    rel = f"{date_s}/{stem}.html"
    full_path = os.path.join(reports_root(), rel)
    n = 1
    while os.path.exists(full_path):
        n += 1
        rel = f"{date_s}/{stem}-{n}.html"
        full_path = os.path.join(reports_root(), rel)
    if n > 1:
        stem = f"{stem}-{n}"
    frag = _localize_images(frag, final_url, date_s, stem)
    page = _article_page(title, frag, url)
    with open(full_path, "w", encoding="utf-8") as fh:
        fh.write(page)
    entry = {"id": f"article-{date_s}-{time_s}-{title}", "date": date_s,
             "time": time_s, "subject": title, "full": rel, "body": rel,
             "read": False}
    with LOCK:
        mlist = load_manifest()
        mlist.append(entry)
        save_manifest(mlist)
    return {"ok": True, "title": title}


def _header_info(raw_header):
    """从邮件头字节里取发件人、Message-ID、日期。"""
    msg = email.message_from_bytes(raw_header, policy=email.policy.default)
    frm = str(msg["From"] or "")
    mid = (msg["Message-ID"] or "").strip()
    dt = None
    try:
        dt = msg["Date"].datetime if msg["Date"] else None
    except Exception:
        dt = None
    return frm, mid, dt


def fetch_once(limit=0, cap=0):
    """不用服务器端 FROM 搜索（163 的 IMAP 按发件人搜索实测搜不到）：
    拉最近 400 封的邮件头、在本地按发件人＋日期过滤，再逐封拉正文。"""
    state = load_json(STATE_PATH, {"done": []})
    done = set(state.get("done", []))
    # 连接带 30 秒超时：网络僵住时这一趟自动断开、下趟重来，
    # 不让一次卡住的连接把后续自动拉取全堵死（老 Python 无此参数时降级）
    try:
        conn = imaplib.IMAP4_SSL(CFG.get("imap_host", "imap.163.com"),
                                 int(CFG.get("imap_port", 993)), timeout=30)
    except TypeError:
        conn = imaplib.IMAP4_SSL(CFG.get("imap_host", "imap.163.com"),
                                 int(CFG.get("imap_port", 993)))
    new_entries = []
    try:
        conn.login(CFG["email"], CFG["auth_code"])
        send_imap_id(conn)
        conn.select("INBOX", readonly=True)
        _typ, data = conn.uid("search", None, "ALL")
        uids = data[0].split() if data and data[0] else []
        uids = uids[-400:]
        if limit:
            uids = uids[-limit:]
        cutoff = datetime.now().astimezone() - timedelta(
            days=int(CFG.get("lookback_days", 4)) + 1)
        targets = []
        for i in range(0, len(uids), 100):
            seq = b",".join(uids[i:i + 100]).decode()
            _typ, hdata = conn.uid(
                "fetch", seq,
                "(BODY.PEEK[HEADER.FIELDS (FROM DATE MESSAGE-ID)])")
            for item in hdata or []:
                if not isinstance(item, tuple):
                    continue
                frm, mid, dt = _header_info(item[1])
                if not mid or mid in done:
                    continue
                if CFG["from_addr"].lower() not in frm.lower():
                    continue
                if dt is not None and dt < cutoff:
                    continue
                m = re.search(rb"UID (\d+)", item[0])
                if m:
                    targets.append(m.group(1))
        if cap and len(targets) > cap:
            targets = targets[-cap:]
        for uid in targets:
            _typ, msg_data = conn.uid("fetch", uid, "(RFC822)")
            if not msg_data or not msg_data[0]:
                continue
            msg = email.message_from_bytes(msg_data[0][1],
                                           policy=email.policy.default)
            mid = (msg["Message-ID"] or "").strip()
            with LOCK:
                entry = process_message(msg, done)
                if mid:
                    done.add(mid)
                if entry:
                    m2 = load_manifest()
                    m2.append(entry)
                    save_manifest(m2)
                    new_entries.append(entry)
        state["done"] = sorted(done)
        save_json(STATE_PATH, state)
        log(f"拉取完成：收件箱近 {len(uids)} 封，匹配 {len(targets)} 封，"
            f"新增 {len(new_entries)} 封")
    finally:
        try:
            conn.logout()
        except Exception:
            pass
    return new_entries


SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")
_DEF_SETTINGS = {"mail_auto": True, "mail_interval_sec": 300,
                 "kpl_auto": True, "kpl_interval_sec": 300, "news_auto": True, "news_interval_sec": 120, "kpl_enabled": True, "news_enabled": True}


def _load_settings():
    s = dict(_DEF_SETTINGS)
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as fh:
            got = json.load(fh)
        if isinstance(got, dict):
            s.update({k: got[k] for k in _DEF_SETTINGS if k in got})
    except Exception:
        pass
    return s


def _save_settings(s):
    try:
        os.makedirs(APP_DIR, exist_ok=True)
        with open(SETTINGS_PATH, "w", encoding="utf-8") as fh:
            json.dump(s, fh, ensure_ascii=False)
    except Exception as exc:
        log(f"保存设置失败：{exc}")


def _sleep_interval(base):
    """竞价关键窗口加密拉取：周一至周五 09:27–09:45 每 30 秒一次，
    其余时间用配置间隔（默认 300 秒）。周六、周日放缓到至少
    30 分钟一次。Mac 侧只有星期概念，节假日落到工作日会白跑几趟，
    无害。"""
    now = datetime.now()
    if now.weekday() >= 5:
        return max(base, 1800)
    if now.weekday() < 5:
        secs = now.hour * 3600 + now.minute * 60 + now.second
        if 9 * 3600 + 27 * 60 <= secs <= 9 * 3600 + 45 * 60 + 59:
            return 30
    return base


def fetcher_loop():
    auth_fails = 0      # 连续登录失败次数
    auth_alerted = False  # 本轮失败是否已经报过警（防重复弹）
    while True:
        s = _load_settings()
        if not s.get("mail_auto", True):
            time.sleep(15)
            continue
        try:
            new = fetch_once(cap=5)
            auth_fails = 0
            auth_alerted = False
            if new:
                log(f"新增 {len(new)} 封：{[e['subject'] for e in new]}")
                for e in new:
                    notify("Muse 选股报告 · 新报告", e["subject"])
        except imaplib.IMAP4.error as exc:
            auth_fails += 1
            log(f"IMAP 登录失败：{exc}（授权码不对或 IMAP 未开启）")
            if auth_fails >= 3 and not auth_alerted:
                auth_alerted = True
                log("已连续 3 次登录失败，弹出授权码报警")
                notify("⚠️ 报告服务提醒",
                       "163 邮箱连续 3 次登录失败，报告已暂停下载——"
                       "多半是 IMAP 授权码失效，请检查邮箱设置或更新授权码。")
        except Exception as exc:
            log(f"抓取出错：{type(exc).__name__}: {exc}")
        # 分段睡：开关/频率改动 10 秒内生效
        total = _sleep_interval(int(s.get("mail_interval_sec", 300) or 300))
        slept = 0
        while slept < total:
            step = min(10, total - slept)
            time.sleep(step)
            slept += step
            s2 = _load_settings()
            if not s2.get("mail_auto", True):
                break
            total = _sleep_interval(
                int(s2.get("mail_interval_sec", 300) or 300))





# ---------------- 网页 ----------------

def _esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _row_html(e, alt=False):
    main = e.get("full") or e.get("body")
    fname = main.rsplit("/", 1)[-1] if main else ""
    acts = []
    if main:
        acts.append(
            f'<a class="btn newtab tabopen" style="display:none" '
            f'href="javascript:void(0)" '
            f'onclick="event.stopPropagation();'
            f'openInBrowser(\'{main}\', this)">在新标签打开</a>')
    if e.get("full") and e.get("body"):
        acts.append(
            f'<a class="btn bodybtn" href="javascript:void(0)" '
            f'onclick="event.stopPropagation();'
            f'openReport(\'{e["body"]}\')">正文</a>')
    elif main:
        acts.append('<span class="slotph"></span>')
    if main:
        acts.append(
            f'<a class="btn src" href="javascript:void(0)" '
            f'onclick="event.stopPropagation();'
            f'reveal(\'{main}\',\'{fname}\')">源文件</a>')
    if main:
        qmain = quote(main, safe="")
        fav_label = f"⭐ {_esc(e['fav'])}" if e.get("fav") else "☆ 收藏"
        acts.append(
            f'<a class="btn favbtn" href="javascript:void(0)" '
            f'onclick="event.stopPropagation();'
            f'goNav(\'/fav?f={qmain}\')">{fav_label}</a>')
        acts.append(
            f'<a class="btn delbtn" href="javascript:void(0)" '
            f'onclick="event.stopPropagation();'
            f'delReport(\'{qmain}\', this)">删除</a>')
    cls = "row alt" if alt else "row"
    return (f'<div class="{cls}" onclick="openReport(\'{main}\')">'
            f'<span class="time">{e["time"]}</span>'
            f'<span class="subj">{_esc(e["subject"])}</span>'
            f'<span class="acts">{"".join(acts)}</span></div>')


_PAGE_CSS = """
* { box-sizing: border-box; }
body { font-family:'PingFang SC','Microsoft YaHei',sans-serif;
       background:#e8edf3; margin:0; padding:28px 20px; color:#222; }
.card { background:#fff; border-radius:14px; max-width:1280px; margin:0 auto;
        box-shadow:0 2px 10px rgba(31,58,95,.08); overflow:hidden; }
.head { background:#1f3a5f; color:#fff;
        padding:16px 22px; display:flex; align-items:center; }
.head .ttl { font-size:21px; font-weight:bold; }
.head .r { text-align:right; font-size:13px; white-space:nowrap;
           opacity:.85; }
.note { padding:12px 22px 4px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.notehint { color:#7a8699; font-size:13px; margin-left:4px; }
.datebar { display:flex; align-items:center; margin:10px 22px 4px; }
.datebar .d { font-size:16px; font-weight:bold; color:#1f3a5f; }
.datebar .cnt { margin-left:10px; background:#eaf1fb; color:#1f3a5f;
                font-size:12px; border-radius:20px; padding:2px 10px; }
.datebar .line { flex:1; height:2px; background:#dbe4f0; margin-left:14px; }
.rows { padding:6px 12px 10px; }
.row { display:flex; align-items:center; gap:12px; padding:11px 12px;
       border-radius:9px; border-left:4px solid transparent; cursor:pointer; }
.row.alt { background:#f6f9fc; }
.row:hover { background:#fff3df; border-left-color:#e67e22; }
.time { background:#eaf1fb; color:#1f3a5f; font-weight:bold; font-size:14px;
        border-radius:6px; padding:3px 9px; min-width:52px; text-align:center; }
.subj { font-size:16px; font-weight:bold; }
.acts { margin-left:auto; display:flex; gap:8px; }
.btn { font-size:13px; color:#1f3a5f; background:#eef4fb;
       border:1px solid #cfdff2; border-radius:20px; padding:6px 15px;
       text-decoration:none; white-space:nowrap;
       box-shadow:0 1px 4px rgba(31,58,95,.12); }
.btn.tabopen { background:#1f3a5f; color:#fff; font-weight:bold;
               border-color:#1f3a5f; width:118px; text-align:center;
               padding-left:0; padding-right:0; }
.btn.bodybtn { background:#fff; border-color:#b9cee8; width:58px;
               text-align:center; padding-left:0; padding-right:0; }
.btn.src { width:74px; text-align:center; padding-left:0;
           padding-right:0; }
.slotph { width:58px; flex:none; }
.btn.src { background:#fdf0e2; border-color:#f0d9bd; color:#9a5a12; }
.foot { margin:6px 12px 14px; background:#f2f5f9; border-radius:9px;
        padding:12px 16px; display:flex; align-items:center; cursor:pointer;
        font-size:15px; font-weight:bold; color:#1f3a5f; }
.foot .arrow { margin-left:auto; }
.toast { position:fixed; left:50%; bottom:26px; transform:translateX(-50%);
         background:rgba(28,38,52,.94); color:#fff; font-size:14px;
         border-radius:24px; padding:10px 22px; display:none;
         box-shadow:0 6px 20px rgba(0,0,0,.25); }
.fetchbtn { position:fixed; right:22px; bottom:22px; z-index:60;
            background:#1f3a5f; color:#fff; border:none; border-radius:24px;
            padding:12px 20px; font-size:15px; font-weight:bold;
            cursor:pointer; box-shadow:0 6px 18px rgba(31,58,95,.35); }
.fetchbtn:disabled { opacity:.55; cursor:default; }
.grabbtn { position:fixed; right:22px; bottom:78px; z-index:60;
           background:#fff; color:#1f3a5f; border:1.5px solid #1f3a5f;
           border-radius:24px; padding:11px 18px; font-size:14.5px;
           font-weight:bold; cursor:pointer;
           box-shadow:0 6px 18px rgba(31,58,95,.25); }
.grabbtn:disabled { opacity:.55; cursor:default; }
.fetchbtn, .grabbtn { width:150px; text-align:center; }
.grabbtn::after, .fetchbtn::after {
  content: attr(data-tip); position:absolute; right:calc(100% + 12px);
  top:50%; width:232px; white-space:normal; text-align:left;
  background:#ffb020; color:#231a00; font-size:13.5px; font-weight:700;
  line-height:1.55; padding:9px 12px; border-radius:10px;
  box-shadow:0 8px 22px rgba(0,0,0,.45); opacity:0; pointer-events:none;
  transform:translate(4px, -50%); transition:opacity .15s, transform .15s;
  z-index:70; }
.grabbtn:hover::after, .fetchbtn:hover::after {
  opacity:1; transform:translate(0, -50%); }
.quitbtn { position:fixed; right:22px; bottom:134px; z-index:60;
           width:150px; text-align:center;
           background:#1f3a5f; color:#fff; border:none; border-radius:24px;
           padding:11px 18px; font-size:14px; font-weight:bold;
           cursor:pointer; box-shadow:0 6px 18px rgba(31,58,95,.35); }
.head .rwrap { margin-left:auto; display:flex; align-items:center;
               gap:16px; }
.webbtn { flex:none; white-space:nowrap; background:#fff; color:#1f3a5f;
          font-weight:bold; font-size:15px; border-radius:20px;
          padding:8px 20px; text-decoration:none; cursor:pointer;
          box-shadow:0 2px 8px rgba(0,0,0,.25); }
.homelink { display:inline-block; background:#fff; border:1px solid #c9d6e8; border-radius:999px; padding:5px 13px; color:#1f3a5f; font-weight:600; font-size:13px; text-decoration:none; line-height:1.5; }
.homelink:hover { background:#eaf1fa; border-color:#9db8d9; }
.searchbar { display:flex; gap:10px; padding:14px 22px 2px; }
.searchbar input { flex:1; font-size:16px; padding:11px 18px;
  border:1.5px solid #b9cee8; border-radius:24px; outline:none;
  font-family:inherit; min-width:0; }
.searchbar input:focus { border-color:#1f3a5f; }
.searchbar button { background:#1f3a5f; color:#fff; border:none;
  border-radius:24px; padding:0 28px; font-size:15px; font-weight:bold;
  cursor:pointer; font-family:inherit; flex:none; }
.srow { border-radius:9px; }
.srow.alt { background:#f6f9fc; }
.srow:hover { background:#fff3df; }
.snipline { padding:0 14px 10px 76px; color:#5b6b82; font-size:14px;
  line-height:1.55; margin-top:-5px; }
.snipline .hl { color:#b91c1c; background:#fee2e2; border-radius:4px;
  padding:0 2px; font-weight:bold; }
"""


def _page(title, head_right, note, inner):
    _s = _load_settings()
    kpl_link = ('<a class="homelink" href="javascript:void(0)" '
                'onclick="goNav(\'/kpl\')">📊 开盘啦</a>'
                if _s.get("kpl_enabled", True) else "")
    news_link = ('<a class="homelink" href="javascript:void(0)" '
                 'onclick="goNav(\'/news\')">📰 快讯</a>'
                 if _s.get("news_enabled", True) else "")
    return f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<title>{title}</title>
<style>{_PAGE_CSS}</style></head><body>
<div class="card">
  <div class="head"><div class="ttl">📈 {title}</div>
  <div class="rwrap"><div class="r">{head_right}</div><a class="webbtn"
    style="display:none" href="javascript:void(0)"
    onclick="fetch('/webhome')">🌐 网页版主页</a></div></div>
  <div class="note"><a class="homelink" href="javascript:void(0)"
    onclick="goNav('/home')">🏠 导航主页</a><a class="homelink"
    href="javascript:void(0)" onclick="goNav('/favs')">⭐ 收藏夹</a>{kpl_link}{news_link}<a class="homelink"
    href="javascript:void(0)"
    onclick="goNav('/settings')">⚙ 设置</a><span class="notehint">{note}</span></div>
  {inner}
</div>
<div class="toast" id="toast"></div>
<button class="grabbtn" id="grabbtn" data-tip="先复制文章链接，再点这里：抓取正文＋评论，存进报告列表" onclick="grabArticle()">🔗 抓文章</button>
<button class="fetchbtn" id="fetchbtn" data-tip="立即拉取最新报告邮件（最近 5 封）" onclick="manualFetch()">↻ 获取邮件</button>
<button class="quitbtn" onclick="goNav('/home')">🏠 主页</button>
<script>
var toastTimer = null;
function goNav(p) {{
  var q = [];
  if (/[?&]panel=1/.test(location.search)) q.push('panel=1');
  if (/[?&]embed=1/.test(location.search)) q.push('embed=1');
  location.href = p + (q.length ? (p.indexOf('?') >= 0 ? '&' : '?') + q.join('&') : '');
}}
function doSearch(f) {{
  var v = f.querySelector('input').value.trim();
  if (!v) return false;
  goNav('/search?q=' + encodeURIComponent(v));
  return false;
}}
function openReport(f) {{
  if (/[?&]panel=1/.test(location.search)) {{
    location.href = '/view?f=' + encodeURIComponent(f) + '&panel=1' +
      (/[?&]embed=1/.test(location.search) ? '&embed=1' : '');
    return;
  }}
  fetch('/popout?f=' + encodeURIComponent(f)).then(function(r) {{
    if (!r.ok) location.href = '/view?f=' + encodeURIComponent(f);
  }}).catch(function() {{
    location.href = '/view?f=' + encodeURIComponent(f);
  }});
}}
function openInBrowser(f, el) {{
  if (/[?&]embed=1/.test(location.search)) {{
    var t = '';
    var row = (el && el.closest) ? el.closest('.row') : null;
    if (row) {{
      var s = row.querySelector('.subj');
      if (s) t = s.textContent;
    }}
    parent.postMessage({{museTab: f, museTitle: t}}, '*');
    return;
  }}
  fetch('/popout?f=' + encodeURIComponent(f));
}}
if (/[?&]panel=1/.test(location.search)) {{
  document.querySelectorAll('.newtab').forEach(function(el) {{
    el.style.display = '';
  }});
  document.querySelectorAll('.webbtn').forEach(function(el) {{
    el.style.display = 'inline-block';
  }});
}}
function showToast(msg, ms) {{
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.style.display = 'block';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function() {{ t.style.display = 'none'; }}, ms || 2600);
}}
function grabArticle() {{
  var b = document.getElementById('grabbtn');
  if (!b || b.disabled) {{ return; }}
  b.disabled = true;
  b.textContent = '抓取中…';
  fetch('/grab').then(function(r) {{ return r.json(); }})
    .then(function(j) {{
      b.disabled = false;
      b.textContent = '🔗 抓文章';
      if (j.ok) {{
        showToast('已抓取：《' + j.title + '》，正在打开…', 1800);
        setTimeout(function() {{ goNav('/report'); }}, 1000);
      }} else {{
        showToast(j.msg || '抓取失败', 4500);
      }}
    }}, function() {{
      b.disabled = false;
      b.textContent = '🔗 抓文章';
      showToast('抓取失败，稍后再试', 3000);
    }});
}}
function delReport(qf, btn) {{
  if (btn.dataset.arm) {{
    fetch('/del?f=' + qf).then(function(r) {{ return r.json(); }})
      .then(function(j) {{
        if (j.ok) {{
          showToast('已删除');
          setTimeout(function() {{ location.reload(); }}, 500);
        }} else {{ showToast('删除失败'); }}
      }}, function() {{ showToast('删除失败'); }});
    return;
  }}
  btn.dataset.arm = '1';
  var old = btn.textContent;
  btn.textContent = '确认删除？';
  btn.style.background = '#dc2626';
  btn.style.color = '#fff';
  btn.style.borderColor = '#dc2626';
  setTimeout(function() {{
    btn.dataset.arm = '';
    btn.textContent = old;
    btn.style.background = '';
    btn.style.color = '';
    btn.style.borderColor = '';
  }}, 2600);
}}
function manualFetch() {{
  var b = document.getElementById('fetchbtn');
  b.disabled = true;
  b.textContent = '获取中…';
  fetch('/fetch').then(function() {{ pollFetch(b); }},
                       function() {{ pollFetch(b); }});
}}
function pollFetch(b) {{
  fetch('/fetch-status').then(function(r) {{ return r.json(); }})
    .then(function(s) {{
      if (s.running) {{
        setTimeout(function() {{ pollFetch(b); }}, 2500);
        return;
      }}
      b.disabled = false;
      b.textContent = '↻ 获取邮件';
      if (s.new > 0) {{
        showToast('新到 ' + s.new + ' 份报告，正在刷新…', 1600);
        setTimeout(function() {{ location.reload(); }}, 900);
      }} else if (s.new === 0) {{
        showToast('没有新邮件');
      }} else {{
        showToast('获取失败，稍后再试');
      }}
    }}, function() {{
      setTimeout(function() {{ pollFetch(b); }}, 3000);
    }});
}}
function quitService() {{
  if (!confirm('退出报告服务？退出后就不再自动下载新报告了。')) return;
  var done = function() {{
    showToast('服务已退出，这个页面可以关了', 5000);
  }};
  fetch('/quit').then(done, done);
}}
function reveal(f, name) {{
  fetch('/reveal?f=' + encodeURIComponent(f)).then(function(r) {{
    var t = document.getElementById('toast');
    t.textContent = r.ok ? ('已在访达定位源文件：' + name)
                         : '定位失败，文件可能已被移动';
    t.style.display = 'block';
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function() {{ t.style.display = 'none'; }}, 2600);
  }});
}}
</script>
</body></html>"""


def _datebar(date_s, count):
    return (f'<div class="datebar"><span class="d">{date_s}</span>'
            f'<span class="cnt">共 {count} 份</span>'
            f'<div class="line"></div></div>')


def _rows_html(entries):
    return ('<div class="rows">' + "".join(
        _row_html(e, alt=(i % 2 == 1))
        for i, e in enumerate(entries)) + "</div>")


def _grouped(entries):
    groups = {}
    for e in entries:
        groups.setdefault(e["date"], []).append(e)
    return groups



def _localize_images(frag, base_url, date_s, stem):
    """把文章里的远程图片下载到本地（带文章页作 Referer，绕开图床
    防盗链），frag 里的地址换成本地 /files/ 路径；下不到的保持原样。"""
    import html as _html
    found = []
    for m in re.finditer(r'<img src="(https?://[^"]+)"', frag):
        if m.group(1) not in found:
            found.append(m.group(1))
    if not found:
        return frag
    imgdir = os.path.join(reports_root(), date_s, "artimg")
    for i, esc_url in enumerate(found[:40]):
        try:
            req = urllib.request.Request(
                _html.unescape(esc_url),
                headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel "
                                       "Mac OS X 10_15_7) "
                                       "AppleWebKit/605.1.15 (KHTML, "
                                       "like Gecko) Version/14.1.2 "
                                       "Safari/605.1.15",
                         "Referer": base_url})
            with urllib.request.urlopen(req, timeout=12) as resp:
                data = resp.read()
                ctype = resp.headers.get_content_type()
            if len(data) < 100 or len(data) > 8 * 1024 * 1024:
                continue
            ext = {"image/jpeg": "jpg", "image/png": "png",
                   "image/gif": "gif", "image/webp": "webp"}.get(ctype)
            if not ext:
                mu = re.search(r"\.(jpg|jpeg|png|gif|webp)([?#]|$)",
                               esc_url, re.I)
                ext = (mu.group(1).lower().replace("jpeg", "jpg")
                       if mu else "jpg")
            os.makedirs(imgdir, exist_ok=True)
            name = f"{stem}-{i + 1}.{ext}"
            with open(os.path.join(imgdir, name), "wb") as fh:
                fh.write(data)
            frag = frag.replace(
                f'src="{esc_url}"',
                f'src="/files/{date_s}/artimg/{name}"')
        except Exception:
            continue
    return frag


def delete_entry(rel):
    with LOCK:
        mlist = load_manifest()
        gone, keep = None, []
        for e in mlist:
            if gone is None and rel in (e.get("full"), e.get("body")):
                gone = e
            else:
                keep.append(e)
        if gone is None:
            return False
        save_manifest(keep)
    root = os.path.realpath(reports_root())
    for key in ("full", "body"):
        r = gone.get(key)
        if not r:
            continue
        p = os.path.realpath(os.path.join(root, r))
        if p.startswith(root) and os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass
    full = gone.get("full") or ""
    if "/" in full:
        date_s, fname = full.rsplit("/", 1)
        stem = fname[:-5] if fname.endswith(".html") else fname
        imgdir = os.path.join(root, date_s, "artimg")
        if os.path.isdir(imgdir):
            for fn in os.listdir(imgdir):
                if fn.startswith(stem + "-"):
                    try:
                        os.remove(os.path.join(imgdir, fn))
                    except OSError:
                        pass
    return True


def _nav_suffix(qs):
    parts = []
    if qs.get("panel", [""])[0] == "1":
        parts.append("panel=1")
    if qs.get("embed", [""])[0] == "1":
        parts.append("embed=1")
    return ("&" + "&".join(parts)) if parts else ""


def render_favs():
    with LOCK:
        manifest = load_manifest()
    favs = [e for e in manifest if e.get("fav")]
    favs.sort(key=lambda e: (e["date"], e["time"]), reverse=True)
    if not favs:
        return _page("⭐ 收藏夹", "", "收藏的报告和文章按分类放在这里。",
                     _searchbar_html() +
                     '<p style="padding:4px 22px 18px">还没有收藏。'
                     '在列表里点任意一行的「☆ 收藏」，就能把它归到'
                     '一个分类里收好。</p>')
    groups = {}
    for e in favs:
        groups.setdefault(e["fav"], []).append(e)
    parts = [_searchbar_html()]
    for cat in sorted(groups):
        parts.append(_datebar(f"⭐ {cat}", len(groups[cat])))
        parts.append(_rows_html(groups[cat]))
    return _page("⭐ 收藏夹", f"共 {len(favs)} 份",
                 "收藏的报告和文章按分类放在这里。", "".join(parts))


def render_fav_chooser(rel, suffix):
    with LOCK:
        manifest = load_manifest()
    entry, cats = None, []
    for e in manifest:
        if rel in (e.get("full"), e.get("body")):
            entry = e
        if e.get("fav") and e["fav"] not in cats:
            cats.append(e["fav"])
    if entry is None:
        return _page("收藏", "", "没找到这份报告。",
                     '<p style="padding:4px 22px 18px">没找到这份报告。</p>')
    q = quote(rel, safe="")
    cur = entry.get("fav") or ""
    parts = [f'<div style="padding:16px 22px 4px;font-size:17px;'
             f'font-weight:bold">{_esc(entry["subject"])}</div>']
    if cur:
        parts.append(f'<p style="padding:2px 22px 0;color:#7a8794">'
                     f'现在收藏在「{_esc(cur)}」分类里。</p>')
    parts.append('<div style="padding:10px 22px 20px;display:flex;'
                 'flex-wrap:wrap;gap:10px;align-items:center">')
    for c in sorted(cats):
        if c == cur:
            continue
        parts.append(
            f'<a class="btn" style="font-size:15px;padding:9px 18px" '
            f'href="/fav/set?f={q}&cat={quote(c, safe="")}{suffix}">'
            f'⭐ {c}</a>')
    hid = ""
    if "panel=1" in suffix:
        hid += '<input type="hidden" name="panel" value="1">'
    if "embed=1" in suffix:
        hid += '<input type="hidden" name="embed" value="1">'
    parts.append(
        f'<form action="/fav/set" method="get" '
        f'style="display:flex;gap:8px;align-items:center;margin:0">'
        f'<input type="hidden" name="f" value="{_esc(rel)}">{hid}'
        f'<input name="cat" placeholder="新分类名，如：柚子哥" '
        f'style="font-size:15px;padding:9px 14px;'
        f'border:1.5px solid #c8d2dd;border-radius:20px;outline:none">'
        f'<button type="submit" style="font-size:15px;padding:9px 18px;'
        f'border:none;border-radius:20px;background:#1f3a5f;'
        f'color:#fff;font-weight:bold;cursor:pointer">收藏</button>'
        f'</form>')
    if cur:
        parts.append(
            f'<a class="btn" style="font-size:15px;padding:9px 18px;'
            f'color:#b91c1c;border-color:#f0c6c6" '
            f'href="/fav/set?f={q}&cat={suffix}">取消收藏</a>')
    parts.append('</div>')
    return _page("收藏到分类", "", "点一个已有分类，或新建一个分类。",
                 "".join(parts))



# ---------------- 开盘啦实时数据 ----------------
_KPL_UA = ("Dalvik/2.1.0 (Linux; U; Android 9; SHARK PRS-A0 "
           "Build/PQ3A.190605.01141736)")
_KPL_CACHE = {"ts": 0, "data": None}

_KPL_CSS = """
.kplwrap { max-width:1280px; margin:0 auto; }
.kplgrid { display:grid; grid-template-columns:5fr 7fr; gap:18px;
  margin-top:18px; }
@media (max-width: 900px) { .kplgrid { grid-template-columns:1fr; } }
.kcard { background:rgba(9,17,33,.60); border:1px solid rgba(255,255,255,.22);
  border-radius:18px; padding:18px 20px; color:#e8eef7; }
.khead { display:flex; align-items:baseline; gap:10px; font-size:19px;
  font-weight:bold; margin-bottom:12px; }
.kday { margin-left:auto; font-size:13px; color:#a9bcd4; font-weight:normal; }
.bignum { display:flex; gap:34px; font-size:16.5px; margin-bottom:6px; }
.bignum b { font-size:25px; }
.knote { color:#93a7bf; font-size:13px; margin-left:4px; }
.bks { display:flex; align-items:flex-end; gap:5px; padding-top:10px; }
.bk { flex:1; text-align:center; min-width:0; }
.bkn { font-size:13px; color:#fff; margin-bottom:4px; }
.bkbar { width:78%; margin:0 auto; border-radius:3px 3px 0 0; }
.bkl { font-size:12px; color:#a9bcd4; margin-top:5px; white-space:nowrap; }
.split { display:flex; height:12px; border-radius:6px; overflow:hidden;
  margin-top:16px; background:#2f9e5f; }
.split .sup { background:#d64545; height:100%; }
.splitlab { display:flex; justify-content:space-between; font-size:14.5px;
  margin-top:6px; }
.up { color:#ff7a72; } .dn { color:#5fc98a; }
.vol { margin-top:14px; border-top:1px solid rgba(255,255,255,.14);
  padding-top:12px; font-size:15.5px; }
.volrow { padding:3px 0; }
.voltag { display:inline-block; border-radius:5px; color:#fff;
  font-size:12.5px; padding:1px 7px; margin-right:8px; }
.voltag.a { background:#7a8794; } .voltag.b { background:#e07b39; }
.ptab { width:100%; border-collapse:collapse; font-size:15.5px; }
.ptab th { text-align:left; color:#a9bcd4; font-weight:normal;
  font-size:13.5px; padding:6px 8px;
  border-bottom:1px solid rgba(255,255,255,.18); white-space:nowrap; }
.ptab td { padding:8px; border-bottom:1px solid rgba(255,255,255,.08);
  white-space:nowrap; }
.ptab .qd { font-weight:bold; font-size:17px; }
.pname { font-weight:bold; }
.pcode { color:#8fa5bd; font-size:12.5px; }
.kplgrid2 { display:grid; grid-template-columns:5fr 3fr 4fr; gap:18px;
  margin-top:18px; align-items:start; }
.kplgrid3 { display:grid; grid-template-columns:5fr 7fr; gap:18px;
  margin-top:18px; align-items:start; }
@media (max-width: 900px) { .kplgrid2, .kplgrid3 { grid-template-columns:1fr; } }
.ptab td.qd { text-align:right; }
.ptab tr.sep td { text-align:center; color:#8fa5bd; font-size:12.5px;
  padding:5px 8px; }
a.stk { color:#eef3f9; text-decoration:none; font-weight:bold; }
a.stk:hover { color:#ffb86b; }
.pills { display:inline-flex; gap:6px; }
.spill { cursor:pointer; font-size:12px; font-weight:bold; padding:3px 12px;
  white-space:nowrap;
  border-radius:999px; background:rgba(255,255,255,.08); color:#aeb9cc;
  border:1px solid rgba(255,255,255,.14); }
.spill.on { background:#2f6fb4; color:#fff; border-color:#2f6fb4; }
.fptabs { display:inline-flex; gap:6px; }
.ftab { cursor:pointer; font-size:13px; font-weight:bold; padding:4px 14px;
  white-space:nowrap;
  border-radius:999px; background:rgba(255,255,255,.08); color:#aeb9cc;
  border:1px solid rgba(255,255,255,.14); }
.ftab.on { background:#b4532a; color:#fff; border-color:#b4532a; }
.fpane { max-height:600px; overflow-y:auto; padding-right:2px; }
.subpills { margin:10px 0 4px; display:flex; gap:6px; }
.subhead { font-size:13.5px; font-weight:bold; color:#ffd9a0; margin:12px 0 4px; }
.tlitem { display:flex; gap:10px; padding:9px 2px;
  border-bottom:1px solid rgba(255,255,255,.08); }
.tltime { color:#ffb86b; font-weight:bold; font-size:13px; min-width:40px; }
.tlmain { flex:1; min-width:0; }
.tlrow { font-size:13.5px; line-height:1.5; }
.tldetail { color:#aeb9cc; font-size:12.5px; line-height:1.55; margin-top:2px; }
.tlplate { color:#8fd0ff; font-size:12px; white-space:nowrap; padding-top:2px; }
.tag { display:inline-block; font-size:11px; font-weight:bold; padding:1px 8px;
  border-radius:999px; margin-right:6px; }
.tag.up { background:rgba(224,49,49,.2); color:#ff9d9d; }
.tag.dn { background:rgba(30,160,90,.2); color:#7fe0ae; }
.tag.mid { background:rgba(90,140,220,.2); color:#a9c8ff; }
.zbfeed { max-height:600px; overflow-y:auto; padding-right:2px; }
.zbitem { display:flex; gap:10px; padding:9px 2px;
  border-bottom:1px solid rgba(255,255,255,.08); }
.zbtime { color:#ffb86b; font-weight:bold; font-size:13px; min-width:40px; }
.zbtext { font-size:13.5px; line-height:1.65; color:#e8edf5; }
.zbchips { margin-top:5px; display:flex; flex-wrap:wrap; gap:5px; }
.zbchip { font-size:12px; background:rgba(255,255,255,.07);
  border:1px solid rgba(255,255,255,.1); border-radius:8px; padding:2px 8px; }
.ghead { font-size:13.5px; font-weight:bold; color:#ffd9a0; margin:10px 0 2px; }
.zttable { width:100%; border-collapse:collapse; font-size:13.5px; }
.zttable td { padding:6px 4px; border-bottom:1px solid rgba(255,255,255,.08);
  vertical-align:top; }
.zttable .r { text-align:right; white-space:nowrap; }
.zttable .mut { color:#8fa5bd; font-size:12px; }
.zyreason { color:#aeb9cc; font-size:12.5px; line-height:1.55; }
.lhbline { font-size:13px; line-height:1.95; }
.lhbline .amt { color:#ffb86b; font-weight:bold; }
.capnote { color:#8fa5bd; font-size:12px; margin:6px 0 2px; }
.ztscroll { max-height:560px; overflow-y:auto; padding-right:2px; }
.lstack { display:flex; flex-direction:column; gap:18px; min-width:0; }
.idxbig { display:flex; gap:12px; align-items:baseline; margin:2px 0 6px; }
.idxbig b { font-size:26px; }
.idxbig span { font-size:15px; font-weight:bold; }
::-webkit-scrollbar { width:9px; height:9px; }
::-webkit-scrollbar-track { background:transparent; }
::-webkit-scrollbar-thumb { background:rgba(150,170,200,.30);
  border-radius:99px; }
::-webkit-scrollbar-thumb:hover { background:rgba(170,195,225,.5); }
.kplgrid2, .kplgrid3 { align-items:stretch; }
.kcard { display:flex; flex-direction:column; }
.lstack > .kcard:last-child { flex:1; }
.ztscroll { flex:1 1 auto; min-height:200px; max-height:686px;
  overflow-y:auto; }
.zbfeed { flex:1 1 auto; min-height:200px; max-height:600px;
  overflow-y:auto; }
.fpane { flex:1 1 auto; min-height:200px; max-height:600px;
  overflow-y:auto; }
.kplrefresh { background:#2f6fb4; color:#fff; border:none; padding:6px 16px;
  border-radius:999px; font-size:13px; font-weight:bold; cursor:pointer;
  white-space:nowrap; }
.kplrefresh:disabled { opacity:.55; cursor:default; }
.kplsw { color:#cfe0f2; font-size:13px; display:flex; align-items:center;
  gap:5px; white-space:nowrap; }
.kplsel { background:#12233d; color:#dbe4f0; border:1px solid #33507a;
  border-radius:8px; padding:5px 8px; font-size:13px; }
.secnav { display:flex; gap:8px; flex-wrap:wrap; margin:14px 2px 0; }
.secnav a { display:inline-block; background:rgba(255,255,255,.13);
  border:1px solid rgba(255,255,255,.32); border-radius:999px;
  padding:6px 14px; color:#e8eefc; font-weight:600; font-size:13.5px;
  text-decoration:none; }
.secnav a:hover { background:rgba(255,255,255,.24); }
.secnav a.on { background:#2f6fb3; border-color:#2f6fb3; color:#fff; }
.ttstep { margin:10px 0 4px; border:1px solid rgba(255,255,255,.09);
  border-radius:12px; padding:8px 10px; background:rgba(255,255,255,.03); }
.ttstep .shead { display:flex; align-items:center; gap:8px; margin-bottom:7px; }
.tierbadge { min-width:30px; height:30px; border-radius:9px;
  display:inline-flex; align-items:center; justify-content:center;
  font-size:16px; font-weight:800; color:#fff; padding:0 4px; }
.tier1 { background:#5b7a99; } .tier2 { background:#b8862a; }
.tier3 { background:#c25e2a; } .tier4 { background:#c23a2a; }
.tier5 { background:#a61e1e; }
.stkrow { display:flex; flex-wrap:wrap; gap:6px; }
.stkchip { background:rgba(255,255,255,.06);
  border:1px solid rgba(255,255,255,.11); border-radius:9px;
  padding:4px 9px; font-size:13px; }
.stkchip .tm { color:#ffb86b; font-size:12px; margin-right:5px; }
.stkchip .pl { color:#8fd0ff; font-size:12px; margin-left:5px; }
.stkchip .lb { color:#ff9d9d; font-size:12px; margin-left:5px; }
.zycard { border:1px solid rgba(255,255,255,.09);
  border-left:3px solid #b4532a; border-radius:10px; padding:8px 10px;
  margin:10px 0; background:rgba(255,255,255,.03); }
.zycard .zhead { display:flex; gap:8px; align-items:center;
  margin-bottom:3px; }
.zycard .zname { font-size:14.5px; font-weight:800; }
.zycard .zcnt { background:rgba(224,49,49,.2); color:#ff9d9d;
  font-size:11.5px; font-weight:700; border-radius:99px; padding:1px 8px; }
.zyrow { padding:5px 0; border-top:1px dashed rgba(255,255,255,.08); }
.concept { background:rgba(143,208,255,.13); color:#8fd0ff;
  border-radius:6px; padding:0 6px; font-size:11.5px; margin-left:4px; }
.lhbgrid { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
.lhbcol { border:1px solid rgba(255,255,255,.09); border-radius:10px;
  padding:8px 10px; background:rgba(255,255,255,.03); }
.lhbrow { display:flex; justify-content:space-between; font-size:13px;
  padding:2px 0; }
.lhbrow .amt { color:#ffb86b; font-weight:700; }
.wwgrid { display:grid; gap:10px;
  grid-template-columns:repeat(auto-fit, minmax(105px, 1fr)); }
@media (max-width: 900px) { .wwgrid { grid-template-columns:repeat(2, 1fr); } }
.wwtile { background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.09); border-radius:11px;
  padding:9px 10px; }
.wwname { font-size:13px; font-weight:700; color:#dbe4f0; }
.wwprice { font-size:19px; font-weight:800; margin-top:3px; }
.wwchg { font-size:13px; font-weight:700; }
.wwsub { color:#8fa5bd; font-size:11.5px; line-height:1.55; margin-top:4px; }
.zygrid { display:grid; grid-template-columns:1fr 1fr; gap:10px;
  align-items:start; }
@media (max-width: 900px) { .zygrid { grid-template-columns:1fr; } }
.zygrid .zycard { margin:0; }
.ttgroup { margin:10px 0 4px; }
.tttab { width:100%; border-collapse:collapse; }
.tttab td { padding:4px 6px; font-size:13.5px;
  border-top:1px solid rgba(255,255,255,.06); }
.tttab .tttime { color:#ffb86b; width:52px; white-space:nowrap; }
.tttab .ttname { font-weight:700; white-space:nowrap; }
.tttab .ttlab { color:#ff9d9d; white-space:nowrap; }
.tttab .ttpl { color:#8fd0ff; text-align:right; }
"""

_KPL_JS = """
var KPL_IDS = {mood: 'moodcard', plate: 'plate-top',
  plateall: 'plate-all', idx: 'idx-sh',
  idxcy: 'idx-cy', idxkc: 'idx-kc', kospi: 'idx-kp',
  yidong: 'yidongwrap',
  ww: 'futureswrap', zt: 'ztwrap',
  dt: 'dtwrap', zb: 'zbwrap', ld: 'sub-ld', hc: 'sub-hc', tt: 'ttwrap',
  lhb: 'lhbwrap', zy: 'zywrap'};
function kplPoll(fresh) {
  return fetch(fresh ? '/kpl.json?fresh=1' : '/kpl.json')
    .then(function(r) { return r.json(); })
    .then(function(j) {
      for (var k in KPL_IDS) {
        var el = document.getElementById(KPL_IDS[k]);
        if (el && j[k]) el.innerHTML = j[k];
      }
      var dl = document.getElementById('kpldate');
      if (dl && j.t) {
        var n = new Date(),
          p = function(x) { return (x < 10 ? '0' : '') + x; };
        dl.textContent = n.getFullYear() + '-' + p(n.getMonth() + 1) + '-' +
          p(n.getDate()) + ' · 更新 ' + j.t;
      }
    }, function() {});
}
var kplTimer = null;
function kplStopTimer() {
  if (kplTimer) { clearInterval(kplTimer); kplTimer = null; }
}
function kplSecVal() {
  var el = document.getElementById('kplsec');
  var v = el ? parseFloat(String(el.value).replace(/[^0-9.]/g, '')) : NaN;
  if (!v || v < 15) v = 15;
  if (v > 1800) v = 1800;
  return Math.round(v);
}
function kplStartTimer() {
  kplStopTimer();
  var cb = document.getElementById('kplauto');
  if (cb && cb.checked) {
    kplTimer = setInterval(function() { kplPoll(false); },
      kplSecVal() * 1000);
  }
}
function kplAutoSet(cb) {
  fetch('/settings/set?kpl_auto=' + (cb.checked ? '1' : '0'))
    .then(function() {
      kplStartTimer();
      if (window.showToast) {
        showToast(cb.checked ? '开盘啦自动更新已开' : '开盘啦自动更新已关');
      }
    });
}
function kplSecSet(el) {
  var v = kplSecVal();
  el.value = v;
  fetch('/settings/set?kpl_sec=' + v).then(function() {
    kplStartTimer();
    if (window.showToast) showToast('开盘啦刷新频率：' + v + ' 秒');
  });
}
kplStartTimer();
var fsTried = false;
function toggleFS() {
  if (document.fullscreenElement) { document.exitFullscreen(); return; }
  function fb() { if (!fsTried) { fsTried = true; fetch('/fs'); } }
  try {
    var r = document.documentElement.requestFullscreen();
    if (r && r.catch) { r.catch(fb); }
    setTimeout(function() {
      if (!document.fullscreenElement) fb();
    }, 700);
  } catch (e) { fb(); }
}
function kplRefresh() {
  var b = document.getElementById('kplrefresh');
  if (b) { b.disabled = true; b.textContent = '↻ 刷新中…'; }
  kplPoll(true).then(function() {
    if (b) { b.disabled = false; b.textContent = '↻ 刷新数据'; }
    showToast('数据已刷新');
  });
}
function plateTab(name) {
  document.getElementById('plate-top').style.display =
    (name === 'top') ? '' : 'none';
  document.getElementById('plate-all').style.display =
    (name === 'all') ? '' : 'none';
  document.getElementById('pill-pt').classList.toggle('on', name === 'top');
  document.getElementById('pill-pa').classList.toggle('on', name === 'all');
}
function idxTab(name, btn) {
  ['sh', 'cy', 'kc', 'kp'].forEach(function(n) {
    document.getElementById('idx-' + n).style.display =
      (n === name) ? '' : 'none';
  });
  var bs = btn.parentNode.querySelectorAll('.spill');
  for (var i = 0; i < bs.length; i++) bs[i].classList.remove('on');
  btn.classList.add('on');
}
function ztTab(name) {
  document.getElementById('ztwrap').style.display =
    (name === 'zt') ? '' : 'none';
  document.getElementById('dtwrap').style.display =
    (name === 'dt') ? '' : 'none';
  document.getElementById('pill-zt').classList.toggle('on', name === 'zt');
  document.getElementById('pill-dt').classList.toggle('on', name === 'dt');
}
function fpTab(name, btn) {
  ['shuli', 'dongxiang', 'yuanyin'].forEach(function(n) {
    document.getElementById('pane-' + n).style.display =
      (n === name) ? '' : 'none';
  });
  var bs = btn.parentNode.querySelectorAll('.ftab');
  for (var i = 0; i < bs.length; i++) bs[i].classList.remove('on');
  btn.classList.add('on');
}
function fpSub(name) {
  document.getElementById('sub-ld').style.display =
    (name === 'ld') ? '' : 'none';
  document.getElementById('sub-hc').style.display =
    (name === 'hc') ? '' : 'none';
  document.getElementById('pill-ld').classList.toggle('on', name === 'ld');
  document.getElementById('pill-hc').classList.toggle('on', name === 'hc');
}
"""


def _kpl_post(host, a, c, extra=None, ver="5.21.0.2", apiv="w42"):
    import uuid
    pub = {"apiv": apiv, "PhoneOSNew": "1", "VerSion": ver}
    body = dict(pub)
    body.update({"a": a, "c": c, "DeviceID": str(uuid.uuid4())})
    if extra:
        body.update(extra)
    url = (f"https://{host}.longhuvip.com/w1/api/index.php?"
           + urlencode(pub))
    req = urllib.request.Request(
        url, data=urlencode(body).encode(),
        headers={"User-Agent": _KPL_UA,
                 "Content-Type":
                     "application/x-www-form-urlencoded; charset=UTF-8"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode("utf-8", "replace"))


def _kpl_prev_day(dstr):
    d = datetime.strptime(dstr, "%Y-%m-%d") - timedelta(days=1)
    return f"{d:%Y-%m-%d}"


def _kpl_zhfu(day):
    try:
        r = _kpl_post("apphis", "HisZhangFuDetail", "HisHomeDingPan",
                      {"Day": day})
        info = r.get("info") or {}
        if info.get("ZT") is not None:
            return info
    except Exception:
        pass
    return None


def _kpl_turnover(day):
    try:
        r = _kpl_post("apphis", "GetZsReal", "StockL2History",
                      {"Day": day})
        tot = 0
        for s in r.get("StockList") or []:
            if s.get("StockID") in ("SH000001", "SZ399001"):
                tot += int(s.get("turnover") or 0)
        return tot or None
    except Exception:
        return None


def _kpl_progress(now):
    mins = now.hour * 60 + now.minute
    done = 0
    for a, b in ((570, 690), (780, 900)):
        if mins >= b:
            done += b - a
        elif mins > a:
            done += mins - a
    return done / 240.0


def _tx_index_minute(code="sh000001"):
    u = ("https://web.ifzq.gtimg.cn/appstock/app/minute/query?code=" + code)
    req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15) as r:
        dd = json.loads(r.read().decode("utf-8"))
    node = ((dd.get("data") or {}).get(code)) or {}
    return {"mins": ((node.get("data") or {}).get("data")) or [],
            "qt": ((node.get("qt") or {}).get(code)) or []}


def _yahoo_chart(enc):
    u = ("https://query1.finance.yahoo.com/v8/finance/chart/" + enc +
         "?interval=5m&range=1d")
    req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15) as r:
        dd = json.loads(r.read().decode("utf-8"))
    res = ((dd.get("chart") or {}).get("result") or [None])[0]
    if not res:
        return None
    meta = res.get("meta") or {}
    closes = ((((res.get("indicators") or {}).get("quote") or [{}])[0]
               .get("close")) or [])
    return {"price": meta.get("regularMarketPrice"),
            "pre": meta.get("chartPreviousClose"),
            "closes": [c for c in closes if c]}


def _sina_hf(codes):
    u = "https://hq.sinajs.cn/list=" + ",".join(codes)
    req = urllib.request.Request(
        u, headers={"User-Agent": "Mozilla/5.0",
                    "Referer": "https://finance.sina.com.cn"})
    with urllib.request.urlopen(req, timeout=15) as r:
        text = r.read().decode("gbk", "replace")
    out = {}
    for code in codes:
        m = re.search(r'hq_str_' + code + r'="([^"]*)"', text)
        if m:
            out[code] = m.group(1).split(",")
    return out


_WW_CODES = (("hf_NK", "日经225期货"), ("hf_NQ", "纳指期货"),
             ("hf_CHA50CFD", "A50期货"), ("hf_GC", "黄金"),
             ("hf_CL", "原油"))


def _waiwei():
    raw = _sina_hf([c for c, _ in _WW_CODES])
    items = []
    for code, name in _WW_CODES:
        f = raw.get(code) or []
        if len(f) > 8:
            try:
                items.append({"name": name, "price": float(f[0]),
                              "pre": float(f[7]), "open": float(f[8]),
                              "hi": float(f[4]), "lo": float(f[5])})
            except ValueError:
                pass
    return {"items": items, "kospi": _yahoo_chart("%5EKS11")}


def _em_dtpool(day):
    u = ("https://push2ex.eastmoney.com/getTopicDTPool"
         "?ut=7eea3edcaed734bea9cbfc24409ed989&dpt=wz.ztzt&Pageindex=0"
         "&pagesize=100&sort=fund:asc&date=" + day.replace("-", ""))
    req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        dd = json.loads(r.read().decode("utf-8"))
    return ((dd.get("data") or {}).get("pool")) or []


def _kpl_fetch_all():
    now = datetime.now(LOCAL_TZ)
    today = f"{now:%Y-%m-%d}"
    mood = {}
    try:
        r = _kpl_post("apphwhq", "MoodNumCount", "MarketMood",
                      {"Red": "0"}, ver="6.2.20.2", apiv="w47")
        mood = r.get("list") or {}
    except Exception:
        mood = {}
    day, info = today, None
    for _ in range(5):
        info = _kpl_zhfu(day)
        if info:
            break
        day = _kpl_prev_day(day)
    if not info:
        info, day = {}, today
    turn = _kpl_turnover(day)
    yday, yturn = _kpl_prev_day(day), None
    for _ in range(5):
        yturn = _kpl_turnover(yday)
        if yturn:
            break
        yday = _kpl_prev_day(yday)
    prog = _kpl_progress(now)
    forecast = None
    if turn:
        forecast = turn / prog if 0.02 < prog < 1 else turn
    plates = []
    try:
        # 分页拉全量：st≥100 会回落到 8 条，只能每页 50 翻页；
        # 精选板块共约 240 个，不翻页后五名就是假的
        _seen_plate = set()
        for _ix in range(0, 300, 50):
            url = ("https://apphq.longhuvip.com/w1/api/index.php?Order=1"
                   f"&a=RealRankingInfo&st=50&Index={_ix}&apiv=w21&Type=1"
                   "&c=ZhiShuRanking&ZSType=7")
            req = urllib.request.Request(url,
                                         headers={"User-Agent": _KPL_UA})
            try:
                with urllib.request.urlopen(req, timeout=10) as resp:
                    r = json.loads(resp.read().decode("utf-8", "replace"))
            except Exception:
                if _ix == 0:
                    raise
                break
            rows = r.get("list") or []
            for row in rows:
                if len(row) >= 15 and row[0] not in _seen_plate:
                    _seen_plate.add(row[0])
                    plates.append({"code": row[0], "name": row[1],
                                   "qd": row[2], "zf": row[3],
                                   "zl": row[6], "jg": row[14]})
            if len(rows) < 50:
                break
        plates.sort(key=lambda p: -(p["qd"] or 0))
    except Exception:
        plates = []
    def _try(fn):
        try:
            return fn()
        except Exception:
            return None
    pmld = _try(lambda: _kpl_post("apphwhq", "GetPMSL_PMLD", "FuPanLa",
                                  {"st": "60", "Index": "0", "Date": day,
                                   "Red": "0"}, ver="6.2.20.2", apiv="w47")) or {}
    kqxy = _try(lambda: _kpl_post("apphwhq", "GetPMSL_KQXY", "FuPanLa",
                                  {"Date": day, "Red": "0"},
                                  ver="6.2.20.2", apiv="w47")) or {}
    tianti = _try(lambda: _kpl_post("apphwhq", "GetZhangTingTianTi_W47",
                                    "FuPanLa", {"Red": "0"},
                                    ver="6.2.20.2", apiv="w47")) or {}
    lhb = _try(lambda: _kpl_post("apphwhq", "GetYTFP_LHBDX", "FuPanLa",
                                 {"st": "30", "Index": "0", "Red": "0"},
                                 ver="6.2.20.2", apiv="w47")) or {}
    ztyy = _try(lambda: _kpl_post("apphis", "GetPlateInfo_w38",
                                  "HisLimitResumption",
                                  {"st": "100", "Index": "0", "Date": day,
                                   "Red": "0"}, ver="6.2.20.2", apiv="w47")) or {}
    zhibo = _try(lambda: _kpl_post("apphwhq", "ZhiBoContent", "ConceptionPoint",
                                   {"index": "0", "Red": "0"},
                                   ver="6.2.20.2", apiv="w47")) or {}
    dtpool = _try(lambda: _em_dtpool(day)) or []
    idx_min = _try(lambda: _tx_index_minute("sh000001")) or {}
    idx_cy = _try(lambda: _tx_index_minute("sz399006")) or {}
    idx_kc = _try(lambda: _tx_index_minute("sh000688")) or {}
    idx_rh = _try(_waiwei) or {}
    return {"t": f"{now:%H:%M:%S}", "day": day, "mood": mood,
            "info": info, "turn": turn, "yturn": yturn,
            "forecast": forecast, "plates": plates,
            "pmld": pmld.get("List") or [], "kqxy": kqxy.get("List") or [],
            "tianti": tianti.get("StockList") or [],
            "lhb": lhb.get("List") or [], "ztyy": ztyy,
            "zhibo": zhibo.get("List") or [], "dtpool": dtpool, "idx": idx_min,
            "idxcy": idx_cy, "idxkc": idx_kc, "idxrh": idx_rh}


def _kpl_data():
    ts = time.time()
    if _KPL_CACHE["data"] and ts - _KPL_CACHE["ts"] < 10:
        return _KPL_CACHE["data"]
    data = _kpl_fetch_all()
    if data["info"] or data["plates"] or data["mood"]:
        _KPL_CACHE.update(ts=ts, data=data)
        return data
    return _KPL_CACHE["data"] or data


def _kpl_buckets(info):
    def g(k):
        try:
            return int(info.get(k) or 0)
        except (TypeError, ValueError):
            return 0
    return [("涨停", g("ZT"), "#a61e1e"),
            (">7%", g("8") + g("9") + g("10"), "#c22f2f"),
            ("7~5%", g("6") + g("7"), "#d64545"),
            ("5~2%", g("3") + g("4") + g("5"), "#e0655f"),
            ("2~0%", g("1") + g("2"), "#e88a84"),
            ("平", g("0"), "#8d99a8"),
            ("0~2%", g("-1") + g("-2"), "#7fb98a"),
            ("2~5%", g("-3") + g("-4") + g("-5"), "#4da763"),
            ("5~7%", g("-6") + g("-7"), "#2f9257"),
            ("7%<", g("-8") + g("-9") + g("-10"), "#1e7d46"),
            ("跌停", g("DT"), "#145c34")]


def _stk(code, name):
    return (f'<a class="stk" href="/jump?code={_esc(str(code or ""))}" '
            f'title="在经典版打开">{_esc(str(name or ""))}</a>')


def _hm(ts):
    try:
        return datetime.fromtimestamp(int(ts), LOCAL_TZ).strftime("%H:%M")
    except (TypeError, ValueError, OSError):
        return ""


def _kpl_fragments(d):
    info, mood = d["info"], d["mood"]

    def mi(k, ik):
        v = mood.get(k)
        if v is None:
            v = info.get(ik)
        try:
            return int(v)
        except (TypeError, ValueError):
            return 0

    zt_n, dt_n = mi("ZTJS", "SJZT"), mi("DTJS", "SJDT")
    sz, xd = mi("SZJS", "SZJS"), mi("XDJS", "XDJS")
    buckets = _kpl_buckets(info)
    mx = max([b[1] for b in buckets] + [1])
    cols = []
    for label, n, color in buckets:
        h = max(3, round(n / mx * 96))
        cols.append(
            f'<div class="bk"><div class="bkn">{n}</div>'
            f'<div class="bkbar" style="height:{h}px;background:{color}"></div>'
            f'<div class="bkl">{label}</div></div>')
    tot = max(sz + xd, 1)
    wpct = round(sz / tot * 100, 1)
    turn, fc, yt = d["turn"], d["forecast"], d["yturn"]
    vol_html = ""
    if turn:
        line = (f'<div class="volrow"><span class="voltag a">沪深</span>'
                f'实际量能 <b>{turn / 1e8:.0f}亿</b></div>')
        if fc and yt:
            diff = fc - yt
            pct = diff / yt * 100
            word = "放量" if diff >= 0 else "缩量"
            cls = "up" if diff >= 0 else "dn"
            line += (f'<div class="volrow"><span class="voltag b">今日</span>'
                     f'预测量能 <b>{fc / 1e8:.0f}亿</b>'
                     f'<span class="{cls}">({pct:+.2f}%，'
                     f'{word}{abs(diff) / 1e8:.0f}亿)</span></div>')
        vol_html = f'<div class="vol">{line}</div>'
    mood_html = (
        f'<div class="khead"><span>📊 市场人气</span>'
        f'<span class="kday">{d["day"]} · 更新 {d["t"]}</span></div>'
        f'<div class="bignum"><div>实际涨停 <b class="up">{zt_n}</b> 家'
        f'<span class="knote">（过滤ST股）</span></div>'
        f'<div>实际跌停 <b class="dn">{dt_n}</b> 家'
        f'<span class="knote">（过滤ST股）</span></div></div>'
        f'<div class="bks">{"".join(cols)}</div>'
        f'<div class="split"><div class="sup" style="width:{wpct}%"></div>'
        f'</div>'
        f'<div class="splitlab"><span class="up">涨 {sz} 家</span>'
        f'<span class="dn">跌 {xd} 家</span></div>'
        + vol_html)

    # ---- 板块强度：只看最强 5 + 最弱 5 ----
    plates = d["plates"]

    def prow(p):
        qd = p["qd"]
        qds = f"{qd:.0f}" if isinstance(qd, (int, float)) else str(qd)
        return (f'<tr><td><div class="pname">{_esc(str(p["name"]))}</div>'
                f'<div class="pcode">{_esc(str(p["code"]))}</div></td>'
                f'<td class="qd">{qds}</td></tr>')

    if len(plates) <= 10:
        prows = "".join(prow(p) for p in plates)
    else:
        prows = "".join(prow(p) for p in plates[:5])
        prows += ('<tr class="sep"><td colspan="2">'
                  f'··· 共 {len(plates)} 个板块 ···</td></tr>')
        prows += "".join(prow(p) for p in plates[-5:])
    if not plates:
        prows = '<tr><td class="mut">暂无数据</td></tr>'
    _pthead = ('<table class="ptab"><tr><th>板块</th>'
               '<th style="text-align:right">强度</th></tr>')
    plate_html = (f'<div class="capnote">更新 {d["t"]} · 最强与最弱</div>'
                  + _pthead + prows + "</table>")
    allrows = "".join(prow(p) for p in plates)
    plateall_html = (f'<div class="capnote">共 {len(plates)} 个板块 · '
                     '按强度从高到低，框内滚动</div>'
                     + _pthead
                     + (allrows or '<tr><td class="mut">暂无数据</td></tr>')
                     + "</table>")

    # ---- 涨停（按封单排，来自涨停原因数据）/ 跌停（东财池，按封单排）----
    ztrows = []
    for sec in (d.get("ztyy") or {}).get("list") or []:
        for s in sec.get("StockList") or []:
            if len(s) >= 10:
                ztrows.append({
                    "code": s[0], "name": s[1], "fd": s[8] or 0,
                    "lb": s[9], "t": s[6],
                    "plate": sec.get("ZSName", "")})
    ztrows.sort(key=lambda r: -(r["fd"] or 0))
    ztrs = "".join(
        f'<tr><td>{_stk(r["code"], r["name"])}'
        f'<div class="mut">{_esc(str(r["code"]))} · '
        f'{_esc(str(r["plate"]))}</div></td>'
        f'<td class="r">{_hm(r["t"])}</td>'
        f'<td class="r">{_esc(str(r["lb"]))}</td>'
        f'<td class="r"><b>{(r["fd"] or 0) / 1e8:.2f}亿</b></td></tr>'
        for r in ztrows)
    zt_html = (f'<div class="capnote">共 {len(ztrows)} 只 · 按封单从大到小</div>'
               '<table class="zttable"><tr><td class="mut">股票</td>'
               '<td class="mut r">封板</td><td class="mut r">连板</td>'
               '<td class="mut r">封单</td></tr>'
               + (ztrs or '<tr><td class="mut">暂无数据</td></tr>') + "</table>")
    dtpool = sorted(d.get("dtpool") or [],
                    key=lambda r: -(r.get("fund") or 0))
    dttrs = "".join(
        f'<tr><td>{_stk(r.get("c"), r.get("n"))}'
        f'<div class="mut">{_esc(str(r.get("c") or ""))} · '
        f'{_esc(str(r.get("hybk") or ""))}</div></td>'
        f'<td class="r dn">{(r.get("zdp") or 0):+.2f}%</td>'
        f'<td class="r"><b>{(r.get("fund") or 0) / 1e8:.2f}亿</b></td></tr>'
        for r in dtpool)
    dt_html = (f'<div class="capnote">共 {len(dtpool)} 只 · 按封单从大到小'
               '（东方财富口径）</div>'
               '<table class="zttable"><tr><td class="mut">股票</td>'
               '<td class="mut r">跌幅</td><td class="mut r">封单</td></tr>'
               + (dttrs or '<tr><td class="mut">暂无数据</td></tr>') + "</table>")

    # ---- 大盘直播 ----
    zb_items = []
    for it in (d.get("zhibo") or [])[:30]:
        chips = ""
        stocks = it.get("Stock") or []
        if stocks:
            cs = []
            for s in stocks[:6]:
                if isinstance(s, (list, tuple)) and len(s) >= 3:
                    zf = s[2]
                    zfs = (f'<span class="{"up" if (zf or 0) >= 0 else "dn"}">'
                           f'{zf:+.2f}%</span>'
                           if isinstance(zf, (int, float)) else "")
                    cs.append(f'<span class="zbchip">{_stk(s[0], s[1])} '
                              f'{zfs}</span>')
            if cs:
                chips = f'<div class="zbchips">{"".join(cs)}</div>'
        zb_items.append(
            f'<div class="zbitem"><div class="zbtime">'
            f'{_hm(it.get("Time"))}</div>'
            f'<div><div class="zbtext">{_esc(str(it.get("Comment") or ""))}'
            f'</div>{chips}</div></div>')
    zb_html = "".join(zb_items) or '<p class="mut">暂无直播内容</p>'

    # ---- 盘面亮点 ----
    ld_items = []
    for it in d.get("pmld") or []:
        sx = it.get("TagShuXing")
        tcls = "dn" if sx == 0 else ("mid" if sx == 1 else "up")
        links = "、".join(
            _stk(s[0], s[1]) for s in (it.get("StockList") or [])
            if isinstance(s, (list, tuple)) and len(s) >= 2)
        ld_items.append(
            f'<div class="tlitem"><div class="tltime">'
            f'{_hm(it.get("TimeMin"))}</div><div class="tlmain">'
            f'<div class="tlrow"><span class="tag {tcls}">'
            f'{_esc(str(it.get("TagName") or ""))}</span>{links}</div>'
            f'<div class="tldetail">'
            f'{_esc(str(it.get("Detail") or ""))}</div></div>'
            f'<div class="tlplate">{_esc(str(it.get("ZSName") or ""))}</div>'
            f'</div>')
    ld_html = "".join(ld_items) or '<p class="mut">暂无数据</p>'

    # ---- 大幅回撤 ----
    hctrs = "".join(
        f'<tr><td>{_stk(r[0], r[1])}'
        f'<div class="mut">{_esc(str(r[6] or ""))}</div></td>'
        f'<td class="r {"up" if str(r[2]).startswith("+") else "dn"}">'
        f'{_esc(str(r[2]))}</td>'
        f'<td class="r dn">{(r[3] or 0):.2f}%</td></tr>'
        for r in (d.get("kqxy") or [])
        if isinstance(r, (list, tuple)) and len(r) >= 7)
    hc_html = ('<table class="zttable"><tr><td class="mut">股票</td>'
               '<td class="mut r">涨跌幅</td><td class="mut r">最大回撤</td>'
               '</tr>'
               + (hctrs or '<tr><td class="mut">暂无数据</td></tr>') + "</table>")

    # ---- 涨停天梯（行式表格，一行一只）----
    best = {}
    for g in d.get("tianti") or []:
        for s in g:
            if isinstance(s, (list, tuple)) and len(s) >= 12:
                cur = best.get(s[0])
                if cur is None or (s[2] or 0) > (cur[2] or 0):
                    best[s[0]] = s
    groups = {}
    for s in best.values():
        groups.setdefault(s[2], []).append(s)
    tt_parts = []
    for lb in sorted(groups, reverse=True):
        rows = sorted(groups[lb], key=lambda s: s[3] or 0)
        is_multi = isinstance(lb, int) and lb >= 2
        head = f"{lb} 连板" if is_multi else "首板"
        tcls = ("tier5" if lb >= 5 else "tier4" if lb == 4 else
                "tier3" if lb == 3 else "tier2" if lb == 2 else "tier1")
        trs = "".join(
            f'<tr><td class="tttime">{_hm(s[3])}</td>'
            f'<td class="ttname">{_stk(s[0], s[1])}</td>'
            f'<td class="ttlab">{_esc(str(s[11] or ""))}</td>'
            f'<td class="ttpl">{_esc(str(s[5] or ""))}</td></tr>'
            for s in rows)
        tt_parts.append(
            f'<div class="ttgroup"><div class="shead">'
            f'<span class="tierbadge {tcls}">{lb if is_multi else "首"}</span>'
            f'<b>{head}</b><span class="mut">{len(rows)} 只</span></div>'
            f'<table class="tttab">{trs}</table></div>')
    tt_html = "".join(tt_parts) or '<p class="mut">暂无数据</p>'

    # ---- 龙虎榜动向（席位卡片，买卖分列）----
    lhb_parts = []
    for g in d.get("lhb") or []:
        def fmt(items, absolute=False):
            def amt(x):
                v = (x.get("Money") or 0) / 1e8
                return abs(v) if absolute else v
            return "".join(
                f'<div class="lhbrow">{_stk(x.get("Sto"), x.get("StoN"))} '
                f'<span class="amt">{amt(x):.2f}亿</span>'
                f'</div>'
                for x in (items or [])[:8])

        buy, sell = fmt(g.get("Buy")), fmt(g.get("Sell"), absolute=True)
        if buy or sell:
            lhb_parts.append(
                f'<div class="ghead">{_esc(str(g.get("BName") or ""))}</div>'
                f'<div class="lhbgrid">'
                f'<div class="lhbcol"><div class="mut" '
                f'style="margin-bottom:4px">买入</div>{buy or "—"}</div>'
                f'<div class="lhbcol"><div class="mut" '
                f'style="margin-bottom:4px">卖出</div>{sell or "—"}</div>'
                f'</div>')
    lhb_html = "".join(lhb_parts) or '<p class="mut">暂无数据</p>'

    # ---- 涨停原因（板块卡片）----
    zy_parts = []
    for sec in (d.get("ztyy") or {}).get("list") or []:
        sl = [s for s in (sec.get("StockList") or []) if len(s) >= 17]
        if not sl:
            continue
        rows = []
        for s in sl:
            concepts = "".join(
                f'<span class="concept">{_esc(c.strip())}</span>'
                for c in str(s[11] or "").replace("、", ",").split(",")
                if c.strip())
            rows.append(
                f'<div class="zyrow"><div class="tlrow">{_stk(s[0], s[1])} '
                f'<span class="tag up">{_esc(str(s[9] or ""))}</span>'
                f'{concepts}</div>'
                f'<div class="zyreason">{_esc(str(s[16] or ""))}</div></div>')
        zy_parts.append(
            f'<div class="zycard"><div class="zhead">'
            f'<span class="zname">{_esc(str(sec.get("ZSName") or ""))}</span>'
            f'<span class="zcnt">{len(sl)} 只涨停</span></div>'
            + "".join(rows) + "</div>")
        zy_html = ('<div class="zygrid">' + "".join(zy_parts) + "</div>"
               if zy_parts else '<p class="mut">暂无数据</p>')

    # ---- 指数分时图（腾讯分钟数据，SVG 手绘）----
    def _idx_pane(idd):
        mins, qt = idd.get("mins") or [], idd.get("qt") or []
        pts = []
        for row in mins:
            f = str(row).split()
            if len(f) >= 2:
                try:
                    pts.append((f[0], float(f[1])))
                except ValueError:
                    pass
        if not pts or len(qt) <= 4:
            return '<p class="mut">暂无数据</p>'
        try:
            pre, nowp = float(qt[4]), float(qt[3])
        except (TypeError, ValueError):
            pre, nowp = pts[0][1], pts[-1][1]

        def _x(hm):
            t = int(hm[:2]) * 60 + int(hm[2:]) - 570
            return t - 90 if t > 120 else t

        xs = [_x(hm) for hm, _ in pts]
        ys = [p for _, p in pts]
        hi0, lo0 = max(ys), min(ys)
        lo, hi = min(ys + [pre]), max(ys + [pre])
        pad = (hi - lo) * 0.1 or 1.0
        lo, hi = lo - pad, hi + pad
        W, H, PT, PB = 560, 190, 8, 18

        def X(x):
            return x / 240 * W

        def Y(v):
            return PT + (hi - v) / (hi - lo) * (H - PT - PB)

        line = " ".join(f"{X(x):.1f},{Y(v):.1f}" for x, v in zip(xs, ys))
        area = (f"{X(xs[0]):.1f},{H - PB} " + line +
                f" {X(xs[-1]):.1f},{H - PB}")
        chg = (nowp / pre - 1) * 100 if pre else 0
        cls = "up" if chg >= 0 else "dn"
        col = "#e0492f" if chg >= 0 else "#1e9e62"
        return (
            f'<div class="idxbig"><b class="{cls}">{nowp:.2f}</b>'
            f'<span class="{cls}">{chg:+.2f}%</span></div>'
            f'<svg viewBox="0 0 {W} {H}" style="width:100%;display:block">'
            f'<polygon points="{area}" fill="{col}" opacity="0.13"/>'
            f'<line x1="0" x2="{W}" y1="{Y(pre):.1f}" y2="{Y(pre):.1f}" '
            f'stroke="#8fa5bd" stroke-dasharray="4 4" stroke-width="1"/>'
            f'<polyline points="{line}" fill="none" stroke="{col}" '
            f'stroke-width="1.8"/>'
            f'<text x="0" y="{H - 4}" fill="#8fa5bd" font-size="11">09:30</text>'
            f'<text x="{W / 2 - 14}" y="{H - 4}" fill="#8fa5bd" '
            f'font-size="11">11:30</text>'
            f'<text x="{W - 30}" y="{H - 4}" fill="#8fa5bd" '
            f'font-size="11">15:00</text></svg>'
            f'<div class="capnote">昨收 {pre:.2f} · 最高 {hi0:.2f} · '
            f'最低 {lo0:.2f}（虚线为昨收）</div>')

    idx_html = _idx_pane(d.get("idx") or {})
    idxcy_html = _idx_pane(d.get("idxcy") or {})
    idxkc_html = _idx_pane(d.get("idxkc") or {})

    # ---- 日韩（Yahoo 5 分钟线）----
    def _spark(closes, pre, H=84):
        if not closes:
            return ""
        lo, hi = min(closes + [pre]), max(closes + [pre])
        pad = (hi - lo) * 0.1 or 1.0
        lo, hi = lo - pad, hi + pad
        W = 560
        n = max(len(closes) - 1, 1)

        def Y(v):
            return 4 + (hi - v) / (hi - lo) * (H - 8)

        line = " ".join(f"{i / n * W:.1f},{Y(v):.1f}"
                        for i, v in enumerate(closes))
        col = "#e0492f" if closes[-1] >= pre else "#1e9e62"
        return (f'<svg viewBox="0 0 {W} {H}" style="width:100%;display:block">'
                f'<line x1="0" x2="{W}" y1="{Y(pre):.1f}" y2="{Y(pre):.1f}" '
                f'stroke="#8fa5bd" stroke-dasharray="4 4" stroke-width="1"/>'
                f'<polyline points="{line}" fill="none" stroke="{col}" '
                f'stroke-width="1.6"/></svg>')

    ww = d.get("idxrh") or {}
    # ---- 韩国KOSPI（独立卡）----
    kp = ww.get("kospi") or {}
    if kp.get("price") and kp.get("pre"):
        price, pre = kp["price"], kp["pre"]
        closes = kp.get("closes") or []
        chg = (price / pre - 1) * 100
        cls = "up" if chg >= 0 else "dn"
        hi_s = f"{max(closes):,.2f}" if closes else "—"
        lo_s = f"{min(closes):,.2f}" if closes else "—"
        kospi_html = (
            f'<div class="idxbig"><b class="{cls}">{price:,.2f}</b>'
            f'<span class="{cls}">{chg:+.2f}%</span></div>'
            + _spark(closes, pre) +
            f'<div class="capnote">昨收 {pre:,.2f} · 最高 {hi_s} · '
            f'最低 {lo_s}（虚线为昨收；现货来自 Yahoo，国内暂无源）</div>')
    else:
        kospi_html = '<p class="mut">暂无数据（网络不通时为空）</p>'
    # ---- 外围期货（新浪 hf_，五格一块）----
    tiles = []
    for it in ww.get("items") or []:
        price, pre = it["price"], it["pre"]
        chg = (price / pre - 1) * 100 if pre else 0
        cls = "up" if chg >= 0 else "dn"
        tiles.append(
            f'<div class="wwtile"><div class="wwname">'
            f'{_esc(it["name"])}</div>'
            f'<div class="wwprice {cls}">{price:,.2f}</div>'
            f'<div class="wwchg {cls}">{chg:+.2f}%</div>'
            f'<div class="wwsub">昨收 {pre:,.2f}<br>'
            f'高 {it["hi"]:,.2f} · 低 {it["lo"]:,.2f}</div></div>')
    ww_html = ('<div class="wwgrid">' + "".join(tiles) + "</div>"
               if tiles else '<p class="mut">暂无数据</p>')

    # ---- 盘中异动标注图（上证分时 + 盘面亮点标注）----
    _mins = d.get("idx") or {}
    _rows = _mins.get("mins") or []
    _meta = _mins.get("qt") or []
    yidong_html = '<p class="mut">暂无数据</p>'
    if _rows and len(_meta) > 35:
        _pre = float(_meta[4])
        _pts = []
        for r in _rows:
            _sp = r.split()
            if len(_sp) >= 2:
                _pts.append((_sp[0], float(_sp[1])))
        _prices = [p for _, p in _pts]

        def _xof(hhmm):
            t = int(hhmm[:2]) * 60 + int(hhmm[2:]) - 570
            if t > 120:
                t -= 90
            return 8 + t / 240 * 544

        def _price_at(hhmm):
            best = _pts[0][1]
            for t, p in _pts:
                if t <= hhmm:
                    best = p
                else:
                    break
            return best

        _evs, _seen = [], set()
        for e in sorted(d.get("pmld") or [],
                        key=lambda x: x.get("TimeMin") or 0):
            if not e.get("TimeMin"):
                continue
            hhmm = time.strftime("%H%M",
                                 time.gmtime(e["TimeMin"] + 8 * 3600))
            if hhmm < "0925" or hhmm > "1500":
                continue
            lab = (e.get("ZSName") or e.get("TagName") or "").strip()
            if not lab or lab in _seen:
                continue
            _seen.add(lab)
            _evs.append((hhmm, lab, e.get("TagShuXing")))
        _evs = _evs[:16]
        lo = min(_prices + [_pre])
        hi = max(_prices + [_pre])
        padv = (hi - lo) * 0.08 or 1.0
        lo, hi = lo - padv, hi + padv
        Y0, Y1 = 64, 196

        def _yof(p):
            return Y0 + (hi - p) / (hi - lo) * (Y1 - Y0)

        is_up = _prices[-1] >= _pre
        colr = "#e03131" if is_up else "#2ea44f"
        path = "M " + " L ".join(
            f"{_xof(t):.0f} {_yof(p):.0f}" for t, p in _pts)
        pre_y = _yof(_pre)
        marks, levels = [], []
        for hhmm, lab, sx in _evs:
            x = _xof(hhmm)
            lvl = 0
            while lvl <= 2:
                while len(levels) <= lvl:
                    levels.append(-100)
                if x - levels[lvl] >= 52:
                    break
                lvl += 1
            if lvl > 2:
                continue
            levels[lvl] = x
            ly = 16 + lvl * 18
            wch = len(lab) * 13 + 12
            cx = min(max(x, wch / 2 + 2), 560 - wch / 2 - 2)
            py = _yof(_price_at(hhmm))
            if sx == 2:
                bg, fg = "rgba(224,49,49,.28)", "#ffb3b3"
            elif sx == 0:
                bg, fg = "rgba(46,160,90,.30)", "#a9dfbf"
            else:
                bg, fg = "rgba(255,255,255,.16)", "#dbe4f0"
            marks.append(
                f'<line x1="{x:.0f}" y1="{ly + 4}" x2="{x:.0f}" '
                f'y2="{py:.0f}" stroke="{fg}" stroke-width="1" '
                f'opacity=".65"/>'
                f'<circle cx="{x:.0f}" cy="{py:.0f}" r="2.4" fill="{fg}"/>'
                f'<rect x="{cx - wch / 2:.0f}" y="{ly - 11}" width="{wch}" '
                f'height="15" rx="4" fill="{bg}"/>'
                f'<text x="{cx:.0f}" y="{ly}" font-size="11.5" '
                f'font-weight="700" fill="{fg}" '
                f'text-anchor="middle">{_esc(lab)}</text>')
        chg = (_prices[-1] / _pre - 1) * 100
        lcls = "up" if is_up else "dn"
        yidong_html = (
            f'<div class="idxbig"><b class="{lcls}">{_prices[-1]:,.2f}</b>'
            f'<span class="{lcls}">{chg:+.2f}%</span>'
            f'<span class="mut" style="font-size:12px">上证 · 标注来自'
            f'开盘啦盘面亮点（板块异动）</span></div>'
            f'<svg viewBox="0 0 560 214" style="width:100%;height:auto;'
            f'display:block">'
            f'<line x1="8" y1="{pre_y:.0f}" x2="552" y2="{pre_y:.0f}" '
            f'stroke="#5b7a99" stroke-width="1" stroke-dasharray="4 4"/>'
            f'<path d="{path}" fill="none" stroke="{colr}" '
            f'stroke-width="1.6"/>'
            + "".join(marks) +
            f'<text x="8" y="210" font-size="10" fill="#5b7a99">09:30</text>'
            f'<text x="268" y="210" font-size="10" fill="#5b7a99">'
            f'11:30 / 13:00</text>'
            f'<text x="528" y="210" font-size="10" fill="#5b7a99">15:00</text>'
            f'</svg>')

    return {"mood": mood_html, "plate": plate_html, "zt": zt_html,
            "plateall": plateall_html, "idx": idx_html, "idxcy": idxcy_html,
            "idxkc": idxkc_html, "kospi": kospi_html, "ww": ww_html,
            "yidong": yidong_html,
            "dt": dt_html, "zb": zb_html, "ld": ld_html, "hc": hc_html,
            "tt": tt_html, "lhb": lhb_html, "zy": zy_html,
            "t": d["t"], "day": d["day"]}


def _disabled_page(name):
    return f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8"><title>模块已关闭</title>
<style>body{{background:#0d1526 url('/wallpaper.jpg') no-repeat center center
fixed;background-size:cover;font-family:sans-serif}}
.box{{max-width:520px;margin:18vh auto 0;background:rgba(255,255,255,.95);
border-radius:16px;padding:34px 30px;text-align:center;
box-shadow:0 12px 36px rgba(0,0,0,.4)}}
.box h2{{margin:0 0 10px;font-size:21px;color:#1c2b3a}}
.box p{{color:#5a6a7d;line-height:1.7}}
.box a{{display:inline-block;margin:8px 6px 0;background:#2f6fb3;
color:#fff;padding:9px 20px;border-radius:999px;text-decoration:none;
font-weight:600}}</style></head>
<body><div class="box"><h2>「{name}」模块已关闭</h2>
<p>入口已全部隐藏。想恢复的话，去设置里把这个模块重新打开。</p>
<a href="/home">🏠 回主页</a><a href="/settings">⚙ 去设置</a>
</div></body></html>"""


def _secnav(active):
    _s = _load_settings()
    items = [("home", "/home", "🏠 主页"),
             ("report", "/report", "📈 报告列表"),
             ("kpl", "/kpl", "📊 开盘啦"),
             ("news", "/news", "📰 快讯"),
             ("favs", "/favs", "⭐ 收藏夹"),
             ("settings", "/settings", "⚙ 设置")]
    items = [it for it in items
             if not (it[0] == "kpl" and not _s.get("kpl_enabled", True))
             and not (it[0] == "news"
                      and not _s.get("news_enabled", True))]
    links = "".join(
        '<a href="#" class="%s" onclick="goNav(\'%s\');return false">%s</a>'
        % ("on" if key == active else "", url, label)
        for key, url, label in items)
    return '<div class="secnav">' + links + "</div>"


def render_kpl():
    d = _kpl_data()
    f = _kpl_fragments(d)
    now = datetime.now(LOCAL_TZ)
    _st = _load_settings()
    _kpl_chk = " checked" if _st.get("kpl_auto", True) else ""
    _kpl_sec = int(_st.get("kpl_interval_sec", 300) or 300)
    _opts = "".join(
        f'<option value="{_v}">{_lab}</option>'
        for _v, _lab in ((15, "15秒"), (30, "30秒"), (45, "45秒"),
                         (60, "1分钟"), (90, "90秒"), (120, "2分钟"),
                         (180, "3分钟"), (240, "4分钟"), (300, "5分钟"),
                         (600, "10分钟"), (900, "15分钟"),
                         (1800, "30分钟")))
    dateline = f"{now:%Y-%m-%d} · 更新 {d['t']}"
    return f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<title>选股报告 · 开盘啦数据</title>
<style>{_PAGE_CSS}{_NAV_CSS}{_KPL_CSS}</style></head><body>
<div class="kplwrap">
  <div class="navtop"><div class="navttl">📊 开盘啦 · 实时数据</div>
  <span style="display:flex;gap:12px;align-items:center">
  <button class="kplrefresh" id="kplrefresh"
    onclick="kplRefresh()">↻ 刷新数据</button>
  <label class="kplsw"><input type="checkbox" id="kplauto"{_kpl_chk}
    onchange="kplAutoSet(this)"> 自动更新</label>
  <input class="kplsel" id="kplsec" list="kplsecs" value="{_kpl_sec}"
    style="width:62px" onchange="kplSecSet(this)">
  <datalist id="kplsecs">{_opts}</datalist>
  <span class="kplsw">秒</span>
  <span class="navdate" id="kpldate">{dateline}</span>
  <button class="kplrefresh" onclick="toggleFS()">⛶ 全屏</button></span></div>
  {_secnav('kpl')}
  <div class="kplgrid2">
    <div class="lstack">
    <div class="kcard" id="moodcard">{f['mood']}</div>
    <div class="kcard" id="idxcard">
      <div class="khead"><span>📈 指数走势</span>
      <span class="pills"><span class="spill on"
        onclick="idxTab('sh', this)">上证</span><span class="spill"
        onclick="idxTab('cy', this)">创业板</span><span class="spill"
        onclick="idxTab('kc', this)">科创50</span><span class="spill"
        onclick="idxTab('kp', this)">韩国</span></span></div>
      <div id="idx-sh">{f['idx']}</div>
      <div id="idx-cy" style="display:none">{f['idxcy']}</div>
      <div id="idx-kc" style="display:none">{f['idxkc']}</div>
      <div id="idx-kp" style="display:none">{f['kospi']}</div>
    </div>
    </div>
    <div class="kcard">
      <div class="khead"><span>🔥 板块强度</span>
      <span class="pills"><span class="spill on" id="pill-pt"
        onclick="plateTab('top')">前五后五</span><span class="spill"
        id="pill-pa" onclick="plateTab('all')">全部</span></span></div>
      <div id="plate-top">{f['plate']}</div>
      <div id="plate-all" class="ztscroll" style="display:none">{f['plateall']}</div>
    </div>
    <div class="kcard">
      <div class="khead"><span>🔒 涨停 · 跌停（按封单排）</span>
      <span class="pills"><span class="spill on" id="pill-zt"
        onclick="ztTab('zt')">涨停</span><span class="spill" id="pill-dt"
        onclick="ztTab('dt')">跌停</span></span></div>
      <div id="ztwrap" class="ztscroll">{f['zt']}</div>
      <div id="dtwrap" class="ztscroll" style="display:none">{f['dt']}</div>
    </div>
  </div>
  <div class="kplgrid3">
    <div class="kcard">
      <div class="khead"><span>🌐 外围期货</span>
      <span class="kday">新浪财经 · 更新 {d['t']}</span></div>
      <div id="futureswrap">{f['ww']}</div>
    </div>
    <div class="kcard">
      <div class="khead"><span>📍 盘中异动标注</span>
      <span class="kday">{d['day']} · 上证分时</span></div>
      <div id="yidongwrap">{f['yidong']}</div>
    </div>
  </div>
  <div class="kplgrid3">
    <div class="kcard">
      <div class="khead"><span>📺 大盘直播</span>
      <span class="kday">{d['day']}</span></div>
      <div class="zbfeed" id="zbwrap">{f['zb']}</div>
    </div>
    <div class="kcard">
      <div class="khead"><span>🧾 复盘啦</span>
      <span class="fptabs"><span class="ftab on"
        onclick="fpTab('shuli', this)">盘面梳理</span><span class="ftab"
        onclick="fpTab('dongxiang', this)">市场动向</span><span class="ftab"
        onclick="fpTab('yuanyin', this)">涨停原因</span></span></div>
      <div class="fpane" id="pane-shuli">
        <div class="subpills"><span class="spill on" id="pill-ld"
          onclick="fpSub('ld')">盘面亮点</span><span class="spill"
          id="pill-hc" onclick="fpSub('hc')">大幅回撤</span></div>
        <div id="sub-ld">{f['ld']}</div>
        <div id="sub-hc" style="display:none">{f['hc']}</div>
      </div>
      <div class="fpane" id="pane-dongxiang" style="display:none">
        <div id="ttwrap">{f['tt']}</div>
      </div>
      <div class="fpane" id="pane-yuanyin" style="display:none">
        <div id="zywrap">{f['zy']}</div>
      </div>
    </div>
  </div>
  <p style="color:#93a7bf;font-size:13px;margin-top:14px">
  数据来自开盘啦公开接口（跌停名单为东方财富口径），每 5 分钟自动刷新；
  盘中为实时值，收盘后为当日收盘值。点股票名跳经典版。</p>
</div>
<div class="toast" id="toast"></div>
<button class="grabbtn" id="grabbtn" data-tip="先复制文章链接，再点这里：抓取正文＋评论，存进报告列表"
  onclick="grabArticle()">🔗 抓文章</button>
<button class="fetchbtn" id="fetchbtn" data-tip="立即拉取最新报告邮件（最近 5 封）"
  onclick="manualFetch()">↻ 获取邮件</button>
<button class="quitbtn" onclick="goNav('/home')">🏠 主页</button>
<script>{_NAV_JS}</script>
<script>{_KPL_JS}</script>
</body></html>"""


def render_index():
    with LOCK:
        manifest = load_manifest()
    entries = sorted(manifest, key=lambda e: (e["date"], e["time"]),
                     reverse=True)
    if not entries:
        return _page("选股报告 · 本地", "",
                     "自动从邮箱下载、按日期存放；打开窗口时自动刷新"
                     "列表，有新报告会弹系统通知。",
                     _searchbar_html() +
                     '<p style="padding:4px 22px 18px">还没有报告，'
                     '第一封到达后会自动出现在这里。</p>')
    groups = _grouped(entries)
    latest = max(groups)
    parts = [_searchbar_html(), _datebar(latest, len(groups[latest])),
             _rows_html(groups[latest])]
    older = [d for d in groups if d != latest]
    if older:
        parts.append(
            f'<div class="foot" onclick="goNav(\'/history\')">'
            f'<span>🗂 历史报告（{len(older)} 天 · '
            f'{sum(len(groups[d]) for d in older)} 份）</span>'
            f'<span class="arrow">查看 →</span></div>')
    return _page("选股报告 · 本地",
                 f"{latest} · 最新一天 · 更新 {datetime.now():%H:%M}",
                 "点一行打开完整版；行尾按钮看正文，或在访达里定位源文件。",
                 "".join(parts))


def render_history():
    with LOCK:
        manifest = load_manifest()
    entries = sorted(manifest, key=lambda e: (e["date"], e["time"]),
                     reverse=True)
    groups = _grouped(entries)
    parts = ['<div class="foot" onclick="goNav(\'/\')">'
             '<span>← 回到最新一天</span></div>']
    parts.insert(1, _searchbar_html())
    for d in sorted(groups, reverse=True):
        parts.append(_datebar(d, len(groups[d])))
        parts.append(_rows_html(groups[d]))
    return _page("选股报告 · 历史",
                 f"共 {len(groups)} 天 · {len(entries)} 份",
                 "全部历史报告，按日期排列。", "".join(parts))



# ---------------- 检索 ----------------

def _searchbar_html(q=""):
    return (f'<form class="searchbar" onsubmit="return doSearch(this)">'
            f'<input type="text" value="{_esc(q)}" '
            f'placeholder="搜索全部报告：股票名、代码、关键词…" '
            f'autocomplete="off">'
            f'<button type="submit">搜索</button></form>')


def _plain_text(path):
    try:
        with open(path, encoding="utf-8", errors="replace") as fh:
            raw = fh.read()
    except OSError:
        return ""
    raw = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", raw)
    raw = re.sub(r"(?s)<[^>]+>", " ", raw)
    raw = (raw.replace("&nbsp;", " ").replace("&amp;", "&")
              .replace("&lt;", "<").replace("&gt;", ">")
              .replace("&quot;", '"'))
    return re.sub(r"\s+", " ", raw).strip()


def _highlight(seg, q):
    out = []
    low, ql, pos = seg.lower(), q.lower(), 0
    while True:
        i = low.find(ql, pos)
        if i < 0:
            out.append(_esc(seg[pos:]))
            break
        out.append(_esc(seg[pos:i]))
        out.append(f'<span class="hl">{_esc(seg[i:i + len(q)])}</span>')
        pos = i + len(q)
    return "".join(out)


def _snippet(text, q, width=40):
    i = text.lower().find(q.lower())
    if i < 0:
        return ""
    start = max(0, i - width)
    end = min(len(text), i + len(q) + width)
    seg = _highlight(text[start:end], q)
    return ("…" if start > 0 else "") + seg + ("…" if end < len(text) else "")


def render_search(q):
    q = (q or "").strip()
    with LOCK:
        manifest = load_manifest()
    entries = sorted(manifest, key=lambda e: (e["date"], e["time"]),
                     reverse=True)
    results = []
    if q:
        root = reports_root()
        for e in entries:
            main = e.get("full") or e.get("body")
            if not main:
                continue
            snip = _snippet(_plain_text(os.path.join(root, main)), q)
            if not snip and q.lower() in e["subject"].lower():
                snip = "标题命中：" + _highlight(e["subject"], q)
            if snip:
                results.append((e, snip))
    bar = _searchbar_html(q)
    if not q:
        inner = (bar + '<p style="padding:6px 22px 18px">输入关键词后'
                 '点搜索：股票名、代码、任意词都行，搜标题＋正文全文。</p>')
    elif not results:
        inner = (bar + f'<p style="padding:6px 22px 18px">没找到'
                 f'「{_esc(q)}」，换个词试试（比如股票代码）。</p>')
    else:
        rows = []
        for i, (e, snip) in enumerate(results):
            cls = "srow alt" if i % 2 == 1 else "srow"
            rows.append(f'<div class="{cls}">' + _row_html(e)
                        + f'<div class="snipline">{snip}</div></div>')
        inner = (bar + f'<div class="note" style="padding-top:12px">'
                 f'「{_esc(q)}」共找到 {len(results)} 份 · 按日期倒序</div>'
                 + '<div class="rows">' + "".join(rows) + "</div>")
    return _page("选股报告 · 搜索",
                 f"共 {len(results)} 份" if q else "",
                 "搜全部报告（最新＋历史）的标题和正文。", inner)


# ---------------- 导航主页 ----------------

def _tiles_path():
    return os.path.join(reports_root(), "nav_tiles.json")


DEFAULT_TILES = [
    {"name": "东方财富", "url": "https://www.eastmoney.com",
     "color": "#c82828", "ch": "财"},
    {"name": "财联社", "url": "https://www.cls.cn",
     "color": "#d97706", "ch": "电"},
    {"name": "淘股吧", "url": "https://www.tgb.cn",
     "color": "#2563eb", "ch": "淘"},
    {"name": "知识星球", "url": "https://wx.zsxq.com",
     "color": "#0891b2", "ch": "星"},
    {"name": "开盘啦", "url": "https://www.kaipanla.com",
     "color": "#dc2626", "ch": "开"},
    {"name": "163邮箱", "url": "https://mail.163.com",
     "color": "#be1e2d", "ch": "邮"},
    {"name": "雪球", "url": "https://xueqiu.com",
     "color": "#0e74e9", "ch": "雪"},
    {"name": "韭研公社", "url": "https://www.jiuyangongshe.com",
     "color": "#16a34a", "ch": "韭"},
    {"name": "腾讯行情", "url": "https://gu.qq.com",
     "color": "#64748b", "ch": "讯"},
]
_TILE_COLORS = ["#2563eb", "#0891b2", "#16a34a", "#d97706",
                "#dc2626", "#7c3aed"]
_TILE_ICONS = {'东方财富': 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAADcklEQVR4nL2XXYhVVRTHfzZL7aWiD1eZYaESpeJLIYof9WBRvRTjrqBeilIsHEQQcV70IYoJsgKLxoIKAtNa4gdCPU1IUkKRPmWkzSA12mzKolGqYWk9nH1jz/XcM/fKsfWy91nrv/f/f85Ze6+9oWaLQabFINNjkAfS8/VV+Ek1k88EHgVeA04A84E1wFE1Hygbc0WdAhJxl5pfAK4D1gHvAjtikEWXTUAMMikGeQnoBn5O7quAtcAoMADsiUGm1S4g/eOdQG9yHYpBbk1z3wwsBfYANwHbahMQg9wQg/QCx4HHkvsLNR8EFmbQpcBXqf94DJLHkA5JrwZWUCRaNzClCfJialdmvjuAk8BfwJXApjQWqFgFMciUNHgecDewiOLNWonereYhBpkODCYygH1q/kgMMgIo4MAtaj5CPlkM0kWRtfcAdwKzgK5WApvsOLAq9V/IyEmEAGOpFeBB4P1xAoAeYGubhLkNAw+p+W8xyH3AM03xU6nNV8C9DQF5Ej51CeTfAcvU/EQMcjuwqwRzJAa5EZia+eY3OrmAuR2S7wcWq/lQDDKHYq1fW4IbAO5q8s0sEzC5TeJR4Dk1f1jNf49BlgNfAjNKsAfV/CTFJ8/tv9+R58B5Jk66HcAGNT8dg0ymWFJbKsZtT213i/g4AT8Ct7XAHQA2q/kRgBhkCdBP9i9L7BjwcQyyDJjdjoChJgFngQ+BfjX/JhHPAfoYv9G0so1q7jFIT0lsuEzAIMUbfQLsBT5V8z8bwRjkCeA9Lt79yuwjNT8Qg8wDQkl8qEzARmB1KqXjLAaZTbFu20nUU8Dzqf8q5bvt4EUC1PxMxaQr2iQfA1aq+a8xyJPA/S1wXzc67VbDP9rEPavmh2OQGcCbFbhDnQr4vA3MejX/IBWxXcA1LXAjwNGOBKj5T8BnFZBeNX899fuBJRXYvWr+T0cCkr1R4rsArFHzPoAYZDXw9ATz7MwfOhGwn2z9AueAbjXfnsinUuwRVfYDcPCSBKi5A2+nx++BhWq+L4MsoLwY5fZW/vmhwyNZEjAL6FHz0abY+QnGjgLvNDtru5ik7D9NcR8os5fVfFOzs7aLiZqP0fpEdQ54pSxQ981oK/Btib9PzX+57ALU/G+K4pPvnMNUnDXr/gKo+TGKGtAQsT6vqv+bxSCLY5DdE+H+BeFcBIEPiY3nAAAAAElFTkSuQmCC', '财联社': 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABWUlEQVR4nKVTMUgCURj+7jgeImFHRIkKDjfddBCIm5ubCCFOEbg6JLjb1hIEQUtDQ1E4hRA2upTbudnieF14EC5pRRx2eQ1yvz7PQO2DB/e+/77vfe/nf4LBFBnAKYAClsMVgLJgMOVyBbGHB8FgijvLSvEYQgcFME0FADimBftRx+dNzefgM5APS5ArpbnH2U0dvXwRo/47ceL0D8FsmhM7pgW7qdM+kEpi8+KYdzWY4nrr+7nrevi4rhFvJTLuz9uAalYiQzW6AtNURFr3ZPyyvcNF/Qt0BXE9ROSw3VlI7OvBKpC8j9FgciLTVIhyiEshxWOQ4tFxwqdJQkowbHfgmBYJNk4qnGGkVUe4UUW4USUjYGYOgtk0tm7PqeiYFhyzi0AqSZzd1PGa3qM914OvegP9o7Op2FGfuJcvcj2YO8pMU7G2n1t4lP/1mEQAZYyf5rK4A7D7C/jGnxodwm+yAAAAAElFTkSuQmCC', '淘股吧': 'data:image/png;base64,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', '知识星球': 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAC+klEQVR4nL3Xz2tdRRQH8E8uIUgIQUq5dFECDaIgUqUbdSPBtlwQ0Y3WH9Uu7OIuXFhxIUVcdKHFlv4Bj1Ytgr/AH1nYykOhBUWluBJxIS6kFKmXUCSUUEIJLmZeMpnc17427/W7uvfMzPl+Z+bcc84dMyDKbmcaT2A3dmIWd8fh/3ABh5uq/m1QnzA+APE9OIwXcVefaVujuK14eCgCym5nEkdwaBChERO3Qt5XQNz113igz7pl/Ikmvm/BIt5s8fUI7sevbdcz1rJgF75FmQ0t4TQ+xy9NVS/3EZf6eh+vJKa3mqp+t6+AuPOfhbtMcUoIsIWbkSa+5nAuM6/gvqaq/+oZxpMFk/gyI1/C/qaq5wclTrCzxVZE+0YBQsCli5awt6nqn26DHNo+x5XcXrB69IeyyS9vglxT1efxQWZ+Oz3+VQHCd56exummqr+6XfJExEE8ioN4MA9AGCu7nSn8i8lou4YdTVVf3qyAQVDgyYQcPrtT5IRj353ZPh4lYdntPIUPhcT19Lj1kX8dP45SAN4TMucWnCiEqtbDxaaqr41YQBrsDxXWSipruX2UuJI8TxdCcuihcIdRCMHQQ16ARoGUY7GQ5GXMxJowGuaQc2YS09+F9bm5wGOjEhB9p9f8e4Hvskn7Rygg932uwFlcTYz7ym5n+7CZy25nBs8kpus4WzRVvYSPkoEJHB+2gOgz7Rnnm6pe6N3HUaEI9fB82e0cGBZz9LWvRVAIiKaqL+GdbMLJstvZMwTyPTiZmT9pqvrCqoCIY0gbkAmcKbudlzZBfgBnrD/6y3it95I3pduFpjQPwi/wRlPVFwcknsEJ64OO0M5XsVvaKCAunhU+zdlsaDkK+RTnm6q+mq2bwhxeEO47/+dYxnN5g7tBQHS2LRLNtY1HXLJWvEobTy3FAp5Nd35DAVFEITSqRzB1A+c3wzxebar6n7bBvgISISVeF/5wBi1WK/geR9t2fUsCEiHjeBx7sQv3YlroJxeFK/kDP+CbQQP2f2e21HWzJrCEAAAAAElFTkSuQmCC', '开盘啦': 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFQUlEQVR4nL2XTWgVVxTHf3NnJjN5iT5LTBC/0cbUD9xkJUpBjUgKoihi0bS4kEoXQhEXAcF01YWCCzciLiylXbhQEb+QVAz6VFBs/YBEDSQ0TdWYNFFrkjd5797Txbw3b95LikpDDhy49z//c8+5c8+5HxATAwmTSHxj4JyBpwaGDATG87QBidT3pag/ERbaBLkxnho4lxs7wUSiYY2GP7Tvi4ZiLcX+H+cPDWvyfh2ADKwx0Ar4E0Y3uTIfaM3ABhdSKoCEwC8CvgBTpL7ALwEkHEkkmsSY+VF8njc+5lJscjjzUarJMSMjjeOI6fR441JscjiNjoFl41lTJsuUgRoDTKa6P/yAd+kS7pEj7+PWODpWl9NfvMCqrPzg8N9Mmxa1/eZm7Lo6AOz161HV1Zi6Oqy5cwsGts3o/v2Y3t48krAGPU9jWQpgRn//RwUwVF4etSuvXsX9/PP3B11fj2lvDzsixtFBoPIfM3fvYlVWYn/2GWr6dExfH7q7G5RCffIJdm1txMMYdDyhjAFg7No1dGcn/u7dmN5exi5eBNumfN++kBYEcTvFS5C49vm+BDdviojI8PHjEfZ6xw7JS19lpZTaBdevi4jIm7175fXOnWKMEWOMvP7yS+nz/ch2YPnyIjtHl/wiK9YWQOcwE8NNrO9v2oS7ciVqwQIAyhobydy7h+7owF68GLu+nvJcbgD4e/cSXLtG+vx5ABzxS3Zfz0NUblVsG/H9EHPdQgCeh2SzIX37diq++qrgYPNm/M2bo37lgQNFw1fs24c1cyajV6+GAWRz6+GtWoWzaBGW66KqqkL/ixbhbduG5bq49fWFGLdsQdJpMo8eMfb4MerXX/HXrcNSirGHDzFv3+KtXk1w40YYqFKUr1sHQHDnDsGjR+T90gXSBfL21Cn5WBlsaZEukOcNDSIiYoyRP5culdG2NhEReXf6tAx+/728O3MmsulZvFjyPrtAlM6ts/DxIoA1bx7Vp04BMNrWRrqjg8HDhzHpNGrePNyVK6nYujWc/YMHBF1d5H1qIErCfy5cINPbC47DjD17sGfOZCSVYrStDRyHsiVLmJYb6O/Dh2FsjJEbN8gOD6PfvMGePZt3Z8/ib9yI0Zqho0dJ37uHt2IFOA56YICR1lbUwoVkursLs+jwfYnrk2RSMi9fiojIq0OHIqy3qSn6jU+qqopsBo4c+eBl6zt4sMg2SsK8uLNmYdfUAJB+9oxsOh2WYSYTcXQQYGJ2JlcRHyImmyXuU8XXQwPJr7/GsixEhHe3b6MprntK+PH8eXv5Mi9bWsJlOnGCwR9/BOCv775j+NYtoLC35FUVnWK1tVTn6nY4lSLd04PJESfaiAxgJZNY+UPJtsFxCm3bDtuuC7m9xUomsZLJwqTue56+7/vyuK5ORjs7o3J69sUXct/35b7vy2/JpHTFcuD3qqroW9+xYx9dvn3HjoX2nqedbBBkgTIHUInwZO4/eZKhy5ejGVuAjuVANpYDRpdu5u8Xo3U+D7LWbRgCZgD4n37KnOZmur/9tijplO9TNmcOyYYGAAZ++gk9OgpA2ezZuNXVWJ6HBEEh6JJ+HMv09zP2/DnAaysFT4ElcWempDJKscniAM8cA+3xAKA44f4LmyROu2NVVFwRrbdE0NRdy7Fs+4p1PbwTdkj4YpmyJbCgB1iq1sKIgV0C6Sl8GaUN7FoLIwpgPaQENhjomewr+gTaI7BhPaQg9zgFaIBUKyxFqSaBRsIHSw2QEBFHILq8isi447sUExEjkAVGgFdAO0pdAX5uCDEA/gUZaDE5F6a7OAAAAABJRU5ErkJggg==', '163邮箱': 'data:image/png;base64,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', '雪球': 'data:image/png;base64,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', '韭研公社': 'data:image/png;base64,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', '腾讯行情': 'data:image/png;base64,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'}



def load_tiles():
    tiles = load_json(_tiles_path(), None)
    if not isinstance(tiles, list):
        return [dict(t) for t in DEFAULT_TILES]
    return [t for t in tiles if isinstance(t, dict) and t.get("url")]


def _tile_html(t, i):
    name = _esc(t.get("name", ""))
    url = _esc(t.get("url", ""))
    icon = t.get("icon") or _TILE_ICONS.get(t.get("name", ""), "")
    if icon:
        inner = f'<img src="{icon}" alt="">'
    else:
        ch = _esc(t.get("ch") or (t.get("name", "？")[:1]))
        inner = f'<span class="ch">{ch}</span>'
    return (f'<div class="tile" data-url="{url}" '
            f'onclick="tileClick(this, event)">'
            f'<span class="tilex" data-i="{i}" '
            f'onclick="tileDel(this, event)">✕</span>'
            f'<div class="tileicon">{inner}</div>'
            f'<div class="tilename">{name}</div></div>')


_NAV_CSS = """
body { background:#060a14 url('/wallpaper.jpg') center top/cover no-repeat;
       min-height:100vh; }
.navwrap { max-width:1280px; margin:0 auto 26px; }
.navtop { display:flex; align-items:baseline; padding:10px 6px 0; }
.navttl { color:#fff; font-size:25px; font-weight:bold;
  text-shadow:0 2px 12px rgba(0,0,0,.45); }
.navdate { margin-left:auto; color:rgba(230,240,252,.88); font-size:14px;
  text-shadow:0 1px 8px rgba(0,0,0,.5); }
.modes { display:flex; justify-content:center; margin:24px 0 14px; }
.modes .mode { border:1px solid rgba(255,255,255,.35); color:#e6eef8;
  background:rgba(10,22,44,.38); padding:7px 26px; font-size:15px;
  cursor:pointer; }
.modes .mode:first-child { border-radius:19px 0 0 19px; }
.modes .mode:last-child { border-radius:0 19px 19px 0; }
.modes .mode.on { background:rgba(255,255,255,.94); color:#1f3a5f;
  font-weight:bold; }
.navsearch { display:flex; }
.navsearch input { flex:1; min-width:0; font-size:17px; padding:14px 24px;
  border-radius:28px; border:1px solid rgba(255,255,255,.38);
  background:rgba(9,19,38,.46); color:#fff; outline:none;
  font-family:inherit; }
.navsearch input::placeholder { color:rgba(215,230,248,.78); }
.appcard { display:flex; align-items:center; gap:16px; margin-top:26px;
  background:rgba(11,23,45,.55); border:1px solid rgba(255,255,255,.22);
  border-radius:18px; padding:18px 22px; cursor:pointer; }
.appcard:hover { background:rgba(20,38,68,.62); }
.appicon { width:58px; height:58px; flex:none; border-radius:50%;
  background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.3);
  font-size:27px; display:flex; align-items:center; justify-content:center; }
.appname { color:#fff; font-size:21px; font-weight:bold; }
.appsub { color:rgba(205,224,246,.88); font-size:13.5px; margin-top:3px; }
.apparrow { margin-left:auto; color:rgba(230,240,252,.9); font-size:32px;
  flex:none; }
.grid { display:flex; flex-wrap:nowrap; gap:14px 6px; margin-top:28px;
  justify-content:center; }
.tile { flex:1 1 0; min-width:0; text-align:center; cursor:pointer;
  position:relative; padding:6px 0 2px; border-radius:16px; }
.tile:hover { background:rgba(255,255,255,.07); }
.tileicon { width:54px; height:54px; margin:0 auto; border-radius:50%;
  background:rgba(255,255,255,.13); border:1px solid rgba(255,255,255,.32);
  display:flex; align-items:center; justify-content:center;
  overflow:hidden; }
.tileicon img { width:40px; height:40px; border-radius:10px; }
.tileicon .ch { color:#fff; font-size:22px; font-weight:bold; }
.tileicon.add { color:#dbe7f5; font-size:24px; }
.tilename { color:#eef4fc; font-size:12.5px; margin-top:8px;
  text-shadow:0 1px 6px rgba(0,0,0,.55);
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tilex { display:none; position:absolute; top:-2px; right:6px;
  width:22px; height:22px; line-height:22px; border-radius:50%;
  background:#dc2626; color:#fff; font-size:13px; font-weight:bold;
  z-index:2; }
.grid.editing .tilex { display:block; }
.editrow { text-align:center; margin-top:16px; }
.editbtn { color:rgba(205,224,246,.85); font-size:13.5px; cursor:pointer;
  text-decoration:underline; }
.card.listcard { background:rgba(9,17,33,.60);
  border:1px solid rgba(255,255,255,.16);
  box-shadow:0 12px 32px rgba(0,0,0,.35); }
.listcard .datebar .d { color:#fff; }
.listcard .datebar .cnt { background:rgba(255,255,255,.16);
  color:#dbe7f5; }
.listcard .datebar .line { background:rgba(255,255,255,.18); }
.listcard .row.alt { background:rgba(255,255,255,.05); }
.listcard .row:hover { background:rgba(255,255,255,.10);
  border-left-color:#f5a623; }
.listcard .subj { color:#fff; }
.listcard .time { background:rgba(255,255,255,.16); color:#fff; }
.listcard .btn { background:rgba(255,255,255,.12); color:#eaf2fc;
  border-color:rgba(255,255,255,.28); box-shadow:none; }
.listcard .btn.tabopen { background:rgba(255,255,255,.92); color:#1f3a5f;
  border-color:transparent; }
.listcard .btn.src { background:rgba(255,255,255,.12); color:#f3d9ae;
  border-color:rgba(255,255,255,.28); }
.listcard .foot { background:rgba(255,255,255,.08); color:#dbe7f5; }
.listcard p { color:#c9d9ee; }
"""

_NAV_JS = r"""
var toastTimer = null;
function goNav(p) {
  var q = [];
  if (/[?&]panel=1/.test(location.search)) q.push('panel=1');
  if (/[?&]embed=1/.test(location.search)) q.push('embed=1');
  location.href = p + (q.length ? (p.indexOf('?') >= 0 ? '&' : '?') + q.join('&') : '');
}
function doSearch(f) {
  var v = f.querySelector('input').value.trim();
  if (!v) return false;
  goNav('/search?q=' + encodeURIComponent(v));
  return false;
}
function openReport(f) {
  if (/[?&]panel=1/.test(location.search)) {
    location.href = '/view?f=' + encodeURIComponent(f) + '&panel=1' +
      (/[?&]embed=1/.test(location.search) ? '&embed=1' : '');
    return;
  }
  fetch('/popout?f=' + encodeURIComponent(f)).then(function(r) {
    if (!r.ok) location.href = '/view?f=' + encodeURIComponent(f);
  }).catch(function() {
    location.href = '/view?f=' + encodeURIComponent(f);
  });
}
function openInBrowser(f, el) {
  if (/[?&]embed=1/.test(location.search)) {
    var t = '';
    var row = (el && el.closest) ? el.closest('.row') : null;
    if (row) {
      var s = row.querySelector('.subj');
      if (s) t = s.textContent;
    }
    parent.postMessage({museTab: f, museTitle: t}, '*');
    return;
  }
  fetch('/popout?f=' + encodeURIComponent(f));
}
if (/[?&]panel=1/.test(location.search)) {
  document.querySelectorAll('.newtab').forEach(function(el) {
    el.style.display = '';
  });
  document.querySelectorAll('.webbtn').forEach(function(el) {
    el.style.display = 'inline-block';
  });
}
function showToast(msg, ms) {
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.style.display = 'block';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function() { t.style.display = 'none'; },
                           ms || 2600);
}
function grabArticle() {
  var b = document.getElementById('grabbtn');
  if (!b || b.disabled) { return; }
  b.disabled = true;
  b.textContent = '抓取中…';
  fetch('/grab').then(function(r) { return r.json(); })
    .then(function(j) {
      b.disabled = false;
      b.textContent = '🔗 抓文章';
      if (j.ok) {
        showToast('已抓取：《' + j.title + '》，正在打开…', 1800);
        setTimeout(function() { goNav('/report'); }, 1000);
      } else {
        showToast(j.msg || '抓取失败', 4500);
      }
    }, function() {
      b.disabled = false;
      b.textContent = '🔗 抓文章';
      showToast('抓取失败，稍后再试', 3000);
    });
}
function delReport(qf, btn) {
  if (btn.dataset.arm) {
    fetch('/del?f=' + qf).then(function(r) { return r.json(); })
      .then(function(j) {
        if (j.ok) {
          showToast('已删除');
          setTimeout(function() { location.reload(); }, 500);
        } else { showToast('删除失败'); }
      }, function() { showToast('删除失败'); });
    return;
  }
  btn.dataset.arm = '1';
  var old = btn.textContent;
  btn.textContent = '确认删除？';
  btn.style.background = '#dc2626';
  btn.style.color = '#fff';
  btn.style.borderColor = '#dc2626';
  setTimeout(function() {
    btn.dataset.arm = '';
    btn.textContent = old;
    btn.style.background = '';
    btn.style.color = '';
    btn.style.borderColor = '';
  }, 2600);
}
function manualFetch() {
  var b = document.getElementById('fetchbtn');
  b.disabled = true;
  b.textContent = '获取中…';
  fetch('/fetch').then(function() { pollFetch(b); },
                       function() { pollFetch(b); });
}
function pollFetch(b) {
  fetch('/fetch-status').then(function(r) { return r.json(); })
    .then(function(s) {
      if (s.running) {
        setTimeout(function() { pollFetch(b); }, 2500);
        return;
      }
      b.disabled = false;
      b.textContent = '↻ 获取邮件';
      if (s.new > 0) {
        showToast('新到 ' + s.new + ' 份报告，正在刷新…', 1600);
        setTimeout(function() { location.reload(); }, 900);
      } else if (s.new === 0) {
        showToast('没有新邮件');
      } else {
        showToast('获取失败，稍后再试');
      }
    }, function() {
      setTimeout(function() { pollFetch(b); }, 3000);
    });
}
function quitService() {
  if (!confirm('退出报告服务？退出后就不再自动下载新报告了。')) return;
  var done = function() {
    showToast('服务已退出，这个页面可以关了', 5000);
  };
  fetch('/quit').then(done, done);
}
function reveal(f, name) {
  fetch('/reveal?f=' + encodeURIComponent(f)).then(function(r) {
    showToast(r.ok ? ('已在访达定位源文件：' + name)
                   : '定位失败，文件可能已被移动');
  });
}
var searchMode = 'report';
function setMode(m) {
  searchMode = m;
  document.querySelectorAll('.mode').forEach(function(b) {
    b.classList.toggle('on', b.dataset.m === m);
  });
  document.getElementById('navq').placeholder =
    (m === 'report') ? '搜报告：股票名、代码、任意关键词…'
                     : '搜全网…';
}
function navSearch(f) {
  var v = document.getElementById('navq').value.trim();
  if (!v) return false;
  if (searchMode === 'report') {
    goNav('/search?q=' + encodeURIComponent(v));
  } else {
    fetch('/openurl?u=' + encodeURIComponent(
      'https://www.bing.com/search?q=' + encodeURIComponent(v)));
  }
  return false;
}
function tileClick(el, ev) {
  if (document.getElementById('grid').classList.contains('editing'))
    return;
  var u = el.dataset.url;
  if (u) fetch('/openurl?u=' + encodeURIComponent(u));
}
function tileDel(el, ev) {
  ev.stopPropagation();
  if (!confirm('删掉这个快捷方式？')) return;
  fetch('/tiles/del?i=' + el.dataset.i).then(function() {
    location.reload();
  });
}
function tileAdd() {
  var name = prompt('快捷方式名称（比如：同花顺）');
  if (!name) return;
  var url = prompt('网址（比如：https://www.10jqka.com.cn）');
  if (!url) return;
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  fetch('/tiles/add?name=' + encodeURIComponent(name) +
        '&url=' + encodeURIComponent(url)).then(function() {
    location.reload();
  });
}
function toggleEdit() {
  var g = document.getElementById('grid');
  g.classList.toggle('editing');
  document.getElementById('editbtn').textContent =
    g.classList.contains('editing') ? '完成' : '编辑快捷方式';
}
"""


# ---------------- 东财快讯页（财经要闻＋全球7×24） ----------------
_EM_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
          "AppleWebKit/605.1.15 (KHTML, like Gecko) "
          "Version/17.0 Safari/605.1.15")
_NEWS_CACHE = {"t": 0.0, "hot": "", "fast": "", "ts": ""}


def _em_json(url, referer):
    req = urllib.request.Request(url, headers={
        "User-Agent": _EM_UA, "Referer": referer})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def _news_strip(summary, title):
    s = (summary or "").strip()
    t = (title or "").strip()
    if t and s.startswith("【" + t + "】"):
        s = s[len(t) + 2:].strip()
    return s


def _news_hot_html():
    url = ("https://np-listapi.eastmoney.com/comm/web/getNewsByColumns"
           "?client=web&biz=web_news_col&column=350&order=1"
           "&needInteractData=0&page_index=1&page_size=40"
           "&req_trace=" + str(int(time.time() * 1000)) +
           "&fields=code,showTime,title,mediaName,summary,url,uniqueUrl")
    d = _em_json(url, "https://news.eastmoney.com/")
    lst = ((d.get("data") or {}).get("list")) or []
    today = datetime.now().strftime("%Y-%m-%d")
    rows = []
    for it in lst:
        title = (it.get("title") or "").strip()
        if not title:
            continue
        st = (it.get("showTime") or "")
        hm = st[11:16] if len(st) >= 16 else ""
        if st and st[:10] != today:
            hm = st[5:10] + " " + hm
        summ = _news_strip(it.get("summary"), title)
        link = it.get("uniqueUrl") or it.get("url") or ""
        media = (it.get("mediaName") or "").strip()
        t_html = _esc(title)
        if link:
            t_html = ('<a class="nlink" href="#" onclick="openExt(\'' +
                      _esc(link.replace("'", '')) + '\');return false">' +
                      t_html + '</a>')
        rows.append(
            '<div class="nrow"><div class="ntime">' + _esc(hm) +
            '</div><div class="nmain"><div class="ntitle">' + t_html +
            ('<span class="nsrc">' + _esc(media) + '</span>'
             if media else '') + '</div>' +
            ('<div class="nsum">' + _esc(summ) + '</div>'
             if summ else '') + '</div></div>')
    return "".join(rows) or '<div class="nempty">暂时没拉到要闻</div>'


def _news_fast_html():
    url = ("https://np-listapi.eastmoney.com/comm/web/getFastNewsList"
           "?client=web&biz=web_724&fastColumn=102&sortEnd=&pageSize=50"
           "&req_trace=" + str(int(time.time() * 1000)))
    d = _em_json(url, "https://kuaixun.eastmoney.com/7_24.html")
    lst = ((d.get("data") or {}).get("fastNewsList")) or []
    rows = []
    for it in lst:
        title = (it.get("title") or "").strip()
        if not title:
            continue
        st = (it.get("showTime") or "")
        hm = st[11:16] if len(st) >= 16 else ""
        hot = " hot" if it.get("titleColor") else ""
        summ = _news_strip(it.get("summary"), title)
        chips = []
        for sitem in (it.get("stockList") or []):
            parts = str(sitem).split(".")
            if (len(parts) == 2 and len(parts[1]) == 6
                    and parts[1].isdigit()):
                code = parts[1]
                chips.append('<a class="nchip" href="#" onclick="jumpCode(\'' +
                             code + '\');return false">' + code + '</a>')
        rows.append(
            '<div class="nrow' + hot + '"><div class="ntime">' +
            _esc(hm) + '</div><div class="nmain">' +
            '<div class="ntitle">' + _esc(title) +
            ('<span class="nchips">' + "".join(chips) + '</span>'
             if chips else '') + '</div>' +
            ('<div class="nsum">' + _esc(summ) + '</div>'
             if summ else '') + '</div></div>')
    return "".join(rows) or '<div class="nempty">暂时没拉到快讯</div>'


def _news_data(fresh=False):
    now = time.time()
    if not fresh and _NEWS_CACHE["hot"] and now - _NEWS_CACHE["t"] < 10:
        return _NEWS_CACHE
    try:
        hot = _news_hot_html()
    except Exception:
        hot = _NEWS_CACHE.get("hot") or \
            '<div class="nempty">要闻拉取失败，稍后点刷新</div>'
    try:
        fast = _news_fast_html()
    except Exception:
        fast = _NEWS_CACHE.get("fast") or \
            '<div class="nempty">快讯拉取失败，稍后点刷新</div>'
    _NEWS_CACHE.update(t=now, hot=hot, fast=fast,
                       ts=datetime.now().strftime("%H:%M:%S"))
    return _NEWS_CACHE


_NEWS_CSS = """
body{background:#0d1526 url('/wallpaper.jpg') no-repeat center center fixed;
background-size:cover}
.kpltop{background:rgba(10,17,30,.86);color:#fff;position:sticky;top:0;
z-index:60;backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);
border-bottom:1px solid rgba(255,255,255,.16);
box-shadow:0 4px 18px rgba(0,0,0,.35)}
.kpltop .navdate{color:#cfe0f2}
.autowrap{color:#e8eefc}
.secin{background:rgba(255,255,255,.14);color:#fff;
border:1px solid rgba(255,255,255,.4);border-radius:6px}
.backbtn{background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.45);
color:#fff}
.backbtn:hover{background:rgba(255,255,255,.28)}
.newswrap{max-width:1240px;margin:16px auto 36px;padding:0 16px}
.ntabs{display:flex;gap:10px;margin:14px 0}
.ntab{padding:9px 24px;border-radius:999px;background:rgba(255,255,255,.14);
color:#e8eefc;font-size:15px;font-weight:700;cursor:pointer;
border:1px solid rgba(255,255,255,.3);backdrop-filter:blur(8px);
-webkit-backdrop-filter:blur(8px)}
.ntab.on{background:#2f6fb3;color:#fff;border-color:#2f6fb3}
.nfeed{background:rgba(255,255,255,.95);border:1px solid rgba(255,255,255,.55);
border-radius:14px;padding:6px 20px;backdrop-filter:blur(12px);
-webkit-backdrop-filter:blur(12px);box-shadow:0 12px 36px rgba(0,0,0,.4)}
.nrow{display:flex;gap:14px;padding:11px 2px;border-bottom:1px solid #e7ecf3}
.nrow:last-child{border-bottom:none}
.ntime{flex:none;width:88px;color:#8a97a8;font-size:13px;padding-top:3px;
font-variant-numeric:tabular-nums}
.nmain{flex:1;min-width:0}
.ntitle{font-size:15.5px;font-weight:600;color:#1c2b3a;line-height:1.5}
.nrow.hot .ntitle{color:#c0392b}
.nlink{color:inherit;text-decoration:none}
.nlink:hover{text-decoration:underline}
.nsrc{margin-left:8px;font-size:12px;color:#93a2b5;font-weight:400}
.nsum{color:#5a6a7d;font-size:13.5px;line-height:1.65;margin-top:3px;
display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;
overflow:hidden}
.nchips{margin-left:6px}
.nchip{display:inline-block;background:#eef3fa;color:#20518f;border-radius:6px;
padding:0 6px;font-size:12px;margin-right:4px;text-decoration:none}
.nempty{padding:22px;color:#8a97a8;text-align:center}

.kpltop{flex-wrap:nowrap;padding:8px 12px;gap:8px;font-size:15px}
.kpltop .kplrefresh{padding:6px 11px;font-size:13.5px;white-space:nowrap}
.kpltop .autowrap{white-space:nowrap;font-size:13.5px}
.kpltop .secin{width:56px;padding:4px 6px}
.kpltop .navdate{font-size:12.5px;white-space:nowrap;color:#9db4d0}

.newsbar{display:flex;justify-content:space-between;align-items:center;
flex-wrap:wrap;gap:10px 14px;background:rgba(10,17,30,.86);
backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);
border-bottom:1px solid rgba(255,255,255,.16);position:sticky;top:0;
z-index:60;padding:10px 18px;box-shadow:0 4px 18px rgba(0,0,0,.35)}
.nbgroup{display:flex;align-items:center;gap:12px;flex-wrap:nowrap}
.nbtitle{color:#fff;font-size:19px;font-weight:bold;white-space:nowrap}
.newsbar .navdate{margin-left:2px;color:#9db4d0;font-size:12.5px;
white-space:nowrap}
.newswrap{max-width:1600px}
"""


_NEWS_JS = """
function newsBack() {
  if (history.length > 1) { history.back(); }
  else { goNav('/home'); }
}
function newsTab(w) {
  document.getElementById('pane-hot').style.display =
    w === 'hot' ? '' : 'none';
  document.getElementById('pane-fast').style.display =
    w === 'fast' ? '' : 'none';
  document.getElementById('tab-hot').className =
    'ntab' + (w === 'hot' ? ' on' : '');
  document.getElementById('tab-fast').className =
    'ntab' + (w === 'fast' ? ' on' : '');
}
function openExt(u) { fetch('/openurl?u=' + encodeURIComponent(u)); }
function jumpCode(c) { fetch('/jump?code=' + c); }
var fsTriedN = false;
function toggleFS() {
  if (document.fullscreenElement) { document.exitFullscreen(); return; }
  function fb() { if (!fsTriedN) { fsTriedN = true; fetch('/fs'); } }
  try {
    var r = document.documentElement.requestFullscreen();
    if (r && r.catch) { r.catch(fb); }
    setTimeout(function() {
      if (!document.fullscreenElement) fb();
    }, 700);
  } catch (e) { fb(); }
}
var newsTimer = null;
function newsSecVal() {
  var el = document.getElementById('newssec');
  var v = parseInt(String(el.value).replace(/[^0-9]/g, ''), 10);
  if (!v) v = 120;
  if (v < 30) v = 30;
  if (v > 1800) v = 1800;
  el.value = v;
  return v;
}
function newsStartTimer() {
  newsStopTimer();
  if (!document.getElementById('newsauto').checked) return;
  newsTimer = setInterval(function() { newsRefresh(false); },
    newsSecVal() * 1000);
}
function newsStopTimer() {
  if (newsTimer) { clearInterval(newsTimer); newsTimer = null; }
}
function newsAutoChange() {
  var on = document.getElementById('newsauto').checked;
  fetch('/settings/set?news_auto=' + (on ? '1' : '0'));
  newsStartTimer();
  showToast(on ? '快讯自动更新已开' : '快讯自动更新已关');
}
function newsSecChange() {
  var v = newsSecVal();
  fetch('/settings/set?news_sec=' + v);
  newsStartTimer();
}
function newsRefresh(manual) {
  fetch('/news.json' + (manual ? '?fresh=1' : '')).then(function(r) {
    return r.json();
  }).then(function(d) {
    if (d.hot) document.getElementById('pane-hot').innerHTML = d.hot;
    if (d.fast) document.getElementById('pane-fast').innerHTML = d.fast;
    if (d.ts) document.getElementById('newsdate').textContent =
      '更新 ' + d.ts;
    if (manual) showToast('已刷新');
  }).catch(function() {});
}
newsStartTimer();
"""


def render_news():
    s = _load_settings()
    d = _news_data()
    nauto = " checked" if s.get("news_auto", True) else ""
    nsec = int(s.get("news_interval_sec", 120) or 120)
    dateline = "更新 " + (d["ts"] or datetime.now().strftime("%H:%M:%S"))
    secopts = "".join(
        '<option value="%d">%s</option>' % (v, lab) for v, lab in
        [(15, "15 秒"), (30, "30 秒"), (45, "45 秒"), (60, "1 分钟"),
         (90, "90 秒"), (120, "2 分钟"), (180, "3 分钟"), (240, "4 分钟"),
         (300, "5 分钟"), (600, "10 分钟"), (900, "15 分钟"),
         (1800, "30 分钟")])
    return f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<title>东财快讯</title><style>{_PAGE_CSS}{_NAV_CSS}{_KPL_CSS}{_NEWS_CSS}</style></head>
<body>
<div class="newsbar"><span class="nbgroup">
  <button class="kplrefresh backbtn" onclick="newsBack()">← 返回</button>
  <span class="nbtitle">📰 东方财富 · 新闻</span>
  <span class="navdate" id="newsdate">{dateline}</span></span>
<span class="nbgroup">
  <button class="kplrefresh" onclick="newsRefresh(true)">↻ 刷新数据</button>
  <label class="kplsw"><input type="checkbox" id="newsauto"{nauto}
    onchange="newsAutoChange()"> 自动更新</label>
  <input class="kplsel" id="newssec" list="newssecs" value="{nsec}"
    style="width:62px" onchange="newsSecChange()">
  <datalist id="newssecs">{secopts}</datalist>
  <span class="kplsw">秒</span>
  <button class="kplrefresh" onclick="toggleFS()">⛶ 全屏</button></span></div>
<div class="newswrap">
  {_secnav('news')}
  <div class="ntabs">
    <span class="ntab on" id="tab-hot" onclick="newsTab('hot')">🔥 财经要闻</span>
    <span class="ntab" id="tab-fast" onclick="newsTab('fast')">⚡ 全球7×24</span>
  </div>
  <div class="nfeed" id="pane-hot">{d['hot']}</div>
  <div class="nfeed" id="pane-fast" style="display:none">{d['fast']}</div>
</div>
<div class="toast" id="toast"></div>
<button class="grabbtn" id="grabbtn" data-tip="先复制文章链接，再点这里：抓取正文＋评论，存进报告列表"
  onclick="grabArticle()">🔗 抓文章</button>
<button class="fetchbtn" id="fetchbtn" data-tip="立即拉取最新报告邮件（最近 5 封）"
  onclick="manualFetch()">↻ 获取邮件</button>
<button class="quitbtn" onclick="goNav('/home')">🏠 主页</button>
<script>{_NAV_JS}</script>
<script>{_NEWS_JS}</script>
</body></html>"""



def render_settings():
    s = _load_settings()
    mail_min = int(s.get("mail_interval_sec", 300) or 300) // 60
    kpl_sec = int(s.get("kpl_interval_sec", 300) or 300)

    def opts(choices):
        return "".join(f'<option value="{v}">{lab}</option>'
                       for v, lab in choices)

    mail_opts = opts([(1, "1 分钟"), (2, "2 分钟"), (3, "3 分钟"),
                      (5, "5 分钟"), (10, "10 分钟"), (15, "15 分钟"),
                      (20, "20 分钟"), (30, "30 分钟"), (45, "45 分钟"),
                      (60, "60 分钟")])
    kpl_opts = opts([(15, "15 秒"), (30, "30 秒"), (45, "45 秒"),
                     (60, "1 分钟"), (90, "90 秒"), (120, "2 分钟"),
                     (180, "3 分钟"), (240, "4 分钟"), (300, "5 分钟"),
                     (600, "10 分钟"), (900, "15 分钟"),
                     (1800, "30 分钟")])
    mail_chk = " checked" if s.get("mail_auto", True) else ""
    kpl_chk = " checked" if s.get("kpl_auto", True) else ""
    news_sec = int(s.get("news_interval_sec", 120) or 120)
    news_chk = " checked" if s.get("news_auto", True) else ""
    kpl_en_chk = " checked" if s.get("kpl_enabled", True) else ""
    news_en_chk = " checked" if s.get("news_enabled", True) else ""
    card = ("background:#fff;border:1px solid #dbe2ec;border-radius:14px;"
            "padding:16px 18px;margin:14px 0")
    ttl = "font-size:16px;font-weight:800;margin-bottom:6px"
    row = ("display:flex;align-items:center;gap:14px;flex-wrap:wrap;"
           "margin-top:10px;font-size:14px")
    sel = ("padding:6px 10px;border:1px solid #c9d4e4;border-radius:8px;"
           "font-size:14px;background:#fff")
    tip = "color:#7286a0;font-size:12.5px;margin-top:8px;line-height:1.6"
    inner = f"""
<div style="{card}">
  <div style="{ttl}">📬 邮件自动拉取</div>
  <div style="{row}">
    <label><input type="checkbox"{mail_chk}
      onchange="setOpt('mail_auto=' + (this.checked ? '1' : '0'),
        this.checked ? '邮件自动拉取已开' : '邮件自动拉取已关')"> 自动拉取</label>
    <span>间隔（分钟，可下拉选、也可直接填）
      <input style="{sel}width:110px" list="mailmins" value="{mail_min}"
        onchange="setMin(this, 'mail_min', 1, 120, '邮件拉取间隔', ' 分钟')">
      <datalist id="mailmins">{mail_opts}</datalist>
    </span>
  </div>
  <div style="{tip}">关掉后不再自动查邮件；右下角「↻ 获取邮件」手动拉仍然可用。
  工作日 9:27–9:45 竞价时段仍会加密到 30 秒一次（之前定的竞价加密）；周六、周日自动放缓到 30 分钟一次，自动拉取每趟最多取最近 5 封新邮件。</div>
</div>
<div style="{card}">
  <div style="{ttl}">📊 开盘啦数据自动刷新</div>
  <div style="{row}">
    <label><input type="checkbox"{kpl_en_chk}
      onchange="setOpt('kpl_enabled=' + (this.checked ? '1' : '0'),
        this.checked ? '开盘啦模块已启用' : '开盘啦模块已关闭')"> <b>启用本模块</b></label>
    <span style="color:#7286a0;font-size:12.5px">关掉后所有入口隐藏、页面停用（不只是停刷新）</span>
  </div>
  <div style="{row}">
    <label><input type="checkbox"{kpl_chk}
      onchange="setOpt('kpl_auto=' + (this.checked ? '1' : '0'),
        this.checked ? '开盘啦自动刷新已开' : '开盘啦自动刷新已关')"> 自动刷新</label>
    <span>间隔（秒，可下拉选、也可直接填）
      <input style="{sel}width:110px" list="kplsecs2" value="{kpl_sec}"
        onchange="setMin(this, 'kpl_sec', 15, 1800, '开盘啦刷新间隔', ' 秒')">
      <datalist id="kplsecs2">{kpl_opts}</datalist>
    </span>
  </div>
  <div style="{tip}">关掉后开盘啦页面不再自动更新；页面右上角「↻ 刷新数据」仍可手动刷。
  开盘啦页面顶栏也有同样的开关和频率，改哪边都一样。</div>
</div>
<div style="{card}">
  <div style="{ttl}">📰 东财快讯自动刷新</div>
  <div style="{row}">
    <label><input type="checkbox"{news_en_chk}
      onchange="setOpt('news_enabled=' + (this.checked ? '1' : '0'),
        this.checked ? '快讯模块已启用' : '快讯模块已关闭')"> <b>启用本模块</b></label>
    <span style="color:#7286a0;font-size:12.5px">关掉后所有入口隐藏、页面停用（不只是停刷新）</span>
  </div>
  <div style="{row}">
    <label><input type="checkbox"{news_chk}
      onchange="setOpt('news_auto=' + (this.checked ? '1' : '0'),
        this.checked ? '快讯自动刷新已开' : '快讯自动刷新已关')"> 自动刷新</label>
    <span>间隔（秒，可下拉选、也可直接填）
      <input style="{sel}width:110px" list="kplsecs2" value="{news_sec}"
        onchange="setMin(this, 'news_sec', 30, 1800, '快讯刷新间隔', ' 秒')">
    </span>
  </div>
  <div style="{tip}">财经要闻＋全球7×24 快讯页的自动刷新；关掉后页面不再自动更新，
  顶栏「↻ 刷新数据」仍可手动刷。快讯页顶栏也有同样的开关和频率。</div>
</div>
<script>
function setOpt(q, msg) {{
  fetch('/settings/set?' + q).then(function() {{ showToast(msg); }});
}}
function setMin(el, key, lo, hi, label, unit) {{
  var v = parseFloat(String(el.value).replace(/[^0-9.]/g, ''));
  if (!v) {{ return; }}
  if (v < lo) v = lo;
  if (v > hi) v = hi;
  el.value = v;
  setOpt(key + '=' + v, label + '：' + v + unit);
}}
</script>"""
    return _page("⚙ 设置", "", "开关与频率改完即生效", inner)

def render_home():
    with LOCK:
        manifest = load_manifest()
    entries = sorted(manifest, key=lambda e: (e["date"], e["time"]),
                     reverse=True)
    groups = _grouped(entries)
    now = datetime.now()
    wd = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六",
          "星期日"][now.weekday()]
    dateline = f"{now:%Y-%m-%d} {wd} · 更新 {now:%H:%M}"
    if entries:
        latest = max(groups)
        sub = (f"今日 {len(groups[latest])} 份 · 历史 {len(groups)} 天 "
               f"{len(entries)} 份")
        list_inner = _datebar(latest, len(groups[latest])) \
            + _rows_html(groups[latest])
        older = [d for d in groups if d != latest]
        if older:
            list_inner += (
                f'<div class="foot" onclick="goNav(\'/history\')">'
                f'<span>🗂 历史报告（{len(older)} 天 · '
                f'{sum(len(groups[d]) for d in older)} 份）</span>'
                f'<span class="arrow">查看 →</span></div>')
    else:
        sub = "还没有报告，第一封到达后会自动出现"
        list_inner = ('<p style="padding:14px 22px 18px">还没有报告，'
                      '第一封到达后会自动出现在这里。</p>')
    tiles = [_tile_html(t, i) for i, t in enumerate(load_tiles())]
    _s = _load_settings()
    kpl_tile = ('<div class="tile" onclick="goNav(\'/kpl\')">'
                '<div class="tileicon"><span class="ch">盘</span></div>'
                '<div class="tilename">开盘啦数据</div></div>'
                if _s.get("kpl_enabled", True) else "")
    news_tile = ('<div class="tile" onclick="goNav(\'/news\')">'
                 '<div class="tileicon"><span class="ch">讯</span></div>'
                 '<div class="tilename">东财快讯</div></div>'
                 if _s.get("news_enabled", True) else "")
    tiles.append(
        '<div class="tile" onclick="goNav(\'/settings\')">'
        '<div class="tileicon"><span class="ch">⚙</span></div>'
        '<div class="tilename">设置</div></div>')
    tiles.append(
        '<div class="tile" onclick="tileAdd()">'
        '<div class="tileicon add">＋</div>'
        '<div class="tilename">添加</div></div>')
    return f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<title>选股报告 · 导航主页</title>
<style>{_PAGE_CSS}{_NAV_CSS}</style></head><body>
<div class="navwrap">
  <div class="navtop"><div class="navttl">📈 选股报告 · 导航</div>
  <div class="navdate">{dateline}</div></div>
  <div class="modes">
    <span class="mode on" data-m="report"
      onclick="setMode('report')">搜报告</span>
    <span class="mode" data-m="web"
      onclick="setMode('web')">搜全网</span>
  </div>
  <form class="navsearch" onsubmit="return navSearch(this)">
    <input id="navq" type="text" autocomplete="off"
      placeholder="搜报告：股票名、代码、任意关键词…">
  </form>
  <div class="appcard" onclick="goNav('/report')">
    <div class="appicon">📈</div>
    <div><div class="appname">报告列表</div>
    <div class="appsub">{sub}</div></div>
    <div class="apparrow">›</div>
  </div>
  <div class="grid" id="grid">
    <div class="tile" onclick="goNav('/history')">
      <div class="tileicon"><span class="ch">史</span></div>
      <div class="tilename">历史报告</div></div>
    <div class="tile" onclick="goNav('/favs')">
      <div class="tileicon"><span class="ch" style="color:#f5b301">★</span></div>
      <div class="tilename">收藏夹</div></div>
    {kpl_tile}
    {news_tile}
    {''.join(tiles)}
  </div>
  <div class="editrow"><span class="editbtn" id="editbtn"
    onclick="toggleEdit()">编辑快捷方式</span></div>
</div>
<div class="card listcard">{list_inner}</div>
<div class="toast" id="toast"></div>
<button class="grabbtn" id="grabbtn" data-tip="先复制文章链接，再点这里：抓取正文＋评论，存进报告列表"
  onclick="grabArticle()">🔗 抓文章</button>
<button class="fetchbtn" id="fetchbtn" data-tip="立即拉取最新报告邮件（最近 5 封）"
  onclick="manualFetch()">↻ 获取邮件</button>
<button class="quitbtn" onclick="goNav('/home')">🏠 主页</button>
<script>{_NAV_JS}</script>
</body></html>"""


SHELL_HTML = r"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<title>Muse 选股报告</title>
<style>
* { box-sizing: border-box; }
html, body { height: 100%; }
body { margin:0; font-family:'PingFang SC','Microsoft YaHei',sans-serif;
       background:#e8edf3; display:flex; flex-direction:column; }
#tabs { display:flex; align-items:flex-end; gap:6px;
        background:#152a45;
        padding:8px 10px 0; overflow-x:auto; flex:none; }
.tab { flex:none; display:flex; align-items:center; gap:8px;
       max-width:300px; background:rgba(255,255,255,.18); color:#fff;
       font-size:14px; font-weight:bold; padding:8px 14px;
       border-radius:10px 10px 0 0; cursor:pointer; }
.tab.active { background:#e8edf3; color:#1f3a5f; }
.tab .t { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tab .x { opacity:.7; padding:0 3px; font-size:13px; }
.tab .x:hover { opacity:1; }
#frames { flex:1; position:relative; }
.addbtn { flex:none; align-self:center; margin:0 2px 4px 4px;
          background:rgba(255,255,255,.18); color:#fff; font-size:17px;
          font-weight:bold; width:30px; height:30px; line-height:30px;
          text-align:center; border-radius:50%; cursor:pointer; }
.addbtn:hover { background:rgba(255,255,255,.35); }
iframe { position:absolute; inset:0; width:100%; height:100%;
         border:0; background:#e8edf3; }
</style></head><body>
<div id="tabs"></div>
<div id="frames"></div>
<script>
var seq = 0;
var addBtn = null;
var tabsEl = document.getElementById('tabs');
var framesEl = document.getElementById('frames');
function activate(id) {
  document.querySelectorAll('.tab').forEach(function(t) {
    t.classList.toggle('active', t.dataset.id === id);
  });
  document.querySelectorAll('#frames iframe').forEach(function(f) {
    f.style.display = (f.dataset.fid === id) ? 'block' : 'none';
  });
}
function closeTab(id) {
  var tab = null, fr = null, wasActive = false;
  document.querySelectorAll('.tab').forEach(function(t) {
    if (t.dataset.id === id) { tab = t; wasActive = t.classList.contains('active'); }
  });
  document.querySelectorAll('#frames iframe').forEach(function(f) {
    if (f.dataset.fid === id) fr = f;
  });
  if (tab) tab.remove();
  if (fr) fr.remove();
  if (wasActive) {
    var rest = tabsEl.querySelectorAll('.tab');
    if (rest.length) activate(rest[rest.length - 1].dataset.id);
  }
}
function syncTitle(id, fr) {
  var tab = null;
  document.querySelectorAll('.tab').forEach(function(t) {
    if (t.dataset.id === id) tab = t;
  });
  if (!tab) return;
  var path = '', f = '';
  try {
    path = fr.contentWindow.location.pathname;
    f = new URLSearchParams(fr.contentWindow.location.search).get('f') || '';
  } catch (e) { return; }
  var label = null;
  if (path === '/' || path === '/report') label = '🏠 报告列表';
  else if (path === '/kpl') label = '📊 开盘啦数据';
  else if (path === '/settings') label = '⚙ 设置';
  else if (path === '/news') label = '📰 东财快讯';
  else if (path === '/favs') label = '⭐ 收藏夹';
  else if (path === '/fav') label = '⭐ 收藏';
  else if (path === '/home') label = '🏠 主页';
  else if (path === '/search') label = '🔍 搜索结果';
  else if (path === '/history') label = '🗂 历史报告';
  else if (path === '/view') {
    if (f) tab.dataset.src = f;
    var nav = null;
    try {
      nav = fr.contentDocument.querySelector('div[style*="sticky"] span');
    } catch (e) {}
    if (nav && nav.textContent) label = nav.textContent;
    else if (f) {
      label = f.split('/').pop().replace(/\.html$/, '')
               .replace(/^\d{4}_/, '').replace(/-正文$/, '');
    }
  }
  if (label) {
    var t = tab.querySelector('.t');
    if (t) { t.textContent = label; t.title = label; }
  }
}
function makeTab(id, title, src, closable) {
  var tab = document.createElement('div');
  tab.className = 'tab';
  tab.dataset.id = id;
  var t = document.createElement('span');
  t.className = 't';
  t.textContent = title;
  t.title = title;
  tab.appendChild(t);
  if (closable) {
    var x = document.createElement('span');
    x.className = 'x';
    x.textContent = '✕';
    x.title = '关闭这个标签';
    x.addEventListener('click', function(ev) {
      ev.stopPropagation();
      closeTab(id);
    });
    tab.appendChild(x);
  }
  tab.addEventListener('click', function() { activate(id); });
  if (addBtn) tabsEl.insertBefore(tab, addBtn);
  else tabsEl.appendChild(tab);
  var fr = document.createElement('iframe');
  fr.dataset.fid = id;
  fr.allowFullscreen = true;
  fr.addEventListener('load', function() { syncTitle(id, fr); });
  if (src) fr.src = src;
  fr.style.display = 'none';
  framesEl.appendChild(fr);
  return tab;
}
function openTab(f, title) {
  var found = null;
  document.querySelectorAll('.tab').forEach(function(t) {
    if (t.dataset.src === f) found = t;
  });
  if (found) { activate(found.dataset.id); return; }
  var id = 't' + (++seq);
  var tab = makeTab(id, title || '报告',
    '/view?f=' + encodeURIComponent(f) + '&panel=1&embed=1', true);
  tab.dataset.src = f;
  activate(id);
}
window.addEventListener('message', function(e) {
  if (e.origin !== location.origin) return;
  var d = e.data || {};
  if (d.museTab) openTab(d.museTab, d.museTitle);
});
addBtn = document.createElement('div');
addBtn.className = 'addbtn';
addBtn.textContent = '＋';
addBtn.title = '开一个新标签（报告列表）';
addBtn.addEventListener('click', function() {
  var id = 't' + (++seq);
  makeTab(id, '🏠 报告列表', '/report?panel=1&embed=1', true);
  activate(id);
});
tabsEl.appendChild(addBtn);
makeTab('home', '🏠 主页', '/home?panel=1&embed=1', false);
activate('home');
</script>
</body></html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path in ("/", "/report", "/history"):
            html = (render_index() if parsed.path in ("/", "/report")
                    else render_history()).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(html)))
            self.end_headers()
            self.wfile.write(html)
            return
        if parsed.path == "/grab":
            payload = grab_article()
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path in ("/fetch", "/fetch-status"):
            if parsed.path == "/fetch":
                started = trigger_fetch()
                payload = {"started": started,
                           "running": FETCH_STATE["running"]}
            else:
                payload = FETCH_STATE
            data = json.dumps(payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/del":
            rel = parse_qs(parsed.query).get("f", [""])[0]
            payload = {"ok": delete_entry(rel)}
            data = json.dumps(payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/fav":
            qs = parse_qs(parsed.query)
            data = render_fav_chooser(
                qs.get("f", [""])[0], _nav_suffix(qs)).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/fav/set":
            qs = parse_qs(parsed.query)
            rel = qs.get("f", [""])[0]
            cat = (qs.get("cat", [""])[0] or "").strip()[:24]
            with LOCK:
                mlist = load_manifest()
                for e in mlist:
                    if rel in (e.get("full"), e.get("body")):
                        if cat:
                            e["fav"] = cat
                        else:
                            e.pop("fav", None)
                        break
                save_manifest(mlist)
            suffix = _nav_suffix(qs)
            self.send_response(302)
            self.send_header("Location",
                             "/favs" + (("?" + suffix[1:])
                                        if suffix else ""))
            self.end_headers()
            return
        if parsed.path == "/favs":
            data = render_favs().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path.startswith("/files/"):
            import mimetypes
            rel = unquote(parsed.path[len("/files/"):])
            root = os.path.realpath(reports_root())
            target = os.path.realpath(os.path.join(root, rel))
            if target.startswith(root) and os.path.isfile(target):
                with open(target, "rb") as fh:
                    data = fh.read()
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    mimetypes.guess_type(target)[0]
                    or "application/octet-stream")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            else:
                self.send_response(404)
                self.end_headers()
            return
        if parsed.path in ("/kpl", "/kpl.json") and \
                not _load_settings().get("kpl_enabled", True):
            if parsed.path.endswith(".json"):
                data = json.dumps({"error": "disabled"}).encode("utf-8")
                self.send_response(403)
                self.send_header("Content-Type", "application/json")
            else:
                data = _disabled_page("开盘啦数据").encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",
                                 "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/kpl":
            data = render_kpl().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/kpl.json":
            if parse_qs(parsed.query).get("fresh"):
                fresh_d = _kpl_fetch_all()
                _KPL_CACHE.update(ts=time.time(), data=fresh_d)
                data = json.dumps(
                    _kpl_fragments(fresh_d)).encode("utf-8")
            else:
                data = json.dumps(
                    _kpl_fragments(_kpl_data())).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path in ("/news", "/news.json") and \
                not _load_settings().get("news_enabled", True):
            if parsed.path.endswith(".json"):
                data = json.dumps({"error": "disabled"}).encode("utf-8")
                self.send_response(403)
                self.send_header("Content-Type", "application/json")
            else:
                data = _disabled_page("东财快讯").encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",
                                 "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/news":
            data = render_news().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/news.json":
            fresh = bool(parse_qs(parsed.query).get("fresh"))
            nd = _news_data(fresh)
            data = json.dumps({"hot": nd["hot"], "fast": nd["fast"],
                               "ts": nd["ts"]}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/settings/set":
            q = parse_qs(parsed.query)
            s = _load_settings()
            if "mail_auto" in q:
                s["mail_auto"] = q["mail_auto"][0] in ("1", "true", "on")
            if "mail_min" in q:
                try:
                    s["mail_interval_sec"] = max(
                        60, min(3600, int(float(q["mail_min"][0]) * 60)))
                except Exception:
                    pass
            if "kpl_auto" in q:
                s["kpl_auto"] = q["kpl_auto"][0] in ("1", "true", "on")
            if "kpl_sec" in q:
                try:
                    s["kpl_interval_sec"] = max(
                        15, min(1800, int(q["kpl_sec"][0])))
                except Exception:
                    pass
            if "news_auto" in q:
                s["news_auto"] = q["news_auto"][0] in ("1", "true", "on")
            if "news_sec" in q:
                try:
                    s["news_interval_sec"] = max(
                        30, min(1800, int(q["news_sec"][0])))
                except Exception:
                    pass
            if "kpl_enabled" in q:
                s["kpl_enabled"] = q["kpl_enabled"][0] in ("1", "true", "on")
            if "news_enabled" in q:
                s["news_enabled"] = q["news_enabled"][0] in ("1", "true", "on")
            _save_settings(s)
            data = json.dumps({"ok": True,
                               "settings": s}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/settings":
            data = render_settings().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/home":
            data = render_home().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/search":
            q = parse_qs(parsed.query).get("q", [""])[0]
            data = render_search(q).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/openurl":
            u = parse_qs(parsed.query).get("u", [""])[0]
            if u.startswith("http://") or u.startswith("https://"):
                try:
                    subprocess.Popen(["open", u])
                    self.send_response(204)
                except Exception:
                    self.send_response(500)
                self.end_headers()
                return
        if parsed.path in ("/tiles/add", "/tiles/del"):
            qs = parse_qs(parsed.query)
            tiles = load_tiles()
            if parsed.path == "/tiles/add":
                name = qs.get("name", [""])[0].strip()[:20]
                url = qs.get("url", [""])[0].strip()
                if name and (url.startswith("http://")
                             or url.startswith("https://")):
                    tiles.append({
                        "name": name, "url": url,
                        "ch": name[:1],
                        "color": _TILE_COLORS[len(tiles)
                                              % len(_TILE_COLORS)]})
                    save_json(_tiles_path(), tiles)
            else:
                try:
                    i = int(qs.get("i", ["-1"])[0])
                except ValueError:
                    i = -1
                if 0 <= i < len(tiles):
                    tiles.pop(i)
                    save_json(_tiles_path(), tiles)
            data = b'{"ok": true}'
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/wallpaper.jpg":
            try:
                with open(os.path.join(APP_DIR, "nav_wallpaper.jpg"),
                          "rb") as fh:
                    data = fh.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except OSError:
                self.send_response(404)
                self.end_headers()
            return
        if parsed.path == "/shell":
            data = SHELL_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/jump":
            code = parse_qs(parsed.query).get("code", [""])[0]
            if code.isdigit() and len(code) == 6:
                try:
                    subprocess.Popen(
                        ["open", f"hammerspoon://emjump?code={code}"])
                    self.send_response(204)
                except Exception:
                    self.send_response(500)
                self.end_headers()
                return
        if parsed.path == "/fs":
            try:
                subprocess.Popen(["open", "hammerspoon://togglefs"])
                self.send_response(204)
            except Exception:
                self.send_response(500)
            self.end_headers()
            return
        if parsed.path == "/webhome":
            try:
                subprocess.Popen(["open", "http://127.0.0.1:8765/home"])
                self.send_response(204)
            except Exception:
                self.send_response(500)
            self.end_headers()
            return
        if parsed.path == "/quit":
            data = "已退出".encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

            def _quit():
                # 先卸掉 launchd 自启，否则 KeepAlive 会立刻把进程拉回来
                plist = os.path.expanduser(
                    "~/Library/LaunchAgents/com.muse.mailreports.plist")
                try:
                    subprocess.run(["launchctl", "unload", plist],
                                   timeout=10)
                except Exception:
                    pass
                os._exit(0)

            threading.Timer(0.6, _quit).start()
            return
        if parsed.path == "/popout":
            rel = unquote(parse_qs(parsed.query).get("f", [""])[0])
            root = os.path.realpath(reports_root())
            target = os.path.realpath(os.path.join(root, rel))
            if target.startswith(root) and os.path.isfile(target) \
                    and target.endswith(".html"):
                url = ("http://127.0.0.1:8765/view?f="
                       + quote(rel, safe=""))
                try:
                    subprocess.Popen(["open", url])
                    self.send_response(204)
                except Exception:
                    self.send_response(500)
                self.end_headers()
                return
        if parsed.path == "/reveal":
            rel = unquote(parse_qs(parsed.query).get("f", [""])[0])
            root = os.path.realpath(reports_root())
            target = os.path.realpath(os.path.join(root, rel))
            if target.startswith(root) and os.path.isfile(target):
                try:
                    subprocess.run(["open", "-R", target], timeout=10)
                    self.send_response(204)
                except Exception:
                    self.send_response(500)
                self.end_headers()
                return
        if parsed.path == "/view":
            rel = unquote(parse_qs(parsed.query).get("f", [""])[0])
            panel = parse_qs(parsed.query).get("panel", [""])[0] == "1"
            embed = parse_qs(parsed.query).get("embed", [""])[0] == "1"
            root = os.path.realpath(reports_root())
            target = os.path.realpath(os.path.join(root, rel))
            if target.startswith(root) and os.path.isfile(target) \
                    and target.endswith(".html"):
                subj = ""
                with LOCK:
                    for e in load_manifest():
                        if rel in (e.get("full"), e.get("body")):
                            subj = _esc(e["subject"])
                            break
                # 服务内看报告时顶部加一条导航：回索引＋邮件标题＋全屏按钮，
                # 并放宽版心（磁盘上的原文件不动）
                if panel and embed:
                    index_href = "/report?panel=1&embed=1"
                elif panel:
                    index_href = "/report?panel=1"
                else:
                    index_href = "/report"
                webbtn = ""
                fs_margin = ""
                if panel:
                    webbtn = ('<a href="javascript:void(0)" '
                              'onclick="fetch(\'/webhome\')" '
                              'style="color:#1f3a5f;font-weight:bold;flex:none;'
                              'text-decoration:none;font-size:15px;'
                              'display:inline-block;'
                              'background:#fff;'
                              'padding:8px 18px;border-radius:20px;">'
                              '🌐 网页版主页</a>')
                nav = ('<div style="position:sticky;top:0;z-index:999;'
                       'background:#1f3a5f;padding:12px 20px;display:flex;'
                       'align-items:center;gap:16px;">'
                       f'<a href="{index_href}" style="color:#fff;'
                       'font-weight:bold;'
                       'text-decoration:none;font-size:18px;'
                       'display:inline-block;flex:none;'
                       'background:rgba(255,255,255,.16);'
                       'padding:8px 18px;border-radius:8px;">← 报告索引</a>'
                       f'<span style="color:#fff;font-size:17px;flex:1;'
                       f'text-align:center;'
                       f'font-weight:bold;">{subj}</span>'
                       + webbtn +
                       f'<a href="javascript:void(0)" onclick="toggleFS()" '
                       f'style="color:#fff;font-weight:bold;flex:none;'
                       f'text-decoration:none;font-size:16px;'
                       f'{fs_margin}display:inline-block;'
                       f'background:rgba(255,255,255,.16);'
                       f'padding:8px 16px;border-radius:8px;">⛶ 全屏</a>'
                       '</div>'
                       '<script>var fsTried=false;'
                       'function toggleFS(){'
                       'if(document.fullscreenElement){'
                       'document.exitFullscreen();return}'
                       'function fb(){if(!fsTried){fsTried=true;'
                       'fetch("/fs")}}'
                       'try{'
                       'var r=document.documentElement.requestFullscreen();'
                       'if(r&&r.catch){r.catch(fb)}'
                       'setTimeout(function(){'
                       'if(!document.fullscreenElement)fb()},700);'
                       '}catch(e){fb()}}'
                       'function jumpStock(c){'
                       'fetch("/jump?code=" + c)}'
                       '</script>'
                       '<style>div[style*="max-width"]'
                       '{max-width:1760px !important}</style>'
                       ).encode("utf-8")
                with open(target, "rb") as fh:
                    raw = fh.read()
                if panel:
                    # 面板内点股票名不走 webview 的协议交接（会卡），改走
                    # 本地服务 /jump 用系统方式打开，和浏览器里点一致
                    raw = re.sub(
                        rb'href="hammerspoon://emjump\?code=(\d+)"',
                        lambda m: b"href='javascript:jumpStock(\""
                                  + m.group(1) + b"\")'", raw)
                data = nav + raw
                # 补齐整行点击＋悬停（旧文件/正文版没带这段脚本时靠它；
                # 已自带脚本的行有 data-jump 标记、自动跳过不重复绑）
                shim = ('<style>tbody tr:hover{filter:brightness(1.15)}'
                        'tbody tr:hover td{background-color:'
                        'rgba(255,183,28,.38)!important}'
                        'tbody tr:hover td:first-child{'
                        'box-shadow:inset 5px 0 0 #f59e0b}'
                        'tbody tr[data-jump]{cursor:pointer}</style>'
                        '<script>document.querySelectorAll("tbody tr")'
                        '.forEach(function(tr){'
                        'if(tr.getAttribute("data-jump"))return;'
                        'var link=tr.querySelector('
                        '\'a[href^="hammerspoon://"],'
                        'a[href^="javascript:jumpStock"]\');'
                        'if(!link)return;'
                        'tr.setAttribute("data-jump","1");'
                        'tr.addEventListener("click",function(e){'
                        'if(e.target.closest("a"))return;'
                        'link.click();});});</script>'
                        ).encode("utf-8")
                data += shim
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
        self.send_response(404)
        self.end_headers()


def main():
    global CFG
    cfg = load_json(CONFIG_PATH, None)
    if not cfg:
        print(f"缺少配置文件 {CONFIG_PATH}，先运行安装脚本。")
        return 1
    CFG = cfg
    os.makedirs(reports_root(), exist_ok=True)
    t = threading.Thread(target=fetcher_loop, daemon=True)
    t.start()
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    log(f"服务已启动 http://127.0.0.1:{PORT}（每 {CFG.get('interval_sec', 300)} 秒拉一次）")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
