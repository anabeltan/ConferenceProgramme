#!/usr/bin/env python3
"""Generate mobile-friendly per-day HTML schedule pages and QR codes for SMBH 2026.

Output:
  web/index.html        — day-picker landing page
  web/day1.html         — Monday, June 29
  web/day2.html         — Tuesday, June 30
  web/day3.html         — Wednesday, July 1
  web/day4.html         — Thursday, July 2
  web/day5.html         — Friday, July 3
  qr/day1_qr.png … qr/day5_qr.png  — QR code PNGs (printable)
  qr/day1_door.html … qr/day5_door.html  — printable door signs

Hosted at: https://anabeltan.github.io/
"""

from __future__ import annotations

import argparse
import base64
import html as html_mod
import re
import sys
from collections import defaultdict
from io import BytesIO
from pathlib import Path

from generate_programme_tex import Entry, build_entries

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

BASE_URL = "https://anabeltan.github.io"

# Must match exact day_label strings produced by build_entries()
DAY_CONFIG = [
    {
        "slug": "day1",
        "name": "Monday",
        "date": "June 29, 2026",
        "label": "Monday, June 29, 2026",
        "bg": "#0A1B25",
        "panel": "#152433",
        "panel2": "#1A2F3A",
        "accent": "#5AB4CC",
        "glow": "#A8DDE8",
        "text": "#E2F4F9",
    },
    {
        "slug": "day2",
        "name": "Tuesday",
        "date": "June 30, 2026",
        "label": "Tuesday, June 30, 2026",
        "bg": "#1A0B14",
        "panel": "#271220",
        "panel2": "#311A2B",
        "accent": "#F28A64",
        "glow": "#F6C177",
        "text": "#FDF0EC",
    },
    {
        "slug": "day3",
        "name": "Wednesday",
        "date": "July 1, 2026",
        "label": "Wednesday, July 01, 2026",
        "bg": "#071A1F",
        "panel": "#0D2228",
        "panel2": "#102E35",
        "accent": "#63C7BE",
        "glow": "#C6F3EE",
        "text": "#E5F9F7",
    },
    {
        "slug": "day4",
        "name": "Thursday",
        "date": "July 2, 2026",
        "label": "Thursday, July 02, 2026",
        "bg": "#0C1130",
        "panel": "#141936",
        "panel2": "#1B2346",
        "accent": "#90A7FF",
        "glow": "#F0D08F",
        "text": "#EEF0FF",
    },
    {
        "slug": "day5",
        "name": "Friday",
        "date": "July 3, 2026",
        "label": "Friday, July 03, 2026",
        "bg": "#160A0A",
        "panel": "#221010",
        "panel2": "#2D1A1A",
        "accent": "#E07070",
        "glow": "#F5C8A8",
        "text": "#FDF0F0",
    },
]

THEME_LABELS: dict[str, str] = {
    "1": "Black Hole Demographics and Scaling Relations",
    "2": "Black Hole Growth and AGN Feedback",
    "3": "Accretion, Jets and Multi-messenger Signatures",
    "4": "First Black Holes, Seeds and High-redshift",
}

THEME_ACCENT: dict[str, str] = {
    "1": "#C8922A",
    "2": "#5FA8B8",
    "3": "#90A7FF",
    "4": "#F28A64",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def primary_theme(theme_raw: str) -> str:
    m = re.search(r"\d", theme_raw)
    return m.group(0) if m else theme_raw


def theme_order(theme_dict: dict) -> list[str]:
    def sort_key(k: str) -> tuple[int, str]:
        try:
            return (int(k), "")
        except ValueError:
            return (999, k)
    return sorted(theme_dict.keys(), key=sort_key)


def group_by_theme(entries: list[Entry]) -> dict[str, list[Entry]]:
    result: dict[str, list[Entry]] = defaultdict(list)
    for e in entries:
        result[primary_theme(e.theme)].append(e)
    return result


def h(text: str) -> str:
    """HTML-escape a string."""
    return html_mod.escape(text or "")


def abstract_to_html(text: str) -> str:
    """Convert plain-text abstract to simple HTML paragraphs."""
    paras = [p.strip() for p in (text or "").split("\n\n") if p.strip()]
    if not paras:
        return "<p><em>Abstract not provided.</em></p>"
    return "".join(f"<p>{h(p)}</p>" for p in paras)


# ---------------------------------------------------------------------------
# QR code generation
# ---------------------------------------------------------------------------

def make_qr_png(url: str, fg: str = "#07111F", bg: str = "#F8F4E8") -> bytes | None:
    """Return PNG bytes for a QR code, or None if qrcode/PIL unavailable."""
    try:
        import qrcode
        from PIL import Image

        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=12,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color=fg, back_color=bg)
        buf = BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None


def png_to_b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


# ---------------------------------------------------------------------------
# CSS shared base (injected per-day with CSS custom properties)
# ---------------------------------------------------------------------------

BASE_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{
  background:var(--bg);
  color:var(--text);
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;
  font-size:16px;
  line-height:1.55;
  min-height:100vh;
}

/* ---- Header ---- */
.site-header{
  background:var(--panel2);
  border-bottom:2px solid var(--accent);
  padding:0.9rem 1rem 0.75rem;
  position:sticky;top:0;z-index:200;
}
.site-header .conf-name{
  font-size:0.7rem;font-weight:700;
  letter-spacing:0.1em;text-transform:uppercase;
  color:var(--glow);margin-bottom:0.15rem;
}
.site-header .day-name{
  font-size:1.55rem;font-weight:800;color:#fff;line-height:1.15;
}
.site-header .day-date{
  font-size:0.82rem;color:var(--accent);margin-top:0.1rem;
}
.site-header .star-row{
  font-size:0.65rem;letter-spacing:0.18em;color:var(--accent);
  opacity:0.5;margin-top:0.3rem;
}

/* ---- Day navigation ---- */
.day-nav{
  display:flex;gap:0.4rem;
  padding:0.6rem 0.75rem;
  background:var(--bg);
  border-bottom:1px solid rgba(255,255,255,0.06);
  overflow-x:auto;-webkit-overflow-scrolling:touch;
  scrollbar-width:none;
}
.day-nav::-webkit-scrollbar{display:none}
.day-nav a{
  flex-shrink:0;
  padding:0.3rem 0.7rem;
  border-radius:20px;
  font-size:0.78rem;font-weight:600;
  text-decoration:none;
  color:rgba(255,255,255,0.6);
  background:rgba(255,255,255,0.05);
  border:1px solid rgba(255,255,255,0.1);
  transition:background 0.15s,color 0.15s;
}
.day-nav a.active{
  background:var(--accent);
  color:#07111F;
  border-color:var(--accent);
  font-weight:700;
}
.day-nav a:hover:not(.active){
  background:rgba(255,255,255,0.1);
  color:#fff;
}

/* ---- Container ---- */
.container{max-width:640px;margin:0 auto;padding:0 0.75rem 3rem}

/* ---- Theme section ---- */
.theme-section{margin-top:1.6rem}
.theme-header{
  display:flex;align-items:flex-start;gap:0.55rem;
  padding:0.55rem 0.85rem;
  border-radius:8px;
  background:var(--panel);
  margin-bottom:0.75rem;
  border-left:3px solid var(--t-accent);
}
.theme-num{
  font-size:0.68rem;font-weight:700;
  color:var(--t-accent);
  letter-spacing:0.06em;text-transform:uppercase;
  white-space:nowrap;margin-top:0.1rem;
}
.theme-label{font-size:0.83rem;font-weight:600;color:var(--glow)}

/* ---- Talk card ---- */
.talk-card{
  background:var(--panel);
  border-radius:10px;
  margin-bottom:0.55rem;
  border:1px solid rgba(255,255,255,0.07);
  overflow:hidden;
}
.talk-card summary{
  padding:0.8rem 0.9rem 0.75rem;
  cursor:pointer;
  list-style:none;
  -webkit-tap-highlight-color:transparent;
  user-select:none;
}
.talk-card summary::-webkit-details-marker{display:none}
.talk-card[open] summary{border-bottom:1px solid rgba(255,255,255,0.07)}
.talk-top{
  display:flex;align-items:center;justify-content:space-between;gap:0.5rem;
  margin-bottom:0.3rem;
}
.talk-speaker{font-weight:700;font-size:0.95rem;color:#fff}
.invited-badge{
  flex-shrink:0;
  font-size:0.62rem;font-weight:700;
  padding:0.12rem 0.42rem;border-radius:12px;
  background:rgba(200,146,42,0.18);
  color:#C8922A;
  border:1px solid rgba(200,146,42,0.45);
  letter-spacing:0.05em;text-transform:uppercase;
}
.talk-title{
  font-size:0.87rem;color:var(--glow);line-height:1.4;
  padding-right:1.4rem;
  position:relative;
}
.chevron{
  position:absolute;right:0;top:0.1rem;
  font-size:0.65rem;color:var(--accent);
  transition:transform 0.2s;
}
details[open] .chevron{transform:rotate(180deg)}
.talk-abstract{
  padding:0.7rem 0.9rem 0.9rem;
  font-size:0.82rem;
  color:rgba(255,255,255,0.68);
  line-height:1.65;
}
.talk-abstract p+p{margin-top:0.6rem}
.talk-affil{
  font-size:0.74rem;
  color:rgba(255,255,255,0.38);
  margin-top:0.15rem;
}

/* ---- Empty day ---- */
.empty-day{
  margin-top:2rem;
  text-align:center;
  color:rgba(255,255,255,0.3);
  font-size:0.9rem;
  font-style:italic;
}

/* ---- Footer ---- */
.site-footer{
  margin-top:2.5rem;
  padding:1.2rem 1rem 2rem;
  border-top:1px solid rgba(255,255,255,0.06);
  text-align:center;
}
.site-footer .star-deco{
  font-size:0.7rem;letter-spacing:0.2em;
  color:var(--accent);opacity:0.4;margin-bottom:0.6rem;
}
.site-footer p{font-size:0.75rem;color:rgba(255,255,255,0.3);line-height:1.7}
.site-footer a{color:var(--accent);text-decoration:none}

/* ---- Index page ---- */
.index-hero{
  padding:2.5rem 1rem 2rem;text-align:center;
}
.index-hero h1{
  font-size:2rem;font-weight:800;color:#fff;
}
.index-hero .subtitle{
  font-size:0.95rem;color:var(--glow);margin-top:0.4rem;font-style:italic;
}
.index-hero .dates{
  font-size:0.85rem;color:var(--accent);margin-top:0.5rem;
}
.day-grid{
  max-width:480px;margin:0 auto;
  padding:0 0.75rem 3rem;
  display:flex;flex-direction:column;gap:0.75rem;
}
.day-btn{
  display:flex;align-items:center;justify-content:space-between;
  padding:1.1rem 1.2rem;
  border-radius:12px;
  text-decoration:none;
  background:var(--panel2);
  border:1px solid rgba(255,255,255,0.08);
  transition:border-color 0.15s,background 0.15s;
}
.day-btn:hover{
  border-color:var(--accent);
  background:var(--panel);
}
.day-btn-left .day-name-big{
  font-size:1.1rem;font-weight:700;color:#fff;
}
.day-btn-left .day-date-small{
  font-size:0.78rem;color:var(--accent);margin-top:0.1rem;
}
.day-btn-right{
  font-size:0.7rem;color:rgba(255,255,255,0.3);
}
.talk-count-badge{
  font-size:0.72rem;font-weight:600;
  padding:0.2rem 0.55rem;border-radius:20px;
  background:rgba(255,255,255,0.07);
  color:rgba(255,255,255,0.45);
}
"""


# ---------------------------------------------------------------------------
# HTML generators
# ---------------------------------------------------------------------------

def _css_vars(day: dict) -> str:
    return (
        f"--bg:{day['bg']};"
        f"--panel:{day['panel']};"
        f"--panel2:{day['panel2']};"
        f"--accent:{day['accent']};"
        f"--glow:{day['glow']};"
        f"--text:{day['text']};"
    )


def _nav_bar(current_slug: str, day_counts: dict[str, int]) -> str:
    links = []
    for dc in DAY_CONFIG:
        active = " active" if dc["slug"] == current_slug else ""
        count = day_counts.get(dc["label"], 0)
        if count == 0:
            continue
        links.append(
            f'<a href="{dc["slug"]}.html" class="{active.strip()}">'
            f'{dc["name"]}</a>'
        )
    return '<nav class="day-nav">' + "".join(links) + "</nav>"


def _talk_card(e: Entry) -> str:
    invited = e.presentation_type == "Invited Speaker"
    badge = '<span class="invited-badge">Invited</span>' if invited else ""
    affil = f'<div class="talk-affil">{h(e.affiliation)}</div>' if e.affiliation else ""
    abstract_html = abstract_to_html(e.abstract)
    return (
        "<details class=\"talk-card\">"
        "<summary>"
        f'<div class="talk-top"><span class="talk-speaker">{h(e.full_name)}</span>{badge}</div>'
        f'<div class="talk-title">{h(e.title)}<span class="chevron">▼</span></div>'
        f"{affil}"
        "</summary>"
        f'<div class="talk-abstract">{abstract_html}</div>'
        "</details>"
    )


def generate_day_html(
    day: dict,
    entries: list[Entry],
    day_counts: dict[str, int],
) -> str:
    css_vars = _css_vars(day)
    nav = _nav_bar(day["slug"], day_counts)

    # Group by theme
    by_theme = group_by_theme(entries)

    sections_html = ""
    for tk in theme_order(by_theme):
        label = THEME_LABELS.get(tk, f"Theme {tk}")
        t_accent = THEME_ACCENT.get(tk, day["accent"])
        talks_html = "\n".join(_talk_card(e) for e in by_theme[tk])
        sections_html += (
            f'<div class="theme-section">'
            f'<div class="theme-header" style="--t-accent:{t_accent}">'
            f'<span class="theme-num">Theme {h(tk)}</span>'
            f'<span class="theme-label">{h(label)}</span>'
            f"</div>"
            f"{talks_html}"
            f"</div>"
        )

    if not sections_html:
        sections_html = '<p class="empty-day">No talks scheduled for this day.</p>'

    total = sum(len(v) for v in by_theme.values())
    invited = sum(1 for e in entries if e.presentation_type == "Invited Speaker")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="theme-color" content="{day['bg']}">
<title>SMBH 2026 — {day['name']}, {day['date']}</title>
<style>
:root{{{css_vars}}}
{BASE_CSS}
</style>
</head>
<body>
<header class="site-header">
  <div class="conf-name">SMBH 2026 · Montr&eacute;al</div>
  <div class="day-name">{day['name']}</div>
  <div class="day-date">{day['date']}</div>
  <div class="star-row">&#9733; &nbsp; &#9679; &nbsp; &#9733; &nbsp; &#9679; &nbsp; &#9733;</div>
</header>
{nav}
<div class="container">
{sections_html}
</div>
<footer class="site-footer">
  <div class="star-deco">&#11819; &nbsp; &#9679; &nbsp; &#11819;</div>
  <p>{total} talks &middot; {invited} invited &middot; <a href="index.html">All days</a></p>
  <p style="margin-top:0.3rem">Tap any talk to expand its abstract</p>
</footer>
</body>
</html>
"""


def generate_index_html(day_counts: dict[str, int]) -> str:
    # Use day1 (Monday) colours for the index
    day = DAY_CONFIG[0]
    css_vars = _css_vars(day)

    buttons = ""
    for dc in DAY_CONFIG:
        count = day_counts.get(dc["label"], 0)
        if count == 0:
            continue
        buttons += (
            f'<a class="day-btn" href="{dc["slug"]}.html">'
            f'<div class="day-btn-left">'
            f'<div class="day-name-big">{dc["name"]}</div>'
            f'<div class="day-date-small">{dc["date"]}</div>'
            f"</div>"
            f'<span class="talk-count-badge">{count} talks</span>'
            f"</a>"
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="theme-color" content="{day['bg']}">
<title>SMBH 2026 — Schedule</title>
<style>
:root{{{css_vars}}}
{BASE_CSS}
</style>
</head>
<body>
<div class="index-hero">
  <div style="font-size:0.72rem;font-weight:700;letter-spacing:0.1em;
              text-transform:uppercase;color:var(--glow);margin-bottom:0.6rem">
    Conference Schedule
  </div>
  <h1>SMBH 2026</h1>
  <div class="subtitle">Supermassive Black Holes: From Seeds to Giants</div>
  <div class="dates">June 29 &ndash; July 3, 2026 &middot; Montr&eacute;al</div>
  <div style="font-size:0.7rem;letter-spacing:0.2em;color:var(--accent);
              opacity:0.45;margin-top:0.9rem">
    &#9733; &nbsp; &#9679; &nbsp; &#9733;
  </div>
</div>
<div class="day-grid">
{buttons}
</div>
<footer class="site-footer">
  <div class="star-deco">&#11819; &nbsp; &#9679; &nbsp; &#11819;</div>
  <p>Tap a day to see its full schedule and abstracts</p>
</footer>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Door sign HTML (printable A4 / Letter)
# ---------------------------------------------------------------------------

def generate_door_html(day: dict, qr_b64: str | None, url: str) -> str:
    qr_section = ""
    if qr_b64:
        qr_section = (
            f'<img src="data:image/png;base64,{qr_b64}" '
            f'alt="QR code" style="width:240px;height:240px;'
            f'border-radius:12px;display:block;margin:0 auto">'
        )
    else:
        qr_section = (
            f'<div style="width:240px;height:240px;margin:0 auto;'
            f'border:2px dashed rgba(255,255,255,0.2);border-radius:12px;'
            f'display:flex;align-items:center;justify-content:center;'
            f'color:rgba(255,255,255,0.4);font-size:0.85rem;text-align:center;padding:1rem">'
            f"QR code<br>(qrcode package needed)</div>"
        )

    accent = day["accent"]
    glow = day["glow"]
    bg = day["bg"]
    panel = day["panel2"]

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>SMBH 2026 Door Sign — {day['name']}</title>
<style>
  @page{{size:A4 portrait;margin:0}}
  *{{box-sizing:border-box;margin:0;padding:0}}
  html,body{{width:210mm;height:297mm;background:{bg};color:#fff;
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif}}
  .page{{
    width:210mm;height:297mm;
    display:flex;flex-direction:column;
    align-items:center;justify-content:center;
    padding:2cm;text-align:center;gap:1.8rem;
  }}
  .conf{{font-size:1rem;font-weight:700;letter-spacing:0.12em;
         text-transform:uppercase;color:{glow};}}
  .day{{font-size:4rem;font-weight:900;color:#fff;line-height:1}}
  .date{{font-size:1.4rem;color:{accent};font-weight:600;margin-top:0.3rem}}
  .rule{{width:5cm;height:3px;background:{accent};border-radius:2px}}
  .scan{{font-size:1.1rem;color:rgba(255,255,255,0.6);letter-spacing:0.02em}}
  .url{{font-size:0.85rem;color:{accent};letter-spacing:0.03em;
        word-break:break-all;opacity:0.75}}
  .star{{font-size:0.8rem;letter-spacing:0.25em;color:{accent};opacity:0.35}}
</style>
</head>
<body>
<div class="page">
  <div class="star">&#9733; &nbsp; &#9679; &nbsp; &#9733; &nbsp; &#9679; &nbsp; &#9733;</div>
  <div>
    <div class="conf">SMBH 2026 &mdash; Today&apos;s Schedule</div>
    <div class="day">{day['name']}</div>
    <div class="date">{day['date']}</div>
  </div>
  <div class="rule"></div>
  {qr_section}
  <div>
    <div class="scan">Scan for talks &amp; abstracts</div>
    <div class="url" style="margin-top:0.4rem">{url}</div>
  </div>
  <div class="star">&#11819; &nbsp; &#9679; &nbsp; &#11819;</div>
</div>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate SMBH 2026 mobile schedule pages and QR codes."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="SMBH 2026 Participant Tracking - MASTER.xlsx",
    )
    parser.add_argument("-s", "--sheet", default=None)
    parser.add_argument(
        "--base-url",
        default=BASE_URL,
        help="Base URL where pages will be hosted (no trailing slash).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Workbook not found: {input_path}", file=sys.stderr)
        return 1

    try:
        all_entries = build_entries(input_path, args.sheet)
    except Exception as exc:
        print(f"Error reading workbook: {exc}", file=sys.stderr)
        return 1

    base_url = args.base_url.rstrip("/")
    scheduled = [e for e in all_entries if e.day_label != "Unscheduled"]

    # Build per-day entry lists
    by_day: dict[str, list[Entry]] = defaultdict(list)
    for e in scheduled:
        by_day[e.day_label].append(e)

    day_counts = {dc["label"]: len(by_day[dc["label"]]) for dc in DAY_CONFIG}

    web_dir = Path("web")
    qr_dir = Path("qr")
    web_dir.mkdir(exist_ok=True)
    qr_dir.mkdir(exist_ok=True)

    # ---- Day pages + QR codes ----
    for dc in DAY_CONFIG:
        entries = by_day.get(dc["label"], [])
        if not entries:
            print(f"  (skipping {dc['slug']} — no entries)")
            continue

        url = f"{base_url}/{dc['slug']}.html"

        # Generate QR PNG
        qr_data = make_qr_png(url, fg="#07111F", bg="#F8F4E8")
        qr_b64: str | None = None
        qr_png_path = qr_dir / f"{dc['slug']}_qr.png"
        if qr_data:
            qr_png_path.write_bytes(qr_data)
            qr_b64 = png_to_b64(qr_data)
            print(f"  QR  {qr_png_path}")
        else:
            print(f"  QR  (skipped — install qrcode+pillow)")

        # Day HTML
        html_content = generate_day_html(dc, entries, day_counts)
        html_path = web_dir / f"{dc['slug']}.html"
        html_path.write_text(html_content, encoding="utf-8")
        print(f"  Web {html_path}  ({len(entries)} talks)")

        # Door sign HTML
        door_html = generate_door_html(dc, qr_b64, url)
        door_path = qr_dir / f"{dc['slug']}_door.html"
        door_path.write_text(door_html, encoding="utf-8")
        print(f"  Door {door_path}")

    # ---- Index page ----
    index_html = generate_index_html(day_counts)
    index_path = web_dir / "index.html"
    index_path.write_text(index_html, encoding="utf-8")
    print(f"  Index {index_path}")

    print()
    print("Done. Deploy the web/ folder to GitHub Pages.")
    print(f"  Pages live at: {base_url}/")
    print()
    print("QR codes and printable door signs are in qr/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
