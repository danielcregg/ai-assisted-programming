#!/usr/bin/env python3
"""Render every lecture for the site: lectures-and-labs/<weekNN>/<topic>-lecture.md ->
OUTPUT_DIR/<topic>/index.html and OUTPUT_DIR/<topic>/slides.pdf.

<topic> is the schedule's `lecture` name, not the week folder, so a lecture keeps
its web address when the semester is renumbered. A week's img/ folder is copied
beside the output. The Marp CLI is `marp` (CI installs it globally); set MARP to
run it another way, e.g. MARP="npx --no-install marp".

A PowerPoint lecture (<topic>-lecture.pptx) is not rendered here: CI cannot run
PowerPoint. Its PDF and text copy were exported beside it on Windows
(scripts/export_decks.py) and committed, so this publishes the PDF as slides.pdf,
the deck as slides.pptx, and an index.html showing the PDF above every slide's
text and speaker notes, taken from the text copy.

stdin is closed for every marp call: marp-cli treats piped stdin as an extra
markdown input, which once turned a loop over the decks into "Converting 2
markdowns" and an output-path error.

Usage:
    python scripts/render_decks.py [OUTPUT_DIR]      # default: build
"""
from __future__ import annotations

import html
import os
import re
import shlex
import shutil
import subprocess
import sys
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent))
from schedule import SITE, Row, frontmatter, load, stale_exports  # noqa: E402

THEME = "themes/aiap.css"
NOTE_RE = re.compile(r"<!--\s*Speaker notes:(.*?)-->", re.S)
COMMENT_RE = re.compile(r"<!--.*?-->", re.S)
# Microsoft's free Office viewer renders a public .pptx in the browser, builds
# included. Microsoft does not support it for production use, hence "experimental".
OFFICE_VIEWER = "https://view.officeapps.live.com/op/view.aspx?src="

DECK_CSS = """<style>
  .wrap { max-width: 1120px; }
  .standfirst { color: var(--slate); max-width: 66ch; margin: 4px 0 0; }
  .deck-actions { display: flex; flex-wrap: wrap; gap: 12px; margin: 18px 0 22px; }
  button.open { background: none; cursor: pointer; }
  object.pdf { display: block; width: 100%; aspect-ratio: 16 / 10; background: #FFFFFF;
               border: 1px solid var(--rule); border-radius: 10px; }
  .nopdf { padding: 20px 24px; }
  .slides { max-width: 780px; }
  .slide { border-top: 1px solid var(--rule); padding-bottom: 10px; }
  .slide-no { font-family: var(--mono); font-size: 13px; margin: 14px 0 0; }
  .slide-no a { color: var(--muted); }
  .slide h3 { color: var(--ink); margin-top: 8px; }
  .slide h4 { font-family: var(--mono); font-size: 16px; color: var(--slate); margin: 16px 0 4px; }
  details.notes { border-color: var(--rule); background: #FDFCF9; }
  details.notes summary { font-size: 14px; color: var(--slate); }
  details.notes p { color: var(--slate); font-size: 15.5px; }
  @media (max-width: 700px) { object.pdf { aspect-ratio: 4 / 3; } }
</style>
"""

# Opens or closes every note at once; without JS the button stays hidden.
NOTES_JS = """<script>
(function () {
  var b = document.getElementById('all-notes');
  if (!b) return;
  b.hidden = false;
  b.addEventListener('click', function () {
    var notes = document.querySelectorAll('details.notes');
    var open = !(notes.length && notes[0].open);
    notes.forEach(function (d) { d.open = open; });
    b.textContent = open ? 'hide every note' : 'show every note';
  });
})();
</script>"""


def split_slides(body: str) -> list[str]:
    """Split on `---` lines, but never inside a code fence (a YAML example can
    start with one)."""
    slides, current, fence = [], [], ""
    for line in body.split("\n"):
        m = re.match(r"\s*(`{3,}|~{3,})", line)
        if m and not fence:
            fence = m.group(1)
        elif m and re.fullmatch(rf"\s*{re.escape(fence[0])}{{{len(fence)},}}\s*", line):
            fence = ""
        elif not fence and re.fullmatch(r"---\s*", line):
            slides.append("\n".join(current))
            current = []
            continue
        current.append(line)
    slides.append("\n".join(current))
    return slides


def demote(body_html: str) -> str:
    """Slide headings sit under the page's own: h1 and h2 become h3, h3 becomes h4."""
    return re.sub(r"<(/?)h([1-3])\b",
                  lambda m: f"<{m.group(1)}h{4 if m.group(2) == '3' else 3}", body_html)


def deck_page(row: Row) -> str:
    """The page for a PowerPoint lecture, built from its text copy."""
    from build_lab_pages import RENDERER, page   # markdown-it-py, as the lab pages use
    fields, body = frontmatter(row.lecture_text.read_text(encoding="utf-8"))
    title = html.escape(fields.get("title") or row.topic)
    sections = []
    for md in split_slides(body):
        notes = [n.strip() for n in NOTE_RE.findall(md)]
        md = COMMENT_RE.sub("", md).strip()
        if not md and not notes:
            continue
        n = len(sections) + 1
        paras = [p for note in notes for p in re.split(r"\n+", note) if p.strip()]
        notes_html = "".join(f"<p>{html.escape(p)}</p>\n" for p in paras)
        sections.append(
            f'<section class="slide" id="slide-{n}">\n'
            f'<p class="slide-no"><a href="slides.pdf#page={n}">slide {n}</a></p>\n'
            f"{demote(RENDERER.render(md))}"
            + (f'<details class="notes"><summary>speaker notes</summary>\n{notes_html}</details>\n'
               if notes_html else "")
            + "</section>\n")
    viewer = OFFICE_VIEWER + quote(f"{SITE}{row.deck}/slides.pptx", safe="")
    body_html = (
        f"<h1>{title}</h1>\n"
        f'<p class="standfirst">The slides as a PDF, with every slide\'s text and speaker '
        f"notes below. Download the PowerPoint to click through it as it is presented.</p>\n"
        f'<p class="deck-actions">'
        f'<a class="open" href="slides.pdf">slides (PDF)</a>'
        f'<a class="open" href="slides.pptx" download>PowerPoint (.pptx)</a>'
        f'<a class="open" href="{html.escape(viewer)}" rel="noopener">'
        f"Microsoft web viewer (experimental)</a></p>\n"
        # navpanes=0: Chrome's viewer otherwise opens with a thumbnail rail that
        # takes a third of the frame. Viewers that do not know it ignore it.
        f'<object class="pdf" data="slides.pdf#navpanes=0" type="application/pdf" '
        f'aria-label="{title}: the slides">\n'
        f'<p class="nopdf">This browser cannot show the PDF inside the page: '
        f'<a href="slides.pdf">open the slides (PDF)</a>.</p>\n</object>\n'
        f'<h2 id="every-slide">Every slide, as text</h2>\n'
        f"<p>All {len(sections)} slides, each with the speaker notes that explain it. "
        f'<button type="button" class="open" id="all-notes" hidden>show every note</button></p>\n'
        f'<div class="slides">\n{"".join(sections)}</div>\n{NOTES_JS}')
    return page(title, '<a href="../">ai-assisted programming</a> · lecture', body_html, False,
                needs_hljs="language-python" in body_html, extra_css=DECK_CSS)


def publish_pptx(row: Row, out: Path) -> None:
    """Publish a PowerPoint lecture's committed exports; refuse stale ones, as
    check_schedule.py does, rather than put an old PDF on the site."""
    stale = stale_exports(row)
    if stale:
        raise SystemExit(f"render_decks: week {row.week}: {'; '.join(stale)}; "
                         f"run scripts/export_decks.py")
    shutil.copyfile(row.lecture_pdf, out / "slides.pdf")
    shutil.copyfile(row.lecture, out / "slides.pptx")
    (out / "index.html").write_text(deck_page(row), encoding="utf-8", newline="\n")


def main() -> None:
    out_root = Path(sys.argv[1] if len(sys.argv) > 1 else "build")
    marp = shlex.split(os.environ.get("MARP", "marp"))
    if os.name == "nt" and marp[0] in ("marp", "npx"):
        marp[0] += ".cmd"           # Windows resolves the npm shims by their .cmd name
    rows = load().teaching
    for row in rows:
        src = row.lecture
        if not src.is_file():
            raise SystemExit(f"render_decks: week {row.week} ({row.topic}) has no {src.as_posix()}")
        out = out_root / row.deck
        out.mkdir(parents=True, exist_ok=True)
        if row.is_pptx:
            publish_pptx(row, out)
            print(f"published {src.as_posix()} and its exports -> {out.as_posix()}/")
            continue
        if (row.path / "img").is_dir():
            shutil.copytree(row.path / "img", out / "img", dirs_exist_ok=True)
        for target in ("index.html", "slides.pdf"):
            subprocess.run([*marp, str(src), "--html", "--allow-local-files",
                            "--theme-set", THEME, "-o", str(out / target)],
                           check=True, stdin=subprocess.DEVNULL)
        print(f"rendered {src.as_posix()} -> {out.as_posix()}/")
    print(f"render_decks: {len(rows)} lectures")


if __name__ == "__main__":
    main()
