#!/usr/bin/env python3
"""Create a self-contained interactive timeline from tables in a Word document."""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from docx import Document

VERSION = "1.0.0"

EVENT_ALIASES = {
    "id": {"id", "event id", "event_id"},
    "start": {"start", "date", "start date", "start_date", "when"},
    "end": {"end", "end date", "end_date", "finish", "finish date"},
    "title": {"title", "event", "label", "name"},
    "description": {"description", "details", "summary", "notes"},
    "category": {"category", "type", "theme"},
    "group": {"group", "track", "lane", "stream"},
    "color": {"color", "colour", "background color", "background colour"},
    "link": {"link", "url", "web link"},
}

SETTING_ALIASES = {
    "title": {"title", "timeline title"},
    "subtitle": {"subtitle", "description"},
    "height": {"height", "timeline height"},
    "orientation": {"orientation", "axis position"},
    "stack": {"stack", "stack events"},
    "zoom_min_days": {"zoom min days", "minimum zoom days"},
    "zoom_max_days": {"zoom max days", "maximum zoom days"},
}


@dataclass
class Issue:
    level: str
    message: str
    row: int | None = None


@dataclass
class ParseResult:
    settings: dict[str, str] = field(default_factory=dict)
    events: list[dict[str, Any]] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)


def normalise(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("_", " ").strip().lower())


def clean_cell(cell) -> str:
    return "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip()).strip()


def canonical_header(raw: str, aliases: dict[str, set[str]]) -> str | None:
    key = normalise(raw)
    for canonical, variants in aliases.items():
        if key == canonical or key in variants:
            return canonical
    return None


def parse_bool(raw: str, default: bool, result: ParseResult, context: str) -> bool:
    if not raw:
        return default
    value = normalise(raw)
    if value in {"true", "yes", "y", "1", "on"}:
        return True
    if value in {"false", "no", "n", "0", "off"}:
        return False
    result.issues.append(Issue("warning", f"{context}: expected yes/no; using {default}."))
    return default


def parse_date(raw: str) -> str | None:
    value = raw.strip()
    if not value:
        return None
    formats = (
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y",
        "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
        "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M",
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed.isoformat(timespec="minutes") if "H" in fmt else parsed.date().isoformat()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).isoformat()
    except ValueError:
        return None


def valid_color(value: str) -> bool:
    return bool(
        re.fullmatch(r"#[0-9a-fA-F]{3,8}", value)
        or re.fullmatch(r"[a-zA-Z]+", value)
        or re.fullmatch(r"(rgb|hsl)a?\([0-9.,%\s]+\)", value)
    )


def valid_link(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def table_rows(table) -> list[list[str]]:
    return [[clean_cell(cell) for cell in row.cells] for row in table.rows]


def detect_table_type(rows: list[list[str]], preceding_heading: str) -> str | None:
    heading = normalise(preceding_heading)
    if heading in {"events", "timeline events", "event data"}:
        return "events"
    if heading in {"settings", "timeline settings", "configuration"}:
        return "settings"
    if not rows:
        return None
    headers = {normalise(x) for x in rows[0]}
    if headers & EVENT_ALIASES["start"] and headers & EVENT_ALIASES["title"]:
        return "events"
    if {"setting", "value"} <= headers or {"key", "value"} <= headers:
        return "settings"
    return None


def iter_blocks(document):
    """Yield paragraphs and tables in document order."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, document)
        elif child.tag.endswith("}tbl"):
            yield Table(child, document)


def read_docx(path: Path) -> ParseResult:
    result = ParseResult()
    document = Document(path)
    heading = ""
    event_tables = 0

    from docx.table import Table

    for block in iter_blocks(document):
        if isinstance(block, Table):
            rows = table_rows(block)
            kind = detect_table_type(rows, heading)
            if kind == "settings":
                parse_settings(rows, result)
            elif kind == "events":
                event_tables += 1
                parse_events(rows, result)
        elif block.text.strip():
            heading = block.text.strip()

    if event_tables == 0:
        result.issues.append(Issue("error", "No EVENTS table found. Add a table with start/date and title columns."))
    if not result.events:
        result.issues.append(Issue("error", "No valid events were found."))
    return result


def parse_settings(rows: list[list[str]], result: ParseResult) -> None:
    if len(rows) < 2:
        return
    headers = [normalise(x) for x in rows[0]]
    try:
        key_i = next(i for i, h in enumerate(headers) if h in {"setting", "key", "option", "name"})
        value_i = next(i for i, h in enumerate(headers) if h in {"value", "setting value"})
    except StopIteration:
        result.issues.append(Issue("warning", "SETTINGS table ignored: it needs Setting and Value columns."))
        return
    for row_num, row in enumerate(rows[1:], start=2):
        if max(key_i, value_i) >= len(row):
            continue
        raw_key, value = normalise(row[key_i]), row[value_i].strip()
        if not raw_key:
            continue
        key = canonical_header(raw_key, SETTING_ALIASES)
        if key:
            result.settings[key] = value
        else:
            result.issues.append(Issue("warning", f"Unknown setting '{row[key_i]}' ignored.", row_num))


def parse_events(rows: list[list[str]], result: ParseResult) -> None:
    if len(rows) < 2:
        result.issues.append(Issue("warning", "An EVENTS table has no data rows."))
        return
    headers: list[str | None] = []
    seen: set[str] = set()
    for raw in rows[0]:
        canonical = canonical_header(raw, EVENT_ALIASES)
        if canonical and canonical in seen:
            result.issues.append(Issue("warning", f"Duplicate '{canonical}' column; later column ignored."))
            canonical = None
        if canonical:
            seen.add(canonical)
        elif raw.strip():
            result.issues.append(Issue("warning", f"Unknown EVENTS column '{raw}' ignored."))
        headers.append(canonical)

    if "start" not in seen or "title" not in seen:
        result.issues.append(Issue("error", "EVENTS table requires a start/date column and a title/event column."))
        return

    existing_ids = {event["id"] for event in result.events}
    for row_num, row in enumerate(rows[1:], start=2):
        record = {headers[i]: value.strip() for i, value in enumerate(row) if i < len(headers) and headers[i]}
        if not any(record.values()):
            continue
        start = parse_date(record.get("start", ""))
        if not start:
            result.issues.append(Issue("error", f"Invalid or missing start date '{record.get('start', '')}'. Row skipped.", row_num))
            continue
        title = record.get("title", "").strip()
        if not title:
            result.issues.append(Issue("error", "Missing event title. Row skipped.", row_num))
            continue
        end = parse_date(record.get("end", ""))
        if record.get("end") and not end:
            result.issues.append(Issue("warning", f"Invalid end date '{record['end']}'; event treated as a point.", row_num))
        if end and end < start:
            result.issues.append(Issue("error", "End date is before start date. Row skipped.", row_num))
            continue
        event_id = record.get("id") or f"event-{len(result.events) + 1}"
        base_id, suffix = event_id, 2
        while event_id in existing_ids:
            event_id = f"{base_id}-{suffix}"
            suffix += 1
        if event_id != base_id:
            result.issues.append(Issue("warning", f"Duplicate id '{base_id}' renamed to '{event_id}'.", row_num))
        existing_ids.add(event_id)

        color = record.get("color", "")
        if color and not valid_color(color):
            result.issues.append(Issue("warning", f"Invalid colour '{color}' ignored.", row_num))
            color = ""
        link = record.get("link", "")
        if link and not valid_link(link):
            result.issues.append(Issue("warning", f"Unsafe or invalid link '{link}' ignored.", row_num))
            link = ""

        event = {
            "id": event_id, "start": start, "title": title,
            "description": record.get("description", ""),
            "category": record.get("category", ""),
            "group": record.get("group", ""),
            "color": color, "link": link,
        }
        if end:
            event["end"] = end
        result.events.append(event)


def get_asset(project: Path, filename: str) -> str:
    path = project / "vendor" / filename
    if not path.exists():
        raise FileNotFoundError(
            f"Missing {path}. Run 'npm install' and 'python3 tools/copy_vendor_assets.py' "
            "or use --cdn."
        )
    return path.read_text(encoding="utf-8")


def build_html(result: ParseResult, project: Path, use_cdn: bool) -> str:
    settings = result.settings
    title = settings.get("title") or "Interactive Timeline"
    subtitle = settings.get("subtitle", "")
    height = settings.get("height", "620px")
    if not re.fullmatch(r"\d+(px|vh|rem|em|%)", height):
        result.issues.append(Issue("warning", f"Invalid height '{height}'; using 620px."))
        height = "620px"
    orientation = normalise(settings.get("orientation", "bottom"))
    if orientation not in {"top", "bottom", "both", "none"}:
        result.issues.append(Issue("warning", f"Invalid orientation '{orientation}'; using bottom."))
        orientation = "bottom"
    stack = parse_bool(settings.get("stack", ""), True, result, "stack")

    categories = sorted({x["category"] for x in result.events if x["category"]})
    groups = sorted({x["group"] for x in result.events if x["group"]})
    group_ids = {name: index + 1 for index, name in enumerate(groups)}
    items = []
    for event in result.events:
        detail = f"<strong>{html.escape(event['title'])}</strong>"
        if event["description"]:
            detail += f"<p>{html.escape(event['description']).replace(chr(10), '<br>')}</p>"
        if event["category"]:
            detail += f"<p><b>Category:</b> {html.escape(event['category'])}</p>"
        if event["link"]:
            detail += f'<p><a href="{html.escape(event["link"])}" target="_blank" rel="noopener">Open link</a></p>'
        item = {
            "id": event["id"], "content": html.escape(event["title"]),
            "start": event["start"], "title": detail,
            "category": event["category"],
        }
        if event.get("end"):
            item["end"] = event["end"]
        if event["group"]:
            item["group"] = group_ids[event["group"]]
        if event["color"]:
            item["style"] = f"background-color:{event['color']};border-color:{event['color']}"
        items.append(item)

    if use_cdn:
        assets = (
            '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/vis-timeline@8.5.0/styles/vis-timeline-graph2d.min.css">\n'
            '<script src="https://cdn.jsdelivr.net/npm/vis-timeline@8.5.0/standalone/umd/vis-timeline-graph2d.min.js"></script>'
        )
    else:
        css = get_asset(project, "vis-timeline-graph2d.min.css")
        js = get_asset(project, "vis-timeline-graph2d.min.js")
        assets = f"<style>{css}</style>\n<script>{js}</script>"

    data_json = json.dumps(items, ensure_ascii=False).replace("</", "<\\/")
    groups_json = json.dumps(
        [{"id": group_ids[name], "content": html.escape(name)} for name in groups],
        ensure_ascii=False,
    ).replace("</", "<\\/")
    categories_json = json.dumps(categories, ensure_ascii=False).replace("</", "<\\/")
    subtitle_html = f"<p class='subtitle'>{html.escape(subtitle)}</p>" if subtitle else ""

    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title>{assets}
<style>
:root{{--ink:#172033;--muted:#5f6b7a;--panel:#fff;--line:#d8dee9;--accent:#2563eb}}
*{{box-sizing:border-box}} body{{margin:0;background:#f6f8fb;color:var(--ink);font:16px/1.5 system-ui,-apple-system,sans-serif}}
main{{max-width:1400px;margin:auto;padding:2rem}} h1{{margin:0 0 .25rem;font-size:clamp(1.7rem,3vw,2.5rem)}}
.subtitle{{margin:.2rem 0 1.5rem;color:var(--muted)}} .toolbar{{display:flex;flex-wrap:wrap;gap:.75rem;align-items:end;margin:1rem 0}}
label{{font-weight:650;font-size:.9rem}} select,input,button{{font:inherit;padding:.55rem .7rem;border:1px solid #aeb7c5;border-radius:.45rem;background:white}}
button{{cursor:pointer}} button:hover{{border-color:var(--accent)}} #timeline{{height:{height};background:var(--panel);border:1px solid var(--line);border-radius:.7rem}}
.summary{{color:var(--muted);margin:.65rem 0}} .vis-item{{border-radius:.35rem}} .vis-item.vis-selected{{border-color:#111;box-shadow:0 0 0 2px white,0 0 0 4px #111}}
.sr-only{{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}}
@media(max-width:700px){{main{{padding:1rem}}#timeline{{height:70vh}}}}
</style></head><body><main>
<h1>{html.escape(title)}</h1>{subtitle_html}
<div class="toolbar" aria-label="Timeline controls">
<label>Search<br><input id="search" type="search" placeholder="Find an event"></label>
<label>Category<br><select id="category"><option value="">All categories</option></select></label>
<button id="fit" type="button">Fit all events</button>
</div>
<p id="summary" class="summary" aria-live="polite"></p>
<div id="timeline" aria-label="{html.escape(title)}"></div>
<p class="sr-only">Use the controls above to filter the timeline. Select an event for details.</p>
</main><script>
const allItems={data_json}, groupData={groups_json}, categories={categories_json};
const dataSet=new vis.DataSet(allItems), groups=new vis.DataSet(groupData);
const timeline=new vis.Timeline(document.getElementById('timeline'),dataSet,groupData.length?groups:null,{{
  stack:{str(stack).lower()},orientation:{json.dumps(orientation)},zoomMin:86400000,
  zoomMax:315576000000,multiselect:false,tooltip:{{followMouse:true,overflowMethod:'cap'}}
}});
const search=document.getElementById('search'), category=document.getElementById('category'), summary=document.getElementById('summary');
categories.forEach(x=>{{const o=document.createElement('option');o.value=x;o.textContent=x;category.appendChild(o)}});
function filter(){{
 const q=search.value.trim().toLowerCase(), c=category.value;
 const shown=allItems.filter(x=>(!c||x.category===c)&&(!q||x.content.toLowerCase().includes(q)||x.title.toLowerCase().includes(q)));
 dataSet.clear();dataSet.add(shown);summary.textContent=`Showing ${{shown.length}} of ${{allItems.length}} events`;
}}
search.addEventListener('input',filter);category.addEventListener('change',filter);
document.getElementById('fit').addEventListener('click',()=>timeline.fit({{animation:true}}));
filter();timeline.fit();
</script></body></html>"""


def write_qa(path: Path, source: Path, result: ParseResult) -> None:
    errors = sum(x.level == "error" for x in result.issues)
    warnings = sum(x.level == "warning" for x in result.issues)
    lines = [
        "# Timeline QA report", "", f"- Source: `{source.name}`",
        f"- Valid events: {len(result.events)}", f"- Errors: {errors}",
        f"- Warnings: {warnings}", "", "## Findings", "",
    ]
    if result.issues:
        for issue in result.issues:
            location = f" (row {issue.row})" if issue.row else ""
            lines.append(f"- **{issue.level.title()}**{location}: {issue.message}")
    else:
        lines.append("- No issues found.")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Word .docx file containing an EVENTS table")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/timeline"))
    parser.add_argument("--cdn", action="store_true", help="Use CDN assets instead of embedding local assets")
    parser.add_argument("--strict", action="store_true", help="Return a failure code when warnings exist")
    parser.add_argument("--version", action="version", version=VERSION)
    args = parser.parse_args(argv)

    if not args.input.is_file() or args.input.suffix.lower() != ".docx":
        parser.error("input must be an existing .docx file")
    result = read_docx(args.input)
    args.output.mkdir(parents=True, exist_ok=True)
    write_qa(args.output / "qa_report.md", args.input, result)
    errors = [x for x in result.issues if x.level == "error"]
    warnings = [x for x in result.issues if x.level == "warning"]
    if errors:
        print(f"Validation failed: {len(errors)} error(s). See {args.output / 'qa_report.md'}", file=sys.stderr)
        return 2
    html_text = build_html(result, Path(__file__).resolve().parent, args.cdn)
    (args.output / "index.html").write_text(html_text, encoding="utf-8")
    (args.output / "data.json").write_text(json.dumps(result.events, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Created {args.output / 'index.html'} from {len(result.events)} events ({len(warnings)} warnings).")
    return 1 if args.strict and warnings else 0


if __name__ == "__main__":
    raise SystemExit(cli())
