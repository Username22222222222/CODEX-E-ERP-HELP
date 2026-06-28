from __future__ import annotations

import html
import json
import re
from pathlib import Path
from urllib.parse import quote, unquote


HELP_ROOT = Path(r"D:\DATEN\HOMEPAGES\x-erp.de\de\help")
TREE_JSON = HELP_ROOT / "xlsx-tree.json"
SCRIPT_VERSION = "20260628-stammdaten-quality"
HELP_CSS_VERSION = "20260628-stammdaten-quality"
CSS_VERSION = "20260627-screenshot-width"
REDIRECTS = {}


def segment_title(segment: str) -> str:
    segment = unquote(segment).strip("/")
    if segment.lower().endswith(".html"):
        segment = segment[:-5]
    if not segment:
        return ""

    special = {
        "faq": "FAQ",
        "api": "API",
        "iis": "IIS",
        "ki": "KI",
        "pwa": "PWA",
        "sql": "SQL",
        "ssl": "SSL",
        "tls": "TLS",
        "xld": "XLD",
        "erp": "ERP",
        "x-erp": "X-ERP",
        "e-mail": "E-Mail",
        "email": "E-Mail",
        "portal-apps": "Portal-Apps",
        "api-schnittstellen": "API-Schnittstellen",
        "email-service": "E-Mail-Service",
    }
    lowered = segment.casefold()
    if lowered in special:
        return special[lowered]

    parts = [part for part in re.split(r"[-_]+", segment) if part]
    words: list[str] = []
    for part in parts:
        key = part.casefold()
        words.append(special.get(key, part[:1].upper() + part[1:]))
    label = " ".join(words)
    return label.replace("X ERP", "X-ERP").replace("E Mail", "E-Mail").replace("Api", "API")


def path_breadcrumb_items(canonical_path: str, title_by_path: dict[str, str], fallback_title: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = [("X-ERP Hilfe", "/de/help/")]
    rel = normalize_url_path(canonical_path).removeprefix("/de/help/").strip("/")
    if not rel:
        return items

    parts = [part for part in rel.split("/") if part]
    cumulative = "/de/help/"
    for index, part in enumerate(parts):
        is_file = part.lower().endswith(".html")
        clean_part = part[:-5] if is_file else part
        if is_file:
            path = cumulative + part
        else:
            path = cumulative + part + "/"
            cumulative = path

        label = title_by_path.get(normalize_url_path(path)) or segment_title(clean_part)
        if index == len(parts) - 1:
            label = fallback_title or label
        if label:
            items.append((label, normalize_url_path(path)))
    return items


def repair_text(value: object) -> str:
    text = "" if value is None else str(value)
    if not any(marker in text for marker in ("Ã", "Â", "â€", "â€“", "â€¢")):
        return text
    try:
        fixed = text.encode("latin1").decode("utf-8")
    except UnicodeError:
        return text
    bad_before = sum(text.count(marker) for marker in ("Ã", "Â", "â€", "â€“", "â€¢"))
    bad_after = sum(fixed.count(marker) for marker in ("Ã", "Â", "â€", "â€“", "â€¢"))
    return fixed if bad_after < bad_before else text


def normalize_url_path(path: str) -> str:
    path = unquote(path).replace("\\", "/")
    if not path.startswith("/"):
        path = "/" + path
    if not path.endswith("/") and not path.lower().endswith(".html"):
        path += "/"
    return path


def flatten(nodes: list[dict], parent: dict | None = None) -> list[dict]:
    result: list[dict] = []
    for node in nodes:
        node["_parent"] = parent
        node["title"] = repair_text(node.get("title", ""))
        node["pageTitle"] = repair_text(node.get("pageTitle", ""))
        node["breadcrumb"] = repair_text(node.get("breadcrumb", ""))
        result.append(node)
        result.extend(flatten(node.get("children") or [], node))
    return result


def strip_markdown(text: str) -> str:
    text = repair_text(text)
    text = re.sub(r"^#{1,6}\s*", "", text, flags=re.M)
    text = re.sub(r"^[\s>*-]*[•-]\s*", "", text, flags=re.M)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def first_markdown_heading(value: str) -> str:
    for line in repair_text(value).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = line.strip()
        if line.startswith("#"):
            return line.lstrip("#").strip()
    return ""


def is_faq_path(path: str) -> bool:
    normalized = normalize_url_path(path).casefold()
    return "/de/help/faq/" in normalized and not normalized.endswith("/faq/")


def display_title_for_node(node: dict | None, fallback: str = "") -> str:
    if not node:
        return fallback
    content = repair_text(node.get("pageTitle", ""))
    path = str(node.get("path") or "")
    if is_faq_path(path):
        question = first_markdown_heading(content)
        if question:
            return question
    return repair_text(node.get("title") or fallback).strip()


def meta_description(text: str, fallback: str) -> str:
    clean = strip_markdown(text) or fallback
    clean = clean.replace(" ,", ",").replace(" .", ".")
    if len(clean) <= 158:
        return clean
    cut = clean[:155].rsplit(" ", 1)[0].rstrip(".,;:")
    return cut + "..."


def normalized_plain(value: str) -> str:
    value = repair_text(value)
    value = re.sub(r"[.。…]+$", "", value.replace("...", ""))
    value = re.sub(r"\s+", " ", value).strip().casefold()
    return value


def visible_summary(description: str, content_source: str) -> str:
    content = normalized_plain(strip_markdown(content_source))
    summary = normalized_plain(description)
    if not summary:
        return ""
    if content.startswith(summary) or summary.startswith(content[: len(summary)]):
        return ""
    return f'<p class="help-view-summary">{html.escape(description)}</p>'


def link_for_text(text: str, link_index: dict[str, str] | None = None) -> str | None:
    if not link_index:
        return None
    key = re.sub(r"\s+", " ", text.strip()).casefold()
    return link_index.get(key)


def inline_markup(text: str, link_index: dict[str, str] | None = None) -> str:
    path = link_for_text(text, link_index)
    if path:
        return f'<a href="{html.escape(path)}">{html.escape(text.strip())}</a>'
    escaped = html.escape(text.strip())
    escaped = re.sub(r"`([^`]+)`", r"<code>\1</code>", escaped)
    return escaped


def markdown_to_html(markdown: str, page_title: str, link_index: dict[str, str] | None = None) -> str:
    markdown = repair_text(markdown).replace("\r\n", "\n").replace("\r", "\n")
    lines = markdown.split("\n")
    chunks: list[str] = []
    paragraph: list[str] = []
    bullet_items: list[str] = []
    numbered_items: list[str] = []
    box_title: str | None = None
    box_lines: list[str] = []
    flow_title: str | None = None
    flow_lines: list[str] = []
    first_heading_seen = False

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            chunks.append(f"<p>{inline_markup(' '.join(paragraph))}</p>")
            paragraph = []

    def flush_bullets() -> None:
        nonlocal bullet_items
        if bullet_items:
            items = "".join(f"<li>{inline_markup(item, link_index)}</li>" for item in bullet_items)
            chunks.append(f"<ul>{items}</ul>")
            bullet_items = []

    def flush_numbered() -> None:
        nonlocal numbered_items
        if numbered_items:
            items = "".join(f"<li>{inline_markup(item, link_index)}</li>" for item in numbered_items)
            chunks.append(f"<ol>{items}</ol>")
            numbered_items = []

    def flush_box() -> None:
        nonlocal box_title, box_lines
        if box_title is not None:
            body = markdown_to_html("\n".join(box_lines).strip(), "__box__", link_index)
            chunks.append(
                f'<section class="static-help-box"><h3>{inline_markup(box_title, link_index)}</h3>{body}</section>'
            )
            box_title = None
            box_lines = []

    def flush_flow() -> None:
        nonlocal flow_title, flow_lines
        if flow_title is not None:
            items: list[str] = []
            for raw_item in flow_lines:
                item = raw_item.strip()
                item = re.sub(r"^(?:[-*]|\d+[.)])\s*", "", item).strip()
                if not item:
                    continue
                if ":" in item:
                    head, body = item.split(":", 1)
                    items.append(
                        f"<li><strong>{inline_markup(head.strip(), link_index)}</strong><span>{inline_markup(body.strip(), link_index)}</span></li>"
                    )
                else:
                    items.append(f"<li><strong>{inline_markup(item, link_index)}</strong></li>")
            body = "".join(items)
            chunks.append(
                f'<section class="static-help-flow"><h3>{inline_markup(flow_title, link_index)}</h3><ol>{body}</ol></section>'
            )
            flow_title = None
            flow_lines = []

    for raw_line in lines:
        line = raw_line.strip()
        if box_title is not None:
            if line == ":::":
                flush_box()
            else:
                box_lines.append(raw_line)
            continue
        if flow_title is not None:
            if line == ":::":
                flush_flow()
            else:
                flow_lines.append(raw_line)
            continue
        if not line:
            flush_paragraph()
            flush_bullets()
            flush_numbered()
            continue
        box = re.match(r"^:::box\s+(.+)$", line)
        if box:
            flush_paragraph()
            flush_bullets()
            flush_numbered()
            box_title = box.group(1).strip()
            box_lines = []
            continue
        flow = re.match(r"^:::flow\s+(.+)$", line)
        if flow:
            flush_paragraph()
            flush_bullets()
            flush_numbered()
            flow_title = flow.group(1).strip()
            flow_lines = []
            continue
        if re.fullmatch(r"-{3,}", line):
            flush_paragraph()
            flush_bullets()
            flush_numbered()
            chunks.append('<hr class="static-help-separator">')
            continue

        image = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)$", line)
        if image:
            flush_paragraph()
            flush_bullets()
            flush_numbered()
            alt_text = image.group(1).strip()
            src = image.group(2).strip()
            chunks.append(
                f'<figure class="static-help-figure"><img src="{html.escape(src)}" alt="{html.escape(alt_text)}"></figure>'
            )
            continue

        bullet = re.match(r"^(?:[•*]|\-\s)(.*)$", line)
        if bullet:
            flush_paragraph()
            flush_numbered()
            bullet_items.append(bullet.group(1).strip())
            continue

        numbered = re.match(r"^\d+[.)]\s+(.+)$", line)
        if numbered:
            flush_paragraph()
            flush_bullets()
            numbered_items.append(numbered.group(1).strip())
            continue

        heading = re.match(r"^(#{1,6})\s+(.+)$", line)
        if heading:
            flush_paragraph()
            flush_bullets()
            flush_numbered()
            heading_text = heading.group(2).strip()
            if not first_heading_seen and heading_text.casefold() == page_title.casefold():
                first_heading_seen = True
                continue
            level = "h2" if len(heading.group(1)) <= 1 else "h3"
            chunks.append(f"<{level}>{inline_markup(heading_text, link_index)}</{level}>")
            first_heading_seen = True
            continue

        flush_bullets()
        flush_numbered()
        paragraph.append(line)

    flush_paragraph()
    flush_bullets()
    flush_numbered()
    flush_box()
    flush_flow()
    return "\n".join(chunks) or "<p>Dieser Abschnitt wird redaktionell vorbereitet.</p>"


def breadcrumb_items(node: dict, title: str, canonical_path: str, title_by_path: dict[str, str]) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = [("X-ERP Hilfe", "/de/help/")]
    chain: list[dict] = []
    current: dict | None = node
    while current:
        chain.append(current)
        current = current.get("_parent")
    for item in reversed(chain):
        item_title = repair_text(item.get("title", "")).strip()
        item_path = item.get("path") or ""
        if item_title and item_path:
            items.append((item_title, normalize_url_path(item_path)))
    path_items = path_breadcrumb_items(canonical_path, title_by_path, title)
    if len(items) <= 2 and len(path_items) > len(items):
        return path_items
    if len(items) == 1:
        items.append((title, canonical_path))
    return items


def render_breadcrumb(items: list[tuple[str, str]]) -> str:
    rendered: list[str] = []
    for index, (title, path) in enumerate(items):
        if index < len(items) - 1:
            rendered.append(f'<a href="{html.escape(path)}">{html.escape(title)}</a><span class="sep">/</span>')
        else:
            rendered.append(f"<span>{html.escape(title)}</span>")
    return "".join(rendered)


def render_child_cards(node: dict | None) -> str:
    children = [
        child
        for child in (node or {}).get("children") or []
        if child.get("path") and child.get("contentType") != "InlineSection"
    ]
    if not children:
        return ""
    cards: list[str] = []
    for child in children[:24]:
        title = display_title_for_node(child, repair_text(child.get("title", "")).strip())
        path = normalize_url_path(str(child.get("path", "")))
        desc = meta_description(repair_text(child.get("pageTitle", "")), f"{title} in der X-ERP ERP Hilfe.")
        cards.append(
            f'''<a class="view-card-link" href="{html.escape(path)}">
              <strong>{html.escape(title)}</strong>
              <span>{html.escape(desc)}</span>
            </a>'''
        )
    return f'''
        <section class="help-view-panel view-page-section">
          <div class="help-view-section-head">
            <h2>Unterseiten</h2>
            <p>Diese Abschnitte gehören fachlich zu diesem Bereich und führen zu den passenden Detailseiten.</p>
          </div>
          <div class="view-card-grid">
            {''.join(cards)}
          </div>
        </section>'''


def absolute_screenshot_path(node: dict | None) -> str:
    if not node:
        return ""
    path = repair_text(node.get("screenshotWebPath") or node.get("screenshot") or "").strip()
    if not path:
        return ""
    path = path.replace("\\", "/")
    if path.startswith("http://") or path.startswith("https://"):
        return path
    if path.startswith("/de/help/"):
        return path
    return "/de/help/" + path.lstrip("/")


def render_screenshot(node: dict | None, title: str) -> str:
    src = absolute_screenshot_path(node)
    if not src:
        return ""
    alt = repair_text((node or {}).get("imageAlt") or f"Screenshot {title} in X-ERP").strip()
    caption = repair_text((node or {}).get("imageCaption") or title).strip()
    return f'''
        <figure class="help-view-shot view-screenshot-figure static-help-screenshot">
          <a class="lightbox" href="{html.escape(src)}"><img src="{html.escape(src)}" alt="{html.escape(alt)}" loading="lazy"></a>
          <figcaption>{html.escape(caption)}</figcaption>
        </figure>'''


def render_json_ld(title: str, description: str, canonical_url: str, breadcrumbs: list[tuple[str, str]]) -> str:
    breadcrumb_schema = {
        "@context": "https://schema.org",
        "@type": "BreadcrumbList",
        "itemListElement": [
            {
                "@type": "ListItem",
                "position": index + 1,
                "name": item_title,
                "item": "https://x-erp.de" + item_path,
            }
            for index, (item_title, item_path) in enumerate(breadcrumbs)
        ],
    }
    article_schema = {
        "@context": "https://schema.org",
        "@type": "TechArticle",
        "headline": title,
        "description": description,
        "inLanguage": "de-DE",
        "about": "ERP",
        "mainEntityOfPage": canonical_url,
        "isPartOf": {"@type": "CreativeWork", "name": "X-ERP Hilfe"},
    }
    return json.dumps([breadcrumb_schema, article_schema], ensure_ascii=False).replace("</", "<\\/")


def render_page(node: dict | None, rel_dir: Path, link_index: dict[str, str], title_by_path: dict[str, str]) -> str:
    fallback_title = repair_text(segment_title(rel_dir.name)).strip() or "X-ERP Hilfe"
    canonical_path = normalize_url_path((node or {}).get("path") or f"/de/help/{quote(rel_dir.as_posix(), safe='/()')}/")
    title = display_title_for_node(node, fallback_title)
    content_source = repair_text((node or {}).get("pageTitle") or f"# {title}\n\nDieser Abschnitt der X-ERP ERP Hilfe wird redaktionell vorbereitet.")
    canonical_url = "https://x-erp.de" + canonical_path
    description = meta_description(content_source, f"{title} in der X-ERP ERP Hilfe.")
    summary_html = visible_summary(description, content_source)
    breadcrumbs = breadcrumb_items(node or {}, title, canonical_path, title_by_path)
    content_html = markdown_to_html(content_source, title, link_index)
    child_cards = render_child_cards(node)
    screenshot_html = render_screenshot(node, title)
    breadcrumb_html = render_breadcrumb(breadcrumbs)
    schema = render_json_ld(f"{title} | X-ERP ERP Hilfe", description, canonical_url, breadcrumbs)

    return f'''<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)} | X-ERP ERP Hilfe</title>
  <meta name="description" content="{html.escape(description)}">
  <meta name="robots" content="index, follow">
  <link rel="canonical" href="{html.escape(canonical_url)}">
  <link rel="stylesheet" href="/de/help/styles.css">
  <link rel="stylesheet" href="/de/help/help.css?v={HELP_CSS_VERSION}">
  <link rel="stylesheet" href="/de/help/Ansichten/ansichten.css?v={CSS_VERSION}">
  <script type="application/ld+json">{schema}</script>
</head>
<body class="help-page">
  <div class="doc-layout">
    <aside class="doc-sidebar" id="sidebar">
      <div class="doc-sidebar-header">
        <h1><img src="/assets/logo/X-ERP.png" alt="X-ERP Logo"> X-ERP Hilfe</h1>
        <div class="doc-search"><input type="text" id="toc-search" placeholder="Suchen... (Strg+K)"></div>
      </div>
      <nav aria-label="Hilfe-Navigation"><ul class="toc-tree" id="toc"></ul></nav>
    </aside>
    <div class="overlay" id="overlay"></div>
    <main class="doc-content">
      <section class="help-page-shell">
        <nav class="view-breadcrumb" aria-label="Breadcrumb">{breadcrumb_html}</nav>
        <section class="help-view-hero">
          <div class="help-view-eyebrow">ERP Hilfe</div>
          <div class="help-view-hero-grid">
            <div>
              <h1>{html.escape(title)}</h1>
              {summary_html}
              <div class="help-view-meta"><span class="view-chip"><strong>Bereich:</strong> Hilfe</span><span class="view-chip"><strong>Keyword:</strong> ERP</span></div>
            </div>
            <div class="help-view-hero-card">
              <h2>Einordnung</h2>
              <p>{html.escape(" > ".join(title for title, _ in breadcrumbs))}</p>
            </div>
          </div>
        </section>

        {screenshot_html}
        <section class="help-view-panel view-page-section static-help-content">
          {content_html}
        </section>
        {child_cards}
      </section>
    </main>
  </div>
  <script src="/de/help/xlsx-tree.js?v={SCRIPT_VERSION}"></script>
  <script src="/de/help/ansichten-tree.js?v={SCRIPT_VERSION}"></script>
  <script src="/de/help/index-tree.js?v={SCRIPT_VERSION}"></script>
  <script src="/de/help/app.js?v={SCRIPT_VERSION}"></script>
</body>
</html>
'''


def render_home_page(vorwort_node: dict, link_index: dict[str, str], title_by_path: dict[str, str]) -> str:
    home_node = dict(vorwort_node)
    home_node["path"] = "/de/help/"
    home_node["breadcrumb"] = "Vorwort"
    return render_page(home_node, Path("."), link_index, title_by_path)


def render_redirect_page(target_url: str) -> str:
    escaped_target = html.escape(target_url, quote=True)
    json_target = json.dumps(target_url, ensure_ascii=False)
    return f'''<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="robots" content="noindex, follow">
  <meta http-equiv="refresh" content="0; url={escaped_target}">
  <link rel="canonical" href="{escaped_target}">
  <title>Weiterleitung zu Multi-Monitor | X-ERP ERP Hilfe</title>
  <script>location.replace({json_target});</script>
</head>
<body>
  <p><a href="{escaped_target}">Weiter zu Multi-Monitor</a></p>
</body>
</html>
'''


def main() -> None:
    nodes = flatten(json.loads(TREE_JSON.read_text(encoding="utf-8-sig")))
    by_path: dict[str, dict] = {}
    link_index: dict[str, str] = {}
    title_by_path: dict[str, str] = {}
    for node in nodes:
        path = node.get("path")
        if path:
            normalized = normalize_url_path(str(path))
            by_path[normalized] = node
            node_title = display_title_for_node(node, repair_text(node.get("title", "")).strip())
            if node_title:
                title_by_path[normalized] = node_title
            for title_candidate in {repair_text(node.get("title", "")).strip(), node_title}:
                title_key = re.sub(r"\s+", " ", title_candidate).casefold()
                if title_key and (
                    "/ansichten/" not in normalized.casefold()
                    or normalized.casefold() == "/de/help/ansichten/"
                ):
                    current = link_index.get(title_key)
                    if current is None or ("/Module/" in normalized and "/Module/" not in current):
                        link_index[title_key] = normalized

    rewritten = 0
    written_html = 0
    for index_file in HELP_ROOT.rglob("index.html"):
        if "Ansichten" in index_file.relative_to(HELP_ROOT).parts:
            continue
        rel_dir = index_file.parent.relative_to(HELP_ROOT)
        rel_url = normalize_url_path("/de/help/" + rel_dir.as_posix())
        node = by_path.get(rel_url)
        if node is None:
            continue
        index_file.write_text(render_page(node, rel_dir, link_index, title_by_path), encoding="utf-8", newline="\n")
        rewritten += 1

    for normalized, node in by_path.items():
        if not normalized.startswith("/de/help/"):
            continue
        if "/ansichten/" in normalized.casefold() and (
            node.get("contentType") != "HelpPage" or int(node.get("sourceLevel") or 0) < 4
        ):
            continue
        rel_file = normalized.removeprefix("/de/help/").lstrip("/")
        if normalized.endswith("/"):
            rel_file = rel_file + "index.html"
        elif not normalized.lower().endswith(".html"):
            continue
        html_file = HELP_ROOT / rel_file
        html_file.parent.mkdir(parents=True, exist_ok=True)
        html_file.write_text(render_page(node, html_file.parent.relative_to(HELP_ROOT), link_index, title_by_path), encoding="utf-8", newline="\n")
        written_html += 1

    vorwort_node = by_path.get("/de/help/Vorwort/")
    if vorwort_node:
        (HELP_ROOT / "index.html").write_text(
            render_home_page(vorwort_node, link_index, title_by_path),
            encoding="utf-8",
            newline="\n",
        )
        written_html += 1

    written_redirects = 0
    for source, target in REDIRECTS.items():
        redirect_file = HELP_ROOT / source.strip("/").replace("\\", "/") / "index.html"
        redirect_file.parent.mkdir(parents=True, exist_ok=True)
        redirect_file.write_text(render_redirect_page(target), encoding="utf-8", newline="\n")
        written_redirects += 1

    print(f"rewritten_static_pages={rewritten}")
    print(f"written_html_pages={written_html}")
    print(f"written_redirects={written_redirects}")


if __name__ == "__main__":
    main()
