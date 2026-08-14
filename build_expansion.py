from __future__ import annotations

import argparse
import importlib.util
import json
import re
import shutil
import sys
import types
from html import escape
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "data" / "expansion_pages.json"
SCHOOL_DATA_FILE = ROOT / "data" / "school_expansion_pages.json"
BASE_OUTPUT = ROOT / "output"
BASE_URL = "https://www.eduguide.kr"
EXPECTED_BASE_PAGES = 1620
EXPECTED_BASE_URLS = 1621
BODY_SHEETS = [
    "고1과외", "고2과외", "고3과외",
    "고1영어과외", "고2영어과외", "고3영어과외",
    "고1수학과외", "고2수학과외", "고3수학과외",
]


def load_generator():
    """Import rendering helpers without requiring workbook dependencies at build time."""
    fake_openpyxl = types.ModuleType("openpyxl")
    fake_openpyxl.load_workbook = lambda *args, **kwargs: (_ for _ in ()).throw(
        RuntimeError("expansion build does not read workbooks")
    )
    sys.modules.setdefault("openpyxl", fake_openpyxl)
    spec = importlib.util.spec_from_file_location("eduguide_generator", ROOT / "generator.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def related_sections(region: str, meta: dict[str, str], all_slugs: set[str], new_slugs: set[str]):
    sections = []
    parent_links = []
    for area in (meta.get("parent", ""), meta.get("district", ""), meta.get("city", "")):
        slug = f"{area}과외".strip()
        link = (slug, f"/{slug}/")
        if area and slug in all_slugs and link not in parent_links:
            parent_links.append(link)
    if parent_links:
        sections.append({"title": "상위 지역", "links": parent_links[:8]})

    existing = []
    for suffix in ("과외", "영어과외", "수학과외", "고등과외", "고등영어과외", "고등수학과외"):
        slug = f"{region}{suffix}"
        if slug in all_slugs:
            existing.append((slug, f"/{slug}/"))
    if existing:
        sections.append({"title": "같은 지역 기존 콘텐츠", "links": existing[:8]})

    siblings = []
    for sheet in BODY_SHEETS:
        slug = f"{region}{sheet}"
        if slug in new_slugs:
            siblings.append((slug, f"/{slug}/"))
    if siblings:
        sections.append({"title": "같은 지역 학년·과목", "links": siblings[:8]})
    return sections


def load_inputs():
    base_payload = json.loads((ROOT / "data" / "pages.json").read_text(encoding="utf-8"))
    base_pages = base_payload["pages"]
    if len(base_pages) != EXPECTED_BASE_PAGES:
        raise SystemExit(f"STOP: base page count {len(base_pages)} != {EXPECTED_BASE_PAGES}")
    expansion = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    pages = expansion.get("pages", [])
    school = json.loads(SCHOOL_DATA_FILE.read_text(encoding="utf-8"))
    school_pages = school.get("pages", [])
    if not pages or not school_pages:
        raise SystemExit("STOP: expansion inputs are empty")
    return base_pages, expansion, pages, school, school_pages


def validate_plan(base_pages, expansion_pages):
    base_slugs = {page["slug"].strip("/") for page in base_pages}
    new_slugs = [page["slug"].strip("/") for page in expansion_pages]
    if len(set(new_slugs)) != len(new_slugs):
        raise SystemExit("STOP: duplicate expansion slug")
    collisions = sorted(base_slugs & set(new_slugs))
    if collisions:
        raise SystemExit(f"STOP: expansion collides with base slugs: {collisions[:5]}")
    for page in expansion_pages:
        expected = f"{page['region']}{page['page_type']}"
        if page["slug"] != expected or page["keyword"] != expected:
            raise SystemExit(f"STOP: invalid slug mapping at {page['source_sheet']}:{page['row']}")
        if not page["title"].strip() or not page["body"].strip():
            raise SystemExit(f"STOP: missing title/body at {page['source_sheet']}:{page['row']}")
    return base_slugs, set(new_slugs)


def validate_school_plan(base_slugs, regional_slugs, school_payload, school_pages):
    marker = "(학교)"
    school_slugs = [page["slug"].strip("/") for page in school_pages]
    if len(set(school_slugs)) != len(school_slugs):
        raise SystemExit("STOP: duplicate school slug")
    collisions = sorted((base_slugs | regional_slugs) & set(school_slugs))
    if collisions:
        raise SystemExit(f"STOP: school collides with existing slugs: {collisions[:5]}")
    integrated_slugs = base_slugs | regional_slugs
    for page in school_pages:
        expected = re.sub(r"\s+", "", f"{page['school_short_name']}{page['page_type']}")
        values = [page.get(key, "") for key in (
            "school_name", "school_short_name", "page_type", "title", "body_html",
            "slug", "canonical", "parent_slug", "connected_region_name",
            "connected_region_url", "connection_level",
        )]
        if marker in "\n".join(map(str, values)) or page["slug"] != expected:
            raise SystemExit(f"STOP: invalid school page at {page['source_sheet']}:{page['row']}")
        if not page["title"].strip() or not page["body_html"].strip():
            raise SystemExit(f"STOP: missing school title/body at {page['source_sheet']}:{page['row']}")
        target_slug = page["connected_region_url"].replace(BASE_URL, "").strip("/")
        if target_slug not in integrated_slugs:
            raise SystemExit(f"STOP: missing school region target: {target_slug}")
    mappings = school_payload.get("schools", [])
    if len(mappings) != len({item["school_short_name"] for item in mappings}):
        raise SystemExit("STOP: duplicate school mapping")
    return set(school_slugs)


def school_related_sections(source, school_slugs):
    links = []
    school_name = source["school_short_name"]

    def school_label(slug):
        page_type = slug[len(school_name):] if slug.startswith(school_name) else slug
        return f"{school_name} {page_type}"

    region_slug = source["connected_region_url"].replace(BASE_URL, "").strip("/")
    links.append({"title": "연결 지역", "links": [(source["connected_region_name"], f"/{region_slug}/")]})
    family = []
    parent = source.get("parent_slug", "").strip("/")
    if parent and parent in school_slugs:
        family.append((school_label(parent), f"/{parent}/"))
    for child in source.get("child_slugs", []):
        child = child.strip("/")
        if child in school_slugs:
            family.append((school_label(child), f"/{child}/"))
    if family:
        links.append({"title": "학교 학년·과목", "links": family})
    return links


def add_school_links(html, links):
    start = "<!-- school-links:start -->"
    end = "<!-- school-links:end -->"
    block = start + '<section class="related-section school-links"><h3>관련 학교</h3><ul>'
    block += "".join(f'<li><a href="{escape(href)}">{escape(label)}</a></li>' for label, href in links)
    block += "</ul></section>" + end
    if start in html and end in html:
        return html[:html.index(start)] + block + html[html.index(end) + len(end):]
    anchor = '<nav class="related-links"><h2>관련 페이지</h2>'
    if anchor not in html:
        raise SystemExit("STOP: related-links anchor not found")
    return html.replace(anchor, anchor + block, 1)


def prepare_output(output_root: Path):
    output_root = output_root.resolve()
    if output_root != BASE_OUTPUT.resolve():
        if output_root.exists() and any(output_root.iterdir()):
            required = output_root / "sitemap.xml"
            if not required.exists():
                raise SystemExit(f"STOP: non-empty output root is not an EduGuide baseline: {output_root}")
        else:
            output_root.mkdir(parents=True, exist_ok=True)
            shutil.copytree(BASE_OUTPUT, output_root, dirs_exist_ok=True)
    return output_root


def write_if_changed(path: Path, content: str) -> bool:
    # The validated Windows candidate was produced by Path.write_text(), which
    # translated newlines to CRLF. Pin that representation so Vercel/Linux and
    # local builds reproduce the exact same bytes.
    normalized = content.replace("\r\n", "\n").replace("\n", "\r\n")
    encoded = normalized.encode("utf-8")
    if path.is_file() and path.read_bytes() == encoded:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(encoded)
    return True


def build(output_root: Path):
    generator = load_generator()
    base_pages, expansion, source_pages, school_payload, school_sources = load_inputs()
    base_slugs, new_slugs = validate_plan(base_pages, source_pages)
    school_slugs = validate_school_plan(base_slugs, new_slugs, school_payload, school_sources)
    all_slugs = base_slugs | new_slugs | school_slugs

    region_meta = expansion["regions"]
    regions = [
        generator.RegionRow(meta["city"], meta["district"], region)
        for region, meta in region_meta.items()
    ]
    regional_pages = []
    for source in source_pages:
        region = source["region"]
        page = {
            "type": "content",
            "sheet": source["source_sheet"],
            "row": source["row"],
            "slug": source["slug"],
            "keyword": source["keyword"],
            "h1": source["keyword"],
            "title": source["title"],
            "description": generator.quality_description(source["keyword"]),
            "content": source["body"],
        }
        page["related_sections"] = related_sections(region, region_meta[region], all_slugs, new_slugs)
        page["related_links"] = [link for section in page["related_sections"] for link in section["links"]]
        regional_pages.append(page)

    school_pages = []
    for source in school_sources:
        page = {
            "type": "school",
            "sheet": source["source_sheet"],
            "row": source["row"],
            "slug": source["slug"],
            "keyword": source["slug"],
            "h1": f"{source['school_short_name']} {source['page_type']}",
            "title": source["title"],
            "description": generator.quality_description(source["slug"]),
            "content": source["body_html"],
            "related_sections": school_related_sections(source, school_slugs),
        }
        page["related_links"] = [link for section in page["related_sections"] for link in section["links"]]
        school_pages.append(page)

    pages = regional_pages + school_pages

    output_root = prepare_output(output_root)
    changed_files = 0
    for page in pages:
        path = output_root / page["slug"] / "index.html"
        if write_if_changed(path, generator.render_page(page, base_pages + pages, regions)):
            changed_files += 1

    schools_by_region = {}
    for mapping in school_payload["schools"]:
        target_slug = mapping["connected_region_url"].replace(BASE_URL, "").strip("/")
        hub_slug = mapping["hub_slug"]
        schools_by_region.setdefault(target_slug, []).append((f"{mapping['school_short_name']} 과외", f"/{hub_slug}/"))
    for target_slug, links in schools_by_region.items():
        path = output_root / target_slug / "index.html"
        if not path.is_file():
            raise SystemExit(f"STOP: region HTML missing for school links: {target_slug}")
        updated = add_school_links(path.read_text(encoding="utf-8"), sorted(links))
        if write_if_changed(path, updated):
            changed_files += 1

    base_urls = [BASE_URL + "/"] + [f"{BASE_URL}/{page['slug'].strip('/')}/" for page in base_pages]
    regional_urls = [f"{BASE_URL}/{page['slug'].strip('/')}/" for page in regional_pages]
    school_urls = [f"{BASE_URL}/{page['slug'].strip('/')}/" for page in school_pages]
    all_urls = base_urls + regional_urls + school_urls
    if len(base_urls) != EXPECTED_BASE_URLS or len(set(all_urls)) != len(all_urls):
        raise SystemExit("STOP: sitemap count or uniqueness check failed")
    sitemap = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
    sitemap += "<urlset xmlns=\"http://www.sitemaps.org/schemas/sitemap/0.9\">\n"
    sitemap += "\n".join(f"  <url><loc>{escape(url)}</loc></url>" for url in all_urls)
    sitemap += "\n</urlset>\n"
    if write_if_changed(output_root / "sitemap.xml", sitemap):
        changed_files += 1
    print(json.dumps({
        "base_pages": len(base_pages), "regional_pages": len(regional_pages),
        "school_pages": len(school_pages),
        "sitemap_urls": len(all_urls), "changed_files": changed_files,
        "output_root": str(output_root),
    }, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-root", type=Path, default=BASE_OUTPUT)
    args = parser.parse_args()
    build(args.output_root)


if __name__ == "__main__":
    main()
