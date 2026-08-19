from __future__ import annotations

import argparse
import hashlib
import html as html_module
import json
import re
import struct
from pathlib import Path
from urllib.parse import unquote, urlparse
from xml.etree import ElementTree


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output"
BASELINE_DEFAULT = ROOT / "validation" / "pre_thumbnail_baseline.json"
PRODUCTION_URLS = ROOT / "validation" / "production_urls_before.txt"
PUBLIC_PREFIX = "https://www.eduguide.kr/assets/images/search-thumbnails/"
BODY_RE = re.compile(r"<body\b[^>]*>(.*?)</body>", re.I | re.S)
HEAD_RE = re.compile(r"<head\b[^>]*>(.*?)</head>", re.I | re.S)
TAG_RE = re.compile(r"<[^>]+>", re.S)
IMG_RE = re.compile(r"<img\b[^>]*\bsrc=[\"']([^\"']+)[\"'][^>]*>", re.I)
META_RE = re.compile(r"<meta\b[^>]*>", re.I)
ATTR_RE = re.compile(r"([:\w-]+)\s*=\s*([\"'])(.*?)\2", re.I | re.S)
CANONICAL_RE = re.compile(r"<link\b(?=[^>]*\brel=[\"']canonical[\"'])[^>]*\bhref=[\"']([^\"']+)[\"'][^>]*>", re.I)
TITLE_RE = re.compile(r"<title>(.*?)</title>", re.I | re.S)
JSON_LD_RE = re.compile(r"<script\b[^>]*type=[\"']application/ld\+json[\"'][^>]*>.*?</script>", re.I | re.S)


def digest(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def meta_values(html: str, key: str, value: str) -> list[str]:
    found = []
    for tag in META_RE.findall(html):
        attrs = {name.lower(): attr_value for name, _, attr_value in ATTR_RE.findall(tag)}
        if attrs.get(key) == value:
            found.append(attrs.get("content", ""))
    return found


def sitemap_urls() -> list[str]:
    tree = ElementTree.parse(OUTPUT / "sitemap.xml")
    return [node.text.strip() for node in tree.getroot().iter() if node.tag.endswith("loc") and node.text]


def html_path(canonical: str) -> Path:
    path = unquote(urlparse(canonical).path).strip("/")
    return OUTPUT / "index.html" if not path else OUTPUT / path / "index.html"


def record(canonical: str) -> dict[str, object]:
    path = html_path(canonical)
    html = path.read_text(encoding="utf-8")
    head = HEAD_RE.search(html)
    body = BODY_RE.search(html)
    if not head or not body:
        raise SystemExit(f"STOP: malformed HTML: {path}")
    body_html = body.group(1)
    body_text = " ".join(html_module.unescape(TAG_RE.sub(" ", body_html)).split())
    canonical_match = CANONICAL_RE.search(head.group(1))
    title_match = TITLE_RE.search(head.group(1))
    return {
        "canonical": canonical_match.group(1) if canonical_match else "",
        "title": title_match.group(1) if title_match else "",
        "description": (meta_values(head.group(1), "name", "description") or [""])[0],
        "robots": (meta_values(head.group(1), "name", "robots") or [""])[0],
        "body_html_sha256": digest(body_html),
        "body_text_sha256": digest(body_text),
        "body_img_src": IMG_RE.findall(body_html),
        "json_ld_sha256": digest("".join(JSON_LD_RE.findall(head.group(1)))),
    }


def snapshot(path: Path) -> None:
    urls = sitemap_urls()
    payload = {"url_count": len(urls), "pages": {url: record(url) for url in urls}}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="")
    print(json.dumps({"snapshot": str(path), "pages": len(urls)}, ensure_ascii=False))


def png_ok(path: Path) -> bool:
    try:
        header = path.read_bytes()[:24]
        return (
            path.stat().st_size > 0
            and len(header) == 24
            and header[:8] == b"\x89PNG\r\n\x1a\n"
            and header[12:16] == b"IHDR"
            and struct.unpack(">II", header[16:24]) == (1254, 1254)
        )
    except OSError:
        return False


def audit(baseline_path: Path) -> None:
    urls = sitemap_urls()
    manifest_path = OUTPUT / "search_thumbnail_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest_map = {row["canonical"]: row for row in manifest["pages"]}
    baseline = json.loads(baseline_path.read_text(encoding="utf-8")) if baseline_path.exists() else None

    missing = duplicate_og = duplicate_twitter = secure_bad = type_bad = 0
    body_exposure = body_changed = body_img_changed = canonical_changed = title_changed = description_changed = robots_changed = json_ld_changed = 0
    for canonical in urls:
        path = html_path(canonical)
        html = path.read_text(encoding="utf-8")
        head = HEAD_RE.search(html).group(1)
        body = BODY_RE.search(html).group(1)
        expected = manifest_map.get(canonical, {}).get("new_og_image", "")
        og = meta_values(head, "property", "og:image")
        secure = meta_values(head, "property", "og:image:secure_url")
        image_type = meta_values(head, "property", "og:image:type")
        twitter = meta_values(head, "name", "twitter:image")
        card = meta_values(head, "name", "twitter:card")
        if og != [expected] or twitter != [expected] or card != ["summary_large_image"]:
            missing += 1
        duplicate_og += int(len(og) != 1)
        duplicate_twitter += int(len(twitter) != 1)
        secure_bad += int(secure != [expected])
        type_bad += int(image_type != ["image/png"])
        body_exposure += int("/assets/images/search-thumbnails/" in body)
        if baseline:
            before = baseline["pages"][canonical]
            after = record(canonical)
            body_changed += int(before["body_html_sha256"] != after["body_html_sha256"] or before["body_text_sha256"] != after["body_text_sha256"])
            body_img_changed += int(before["body_img_src"] != after["body_img_src"])
            canonical_changed += int(before["canonical"] != after["canonical"])
            title_changed += int(before["title"] != after["title"])
            description_changed += int(before["description"] != after["description"])
            robots_changed += int(before["robots"] != after["robots"])
            json_ld_changed += int(before["json_ld_sha256"] != after["json_ld_sha256"])

    images = sorted((OUTPUT / "assets" / "images" / "search-thumbnails").glob("*.png"))
    broken = sum(not png_ok(path) for path in images)
    candidate = {url.rstrip("/") for url in urls}
    production = candidate
    if PRODUCTION_URLS.exists():
        production = {
            line.strip().rstrip("/")
            for line in PRODUCTION_URLS.read_text(encoding="utf-8").splitlines()
            if line.strip()
        }
    result = {
        "html": len(list(OUTPUT.rglob("*.html"))),
        "indexable": len(urls),
        "sitemap": len(urls),
        "thumbnail_applied": len(urls) - missing,
        "missing": missing,
        "duplicate_og": duplicate_og,
        "duplicate_twitter": duplicate_twitter,
        "secure_bad": secure_bad,
        "type_bad": type_bad,
        "body_thumbnail_exposure": body_exposure,
        "body_changed": body_changed,
        "body_img_changed": body_img_changed,
        "canonical_changed": canonical_changed,
        "title_changed": title_changed,
        "description_changed": description_changed,
        "robots_changed": robots_changed,
        "json_ld_changed": json_ld_changed,
        "production_only": len(production - candidate),
        "candidate_only": len(candidate - production),
        "thumbnail_png": len(images),
        "broken_png": broken,
    }
    failures = [key for key, value in result.items() if key not in {"html", "indexable", "sitemap", "thumbnail_applied", "thumbnail_png"} and value]
    if result["html"] != 4864 or result["indexable"] != 4863 or result["thumbnail_applied"] != 4863 or result["thumbnail_png"] != 13:
        failures.append("count_mismatch")
    print(json.dumps(result, ensure_ascii=False))
    if failures:
        raise SystemExit("STOP: thumbnail audit failed: " + ", ".join(sorted(set(failures))))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--snapshot", type=Path)
    parser.add_argument("--baseline", type=Path, default=BASELINE_DEFAULT)
    args = parser.parse_args()
    if args.snapshot:
        snapshot(args.snapshot)
    else:
        audit(args.baseline)


if __name__ == "__main__":
    main()
