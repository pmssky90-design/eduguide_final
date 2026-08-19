from __future__ import annotations

import hashlib
import json
import re
import shutil
import struct
from collections import Counter
from pathlib import Path
from urllib.parse import unquote, urlparse
from xml.etree import ElementTree


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output"
SOURCE = ROOT / "assets" / "images" / "search-thumbnails-source"
OUTPUT_IMAGES = OUTPUT / "assets" / "images" / "search-thumbnails"
MANIFEST = OUTPUT / "search_thumbnail_manifest.json"
BASE_URL = "https://www.eduguide.kr"
PUBLIC_PREFIX = f"{BASE_URL}/assets/images/search-thumbnails/"

HEAD_RE = re.compile(r"(<head\b[^>]*>)(.*?)(</head>)", re.I | re.S)
BODY_RE = re.compile(r"<body\b[^>]*>.*?</body>", re.I | re.S)
META_RE = re.compile(r"<meta\b[^>]*>", re.I)
ATTR_RE = re.compile(r"([:\w-]+)\s*=\s*([\"'])(.*?)\2", re.I | re.S)
CANONICAL_RE = re.compile(
    r"<link\b(?=[^>]*\brel=[\"']canonical[\"'])[^>]*\bhref=[\"']([^\"']+)[\"'][^>]*>",
    re.I,
)


def meta_value(html: str, key: str, value: str) -> str:
    for tag in META_RE.findall(html):
        attrs = {name.lower(): attr_value for name, _, attr_value in ATTR_RE.findall(tag)}
        if attrs.get(key) == value:
            return attrs.get("content", "")
    return ""


def remove_search_meta(head: str) -> str:
    targets = {
        ("property", "og:image"),
        ("property", "og:image:secure_url"),
        ("property", "og:image:type"),
        ("name", "twitter:card"),
        ("name", "twitter:image"),
    }

    def keep_or_remove(match: re.Match[str]) -> str:
        attrs = {name.lower(): value for name, _, value in ATTR_RE.findall(match.group(0))}
        return "" if any(attrs.get(key) == value for key, value in targets) else match.group(0)

    return META_RE.sub(keep_or_remove, head)


def sitemap_urls() -> list[str]:
    tree = ElementTree.parse(OUTPUT / "sitemap.xml")
    urls = [node.text.strip() for node in tree.getroot().iter() if node.tag.endswith("loc") and node.text]
    if len(urls) != 4863 or len(set(urls)) != len(urls):
        raise SystemExit(f"STOP: expected 4863 unique sitemap URLs, found {len(urls)}")
    return urls


def html_path(canonical: str) -> Path:
    path = unquote(urlparse(canonical).path).strip("/")
    return OUTPUT / "index.html" if not path else OUTPUT / path / "index.html"


def png_dimensions(path: Path) -> tuple[int, int]:
    with path.open("rb") as handle:
        header = handle.read(24)
    if len(header) != 24 or header[:8] != b"\x89PNG\r\n\x1a\n" or header[12:16] != b"IHDR":
        raise SystemExit(f"STOP: invalid PNG: {path}")
    return struct.unpack(">II", header[16:24])


def source_images() -> list[Path]:
    images = sorted(SOURCE.glob("*.png"), key=lambda path: path.name)
    if len(images) != 13:
        raise SystemExit(f"STOP: expected 13 source PNGs, found {len(images)}")
    for image in images:
        if image.stat().st_size <= 0 or png_dimensions(image) != (1254, 1254):
            raise SystemExit(f"STOP: invalid source image: {image}")
    return images


def thumbnail_for(canonical: str, images: list[Path]) -> Path:
    index = int.from_bytes(hashlib.sha256(canonical.encode("utf-8")).digest()[:8], "big") % len(images)
    return images[index]


def main() -> None:
    images = source_images()
    OUTPUT_IMAGES.mkdir(parents=True, exist_ok=True)
    for image in images:
        shutil.copy2(image, OUTPUT_IMAGES / image.name)

    previous = {}
    if MANIFEST.exists():
        previous = {row["canonical"]: row for row in json.loads(MANIFEST.read_text(encoding="utf-8"))["pages"]}

    rows = []
    changed = 0
    distribution: Counter[str] = Counter()
    for canonical in sitemap_urls():
        path = html_path(canonical)
        if not path.is_file():
            raise SystemExit(f"STOP: sitemap HTML missing: {canonical}")
        html = path.read_text(encoding="utf-8")
        canonical_match = CANONICAL_RE.search(html)
        if not canonical_match or canonical_match.group(1) != canonical:
            raise SystemExit(f"STOP: canonical mismatch: {canonical}")
        if meta_value(html, "name", "robots").lower().startswith("noindex"):
            raise SystemExit(f"STOP: sitemap contains noindex page: {canonical}")
        body_before = BODY_RE.search(html)
        if not body_before:
            raise SystemExit(f"STOP: body missing: {canonical}")

        old_og = meta_value(html, "property", "og:image")
        old_twitter = meta_value(html, "name", "twitter:image")
        if canonical in previous:
            old_og = previous[canonical]["old_og_image"]
            old_twitter = previous[canonical]["old_twitter_image"]

        image = thumbnail_for(canonical, images)
        image_url = PUBLIC_PREFIX + image.name
        block = (
            f'<meta property="og:image" content="{image_url}">'
            f'<meta property="og:image:secure_url" content="{image_url}">'
            '<meta property="og:image:type" content="image/png">'
            '<meta name="twitter:card" content="summary_large_image">'
            f'<meta name="twitter:image" content="{image_url}">'
        )
        head_match = HEAD_RE.search(html)
        if not head_match:
            raise SystemExit(f"STOP: head missing: {canonical}")
        clean_head = remove_search_meta(head_match.group(2))
        og_url = re.search(r"<meta\b(?=[^>]*\bproperty=[\"']og:url[\"'])[^>]*>", clean_head, re.I)
        if og_url:
            clean_head = clean_head[: og_url.end()] + block + clean_head[og_url.end() :]
        else:
            clean_head = block + clean_head
        updated = html[: head_match.start()] + head_match.group(1) + clean_head + head_match.group(3) + html[head_match.end() :]
        body_after = BODY_RE.search(updated)
        if not body_after or body_before.group(0) != body_after.group(0):
            raise SystemExit(f"STOP: body changed while applying metadata: {canonical}")
        if updated != html:
            path.write_text(updated, encoding="utf-8", newline="")
            changed += 1
        distribution[image.name] += 1
        rows.append({
            "canonical": canonical,
            "thumbnail": f"/assets/images/search-thumbnails/{image.name}",
            "old_og_image": old_og,
            "new_og_image": image_url,
            "old_twitter_image": old_twitter,
            "new_twitter_image": image_url,
        })

    payload = {
        "schema_version": 1,
        "algorithm": "sha256(canonical UTF-8) first 64 bits modulo 13",
        "source_count": len(images),
        "page_count": len(rows),
        "distribution": dict(sorted(distribution.items())),
        "pages": rows,
    }
    serialized = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    manifest_changed = not MANIFEST.exists() or MANIFEST.read_text(encoding="utf-8") != serialized
    if manifest_changed:
        MANIFEST.write_text(serialized, encoding="utf-8", newline="")
    print(json.dumps({
        "pages": len(rows),
        "changed_html": changed,
        "manifest_changed": manifest_changed,
        "minimum_assignment": min(distribution.values()),
        "maximum_assignment": max(distribution.values()),
        "distribution": dict(sorted(distribution.items())),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
