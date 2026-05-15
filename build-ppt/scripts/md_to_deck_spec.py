#!/usr/bin/env python
"""Convert an Obsidian-style Markdown deck brief into deck_spec.json."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


IMAGE_RE = re.compile(r"!\[\[([^\]]+)\]\]")
HEADING_RE = re.compile(r"^##\s*(?:\d+\s*)?(.+?)\s*$")
META_RE = re.compile(r"^([A-Za-z_][\w-]*)\s*:\s*(.*)$")
KNOWN_KEYS = {"layout_type", "layout", "kicker", "section", "notes", "image", "images"}


def normalize_image_ref(raw: str) -> str:
    value = raw.strip()
    if "|" in value:
        value = value.split("|", 1)[0]
    return value.strip()


def extract_images(text: str) -> list[dict[str, str]]:
    images: list[dict[str, str]] = []
    for match in IMAGE_RE.finditer(text):
        path = normalize_image_ref(match.group(1))
        if path:
            images.append({"path": path, "fit": "contain"})
    return images


def resolve_image_path(path: str, base_dir: Path) -> str:
    candidate = Path(path)
    if candidate.is_absolute():
        return str(candidate)
    direct = base_dir / candidate
    if direct.exists():
        return str(direct.resolve())
    in_images = base_dir / "images" / candidate.name
    if in_images.exists():
        return str(in_images.resolve())
    fuzzy = fuzzy_image_match(candidate.name, [base_dir / "images", base_dir])
    if fuzzy:
        return str(fuzzy.resolve())
    return path


def fuzzy_image_match(name: str, roots: list[Path]) -> Path | None:
    requested = Path(name)
    suffix = requested.suffix.lower()
    stem = requested.stem.lower().strip()
    if not stem or len(stem) < 4:
        return None
    matches: list[Path] = []
    for root in roots:
        if not root.exists() or not root.is_dir():
            continue
        for item in root.iterdir():
            if not item.is_file():
                continue
            if suffix and item.suffix.lower() != suffix:
                continue
            item_stem = item.stem.lower().strip()
            if item_stem.startswith(stem) or stem.startswith(item_stem):
                matches.append(item)
    if not matches:
        return None
    return sorted(matches, key=lambda item: (len(item.name), item.name))[0]


def resolve_slide_images(slide: dict[str, Any], base_dir: Path) -> None:
    images = slide.get("images")
    if not isinstance(images, list):
        return
    for image in images:
        if isinstance(image, dict) and image.get("path"):
            image["path"] = resolve_image_path(str(image["path"]), base_dir)


def strip_images(text: str) -> str:
    return IMAGE_RE.sub("", text).strip()


def parse_frontmatter(lines: list[str]) -> tuple[dict[str, Any], int]:
    meta: dict[str, Any] = {}
    idx = 0
    while idx < len(lines):
        line = lines[idx].strip()
        if line == "---" or line.startswith("## "):
            break
        match = META_RE.match(line)
        if match:
            meta[match.group(1)] = match.group(2).strip()
        idx += 1
    return meta, idx


def split_slide_blocks(text: str) -> list[str]:
    blocks = re.split(r"(?m)^\s*---\s*$", text)
    return [block.strip() for block in blocks if block.strip() and re.search(r"(?m)^##\s+", block)]


def parse_bullet_items(lines: list[str]) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    current: dict[str, str] | None = None
    for raw in lines:
        line = strip_images(raw).rstrip()
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("- "):
            content = stripped[2:].strip()
            if ":" in content and content.split(":", 1)[0].strip() in {"icon", "title", "body", "label", "value"}:
                if current:
                    items.append(current)
                key, value = content.split(":", 1)
                current = {key.strip(): value.strip()}
            else:
                if current:
                    items.append(current)
                if "|" in content:
                    title, body = content.split("|", 1)
                    current = {"title": title.strip(), "body": body.strip()}
                else:
                    current = {"title": content}
            continue
        if current and ":" in stripped:
            key, value = stripped.split(":", 1)
            key = key.strip()
            if key in {"icon", "title", "body", "label", "value"}:
                current[key] = value.strip()
    if current:
        items.append(current)
    return items


def parse_duo_items(lines: list[str]) -> list[dict[str, Any]]:
    groups: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for raw in lines:
        stripped = strip_images(raw).strip()
        if not stripped:
            continue
        if stripped.startswith(("左侧：", "左侧:", "右侧：", "右侧:")):
            label = re.split(r"[:：]", stripped, maxsplit=1)[1].strip()
            current = {"title": label, "items": []}
            groups.append(current)
            continue
        if current and stripped.startswith("- "):
            current["items"].append({"title": stripped[2:].strip()})
    return groups[:2]


def parse_slide(block: str) -> dict[str, Any]:
    lines = block.splitlines()
    heading = ""
    body_lines: list[str] = []
    slide: dict[str, Any] = {}
    for line in lines:
        heading_match = HEADING_RE.match(line.strip())
        if heading_match:
            heading = heading_match.group(1).strip()
            continue
        meta_match = META_RE.match(line.strip())
        if meta_match and meta_match.group(1) in KNOWN_KEYS:
            key, value = meta_match.group(1), meta_match.group(2).strip()
            if key in {"image", "images"}:
                slide.setdefault("images", []).extend(extract_images(value))
            else:
                slide[key] = strip_images(value)
            continue
        body_lines.append(line)

    slide.setdefault("layout_type", "split_statement")
    clean_lines = [strip_images(line) for line in body_lines]
    clean_lines = [line.strip() for line in clean_lines if line.strip()]
    images = extract_images(block)
    if images:
        existing = {img["path"] for img in slide.get("images", [])}
        slide.setdefault("images", []).extend(img for img in images if img["path"] not in existing)

    items = parse_duo_items(body_lines) if slide.get("layout_type") == "duo_compare" else parse_bullet_items(body_lines)
    if items:
        slide["items"] = items

    skip_keys = KNOWN_KEYS | {"icon", "title", "body", "label", "value"}
    plain = [line for line in clean_lines if not line.startswith("- ") and not (":" in line and line.split(":", 1)[0].strip() in skip_keys)]
    if slide.get("layout_type") == "cover" and plain:
        slide["title"] = "\n".join(plain[:2])
        if len(plain) > 2:
            slide["subtitle"] = " ".join(plain[2:])
    else:
        slide["title"] = heading or (plain[0] if plain else "Untitled")
        if plain:
            slide["body"] = "\n".join(plain)
    return slide


def markdown_to_spec(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8-sig")
    lines = text.splitlines()
    meta, _ = parse_frontmatter(lines)
    title = path.stem
    for line in lines:
        if line.startswith("# "):
            title = line[2:].strip()
            break
    slides = [parse_slide(block) for block in split_slide_blocks(text)]
    for slide in slides:
        resolve_slide_images(slide, path.parent)
    spec: dict[str, Any] = {
        "title": title,
        "subtitle": meta.get("subtitle", ""),
        "author": meta.get("author", ""),
        "date": meta.get("date", ""),
        "theme": meta.get("theme", "brutalist_tech"),
        "mode": meta.get("mode", "light"),
        "slides": slides,
    }
    return {key: value for key, value in spec.items() if value not in ("", None) and value != []}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert Obsidian Markdown to build-ppt deck_spec.json")
    parser.add_argument("markdown", type=Path, help="Path to an Obsidian-style Markdown deck brief")
    parser.add_argument("--out", type=Path, required=True, help="Output deck_spec.json path")
    args = parser.parse_args(argv)
    spec = markdown_to_spec(args.markdown.resolve())
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: wrote {args.out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
