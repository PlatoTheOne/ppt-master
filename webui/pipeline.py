from __future__ import annotations

import base64
import json
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Callable

from openai import OpenAI


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_ROOT = REPO_ROOT / "skills" / "ppt-master" / "scripts"
if str(SCRIPT_ROOT) not in sys.path:
    sys.path.insert(0, str(SCRIPT_ROOT))

from project_manager import ProjectManager  # type: ignore


LogFn = Callable[[str], None]

CANVAS = {
    "width": 1280,
    "height": 720,
    "viewbox": "0 0 1280 720",
}

FONT_STACKS = {
    "title": "Georgia, Times New Roman, SimSun, serif",
    "body": "Microsoft YaHei, PingFang SC, Arial, sans-serif",
    "mono": "Consolas, Courier New, monospace",
}

STYLE_PRESETS = {
    "business": {
        "name": "Business Briefing",
        "template": "business",
        "mode": "light",
        "background": "#F5F1EA",
        "surface": "#FFFCF7",
        "panel": "#FFFFFF",
        "primary": "#17324D",
        "accent": "#B46B36",
        "accent_soft": "#D9B178",
        "text": "#162534",
        "muted": "#63717B",
        "border": "#D8CDC0",
    },
    "tech": {
        "name": "Tech Product",
        "template": "tech",
        "mode": "dark",
        "background": "#09131E",
        "surface": "#102030",
        "panel": "#13283C",
        "primary": "#73D8FF",
        "accent": "#2DE0C2",
        "accent_soft": "#4F7CFF",
        "text": "#EAF7FF",
        "muted": "#8DA9BD",
        "border": "#254157",
    },
    "academic": {
        "name": "Academic Research",
        "template": "academic",
        "mode": "light",
        "background": "#F4F1EC",
        "surface": "#FBF8F4",
        "panel": "#FFFFFF",
        "primary": "#334A66",
        "accent": "#8D5B4C",
        "accent_soft": "#A6B5C6",
        "text": "#1E2A38",
        "muted": "#66727D",
        "border": "#D7D8DC",
    },
    "guizang": {
        "name": "Guizang Editorial",
        "template": "guizang",
        "mode": "light",
        "background": "#F1EFEA",
        "surface": "#E7E2D8",
        "panel": "#F8F4EC",
        "primary": "#171819",
        "accent": "#24446A",
        "accent_soft": "#B88B5A",
        "text": "#161514",
        "muted": "#625E57",
        "border": "#C9C1B5",
    },
}

LAYOUT_TYPES = {"cover", "agenda", "bullets", "two_column", "stats", "timeline", "quote", "closing"}


class GenerationError(RuntimeError):
    """Raised when the web pipeline cannot finish."""


@dataclass
class DraftArtifacts:
    project_dir: Path
    slide_svgs: list[Path]
    plan: dict
    theme: dict


@dataclass
class GenerationArtifacts:
    project_dir: Path
    native_pptx: Path | None
    legacy_pptx: Path | None
    web_deck: Path | None
    slide_svgs: list[Path]
    plan: dict


def slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "-", ascii_only).strip("-").lower()
    return cleaned or "pptmaster"


def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def clamp(value: int, lower: int, upper: int) -> int:
    return max(lower, min(upper, value))


def detect_language(text: str) -> str:
    cjk = sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")
    latin = sum(1 for ch in text if ch.isascii() and ch.isalpha())
    return "zh" if cjk >= latin else "en"


def strip_markdown(text: str) -> str:
    text = re.sub(r"```.*?```", "", text, flags=re.S)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"!\[[^\]]*\]\([^)]+\)", "", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = re.sub(r"^#{1,6}\s*", "", text, flags=re.M)
    text = re.sub(r"^\s*[-*+]\s+", "", text, flags=re.M)
    text = re.sub(r"^\s*\d+\.\s+", "", text, flags=re.M)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def split_sentences(text: str) -> list[str]:
    normalized = strip_markdown(text).replace("\r", "")
    parts = re.split(r"(?<=[。！？.!?])\s+|\n+", normalized)
    sentences = []
    for part in parts:
        candidate = re.sub(r"\s+", " ", part).strip(" -•\t")
        if len(candidate) >= 10:
            sentences.append(candidate)
    return sentences


def extract_sections(markdown_text: str) -> list[dict[str, str]]:
    lines = markdown_text.splitlines()
    sections: list[dict[str, str]] = []
    current: dict[str, str] | None = None

    for raw in lines:
        line = raw.rstrip()
        heading = re.match(r"^(#{1,3})\s+(.+)", line)
        if heading:
            if current and current["content"].strip():
                sections.append(current)
            current = {"title": heading.group(2).strip(), "content": ""}
            continue
        if current is None:
            current = {"title": "Overview", "content": ""}
        current["content"] += line + "\n"

    if current and current["content"].strip():
        sections.append(current)

    if sections:
        return sections

    chunks = [chunk.strip() for chunk in re.split(r"\n\s*\n", markdown_text) if chunk.strip()]
    fallback_sections = []
    for idx, chunk in enumerate(chunks[:8], start=1):
        fallback_sections.append({"title": f"Section {idx}", "content": chunk})
    return fallback_sections


def summarize_points(text: str, limit: int = 4) -> list[str]:
    sentences = split_sentences(text)
    if not sentences:
        stripped = strip_markdown(text)
        if stripped:
            return [stripped[:90]]
        return []

    points = []
    for sentence in sentences:
        cleaned = sentence[:110].strip()
        if cleaned and cleaned not in points:
            points.append(cleaned)
        if len(points) >= limit:
            break
    return points


def estimate_units(text: str) -> int:
    total = 0
    for ch in text:
        total += 2 if unicodedata.east_asian_width(ch) in {"W", "F"} else 1
    return total


def wrap_text(text: str, max_units: int) -> list[str]:
    if not text:
        return []
    words = re.split(r"(\s+)", text)
    lines: list[str] = []
    current = ""

    for token in words:
        candidate = f"{current}{token}"
        if current and estimate_units(candidate) > max_units:
            lines.append(current.strip())
            current = token.lstrip()
            continue
        current = candidate

    if current.strip():
        lines.append(current.strip())

    normalized: list[str] = []
    for line in lines:
        if estimate_units(line) <= max_units:
            normalized.append(line)
            continue

        buffer = ""
        for ch in line:
            if buffer and estimate_units(buffer + ch) > max_units:
                normalized.append(buffer)
                buffer = ch
            else:
                buffer += ch
        if buffer:
            normalized.append(buffer)

    return normalized


def xml_text(text: str) -> str:
    return escape(text, quote=False)


def normalize_hex(value: str, fallback: str) -> str:
    if isinstance(value, str) and re.fullmatch(r"#[0-9A-Fa-f]{6}", value.strip()):
        return value.strip().upper()
    return fallback


def safe_filename(name: str, suffix: str) -> str:
    base = slugify(name)[:60]
    return f"{base}{suffix}"


def normalize_export_mode(value: str | None) -> str:
    mode = (value or "pptx").strip().lower()
    if mode not in {"pptx", "web", "both"}:
        return "pptx"
    return mode


def extract_json_object(text: str) -> dict:
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise GenerationError("模型没有返回可解析的 JSON。")
    try:
        return json.loads(text[start : end + 1])
    except json.JSONDecodeError as exc:
        raise GenerationError(f"模型返回的 JSON 无法解析: {exc}") from exc


def read_markdown_sources(project_dir: Path) -> str:
    sources_dir = project_dir / "sources"
    markdown_files = sorted(sources_dir.glob("*.md"))
    if not markdown_files:
        text_files = sorted(sources_dir.glob("*.txt"))
        markdown_files = text_files

    blocks = []
    for path in markdown_files:
        try:
            content = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = path.read_text(encoding="utf-8", errors="ignore")
        if content.strip():
            blocks.append(f"# {path.stem}\n\n{content.strip()}")

    combined = "\n\n---\n\n".join(blocks).strip()
    if not combined:
        raise GenerationError("没有拿到可用于生成 PPT 的文本内容。")
    return combined


def choose_rhythm(slide_type: str) -> str:
    if slide_type in {"cover", "closing"}:
        return "anchor"
    if slide_type in {"quote"}:
        return "breathing"
    return "dense"


def choose_theme(style_key: str) -> dict[str, str]:
    return dict(STYLE_PRESETS.get(style_key, STYLE_PRESETS["business"]))


def apply_style_direction(plan: dict, style_key: str) -> dict:
    if style_key != "guizang":
        return plan

    language = plan.get("language") or "en"
    slides = [dict(slide) for slide in plan.get("slides") or []]
    if not slides:
        return plan

    cover_label = "Guizang Editorial" if language == "en" else "归藏杂志风"
    act_label = "Act" if language == "en" else "章节"
    contents_label = "Contents" if language == "en" else "目录"

    slides[0]["eyebrow"] = cover_label

    if len(slides) > 1 and slides[1].get("type") == "agenda":
        slides[1]["eyebrow"] = contents_label

    body_indices = [idx for idx, slide in enumerate(slides) if slide.get("type") not in {"cover", "agenda", "closing"}]
    for ordinal, idx in enumerate(body_indices, start=1):
        slides[idx]["eyebrow"] = f"{act_label} {ordinal:02d}"

    quote_candidates = [
        idx
        for idx in body_indices
        if slides[idx].get("type") in {"bullets", "two_column"}
        and (
            slides[idx].get("takeaway")
            or slides[idx].get("subtitle")
            or (slides[idx].get("bullets") or [])
            or (slides[idx].get("left_bullets") or [])
        )
    ]
    quote_index: int | None = None
    if quote_candidates:
        quote_index = quote_candidates[len(quote_candidates) // 2]
        quote_slide = slides[quote_index]
        quote_text = (
            quote_slide.get("takeaway")
            or quote_slide.get("subtitle")
            or next(iter(quote_slide.get("bullets") or []), "")
            or next(iter(quote_slide.get("left_bullets") or []), "")
        )
        if quote_text:
            quote_slide["type"] = "quote"
            quote_slide["quote"] = str(quote_text)[:180]
            quote_slide["attribution"] = quote_slide.get("title") or cover_label
            quote_slide["subtitle"] = str(quote_slide.get("subtitle") or quote_text)[:120]

    timeline_candidates = [idx for idx in body_indices if idx != quote_index and slides[idx].get("type") in {"bullets", "two_column"}]
    if len(body_indices) >= 3 and timeline_candidates:
        timeline_index = timeline_candidates[-1]
        timeline_slide = slides[timeline_index]
        raw_points = []
        if timeline_slide.get("type") == "two_column":
            raw_points.extend(timeline_slide.get("left_bullets") or [])
            raw_points.extend(timeline_slide.get("right_bullets") or [])
        else:
            raw_points.extend(timeline_slide.get("bullets") or [])
        steps = []
        for step_number, point in enumerate(raw_points[:4], start=1):
            steps.append(
                {
                    "label": f"Stage {step_number}" if language == "en" else f"阶段 {step_number}",
                    "detail": str(point)[:90],
                }
            )
        if len(steps) >= 3:
            timeline_slide["type"] = "timeline"
            timeline_slide["steps"] = steps
            timeline_slide["subtitle"] = str(timeline_slide.get("takeaway") or timeline_slide.get("subtitle") or "")[:120]

    slides[-1]["eyebrow"] = "Closing" if language == "en" else "收束"
    slides[-1]["closing_note"] = slides[-1].get("closing_note") or (
        "Magazine pacing, local SVG rendering, editable PowerPoint export"
        if language == "en"
        else "杂志化节奏，本地 SVG 渲染，可编辑 PowerPoint 导出"
    )

    return {**plan, "slides": slides}


def build_heuristic_plan(markdown_text: str, request: dict, log: LogFn) -> dict:
    log("Using heuristic planner")
    language = request.get("language") or "auto"
    language = detect_language(markdown_text) if language == "auto" else language
    page_count = clamp(int(request.get("pageCount") or 6), 4, 12)
    requested_title = (request.get("deckTitle") or "").strip()

    sections = extract_sections(markdown_text)
    title = requested_title or sections[0]["title"] if sections else "Generated Presentation"
    summary_lines = summarize_points(markdown_text, limit=3)
    subtitle = " | ".join(summary_lines[:2]) if summary_lines else ("AI-generated deck" if language == "en" else "自动生成演示文稿")

    agenda_items = [section["title"] for section in sections[: min(6, max(3, page_count - 2))]]
    slides: list[dict] = [
        {
            "type": "cover",
            "slug": "cover",
            "title": title,
            "subtitle": subtitle[:120],
            "eyebrow": request.get("styleKey", "business"),
            "bullets": summary_lines[:3],
        }
    ]

    if agenda_items:
        slides.append(
            {
                "type": "agenda",
                "slug": "agenda",
                "title": "Agenda" if language == "en" else "内容概览",
                "items": agenda_items,
            }
        )

    body_budget = max(1, page_count - 3)
    for index, section in enumerate(sections[:body_budget], start=1):
        points = summarize_points(section["content"], limit=5)
        if not points:
            points = summarize_points(section["title"], limit=3)
        numeric_hits = re.findall(r"\b\d+(?:\.\d+)?%?\b", section["content"])

        if numeric_hits and len(numeric_hits) >= 3 and index % 3 == 0:
            stats = []
            seen = []
            for hit in numeric_hits:
                if hit not in seen:
                    seen.append(hit)
                if len(seen) >= 4:
                    break
            for idx, value in enumerate(seen, start=1):
                stats.append(
                    {
                        "value": value,
                        "label": points[idx - 1][:28] if idx - 1 < len(points) else f"Metric {idx}",
                    }
                )
            slides.append(
                {
                    "type": "stats",
                    "slug": slugify(section["title"]),
                    "title": section["title"],
                    "subtitle": points[0] if points else "",
                    "stats": stats,
                    "takeaway": points[-1] if points else "",
                }
            )
            continue

        if len(points) >= 4 and index % 2 == 0:
            midpoint = max(2, len(points) // 2)
            slides.append(
                {
                    "type": "two_column",
                    "slug": slugify(section["title"]),
                    "title": section["title"],
                    "left_title": "Key Points" if language == "en" else "关键点",
                    "left_bullets": points[:midpoint],
                    "right_title": "Details" if language == "en" else "补充说明",
                    "right_bullets": points[midpoint:],
                    "takeaway": points[0],
                }
            )
            continue

        slides.append(
            {
                "type": "bullets",
                "slug": slugify(section["title"]),
                "title": section["title"],
                "subtitle": points[0] if points else "",
                "bullets": points[:5],
                "takeaway": points[-1] if points else "",
            }
        )

    closing_title = "Next Steps" if language == "en" else "下一步"
    closing_bullets = summary_lines[:3] or (
        ["Review the exported PPT", "Adjust wording", "Present"] if language == "en" else ["检查导出的 PPT", "微调文案", "开始演示"]
    )
    slides.append(
        {
            "type": "closing",
            "slug": "closing",
            "title": closing_title,
            "subtitle": requested_title or title,
            "bullets": closing_bullets,
            "closing_note": "Generated locally with PPT Master Web UI" if language == "en" else "由 PPT Master Web UI 本地生成",
        }
    )

    normalized = []
    for idx, slide in enumerate(slides[:page_count], start=1):
        slide["index"] = idx
        slide["slug"] = slide.get("slug") or f"slide-{idx:02d}"
        normalized.append(slide)

    return {
        "deck_title": title,
        "deck_subtitle": subtitle,
        "language": language,
        "slides": normalized,
    }


def build_openai_plan(markdown_text: str, request: dict, log: LogFn) -> dict:
    api_key = (request.get("apiKey") or os.environ.get("OPENAI_API_KEY") or "").strip()
    if not api_key:
        raise GenerationError("选择 OpenAI-compatible 规划器时，需要提供 API Key。")

    base_url = (request.get("baseUrl") or os.environ.get("OPENAI_BASE_URL") or "").strip() or None
    model = (request.get("model") or os.environ.get("OPENAI_MODEL") or "gpt-4.1-mini").strip()
    page_count = clamp(int(request.get("pageCount") or 6), 4, 12)
    language = request.get("language") or "auto"
    title = (request.get("deckTitle") or "").strip()
    source_excerpt = markdown_text[:18000]
    log(f"Calling model {model}")

    client = OpenAI(api_key=api_key, base_url=base_url)
    system_prompt = (
        "You are planning a presentation outline for a deterministic SVG renderer. "
        "Return JSON only. Do not wrap in markdown fences. "
        "Use concise, presentation-ready wording. "
        "Allowed slide types: cover, agenda, bullets, two_column, stats, timeline, quote, closing. "
        "Every deck must start with cover and end with closing. "
        "Use stats only when the source clearly contains numbers. "
        "Schema: "
        "{"
        '"deck_title": string, '
        '"deck_subtitle": string, '
        '"language": "en"|"zh", '
        '"slides": ['
        "{"
        '"type": string, "slug": string, "title": string, '
        '"subtitle": string?, "eyebrow": string?, '
        '"items": [string]?, "bullets": [string]?, '
        '"left_title": string?, "left_bullets": [string]?, '
        '"right_title": string?, "right_bullets": [string]?, '
        '"stats": [{"value": string, "label": string}]?, '
        '"steps": [{"label": string, "detail": string}]?, '
        '"quote": string?, "attribution": string?, '
        '"takeaway": string?, "closing_note": string?'
        "}"
        "]}"
    )
    user_prompt = (
        f"Requested title: {title or 'auto'}\n"
        f"Requested page count: {page_count}\n"
        f"Language: {language}\n"
        f"Style hint: {request.get('styleKey', 'business')}\n\n"
        "Source material:\n"
        f"{source_excerpt}"
    )
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.4,
    )
    content = response.choices[0].message.content or ""
    plan = extract_json_object(content)
    if not isinstance(plan, dict) or "slides" not in plan:
        raise GenerationError("模型没有返回有效的 deck 结构。")

    slides = plan.get("slides") or []
    normalized = []
    for idx, slide in enumerate(slides[:page_count], start=1):
        if not isinstance(slide, dict):
            continue
        slide_type = slide.get("type", "bullets")
        if slide_type not in LAYOUT_TYPES:
            slide_type = "bullets"
        slide["type"] = slide_type
        slide["slug"] = slide.get("slug") or f"slide-{idx:02d}"
        slide["index"] = idx
        normalized.append(slide)

    if not normalized:
        raise GenerationError("模型返回了空白 deck。")
    if normalized[0]["type"] != "cover":
        normalized.insert(
            0,
            {
                "type": "cover",
                "slug": "cover",
                "title": plan.get("deck_title") or title or "Generated Presentation",
                "subtitle": plan.get("deck_subtitle") or "",
                "bullets": [],
                "index": 1,
            },
        )
    normalized = normalized[:page_count]
    for idx, slide in enumerate(normalized, start=1):
        slide["index"] = idx
    if normalized[-1]["type"] != "closing":
        normalized[-1]["type"] = "closing"

    return {
        "deck_title": plan.get("deck_title") or title or "Generated Presentation",
        "deck_subtitle": plan.get("deck_subtitle") or "",
        "language": plan.get("language") or detect_language(markdown_text),
        "slides": normalized,
    }


def render_lines(lines: list[str], x: int, y: int, font_size: int, fill: str, font_family: str, font_weight: str = "400", line_gap: int = 12, anchor: str = "start") -> tuple[str, int]:
    cursor = y
    parts = []
    for line in lines:
        parts.append(
            f'<text x="{x}" y="{cursor}" font-family="{font_family}" font-size="{font_size}" '
            f'font-weight="{font_weight}" text-anchor="{anchor}" fill="{fill}">{xml_text(line)}</text>'
        )
        cursor += font_size + line_gap
    return "\n".join(parts), cursor


def base_defs(theme: dict, slide_number: int) -> str:
    return "<defs></defs>"


def render_chrome(theme: dict, slide: dict, plan: dict) -> str:
    title = plan["deck_title"]
    page_index = slide["index"]
    total = len(plan["slides"])
    top_label = slide.get("eyebrow") or title
    template = theme.get("template", "business")

    if template == "tech":
        backdrop = (
            f'<rect x="0" y="0" width="{CANVAS["width"]}" height="{CANVAS["height"]}" fill="{theme["background"]}"/>'
            f'<rect x="44" y="38" width="1192" height="644" rx="34" fill="{theme["surface"]}" fill-opacity="0.72" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<circle cx="1094" cy="126" r="170" fill="{theme["accent_soft"]}" fill-opacity="0.16"/>'
            f'<circle cx="980" cy="190" r="86" fill="{theme["accent"]}" fill-opacity="0.12"/>'
            f'<path d="M812 86H1184" stroke="{theme["border"]}" stroke-width="1.2"/>'
            f'<path d="M1042 46V216" stroke="{theme["border"]}" stroke-width="1.2"/>'
        )
    elif template == "guizang":
        backdrop = (
            f'<rect x="0" y="0" width="{CANVAS["width"]}" height="{CANVAS["height"]}" fill="{theme["background"]}"/>'
            f'<rect x="34" y="28" width="1212" height="664" rx="20" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1.2"/>'
            f'<rect x="58" y="118" width="544" height="492" rx="14" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="0.8" fill-opacity="0.4"/>'
            f'<circle cx="1110" cy="132" r="112" fill="{theme["accent_soft"]}" fill-opacity="0.10"/>'
            f'<circle cx="1044" cy="188" r="72" fill="{theme["accent"]}" fill-opacity="0.08"/>'
        )
    elif template == "academic":
        backdrop = (
            f'<rect x="0" y="0" width="{CANVAS["width"]}" height="{CANVAS["height"]}" fill="{theme["background"]}"/>'
            f'<rect x="0" y="0" width="{CANVAS["width"]}" height="124" fill="{theme["surface"]}"/>'
            f'<rect x="0" y="640" width="{CANVAS["width"]}" height="80" fill="{theme["surface"]}"/>'
            f'<rect x="70" y="86" width="1140" height="520" rx="24" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
        )
    else:
        backdrop = (
            f'<rect x="0" y="0" width="{CANVAS["width"]}" height="{CANVAS["height"]}" fill="{theme["background"]}"/>'
            f'<circle cx="1080" cy="110" r="180" fill="{theme["accent_soft"]}" fill-opacity="0.18"/>'
            f'<circle cx="1140" cy="180" r="110" fill="{theme["primary"]}" fill-opacity="0.08"/>'
            f'<rect x="880" y="-20" width="420" height="220" rx="110" fill="{theme["primary"]}" fill-opacity="0.05" transform="rotate(8 1090 90)"/>'
        )

    if template == "guizang":
        return (
            f"{backdrop}"
            f'<text x="86" y="82" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">{xml_text(top_label[:44].upper())}</text>'
            f'<text x="1192" y="82" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" text-anchor="end">{page_index:02d} / {total:02d}</text>'
            f'<line x1="84" y1="98" x2="1196" y2="98" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<line x1="84" y1="626" x2="1196" y2="626" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="86" y="656" font-family="{FONT_STACKS["body"]}" font-size="14" fill="{theme["muted"]}">{xml_text(title[:58])}</text>'
            f'<text x="1192" y="656" font-family="{FONT_STACKS["body"]}" font-size="14" fill="{theme["muted"]}" text-anchor="end">Local editorial deck / PPTX export</text>'
        )

    return (
        f"{backdrop}"
        f'<rect x="56" y="44" width="66" height="6" rx="3" fill="{theme["accent"]}"/>'
        f'<text x="56" y="84" font-family="{FONT_STACKS["body"]}" font-size="15" fill="{theme["muted"]}">{xml_text(top_label[:42])}</text>'
        f'<text x="1160" y="674" font-family="{FONT_STACKS["body"]}" font-size="12" fill="{theme["muted"]}" text-anchor="end">{page_index:02d} / {total:02d}</text>'
        f'<line x1="56" y1="646" x2="1224" y2="646" stroke="{theme["border"]}" stroke-width="1"/>'
    )


def render_cover(theme: dict, slide: dict, plan: dict) -> str:
    title_lines = wrap_text(slide.get("title") or plan["deck_title"], 22)[:3]
    subtitle_lines = wrap_text(slide.get("subtitle") or plan["deck_subtitle"], 44)[:2]
    bullets = [item for item in (slide.get("bullets") or []) if item][:3]
    template = theme.get("template", "business")

    title_svg, cursor = render_lines(title_lines, 88, 248, 56, theme["text"], FONT_STACKS["title"], font_weight="700", line_gap=16)
    subtitle_svg, cursor = render_lines(subtitle_lines, 92, cursor + 18, 22, theme["muted"], FONT_STACKS["body"], line_gap=10)

    chips = []
    chip_x = 92
    for bullet in bullets:
        label = bullet[:26]
        chip_width = max(170, min(320, 40 + estimate_units(label) * 8))
        chips.append(
            f'<rect x="{chip_x}" y="548" width="{chip_width}" height="44" rx="22" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="{chip_x + 20}" y="576" font-family="{FONT_STACKS["body"]}" font-size="15" fill="{theme["text"]}">{xml_text(label)}</text>'
        )
        chip_x += chip_width + 14

    if template == "tech":
        hero = (
            f'<rect x="74" y="124" width="1132" height="448" rx="36" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1.2"/>'
            f'<rect x="742" y="168" width="392" height="300" rx="24" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<rect x="770" y="196" width="136" height="136" rx="22" fill="{theme["accent_soft"]}" fill-opacity="0.18"/>'
            f'<rect x="930" y="196" width="176" height="18" rx="9" fill="{theme["accent"]}"/>'
            f'<rect x="930" y="230" width="144" height="12" rx="6" fill="{theme["muted"]}" fill-opacity="0.55"/>'
            f'<rect x="770" y="360" width="336" height="12" rx="6" fill="{theme["accent"]}" fill-opacity="0.92"/>'
            f'<rect x="770" y="392" width="284" height="12" rx="6" fill="{theme["accent_soft"]}" fill-opacity="0.92"/>'
            f'<rect x="770" y="424" width="232" height="12" rx="6" fill="{theme["border"]}"/>'
            f'<text x="770" y="514" font-family="{FONT_STACKS["body"]}" font-size="16" fill="{theme["muted"]}">Local pipeline · deterministic SVG · editable slides</text>'
        )
    elif template == "academic":
        hero = (
            f'<rect x="86" y="146" width="1108" height="392" rx="18" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<line x1="736" y1="176" x2="736" y2="510" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="780" y="216" font-family="{FONT_STACKS["body"]}" font-size="15" fill="{theme["muted"]}">Abstract</text>'
            f'<text x="780" y="256" font-family="{FONT_STACKS["title"]}" font-size="28" font-weight="700" fill="{theme["primary"]}">Structured source to editable presentation</text>'
            f'<text x="780" y="308" font-family="{FONT_STACKS["body"]}" font-size="17" fill="{theme["muted"]}">A calm, report-style deck with restrained hierarchy</text>'
            f'<rect x="780" y="356" width="330" height="126" rx="14" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="808" y="398" font-family="{FONT_STACKS["body"]}" font-size="16" fill="{theme["text"]}">Method</text>'
            f'<text x="808" y="432" font-family="{FONT_STACKS["body"]}" font-size="15" fill="{theme["muted"]}">Normalize source · plan sections · render SVG · export PPTX</text>'
        )
    elif template == "guizang":
        bullet_y = 394
        right_notes = []
        for index, bullet in enumerate(bullets, start=1):
            lines, _ = render_lines(wrap_text(bullet, 18)[:2], 850, bullet_y + 26, 16, theme["text"], FONT_STACKS["body"], line_gap=7)
            right_notes.append(
                f'<text x="814" y="{bullet_y + 24}" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["accent"]}">{index:02d}</text>'
                f"{lines}"
            )
            bullet_y += 84
        hero = (
            f'<line x1="786" y1="142" x2="786" y2="558" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="88" y="154" font-family="{FONT_STACKS["mono"]}" font-size="13" fill="{theme["accent"]}" letter-spacing="1.8">ISSUE 01</text>'
            f'<text x="88" y="186" font-family="{FONT_STACKS["body"]}" font-size="17" fill="{theme["muted"]}">Guizang-inspired editorial direction</text>'
            f'<rect x="814" y="144" width="320" height="176" rx="20" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="846" y="184" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">ABSTRACT</text>'
            f'<text x="846" y="226" font-family="{FONT_STACKS["title"]}" font-size="27" font-weight="700" fill="{theme["accent"]}">From source material to an editable deck</text>'
            f'<text x="846" y="278" font-family="{FONT_STACKS["body"]}" font-size="16" fill="{theme["muted"]}">Borrow the pacing and typography, keep the native PPTX pipeline.</text>'
            f'<text x="814" y="362" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">EDITORIAL NOTES</text>'
            f'{"".join(right_notes)}'
        )
    else:
        hero = (
            f'<rect x="72" y="128" width="640" height="434" rx="32" fill="{theme["surface"]}" fill-opacity="0.78" stroke="{theme["border"]}" stroke-width="1.4"/>'
            f'<rect x="748" y="128" width="460" height="434" rx="32" fill="{theme["primary"]}" fill-opacity="0.10"/>'
            f'<rect x="786" y="170" width="344" height="12" rx="6" fill="{theme["accent"]}" fill-opacity="0.95"/>'
            f'<rect x="786" y="214" width="292" height="12" rx="6" fill="{theme["accent_soft"]}" fill-opacity="0.95"/>'
            f'<rect x="786" y="258" width="252" height="12" rx="6" fill="{theme["border"]}"/>'
            f'<rect x="786" y="346" width="362" height="146" rx="22" fill="{theme["panel"]}" fill-opacity="0.92" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="816" y="388" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">Input</text>'
            f'<text x="816" y="430" font-family="{FONT_STACKS["title"]}" font-size="30" font-weight="700" fill="{theme["text"]}">Document to editable PPTX</text>'
            f'<text x="816" y="474" font-family="{FONT_STACKS["body"]}" font-size="16" fill="{theme["muted"]}">Structured local pipeline with native PowerPoint export</text>'
        )

    return f"{hero}{title_svg}{subtitle_svg}{''.join(chips)}"


def render_agenda(theme: dict, slide: dict, _plan: dict) -> str:
    items = slide.get("items") or slide.get("bullets") or []
    items = [str(item)[:56] for item in items][:6]
    if theme.get("template") == "guizang":
        left_items = items[:3]
        right_items = items[3:]

        def column(items_subset: list[str], x: int, start_index: int) -> str:
            parts = []
            y = 244
            for offset, item in enumerate(items_subset, start=1):
                lines, _ = render_lines(wrap_text(item, 22)[:2], x + 54, y, 20, theme["text"], FONT_STACKS["body"], line_gap=8)
                parts.append(
                    f'<text x="{x}" y="{y}" font-family="{FONT_STACKS["mono"]}" font-size="13" fill="{theme["accent"]}">{start_index + offset - 1:02d}</text>'
                    f"{lines}"
                    f'<line x1="{x}" y1="{y + 38}" x2="{x + 458}" y2="{y + 38}" stroke="{theme["border"]}" stroke-width="1"/>'
                )
                y += 108
            return "".join(parts)

        return (
            f'<text x="88" y="168" font-family="{FONT_STACKS["title"]}" font-size="46" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "Agenda")}</text>'
            f'<text x="88" y="204" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">A calmer chapter map before we enter the deck.</text>'
            f'<line x1="642" y1="228" x2="642" y2="564" stroke="{theme["border"]}" stroke-width="1"/>'
            f"{column(left_items, 108, 1)}"
            f"{column(right_items, 678, 4)}"
        )

    blocks = []
    y = 178
    for idx, item in enumerate(items, start=1):
        blocks.append(
            f'<rect x="84" y="{y}" width="1112" height="68" rx="24" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<circle cx="126" cy="{y + 34}" r="19" fill="{theme["accent"]}"/>'
            f'<text x="126" y="{y + 41}" text-anchor="middle" font-family="{FONT_STACKS["body"]}" font-size="18" font-weight="700" fill="#FFFFFF">{idx}</text>'
            f'<text x="168" y="{y + 42}" font-family="{FONT_STACKS["title"]}" font-size="25" font-weight="700" fill="{theme["text"]}">{xml_text(item)}</text>'
        )
        y += 84
    return (
        f'<text x="84" y="138" font-family="{FONT_STACKS["title"]}" font-size="40" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "Agenda")}</text>'
        f'<text x="84" y="170" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">A clean run-through of the deck structure</text>'
        f'{"".join(blocks)}'
    )


def render_bullet_cards(theme: dict, slide: dict) -> str:
    bullets = [str(item)[:100] for item in (slide.get("bullets") or [])][:5]
    if theme.get("template") == "guizang":
        list_rows = []
        y = 242
        for index, bullet in enumerate(bullets[:4], start=1):
            lines, _ = render_lines(wrap_text(bullet, 34)[:2], 176, y, 18, theme["text"], FONT_STACKS["body"], line_gap=8)
            list_rows.append(
                f'<text x="116" y="{y}" font-family="{FONT_STACKS["mono"]}" font-size="13" fill="{theme["accent"]}">{index:02d}</text>'
                f"{lines}"
                f'<line x1="108" y1="{y + 42}" x2="760" y2="{y + 42}" stroke="{theme["border"]}" stroke-width="1"/>'
            )
            y += 102

        takeaway = slide.get("takeaway") or slide.get("subtitle") or ""
        takeaway_lines, _ = render_lines(wrap_text(takeaway, 20)[:5], 846, 318, 20, theme["text"], FONT_STACKS["body"], line_gap=10)
        return (
            f'<text x="88" y="160" font-family="{FONT_STACKS["title"]}" font-size="42" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "")}</text>'
            f'<text x="88" y="198" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{xml_text((slide.get("subtitle") or "")[:120])}</text>'
            f'{"".join(list_rows)}'
            f'<rect x="810" y="222" width="318" height="318" rx="22" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="846" y="264" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">MARGIN NOTE</text>'
            f'<text x="846" y="300" font-family="{FONT_STACKS["title"]}" font-size="28" font-weight="700" fill="{theme["accent"]}">{xml_text((slide.get("title") or "")[:30])}</text>'
            f"{takeaway_lines}"
        )

    cards = []
    y = 220
    for bullet in bullets:
        wrapped = wrap_text(bullet, 42)[:2]
        card_lines, _ = render_lines(wrapped, 156, y + 40, 20, theme["text"], FONT_STACKS["body"], line_gap=8)
        cards.append(
            f'<rect x="120" y="{y}" width="680" height="84" rx="22" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<rect x="120" y="{y}" width="14" height="84" rx="7" fill="{theme["accent"]}"/>'
            f"{card_lines}"
        )
        y += 98
    takeaway = slide.get("takeaway") or slide.get("subtitle") or ""
    takeaway_lines, _ = render_lines(wrap_text(takeaway, 20)[:5], 874, 286, 20, theme["text"], FONT_STACKS["body"], line_gap=10)
    return (
        f'<text x="84" y="138" font-family="{FONT_STACKS["title"]}" font-size="38" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "")}</text>'
        f'<text x="84" y="172" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{xml_text((slide.get("subtitle") or "")[:120])}</text>'
        f'{"".join(cards)}'
        f'<rect x="844" y="220" width="352" height="378" rx="28" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1"/>'
        f'<text x="874" y="258" font-family="{FONT_STACKS["body"]}" font-size="15" fill="{theme["muted"]}">Takeaway</text>'
        f'<text x="874" y="316" font-family="{FONT_STACKS["title"]}" font-size="28" font-weight="700" fill="{theme["accent"]}">{xml_text((slide.get("title") or "")[:34])}</text>'
        f"{takeaway_lines}"
    )


def render_two_column(theme: dict, slide: dict) -> str:
    left_title = slide.get("left_title") or "Left"
    right_title = slide.get("right_title") or "Right"
    left_bullets = [str(item)[:70] for item in (slide.get("left_bullets") or [])][:4]
    right_bullets = [str(item)[:70] for item in (slide.get("right_bullets") or [])][:4]

    if theme.get("template") == "guizang":
        def column(x: int, title: str, bullets: list[str], accent: str) -> str:
            parts = [
                f'<text x="{x}" y="258" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">COLUMN</text>',
                f'<text x="{x}" y="294" font-family="{FONT_STACKS["title"]}" font-size="30" font-weight="700" fill="{accent}">{xml_text(title)}</text>',
            ]
            y = 338
            for bullet in bullets:
                lines, _ = render_lines(wrap_text(bullet, 24)[:2], x + 30, y, 18, theme["text"], FONT_STACKS["body"], line_gap=8)
                parts.append(f'<circle cx="{x + 10}" cy="{y - 6}" r="4" fill="{accent}"/>')
                parts.append(lines)
                y += 84
            return "".join(parts)

        return (
            f'<text x="88" y="160" font-family="{FONT_STACKS["title"]}" font-size="42" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "")}</text>'
            f'<text x="88" y="198" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{xml_text((slide.get("takeaway") or slide.get("subtitle") or "")[:120])}</text>'
            f'<line x1="642" y1="226" x2="642" y2="566" stroke="{theme["border"]}" stroke-width="1"/>'
            f"{column(108, left_title, left_bullets, theme['accent'])}"
            f"{column(690, right_title, right_bullets, theme['accent_soft'])}"
        )

    def column(x: int, title: str, bullets: list[str], accent: str) -> str:
        parts = [
            f'<rect x="{x}" y="204" width="516" height="374" rx="28" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>',
            f'<text x="{x + 34}" y="252" font-family="{FONT_STACKS["title"]}" font-size="28" font-weight="700" fill="{accent}">{xml_text(title)}</text>',
        ]
        y = 292
        for bullet in bullets:
            lines, _ = render_lines(wrap_text(bullet, 28)[:2], x + 54, y, 18, theme["text"], FONT_STACKS["body"], line_gap=8)
            parts.append(f'<circle cx="{x + 34}" cy="{y - 6}" r="6" fill="{accent}"/>')
            parts.append(lines)
            y += 86
        return "".join(parts)

    return (
        f'<text x="84" y="138" font-family="{FONT_STACKS["title"]}" font-size="38" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "")}</text>'
        f'<text x="84" y="172" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{xml_text((slide.get("takeaway") or slide.get("subtitle") or "")[:120])}</text>'
        f"{column(84, left_title, left_bullets, theme['accent'])}"
        f"{column(680, right_title, right_bullets, theme['accent_soft'])}"
    )


def render_stats(theme: dict, slide: dict) -> str:
    stats = slide.get("stats") or []
    if theme.get("template") == "guizang":
        boxes = []
        coords = [(88, 234), (648, 234), (88, 420), (648, 420)]
        for item, (x, y) in zip(stats[:4], coords):
            value = str(item.get("value", ""))[:18]
            label = str(item.get("label", ""))[:34]
            label_lines, _ = render_lines(wrap_text(label, 18)[:2], x + 34, y + 124, 18, theme["text"], FONT_STACKS["body"], line_gap=8)
            boxes.append(
                f'<rect x="{x}" y="{y}" width="468" height="154" rx="18" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
                f'<text x="{x + 34}" y="{y + 34}" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">STAT</text>'
                f'<text x="{x + 34}" y="{y + 96}" font-family="{FONT_STACKS["title"]}" font-size="52" font-weight="700" fill="{theme["accent"]}">{xml_text(value)}</text>'
                f"{label_lines}"
            )
        takeaway = xml_text((slide.get("takeaway") or slide.get("subtitle") or "")[:120])
        return (
            f'<text x="88" y="160" font-family="{FONT_STACKS["title"]}" font-size="42" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "")}</text>'
            f'<text x="88" y="198" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{takeaway}</text>'
            f'{"".join(boxes)}'
        )

    boxes = []
    coords = [(84, 224), (648, 224), (84, 404), (648, 404)]
    for item, (x, y) in zip(stats[:4], coords):
        value = str(item.get("value", ""))[:18]
        label = str(item.get("label", ""))[:34]
        label_lines, _ = render_lines(wrap_text(label, 18)[:2], x + 36, y + 126, 18, theme["muted"], FONT_STACKS["body"], line_gap=8)
        boxes.append(
            f'<rect x="{x}" y="{y}" width="480" height="150" rx="28" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="{x + 36}" y="{y + 76}" font-family="{FONT_STACKS["title"]}" font-size="54" font-weight="700" fill="{theme["accent"]}">{xml_text(value)}</text>'
            f"{label_lines}"
        )
    takeaway = xml_text((slide.get("takeaway") or slide.get("subtitle") or "")[:120])
    return (
        f'<text x="84" y="138" font-family="{FONT_STACKS["title"]}" font-size="38" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "")}</text>'
        f'<text x="84" y="176" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{takeaway}</text>'
        f'{"".join(boxes)}'
    )


def render_timeline(theme: dict, slide: dict) -> str:
    steps = slide.get("steps") or []
    if not steps:
        bullets = slide.get("bullets") or []
        steps = [{"label": f"Step {idx}", "detail": item} for idx, item in enumerate(bullets[:4], start=1)]

    points = []
    x_positions = [170, 430, 690, 950]
    for idx, (step, x) in enumerate(zip(steps[:4], x_positions), start=1):
        detail_lines, _ = render_lines(wrap_text(str(step.get("detail", ""))[:90], 16)[:3], x - 80, 418, 16, theme["text"], FONT_STACKS["body"], line_gap=6, anchor="middle")
        points.append(
            f'<circle cx="{x}" cy="324" r="34" fill="{theme["panel"]}" stroke="{theme["accent"]}" stroke-width="4"/>'
            f'<text x="{x}" y="334" text-anchor="middle" font-family="{FONT_STACKS["title"]}" font-size="24" font-weight="700" fill="{theme["accent"]}">{idx}</text>'
            f'<text x="{x}" y="380" text-anchor="middle" font-family="{FONT_STACKS["title"]}" font-size="22" font-weight="700" fill="{theme["text"]}">{xml_text(str(step.get("label", ""))[:18])}</text>'
            f"{detail_lines}"
        )
    return (
        f'<text x="84" y="138" font-family="{FONT_STACKS["title"]}" font-size="38" font-weight="700" fill="{theme["text"]}">{xml_text(slide.get("title") or "")}</text>'
        f'<text x="84" y="176" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{xml_text((slide.get("subtitle") or slide.get("takeaway") or "")[:120])}</text>'
        f'<line x1="170" y1="324" x2="950" y2="324" stroke="{theme["border"]}" stroke-width="6"/>'
        f'{"".join(points)}'
    )


def render_quote(theme: dict, slide: dict) -> str:
    quote = slide.get("quote") or slide.get("takeaway") or slide.get("subtitle") or ""
    quote_lines = wrap_text(quote, 34)[:4]
    quote_svg, cursor = render_lines(quote_lines, 164, 284, 34, theme["text"], FONT_STACKS["title"], font_weight="700", line_gap=16)
    attribution = slide.get("attribution") or slide.get("title") or ""
    if theme.get("template") == "guizang":
        return (
            f'<text x="88" y="210" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">PULL QUOTE</text>'
            f'<line x1="116" y1="236" x2="116" y2="544" stroke="{theme["accent"]}" stroke-width="4"/>'
            f"{quote_svg}"
            f'<text x="164" y="{cursor + 30}" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{xml_text(attribution[:60])}</text>'
            f'<rect x="866" y="214" width="260" height="286" rx="18" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="900" y="252" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">EDITORIAL</text>'
            f'<text x="900" y="292" font-family="{FONT_STACKS["title"]}" font-size="28" font-weight="700" fill="{theme["accent"]}">One line worth pausing on</text>'
            f'<text x="900" y="338" font-family="{FONT_STACKS["body"]}" font-size="16" fill="{theme["muted"]}">Use this page to slow the rhythm and let a point land.</text>'
        )

    return (
        f'<text x="108" y="260" font-family="{FONT_STACKS["title"]}" font-size="180" fill="{theme["accent"]}" fill-opacity="0.18">"</text>'
        f'<rect x="96" y="174" width="1088" height="392" rx="36" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1"/>'
        f"{quote_svg}"
        f'<text x="164" y="{cursor + 24}" font-family="{FONT_STACKS["body"]}" font-size="18" fill="{theme["muted"]}">{xml_text(attribution[:60])}</text>'
    )


def render_closing(theme: dict, slide: dict, plan: dict) -> str:
    bullets = [str(item)[:90] for item in (slide.get("bullets") or [])][:4]
    if theme.get("template") == "guizang":
        title_lines = wrap_text(slide.get("title") or "Closing", 18)[:2]
        title_svg, cursor = render_lines(title_lines, 88, 244, 60, theme["text"], FONT_STACKS["title"], font_weight="700", line_gap=18)
        notes = []
        y = 388
        for index, bullet in enumerate(bullets[:3], start=1):
            lines, _ = render_lines(wrap_text(bullet, 28)[:2], 162, y, 18, theme["text"], FONT_STACKS["body"], line_gap=8)
            notes.append(
                f'<text x="108" y="{y}" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["accent"]}">{index:02d}</text>'
                f"{lines}"
                f'<line x1="108" y1="{y + 40}" x2="628" y2="{y + 40}" stroke="{theme["border"]}" stroke-width="1"/>'
            )
            y += 92
        note = slide.get("closing_note") or plan["deck_subtitle"]
        note_lines, _ = render_lines(wrap_text(str(note), 22)[:4], 774, 316, 20, theme["text"], FONT_STACKS["body"], line_gap=10)
        return (
            f"{title_svg}"
            f'<text x="88" y="{cursor + 24}" font-family="{FONT_STACKS["body"]}" font-size="20" fill="{theme["muted"]}">{xml_text((slide.get("subtitle") or plan["deck_title"])[:100])}</text>'
            f'{"".join(notes)}'
            f'<rect x="736" y="208" width="396" height="336" rx="22" fill="{theme["surface"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<text x="774" y="252" font-family="{FONT_STACKS["mono"]}" font-size="12" fill="{theme["muted"]}" letter-spacing="1.6">FINAL NOTE</text>'
            f'<text x="774" y="292" font-family="{FONT_STACKS["title"]}" font-size="30" font-weight="700" fill="{theme["accent"]}">{xml_text(plan["deck_title"][:28])}</text>'
            f"{note_lines}"
        )

    title_lines = wrap_text(slide.get("title") or "Closing", 20)[:2]
    title_svg, cursor = render_lines(title_lines, 100, 244, 52, theme["text"], FONT_STACKS["title"], font_weight="700", line_gap=16)
    cards = []
    y = 360
    for bullet in bullets:
        lines, _ = render_lines(wrap_text(bullet, 34)[:2], 140, y + 34, 18, theme["text"], FONT_STACKS["body"], line_gap=8)
        cards.append(
            f'<rect x="108" y="{y}" width="520" height="76" rx="20" fill="{theme["panel"]}" stroke="{theme["border"]}" stroke-width="1"/>'
            f'<rect x="108" y="{y}" width="14" height="76" rx="7" fill="{theme["accent"]}"/>'
            f"{lines}"
        )
        y += 92
    note = slide.get("closing_note") or plan["deck_subtitle"]
    note_lines, _ = render_lines(wrap_text(str(note), 24)[:4], 764, 300, 21, theme["text"], FONT_STACKS["body"], line_gap=10)
    return (
        f"{title_svg}"
        f'<text x="100" y="{cursor + 18}" font-family="{FONT_STACKS["body"]}" font-size="20" fill="{theme["muted"]}">{xml_text((slide.get("subtitle") or plan["deck_title"])[:100])}</text>'
        f'{"".join(cards)}'
        f'<rect x="724" y="206" width="472" height="352" rx="32" fill="{theme["primary"]}" fill-opacity="0.10"/>'
        f'<text x="764" y="254" font-family="{FONT_STACKS["body"]}" font-size="15" fill="{theme["muted"]}">Ready</text>'
        f'<text x="764" y="296" font-family="{FONT_STACKS["title"]}" font-size="34" font-weight="700" fill="{theme["accent"]}">{xml_text(plan["deck_title"][:32])}</text>'
        f"{note_lines}"
    )


def render_slide(theme: dict, slide: dict, plan: dict) -> str:
    slide_number = slide["index"]
    layout = slide["type"]
    content = ""
    if layout == "cover":
        content = render_cover(theme, slide, plan)
    elif layout == "agenda":
        content = render_agenda(theme, slide, plan)
    elif layout == "bullets":
        content = render_bullet_cards(theme, slide)
    elif layout == "two_column":
        content = render_two_column(theme, slide)
    elif layout == "stats":
        content = render_stats(theme, slide)
    elif layout == "timeline":
        content = render_timeline(theme, slide)
    elif layout == "quote":
        content = render_quote(theme, slide)
    else:
        content = render_closing(theme, slide, plan)

    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{CANVAS["viewbox"]}" width="{CANVAS["width"]}" height="{CANVAS["height"]}">'
        f"{base_defs(theme, slide_number)}"
        f"{render_chrome(theme, slide, plan)}"
        f"{content}"
        "</svg>"
    )


def write_design_files(project_dir: Path, plan: dict, theme: dict) -> None:
    project_name = plan["deck_title"]
    slide_count = len(plan["slides"])
    design_spec = (
        f"# {project_name} - Design Spec\n\n"
        f"## I. Project Information\n\n"
        f"- Project Name: {project_name}\n"
        f"- Canvas Format: PPT 16:9 (1280x720)\n"
        f"- Page Count: {slide_count}\n"
        f"- Language: {plan['language']}\n"
        f"- Style: {theme['name']}\n\n"
        "## II. Visual Theme\n\n"
        f"- Background: {theme['background']}\n"
        f"- Surface: {theme['surface']}\n"
        f"- Primary: {theme['primary']}\n"
        f"- Accent: {theme['accent']}\n"
        f"- Text: {theme['text']}\n\n"
        "## III. Content Outline\n\n"
    )
    for slide in plan["slides"]:
        design_spec += f"- P{slide['index']:02d} [{slide['type']}] {slide.get('title', '')}\n"

    spec_lock_lines = [
        "# Execution Lock",
        "",
        "## canvas",
        f"- viewBox: {CANVAS['viewbox']}",
        "- format: PPT 16:9",
        "",
        "## colors",
        f"- bg: {theme['background']}",
        f"- surface: {theme['surface']}",
        f"- panel: {theme['panel']}",
        f"- primary: {theme['primary']}",
        f"- accent: {theme['accent']}",
        f"- secondary_accent: {theme['accent_soft']}",
        f"- text: {theme['text']}",
        f"- text_secondary: {theme['muted']}",
        f"- border: {theme['border']}",
        "- inverse_text: #FFFFFF",
        "",
        "## typography",
        f'- font_family: {FONT_STACKS["body"]}',
        f'- title_family: {FONT_STACKS["title"]}',
        f'- body_family: {FONT_STACKS["body"]}',
        f'- code_family: {FONT_STACKS["mono"]}',
        "- body: 18",
        "- title: 38",
        "- subtitle: 24",
        "- annotation: 14",
        "",
        "## page_rhythm",
    ]
    for slide in plan["slides"]:
        spec_lock_lines.append(f"- P{slide['index']:02d}: {choose_rhythm(slide['type'])}")
    spec_lock_lines += [
        "",
        "## forbidden",
        "- Mixing icon libraries",
        "- rgba()",
        '- <style>, class, <foreignObject>, textPath, @font-face, <animate*>, <script>, <iframe>, <symbol>+<use>',
        "- <g opacity>",
    ]

    (project_dir / "design_spec.md").write_text(design_spec, encoding="utf-8")
    (project_dir / "spec_lock.md").write_text("\n".join(spec_lock_lines), encoding="utf-8")


def save_slides(project_dir: Path, plan: dict, theme: dict, log: LogFn) -> list[Path]:
    svg_output = project_dir / "svg_output"
    generated: list[Path] = []
    for slide in plan["slides"]:
        svg_name = f"slide_{slide['index']:02d}_{safe_filename(slide['slug'], '.svg')}"
        svg_path = svg_output / svg_name
        svg_path.write_text(render_slide(theme, slide, plan), encoding="utf-8")
        generated.append(svg_path)
    log(f"Rendered {len(generated)} SVG slides")
    return generated


def run_command(args: list[str], log: LogFn, cwd: Path | None = None) -> str:
    process = subprocess.run(
        args,
        cwd=str(cwd or REPO_ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    output = "\n".join(part for part in [process.stdout.strip(), process.stderr.strip()] if part).strip()
    if output:
        for line in output.splitlines():
            log(line)
    if process.returncode != 0:
        raise GenerationError(f"命令执行失败: {' '.join(args)}")
    return output


def ingest_source(project_dir: Path, request: dict, log: LogFn) -> str:
    manager = ProjectManager()
    source_text = (request.get("sourceText") or "").strip()
    source_url = (request.get("sourceUrl") or "").strip()
    uploaded_name = (request.get("fileName") or "").strip()
    uploaded_content = (request.get("fileContentBase64") or "").strip()

    if source_text:
        filename = safe_filename(request.get("deckTitle") or "source", ".md")
        target = project_dir / "sources" / filename
        target.write_text(source_text, encoding="utf-8")
        log(f"Saved pasted content to {target.relative_to(REPO_ROOT)}")
        return read_markdown_sources(project_dir)

    if source_url:
        log(f"Importing URL: {source_url}")
        manager.import_sources(str(project_dir), [source_url], move=False, copy=True)
        return read_markdown_sources(project_dir)

    if uploaded_name and uploaded_content:
        suffix = Path(uploaded_name).suffix or ".txt"
        with tempfile.TemporaryDirectory(prefix="pptmaster-webui-") as tmp_dir:
            temp_path = Path(tmp_dir) / f"upload{suffix}"
            temp_path.write_bytes(base64.b64decode(uploaded_content))
            log(f"Importing uploaded file: {uploaded_name}")
            manager.import_sources(str(project_dir), [str(temp_path)], move=True, copy=False)
        return read_markdown_sources(project_dir)

    raise GenerationError("请上传文件、粘贴内容，或者提供一个 URL。")


def draft_state_path(project_dir: Path) -> Path:
    return project_dir / "webui_draft.json"


def ensure_slide_ids(plan: dict) -> dict:
    normalized = dict(plan)
    slides = []
    for index, slide in enumerate(plan.get("slides") or [], start=1):
        slide_copy = dict(slide)
        slide_copy["id"] = slide_copy.get("id") or f"slide-{index:02d}"
        slide_copy["index"] = index
        slide_copy["slug"] = slide_copy.get("slug") or slide_copy["id"]
        slides.append(slide_copy)
    normalized["slides"] = slides
    return normalized


def save_draft_state(project_dir: Path, plan: dict, theme: dict, source_meta: dict | None = None) -> Path:
    payload = {
        "plan": ensure_slide_ids(plan),
        "theme": theme,
        "sourceMeta": source_meta or {},
        "updatedAt": datetime.now().isoformat(timespec="seconds"),
    }
    state_path = draft_state_path(project_dir)
    state_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return state_path


def load_draft_state(project_dir: Path) -> dict:
    state_path = draft_state_path(project_dir)
    if not state_path.exists():
        raise GenerationError("Draft state not found for this project.")
    return json.loads(state_path.read_text(encoding="utf-8"))


def list_slide_svgs(project_dir: Path, source_dir: str = "svg_output") -> list[Path]:
    svg_dir = project_dir / source_dir
    if not svg_dir.exists():
        return []
    return sorted(svg_dir.glob("*.svg"))


def stage_quality_check(project_dir: Path, log: LogFn) -> None:
    checker_script = SCRIPT_ROOT / "svg_quality_checker.py"
    run_command([sys.executable, str(checker_script), str(project_dir)], log)


def stage_finalize(project_dir: Path, log: LogFn) -> None:
    finalize_script = SCRIPT_ROOT / "finalize_svg.py"
    run_command([sys.executable, str(finalize_script), str(project_dir)], log)


def stage_export(project_dir: Path, log: LogFn) -> tuple[Path, Path]:
    export_script = SCRIPT_ROOT / "svg_to_pptx.py"
    output_base = project_dir / "exports" / f"{project_dir.name}_generated.pptx"
    run_command(
        [
            sys.executable,
            str(export_script),
            str(project_dir),
            "-s",
            "final",
            "--no-notes",
            "-o",
            str(output_base),
        ],
        log,
    )
    return latest_svg_exports(project_dir)


def create_project(request: dict, log: LogFn) -> Path:
    project_title = (request.get("deckTitle") or "webui-presentation").strip()
    project_name = slugify(project_title)[:42]
    if not project_name:
        project_name = "webui-presentation"
    project_name = f"{project_name}_{now_stamp()}"
    manager = ProjectManager()
    project_path = Path(manager.init_project(project_name, "ppt169"))
    if not project_path.is_absolute():
        project_path = (REPO_ROOT / project_path).resolve()
    log(f"Project created at {project_path.relative_to(REPO_ROOT)}")
    return project_path


def latest_svg_exports(project_dir: Path) -> tuple[Path, Path]:
    exports_dir = project_dir / "exports"
    native = exports_dir / f"{project_dir.name}_generated.pptx"
    legacy = exports_dir / f"{project_dir.name}_generated_svg.pptx"
    if not native.exists():
        raise GenerationError("没有找到导出的原生 PPTX。")
    if not legacy.exists():
        raise GenerationError("没有找到导出的 SVG 参考 PPTX。")
    return native, legacy


def web_deck_path(project_dir: Path) -> Path:
    return project_dir / "exports" / f"{project_dir.name}_web_deck.html"


def web_text(value: str) -> str:
    return escape(str(value or ""), quote=True)


def web_join_list(items: list[str], item_tag: str = "li", class_name: str = "") -> str:
    class_attr = f' class="{class_name}"' if class_name else ""
    return "".join(f"<{item_tag}{class_attr}>{web_text(item)}</{item_tag}>" for item in items if str(item or "").strip())


def build_web_note_lines(slide: dict, plan: dict) -> list[str]:
    note_lines: list[str] = []
    title = str(slide.get("title") or "").strip()
    subtitle = str(slide.get("subtitle") or "").strip()
    takeaway = str(slide.get("takeaway") or "").strip()
    if title:
        note_lines.append(f"Focus: {title}")
    if subtitle:
        note_lines.append(subtitle[:180])

    slide_type = slide.get("type") or "bullets"
    if slide_type == "agenda":
        note_lines.extend(str(item)[:140] for item in (slide.get("items") or [])[:4])
    elif slide_type == "two_column":
        note_lines.extend(str(item)[:140] for item in (slide.get("left_bullets") or [])[:2])
        note_lines.extend(str(item)[:140] for item in (slide.get("right_bullets") or [])[:2])
    elif slide_type == "stats":
        for item in (slide.get("stats") or [])[:4]:
            note_lines.append(f"{item.get('label')}: {item.get('value')}")
    elif slide_type == "timeline":
        for step in (slide.get("steps") or [])[:4]:
            note_lines.append(f"{step.get('label')}: {step.get('detail')}")
    elif slide_type == "quote":
        quote = str(slide.get("quote") or "").strip()
        attribution = str(slide.get("attribution") or "").strip()
        if quote:
            note_lines.append(quote[:180])
        if attribution:
            note_lines.append(f"Source: {attribution}")
    else:
        note_lines.extend(str(item)[:140] for item in (slide.get("bullets") or [])[:4])

    if takeaway and takeaway not in note_lines:
        note_lines.append(f"Takeaway: {takeaway[:160]}")

    if not note_lines:
        fallback = str(plan.get("deck_subtitle") or plan.get("deck_title") or "").strip()
        if fallback:
            note_lines.append(fallback[:180])
    return note_lines[:6]


def render_web_slide(theme: dict, slide: dict, plan: dict, style_key: str) -> str:
    slide_type = slide.get("type") or "bullets"
    eyebrow = web_text(slide.get("eyebrow") or plan.get("deck_title") or "")
    title = web_text(slide.get("title") or "")
    subtitle = web_text(slide.get("subtitle") or slide.get("takeaway") or "")
    language = plan.get("language") or "en"
    is_guizang = style_key == "guizang"

    if slide_type == "cover":
        bullets = slide.get("bullets") or []
        chips = "".join(f'<span class="hero-chip">{web_text(item)}</span>' for item in bullets[:3] if str(item or "").strip())
        side_note = (
            '<div class="editorial-note"><span>Abstract</span><p>Editorial pacing, clean chapters, and a downloadable deck from the same local draft.</p></div>'
            if is_guizang
            else ""
        )
        return (
            f'<section class="web-slide web-cover{" is-guizang" if is_guizang else ""}">'
            f'<div class="slide-top"><span>{eyebrow}</span><span>01</span></div>'
            f'<div class="hero-layout">'
            f'<div class="hero-copy">'
            f'<p class="slide-kicker">{eyebrow}</p>'
            f'<h1>{title}</h1>'
            f'<p class="slide-subtitle">{subtitle}</p>'
            f'<div class="hero-chips">{chips}</div>'
            f"</div>"
            f"{side_note}"
            f"</div>"
            f"</section>"
        )

    if slide_type == "agenda":
        items = slide.get("items") or slide.get("bullets") or []
        return (
            f'<section class="web-slide">'
            f'<div class="slide-top"><span>{eyebrow}</span><span>{web_text(slide.get("index"))}</span></div>'
            f'<div class="slide-shell">'
            f'<p class="slide-kicker">{eyebrow}</p>'
            f'<h2>{title}</h2>'
            f'<p class="slide-subtitle">{subtitle or ("Deck map" if language == "en" else "本次内容结构")}</p>'
            f'<ol class="agenda-list">{web_join_list([str(item) for item in items[:6]], "li")}</ol>'
            f"</div></section>"
        )

    if slide_type == "two_column":
        left = slide.get("left_bullets") or []
        right = slide.get("right_bullets") or []
        return (
            f'<section class="web-slide">'
            f'<div class="slide-top"><span>{eyebrow}</span><span>{web_text(slide.get("index"))}</span></div>'
            f'<div class="slide-shell">'
            f'<p class="slide-kicker">{eyebrow}</p>'
            f'<h2>{title}</h2>'
            f'<p class="slide-subtitle">{subtitle}</p>'
            f'<div class="two-col">'
            f'<article class="content-card"><h3>{web_text(slide.get("left_title") or "Left")}</h3><ul>{web_join_list([str(item) for item in left[:4]])}</ul></article>'
            f'<article class="content-card"><h3>{web_text(slide.get("right_title") or "Right")}</h3><ul>{web_join_list([str(item) for item in right[:4]])}</ul></article>'
            f"</div></div></section>"
        )

    if slide_type == "stats":
        stats = slide.get("stats") or []
        cards = "".join(
            f'<article class="stat-card-web"><span>{web_text(item.get("label") or "")}</span><strong>{web_text(item.get("value") or "")}</strong></article>'
            for item in stats[:4]
        )
        return (
            f'<section class="web-slide">'
            f'<div class="slide-top"><span>{eyebrow}</span><span>{web_text(slide.get("index"))}</span></div>'
            f'<div class="slide-shell">'
            f'<p class="slide-kicker">{eyebrow}</p>'
            f'<h2>{title}</h2>'
            f'<p class="slide-subtitle">{subtitle}</p>'
            f'<div class="stats-grid">{cards}</div>'
            f"</div></section>"
        )

    if slide_type == "timeline":
        steps = slide.get("steps") or []
        items = "".join(
            f'<article class="timeline-step"><span>{web_text(step.get("label") or "")}</span><p>{web_text(step.get("detail") or "")}</p></article>'
            for step in steps[:4]
        )
        return (
            f'<section class="web-slide">'
            f'<div class="slide-top"><span>{eyebrow}</span><span>{web_text(slide.get("index"))}</span></div>'
            f'<div class="slide-shell">'
            f'<p class="slide-kicker">{eyebrow}</p>'
            f'<h2>{title}</h2>'
            f'<p class="slide-subtitle">{subtitle}</p>'
            f'<div class="timeline-grid">{items}</div>'
            f"</div></section>"
        )

    if slide_type == "quote":
        quote = web_text(slide.get("quote") or slide.get("takeaway") or slide.get("subtitle") or "")
        attribution = web_text(slide.get("attribution") or slide.get("title") or "")
        return (
            f'<section class="web-slide web-quote{" is-guizang" if is_guizang else ""}">'
            f'<div class="slide-top"><span>{eyebrow}</span><span>{web_text(slide.get("index"))}</span></div>'
            f'<div class="quote-block"><p>{quote}</p><span>{attribution}</span></div>'
            f"</section>"
        )

    if slide_type == "closing":
        bullets = slide.get("bullets") or []
        note = web_text(slide.get("closing_note") or plan.get("deck_subtitle") or "")
        return (
            f'<section class="web-slide web-closing">'
            f'<div class="slide-top"><span>{eyebrow}</span><span>{web_text(slide.get("index"))}</span></div>'
            f'<div class="slide-shell closing-layout">'
            f'<div>'
            f'<p class="slide-kicker">{eyebrow}</p>'
            f'<h2>{title}</h2>'
            f'<p class="slide-subtitle">{web_text(slide.get("subtitle") or plan.get("deck_title") or "")}</p>'
            f'<ul>{web_join_list([str(item) for item in bullets[:4]])}</ul>'
            f"</div>"
            f'<aside class="editorial-note"><span>Final note</span><p>{note}</p></aside>'
            f"</div></section>"
        )

    bullets = slide.get("bullets") or []
    return (
        f'<section class="web-slide">'
        f'<div class="slide-top"><span>{eyebrow}</span><span>{web_text(slide.get("index"))}</span></div>'
        f'<div class="slide-shell">'
        f'<p class="slide-kicker">{eyebrow}</p>'
        f'<h2>{title}</h2>'
        f'<p class="slide-subtitle">{subtitle}</p>'
        f'<div class="content-split">'
        f'<ul>{web_join_list([str(item) for item in bullets[:5]])}</ul>'
        f'<aside class="editorial-note"><span>Takeaway</span><p>{web_text(slide.get("takeaway") or slide.get("subtitle") or "")}</p></aside>'
        f"</div></div></section>"
    )


def build_web_deck_html(plan: dict, theme: dict, style_key: str) -> str:
    slides = plan.get("slides") or []
    slide_markup = "\n".join(render_web_slide(theme, slide, plan, style_key) for slide in slides)
    notes_payload = json.dumps(
        [
            {
                "title": str(slide.get("title") or f"Slide {index + 1}"),
                "type": str(slide.get("type") or "slide"),
                "lines": build_web_note_lines(slide, plan),
            }
            for index, slide in enumerate(slides)
        ],
        ensure_ascii=False,
    )
    dot_markup = "".join(
        f'<button class="nav-dot{" active" if index == 0 else ""}" type="button" data-index="{index}" aria-label="Go to slide {index + 1}"></button>'
        for index, _ in enumerate(slides)
    )
    title_font = '"Cormorant Garamond", Georgia, serif' if style_key == "guizang" else "Georgia, serif"
    body_font = '"IBM Plex Sans", "Microsoft YaHei", sans-serif'
    return f"""<!DOCTYPE html>
<html lang="{'zh-CN' if (plan.get('language') == 'zh') else 'en'}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{web_text(plan.get('deck_title') or 'Web Deck')}</title>
  <style>
    :root {{
      --bg: {theme['background']};
      --surface: {theme['surface']};
      --panel: {theme['panel']};
      --primary: {theme['primary']};
      --accent: {theme['accent']};
      --accent-soft: {theme['accent_soft']};
      --text: {theme['text']};
      --muted: {theme['muted']};
      --border: {theme['border']};
    }}
    * {{ box-sizing: border-box; }}
    html, body {{ margin: 0; height: 100%; background: var(--bg); color: var(--text); font-family: {body_font}; }}
    body {{ overflow: hidden; background: radial-gradient(circle at top right, rgba(184, 139, 90, 0.18), transparent 28%), radial-gradient(circle at bottom left, rgba(36, 68, 106, 0.12), transparent 24%), var(--bg); }}
    .deck {{ display: flex; width: 100vw; height: 100vh; overflow-x: auto; overflow-y: hidden; scroll-snap-type: x mandatory; scroll-behavior: smooth; }}
    .deck::-webkit-scrollbar {{ display: none; }}
    body.presenter-mode .deck {{ width: calc(100vw - min(30vw, 380px)); }}
    .web-slide {{ position: relative; flex: 0 0 100vw; width: 100vw; height: 100vh; padding: 40px 56px 56px; scroll-snap-align: start; background: linear-gradient(180deg, rgba(255,255,255,0.2), rgba(255,255,255,0.02)); }}
    .web-slide::before {{ content: ""; position: absolute; inset: 28px; border: 1px solid var(--border); border-radius: 24px; background: rgba(255,255,255,0.38); pointer-events: none; }}
    .slide-top {{ display: flex; justify-content: space-between; align-items: center; text-transform: uppercase; letter-spacing: .18em; font-size: 12px; color: var(--muted); margin-bottom: 34px; }}
    .slide-shell {{ position: relative; z-index: 1; display: grid; gap: 22px; align-content: start; min-height: calc(100vh - 140px); }}
    .slide-kicker {{ margin: 0; text-transform: uppercase; letter-spacing: .18em; font-size: 12px; color: var(--accent); }}
    h1, h2, h3 {{ margin: 0; font-family: {title_font}; line-height: .95; }}
    h1 {{ font-size: clamp(56px, 8vw, 120px); max-width: 10ch; }}
    h2 {{ font-size: clamp(34px, 5vw, 72px); max-width: 12ch; }}
    h3 {{ font-size: clamp(24px, 2vw, 32px); }}
    .slide-subtitle {{ margin: 0; max-width: 62ch; font-size: 18px; line-height: 1.7; color: var(--muted); }}
    .hero-layout {{ display: grid; grid-template-columns: minmax(0, 1.4fr) minmax(280px, .8fr); gap: 36px; align-items: end; min-height: calc(100vh - 180px); }}
    .hero-copy {{ display: grid; gap: 24px; align-content: center; }}
    .hero-chips {{ display: flex; flex-wrap: wrap; gap: 12px; }}
    .hero-chip {{ padding: 12px 18px; border-radius: 999px; background: rgba(255,255,255,0.55); border: 1px solid var(--border); font-size: 14px; }}
    .editorial-note {{ align-self: stretch; display: grid; gap: 14px; padding: 28px; border-radius: 24px; background: rgba(255,255,255,0.5); border: 1px solid var(--border); }}
    .editorial-note span {{ text-transform: uppercase; letter-spacing: .16em; font-size: 12px; color: var(--muted); }}
    .editorial-note p {{ margin: 0; font-size: 18px; line-height: 1.7; }}
    .agenda-list, .content-split ul, .content-card ul, .closing-layout ul {{ margin: 0; padding-left: 22px; display: grid; gap: 14px; font-size: 20px; line-height: 1.7; }}
    .agenda-list {{ grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px 42px; }}
    .two-col, .content-split, .closing-layout {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 28px; align-items: start; }}
    .content-card, .stat-card-web, .timeline-step {{ padding: 24px; border-radius: 24px; border: 1px solid var(--border); background: rgba(255,255,255,0.48); }}
    .stats-grid, .timeline-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; }}
    .stat-card-web span, .timeline-step span {{ display: block; margin-bottom: 10px; text-transform: uppercase; letter-spacing: .16em; font-size: 12px; color: var(--muted); }}
    .stat-card-web strong {{ font-family: {title_font}; font-size: clamp(34px, 5vw, 64px); color: var(--accent); }}
    .timeline-step p {{ margin: 0; font-size: 18px; line-height: 1.7; }}
    .quote-block {{ display: grid; gap: 18px; align-content: center; min-height: calc(100vh - 160px); max-width: 960px; margin: 0 auto; padding: 0 40px; }}
    .quote-block p {{ margin: 0; font-family: {title_font}; font-size: clamp(42px, 6vw, 82px); line-height: 1.08; }}
    .quote-block span {{ font-size: 16px; letter-spacing: .14em; text-transform: uppercase; color: var(--muted); }}
    .deck-progress {{ position: fixed; inset: 0 0 auto 0; z-index: 24; height: 4px; background: rgba(255,255,255,.18); }}
    .deck-progress-bar {{ width: 0; height: 100%; background: linear-gradient(90deg, var(--accent), var(--accent-soft)); transition: width .25s ease; }}
    .deck-toolbar {{
      position: fixed; top: 20px; right: 24px; z-index: 25; display: flex; align-items: center; gap: 10px;
      padding: 10px 12px; border-radius: 999px; background: rgba(20, 20, 20, .16); backdrop-filter: blur(14px);
      border: 1px solid rgba(255,255,255,.16);
    }}
    .deck-toolbar button {{
      border: 0; border-radius: 999px; padding: 10px 14px; font: inherit; font-size: 13px; color: white;
      background: rgba(255,255,255,.14); cursor: pointer; transition: transform .2s ease, background .2s ease, opacity .2s ease;
    }}
    .deck-toolbar button:hover {{ transform: translateY(-1px); background: rgba(255,255,255,.22); }}
    .deck-toolbar .is-active {{ background: var(--accent); }}
    .deck-status {{
      display: grid; gap: 2px; margin-right: 6px; color: white; min-width: 108px;
    }}
    .deck-status strong {{ font-size: 12px; letter-spacing: .16em; text-transform: uppercase; }}
    .deck-status span {{ font-size: 12px; opacity: .78; }}
    .speaker-panel {{
      position: fixed; top: 0; right: 0; width: min(30vw, 380px); height: 100vh; z-index: 23;
      display: grid; grid-template-rows: auto auto auto 1fr auto; gap: 18px; padding: 24px; color: #f5f4ef;
      background: rgba(18, 18, 18, .86); backdrop-filter: blur(20px); border-left: 1px solid rgba(255,255,255,.12);
      transform: translateX(100%); transition: transform .28s ease;
    }}
    body.notes-open .speaker-panel,
    body.presenter-mode .speaker-panel {{ transform: translateX(0); }}
    .speaker-panel-header {{ display: grid; gap: 8px; }}
    .speaker-panel-header strong {{ font-size: 12px; letter-spacing: .2em; text-transform: uppercase; opacity: .7; }}
    .speaker-panel-header h3 {{ font-size: clamp(26px, 2vw, 34px); line-height: 1.04; color: white; max-width: none; }}
    .speaker-panel-header p {{ margin: 0; font-size: 13px; text-transform: uppercase; letter-spacing: .16em; opacity: .6; }}
    .speaker-stats {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }}
    .speaker-stat {{
      display: grid; gap: 10px; padding: 16px; border-radius: 18px; background: rgba(255,255,255,.06);
      border: 1px solid rgba(255,255,255,.08);
    }}
    .speaker-stat span,
    .next-slide-header strong {{
      font-size: 11px; letter-spacing: .18em; text-transform: uppercase; opacity: .62;
    }}
    .speaker-stat strong {{ font-family: {title_font}; font-size: 28px; color: white; line-height: 1; }}
    .speaker-timer-meta {{ display: grid; gap: 10px; }}
    .speaker-timer-actions {{ display: flex; gap: 8px; flex-wrap: wrap; }}
    .speaker-timer-actions button {{
      border: 0; border-radius: 999px; padding: 7px 11px; font: inherit; font-size: 12px; color: white;
      background: rgba(255,255,255,.12); cursor: pointer; transition: background .2s ease, transform .2s ease;
    }}
    .speaker-timer-actions button:hover {{ background: rgba(255,255,255,.2); transform: translateY(-1px); }}
    .speaker-timer-actions button.is-active {{ background: var(--accent); }}
    .next-slide-card {{
      display: grid; gap: 12px; padding: 16px; border-radius: 20px; background: rgba(255,255,255,.05);
      border: 1px solid rgba(255,255,255,.08);
    }}
    .next-slide-header {{ display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }}
    .next-slide-header span {{ font-size: 12px; opacity: .7; }}
    .next-slide-card h4 {{ margin: 0; font-family: {title_font}; font-size: 24px; line-height: 1.12; color: white; }}
    .next-slide-lines {{ margin: 0; padding-left: 18px; display: grid; gap: 10px; font-size: 14px; line-height: 1.55; opacity: .9; }}
    .next-slide-lines li::marker {{ color: var(--accent-soft); }}
    .speaker-notes {{ margin: 0; padding-left: 18px; display: grid; gap: 14px; align-content: start; font-size: 17px; line-height: 1.7; overflow: auto; }}
    .speaker-notes li::marker {{ color: var(--accent-soft); }}
    .speaker-help {{ display: grid; gap: 8px; font-size: 12px; opacity: .72; }}
    .speaker-help span {{ display: flex; justify-content: space-between; gap: 12px; }}
    .speaker-help code {{
      padding: 4px 8px; border-radius: 8px; background: rgba(255,255,255,.08); color: white; font-family: {body_font};
    }}
    .presenter-console-app {{
      display: none; min-height: 100vh; padding: 28px; gap: 20px; color: #f5f4ef;
      background:
        radial-gradient(circle at top left, rgba(36, 68, 106, .28), transparent 34%),
        radial-gradient(circle at bottom right, rgba(184, 139, 90, .2), transparent 28%),
        #0e1116;
    }}
    .presenter-console-shell {{
      display: grid; grid-template-columns: minmax(340px, 1.1fr) minmax(320px, .9fr); gap: 18px; min-height: calc(100vh - 56px);
    }}
    .presenter-console-card {{
      display: grid; gap: 16px; padding: 22px; border-radius: 26px;
      background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.08);
      box-shadow: 0 20px 80px rgba(0,0,0,.24);
    }}
    .presenter-console-top {{
      display: flex; justify-content: space-between; align-items: flex-start; gap: 16px;
    }}
    .presenter-console-kicker,
    .presenter-console-top strong,
    .presenter-console-rail strong,
    .presenter-console-stat span,
    .presenter-console-next-header strong {{
      font-size: 11px; letter-spacing: .18em; text-transform: uppercase; opacity: .64;
    }}
    .presenter-console-top h2,
    .presenter-console-next h3,
    .presenter-console-current h3 {{
      margin: 0; font-family: {title_font}; color: white;
    }}
    .presenter-console-top h2 {{ font-size: clamp(36px, 3.2vw, 58px); line-height: .98; }}
    .presenter-console-top p,
    .presenter-console-current p,
    .presenter-console-status p {{
      margin: 0; font-size: 14px; color: rgba(245,244,239,.72); letter-spacing: .08em; text-transform: uppercase;
    }}
    .presenter-console-current,
    .presenter-console-next {{
      display: grid; gap: 12px;
    }}
    .presenter-console-current h3,
    .presenter-console-next h3 {{ font-size: clamp(28px, 2.4vw, 42px); line-height: 1.02; }}
    .presenter-console-grid {{
      display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px;
    }}
    .presenter-console-stat {{
      display: grid; gap: 10px; padding: 16px; border-radius: 20px;
      background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.07);
    }}
    .presenter-console-stat strong {{ font-family: {title_font}; font-size: clamp(30px, 2.4vw, 44px); line-height: 1; color: white; }}
    .presenter-console-actions {{
      display: flex; flex-wrap: wrap; gap: 10px;
    }}
    .presenter-console-actions button {{
      border: 0; border-radius: 999px; padding: 11px 16px; font: inherit; font-size: 13px; color: white;
      background: rgba(255,255,255,.12); cursor: pointer; transition: transform .2s ease, background .2s ease;
    }}
    .presenter-console-actions button:hover {{ transform: translateY(-1px); background: rgba(255,255,255,.2); }}
    .presenter-console-actions button.is-active {{ background: var(--accent); }}
    .presenter-console-notes,
    .presenter-console-next-lines {{
      margin: 0; padding-left: 18px; display: grid; gap: 12px; font-size: 16px; line-height: 1.6;
    }}
    .presenter-console-notes li::marker,
    .presenter-console-next-lines li::marker {{ color: var(--accent-soft); }}
    .presenter-console-rail {{
      display: grid; grid-template-rows: auto auto 1fr; gap: 18px;
    }}
    .presenter-console-next-header,
    .presenter-console-status {{
      display: flex; justify-content: space-between; gap: 12px; align-items: baseline;
    }}
    .presenter-console-badge {{
      padding: 7px 10px; border-radius: 999px; background: rgba(255,255,255,.08); font-size: 12px; color: white;
    }}
    .presenter-console-badge.is-live {{ background: rgba(45, 224, 194, .18); color: #d9fff7; }}
    body.presenter-console-window {{ overflow: hidden; }}
    body.presenter-console-window .deck,
    body.presenter-console-window .deck-progress,
    body.presenter-console-window .deck-toolbar,
    body.presenter-console-window .speaker-panel,
    body.presenter-console-window .nav-dots,
    body.presenter-console-window .hint {{ display: none !important; }}
    body.presenter-console-window .presenter-console-app {{ display: grid; }}
    .nav-dots {{ position: fixed; left: 50%; bottom: 24px; transform: translateX(-50%); z-index: 20; display: flex; gap: 10px; padding: 10px 14px; border-radius: 999px; background: rgba(20, 20, 20, .14); backdrop-filter: blur(12px); }}
    .nav-dot {{ width: 10px; height: 10px; border: 0; border-radius: 999px; background: rgba(255,255,255,.45); cursor: pointer; transition: transform .2s ease, width .2s ease, background .2s ease; }}
    .nav-dot.active {{ width: 26px; background: var(--accent); }}
    .hint {{ position: fixed; right: 28px; bottom: 28px; z-index: 20; font-size: 12px; color: var(--muted); letter-spacing: .08em; }}
    body.chrome-hidden .deck-toolbar,
    body.chrome-hidden .nav-dots,
    body.chrome-hidden .hint,
    body.chrome-hidden .deck-progress {{ opacity: 0; pointer-events: none; }}
    body.presenter-mode .hint {{ right: calc(min(30vw, 380px) + 24px); }}
    @media (max-width: 900px) {{
      .web-slide {{ padding: 24px 24px 76px; }}
      .web-slide::before {{ inset: 16px; }}
      .hero-layout, .two-col, .content-split, .closing-layout, .agenda-list, .stats-grid, .timeline-grid {{ grid-template-columns: 1fr; }}
      .quote-block {{ padding: 0 12px; }}
      body.presenter-mode .deck {{ width: 100vw; }}
      .speaker-panel {{ width: min(92vw, 420px); }}
      .speaker-stats {{ grid-template-columns: 1fr; }}
      .presenter-console-shell,
      .presenter-console-grid {{ grid-template-columns: 1fr; }}
      body.presenter-mode .hint {{ right: 20px; }}
    }}
  </style>
</head>
<body>
  <div class="deck-progress"><div id="deckProgressBar" class="deck-progress-bar"></div></div>
  <div class="deck-toolbar">
    <div class="deck-status">
      <strong id="deckStatusIndex">01 / {max(1, len(slides)):02d}</strong>
      <span id="deckStatusLabel">{web_text(plan.get('deck_title') or 'Web Deck')}</span>
    </div>
    <button id="presenterToggle" type="button">Presenter</button>
    <button id="presenterWindowToggle" type="button">Console</button>
    <button id="notesToggle" type="button">Notes</button>
    <button id="chromeToggle" type="button">Hide UI</button>
    <button id="fullscreenToggle" type="button">Fullscreen</button>
  </div>
  <main id="deck" class="deck">
    {slide_markup}
  </main>
  <aside id="speakerPanel" class="speaker-panel" aria-live="polite">
    <div class="speaker-panel-header">
      <strong>Speaker notes</strong>
      <h3 id="speakerTitle">{web_text(plan.get('deck_title') or 'Web Deck')}</h3>
      <p id="speakerMeta">slide</p>
    </div>
    <div class="speaker-stats">
      <section class="speaker-stat">
        <span>Clock</span>
        <strong id="speakerClock">--:--</strong>
      </section>
      <section class="speaker-stat">
        <span>Timer</span>
        <div class="speaker-timer-meta">
          <strong id="speakerTimer">00:00</strong>
          <div class="speaker-timer-actions">
            <button id="timerStartToggle" type="button">Start</button>
            <button id="timerReset" type="button">Reset</button>
          </div>
        </div>
      </section>
    </div>
    <section class="next-slide-card" aria-live="polite">
      <div class="next-slide-header">
        <strong>Next slide</strong>
        <span id="nextSlideMeta">End of deck</span>
      </div>
      <h4 id="nextSlideTitle">No next slide</h4>
      <ul id="nextSlideLines" class="next-slide-lines"></ul>
    </section>
    <ol id="speakerNotesList" class="speaker-notes"></ol>
    <div class="speaker-help">
      <span><code>Left / Right</code><em>navigate</em></span>
      <span><code>F</code><em>fullscreen</em></span>
      <span><code>P</code><em>presenter mode</em></span>
      <span><code>O</code><em>open console</em></span>
      <span><code>N</code><em>notes</em></span>
      <span><code>H</code><em>hide UI</em></span>
      <span><code>T / R</code><em>timer</em></span>
    </div>
  </aside>
  <section id="presenterConsoleApp" class="presenter-console-app" aria-live="polite">
    <div class="presenter-console-shell">
      <article class="presenter-console-card">
        <div class="presenter-console-top">
          <div>
            <strong id="presenterConsoleDeckLabel">{web_text(plan.get('deck_title') or 'Web Deck')}</strong>
            <h2 id="presenterConsoleCurrentTitle">{web_text(plan.get('deck_title') or 'Web Deck')}</h2>
            <p id="presenterConsoleCurrentMeta">01 / {max(1, len(slides)):02d} | slide</p>
          </div>
          <span id="presenterConsoleConnection" class="presenter-console-badge">Waiting</span>
        </div>
        <div class="presenter-console-grid">
          <section class="presenter-console-stat">
            <span>Clock</span>
            <strong id="presenterConsoleClock">--:--</strong>
          </section>
          <section class="presenter-console-stat">
            <span>Timer</span>
            <strong id="presenterConsoleTimer">00:00</strong>
          </section>
        </div>
        <div class="presenter-console-actions">
          <button id="presenterConsolePrev" type="button">Previous</button>
          <button id="presenterConsoleNext" type="button">Next</button>
          <button id="presenterConsoleTimerToggle" type="button">Start</button>
          <button id="presenterConsoleTimerReset" type="button">Reset</button>
          <button id="presenterConsoleFullscreen" type="button">Fullscreen</button>
        </div>
        <section class="presenter-console-current">
          <span class="presenter-console-kicker">Current notes</span>
          <h3 id="presenterConsoleNotesTitle">{web_text(plan.get('deck_title') or 'Web Deck')}</h3>
          <ul id="presenterConsoleNotes" class="presenter-console-notes"></ul>
        </section>
      </article>
      <aside class="presenter-console-rail">
        <section class="presenter-console-card presenter-console-next">
          <div class="presenter-console-next-header">
            <strong>Next slide</strong>
            <span id="presenterConsoleNextMeta">End of deck</span>
          </div>
          <h3 id="presenterConsoleNextTitle">No next slide</h3>
          <ul id="presenterConsoleNextLines" class="presenter-console-next-lines"></ul>
        </section>
        <section class="presenter-console-card presenter-console-status">
          <div>
            <strong>Display</strong>
            <p id="presenterConsoleStatus">Open this window on your presenter screen.</p>
          </div>
          <span id="presenterConsoleLiveBadge" class="presenter-console-badge">Local</span>
        </section>
      </aside>
    </div>
  </section>
  <div class="nav-dots">{dot_markup}</div>
  <div class="hint">Arrow keys / scroll / swipe / F / P / O / N / H / T / R</div>
  <script>
    const deck = document.getElementById("deck");
    const isPresenterConsole = new URLSearchParams(window.location.search).get("presenter") === "1";
    const dots = Array.from(document.querySelectorAll(".nav-dot"));
    const speakerNotes = {notes_payload};
    const progressBar = document.getElementById("deckProgressBar");
    const deckStatusIndex = document.getElementById("deckStatusIndex");
    const deckStatusLabel = document.getElementById("deckStatusLabel");
    const speakerTitle = document.getElementById("speakerTitle");
    const speakerMeta = document.getElementById("speakerMeta");
    const speakerClock = document.getElementById("speakerClock");
    const speakerTimer = document.getElementById("speakerTimer");
    const speakerNotesList = document.getElementById("speakerNotesList");
    const nextSlideTitle = document.getElementById("nextSlideTitle");
    const nextSlideMeta = document.getElementById("nextSlideMeta");
    const nextSlideLines = document.getElementById("nextSlideLines");
    const presenterToggle = document.getElementById("presenterToggle");
    const presenterWindowToggle = document.getElementById("presenterWindowToggle");
    const notesToggle = document.getElementById("notesToggle");
    const chromeToggle = document.getElementById("chromeToggle");
    const fullscreenToggle = document.getElementById("fullscreenToggle");
    const timerStartToggle = document.getElementById("timerStartToggle");
    const timerReset = document.getElementById("timerReset");
    const presenterConsoleConnection = document.getElementById("presenterConsoleConnection");
    const presenterConsoleDeckLabel = document.getElementById("presenterConsoleDeckLabel");
    const presenterConsoleCurrentTitle = document.getElementById("presenterConsoleCurrentTitle");
    const presenterConsoleCurrentMeta = document.getElementById("presenterConsoleCurrentMeta");
    const presenterConsoleClock = document.getElementById("presenterConsoleClock");
    const presenterConsoleTimer = document.getElementById("presenterConsoleTimer");
    const presenterConsoleNotesTitle = document.getElementById("presenterConsoleNotesTitle");
    const presenterConsoleNotes = document.getElementById("presenterConsoleNotes");
    const presenterConsolePrev = document.getElementById("presenterConsolePrev");
    const presenterConsoleNext = document.getElementById("presenterConsoleNext");
    const presenterConsoleTimerToggle = document.getElementById("presenterConsoleTimerToggle");
    const presenterConsoleTimerReset = document.getElementById("presenterConsoleTimerReset");
    const presenterConsoleFullscreen = document.getElementById("presenterConsoleFullscreen");
    const presenterConsoleNextMeta = document.getElementById("presenterConsoleNextMeta");
    const presenterConsoleNextTitle = document.getElementById("presenterConsoleNextTitle");
    const presenterConsoleNextLines = document.getElementById("presenterConsoleNextLines");
    const presenterConsoleStatus = document.getElementById("presenterConsoleStatus");
    const presenterConsoleLiveBadge = document.getElementById("presenterConsoleLiveBadge");
    let timerRunning = false;
    let timerElapsedMs = 0;
    let timerStartedAt = 0;
    let presenterWindowRef = null;
    function goTo(index) {{
      deck.scrollTo({{ left: window.innerWidth * index, behavior: "smooth" }});
    }}
    function currentIndex() {{
      return Math.max(0, Math.min(dots.length - 1, Math.round(deck.scrollLeft / window.innerWidth)));
    }}
    function formatDuration(totalMs) {{
      const totalSeconds = Math.max(0, Math.floor(totalMs / 1000));
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      if (hours > 0) {{
        return `${{String(hours).padStart(2, "0")}}:${{String(minutes).padStart(2, "0")}}:${{String(seconds).padStart(2, "0")}}`;
      }}
      return `${{String(minutes).padStart(2, "0")}}:${{String(seconds).padStart(2, "0")}}`;
    }}
    function updateClock() {{
      speakerClock.textContent = new Date().toLocaleTimeString([], {{ hour: "2-digit", minute: "2-digit" }});
    }}
    function currentElapsedMs() {{
      return timerRunning ? timerElapsedMs + (Date.now() - timerStartedAt) : timerElapsedMs;
    }}
    function syncTimerUi() {{
      speakerTimer.textContent = formatDuration(currentElapsedMs());
      timerStartToggle.textContent = timerRunning ? "Pause" : "Start";
      timerStartToggle.classList.toggle("is-active", timerRunning);
    }}
    function toggleTimer(force) {{
      const next = typeof force === "boolean" ? force : !timerRunning;
      if (next === timerRunning) {{
        syncTimerUi();
        return;
      }}
      if (next) {{
        timerStartedAt = Date.now();
      }} else {{
        timerElapsedMs += Date.now() - timerStartedAt;
        timerStartedAt = 0;
      }}
      timerRunning = next;
      syncTimerUi();
      publishPresenterState();
    }}
    function resetTimerValue() {{
      timerRunning = false;
      timerElapsedMs = 0;
      timerStartedAt = 0;
      syncTimerUi();
      publishPresenterState();
    }}
    function presenterWindowOpen() {{
      return Boolean(presenterWindowRef && !presenterWindowRef.closed);
    }}
    function syncPresenterWindowButton() {{
      presenterWindowToggle.classList.toggle("is-active", presenterWindowOpen());
      presenterWindowToggle.textContent = presenterWindowOpen() ? "Console On" : "Console";
    }}
    function buildPresenterState() {{
      const index = currentIndex();
      const note = speakerNotes[index] || {{ title: "Slide", type: "slide", lines: [] }};
      const next = speakerNotes[index + 1] || null;
      return {{
        deckTitle: "{web_text(plan.get('deck_title') or 'Web Deck')}",
        index,
        total: dots.length,
        note,
        next,
        clock: new Date().toLocaleTimeString([], {{ hour: "2-digit", minute: "2-digit" }}),
        timerLabel: formatDuration(currentElapsedMs()),
        timerRunning,
        fullscreen: Boolean(document.fullscreenElement),
      }};
    }}
    function renderPresenterConsole(state) {{
      if (!presenterConsoleCurrentTitle) return;
      const safeState = state || buildPresenterState();
      presenterConsoleDeckLabel.textContent = safeState.deckTitle || "{web_text(plan.get('deck_title') or 'Web Deck')}";
      presenterConsoleCurrentTitle.textContent = safeState.note?.title || safeState.deckTitle || "Slide";
      presenterConsoleCurrentMeta.textContent = `${{String((safeState.index || 0) + 1).padStart(2, "0")}} / ${{String(Math.max(safeState.total || 1, 1)).padStart(2, "0")}} | ${{safeState.note?.type || "slide"}}`;
      presenterConsoleClock.textContent = safeState.clock || "--:--";
      presenterConsoleTimer.textContent = safeState.timerLabel || "00:00";
      presenterConsoleNotesTitle.textContent = safeState.note?.title || "Current slide";
      presenterConsoleNotes.innerHTML = "";
      (safeState.note?.lines?.length ? safeState.note.lines : ["No notes for this slide."]).forEach((line) => {{
        const item = document.createElement("li");
        item.textContent = line;
        presenterConsoleNotes.appendChild(item);
      }});
      presenterConsoleNextLines.innerHTML = "";
      if (safeState.next) {{
        presenterConsoleNextMeta.textContent = `${{String((safeState.index || 0) + 2).padStart(2, "0")}} / ${{String(Math.max(safeState.total || 1, 1)).padStart(2, "0")}} | ${{safeState.next.type || "slide"}}`;
        presenterConsoleNextTitle.textContent = safeState.next.title || "Next slide";
        (safeState.next.lines?.length ? safeState.next.lines.slice(0, 3) : ["Advance when you are ready."]).forEach((line) => {{
          const item = document.createElement("li");
          item.textContent = line;
          presenterConsoleNextLines.appendChild(item);
        }});
      }} else {{
        presenterConsoleNextMeta.textContent = "End of deck";
        presenterConsoleNextTitle.textContent = "No next slide";
        const item = document.createElement("li");
        item.textContent = "You are on the final slide.";
        presenterConsoleNextLines.appendChild(item);
      }}
      presenterConsoleTimerToggle.textContent = safeState.timerRunning ? "Pause" : "Start";
      presenterConsoleTimerToggle.classList.toggle("is-active", Boolean(safeState.timerRunning));
      presenterConsoleFullscreen.classList.toggle("is-active", Boolean(safeState.fullscreen));
    }}
    function publishPresenterState() {{
      if (isPresenterConsole) return;
      const state = buildPresenterState();
      window.__pptMasterDeckState = state;
      renderPresenterConsole(state);
      if (presenterWindowOpen()) {{
        presenterWindowRef.postMessage({{ type: "pptmaster-deck-state", state }}, "*");
      }}
      syncPresenterWindowButton();
    }}
    function executeDeckCommand(command) {{
      if (!command || typeof command !== "object") return;
      if (command.type === "goto" && Number.isFinite(command.index)) {{
        goTo(Math.max(0, Math.min(Number(command.index), dots.length - 1)));
        return;
      }}
      if (command.type === "next") {{
        goTo(Math.min(currentIndex() + 1, dots.length - 1));
        return;
      }}
      if (command.type === "prev") {{
        goTo(Math.max(currentIndex() - 1, 0));
        return;
      }}
      if (command.type === "toggle-timer") {{
        toggleTimer();
        return;
      }}
      if (command.type === "reset-timer") {{
        resetTimerValue();
        return;
      }}
      if (command.type === "fullscreen") {{
        toggleFullscreen();
      }}
    }}
    window.__pptMasterDeckCommand = executeDeckCommand;
    function openPresenterWindow() {{
      if (isPresenterConsole) return;
      const url = new URL(window.location.href);
      url.searchParams.set("presenter", "1");
      presenterWindowRef = window.open(url.toString(), "pptmaster-presenter-console", "popup=yes,width=1480,height=920");
      if (presenterWindowRef) {{
        presenterWindowRef.focus();
        syncPresenterWindowButton();
        window.setTimeout(() => publishPresenterState(), 120);
      }}
    }}
    function requestDeckCommand(command) {{
      if (!isPresenterConsole) return;
      if (window.opener && !window.opener.closed && typeof window.opener.__pptMasterDeckCommand === "function") {{
        window.opener.__pptMasterDeckCommand(command);
      }}
    }}
    function syncPresenterConsoleConnection() {{
      if (!isPresenterConsole) return;
      const connected = Boolean(window.opener && !window.opener.closed && window.opener.__pptMasterDeckState);
      presenterConsoleConnection.textContent = connected ? "Connected" : "Waiting";
      presenterConsoleConnection.classList.toggle("is-live", connected);
      presenterConsoleLiveBadge.textContent = connected ? "Live" : "Offline";
      presenterConsoleLiveBadge.classList.toggle("is-live", connected);
      presenterConsoleStatus.textContent = connected
        ? "Use this window on your presenter display while the deck stays on the audience screen."
        : "Return to the deck window and press Console again to reconnect.";
    }}
    function pullPresenterStateFromOpener() {{
      if (!isPresenterConsole) return;
      if (window.opener && !window.opener.closed && window.opener.__pptMasterDeckState) {{
        renderPresenterConsole(window.opener.__pptMasterDeckState);
      }}
      syncPresenterConsoleConnection();
    }}
    function renderNotes(index) {{
      const note = speakerNotes[index] || {{ title: "Slide", type: "slide", lines: [] }};
      speakerTitle.textContent = note.title || "Slide";
      speakerMeta.textContent = `${{String(index + 1).padStart(2, "0")}} / ${{String(dots.length).padStart(2, "0")}} | ${{note.type || "slide"}}`;
      speakerNotesList.innerHTML = "";
      const lines = (note.lines && note.lines.length) ? note.lines : ["No notes for this slide."];
      lines.forEach((line) => {{
        const item = document.createElement("li");
        item.textContent = line;
        speakerNotesList.appendChild(item);
      }});
    }}
    function renderNextSlide(index) {{
      const next = speakerNotes[index + 1];
      nextSlideLines.innerHTML = "";
      if (!next) {{
        nextSlideTitle.textContent = "No next slide";
        nextSlideMeta.textContent = "End of deck";
        const item = document.createElement("li");
        item.textContent = "You are on the final slide.";
        nextSlideLines.appendChild(item);
        return;
      }}
      nextSlideTitle.textContent = next.title || "Next slide";
      nextSlideMeta.textContent = `${{String(index + 2).padStart(2, "0")}} / ${{String(dots.length).padStart(2, "0")}} | ${{next.type || "slide"}}`;
      const lines = (next.lines && next.lines.length) ? next.lines.slice(0, 3) : ["Advance when you are ready."];
      lines.forEach((line) => {{
        const item = document.createElement("li");
        item.textContent = line;
        nextSlideLines.appendChild(item);
      }});
    }}
    async function toggleFullscreen() {{
      if (document.fullscreenElement) {{
        await document.exitFullscreen();
      }} else {{
        await document.documentElement.requestFullscreen();
      }}
      syncDots();
    }}
    function togglePresenter(force) {{
      const next = typeof force === "boolean" ? force : !document.body.classList.contains("presenter-mode");
      document.body.classList.toggle("presenter-mode", next);
      presenterToggle.classList.toggle("is-active", next);
      if (next) {{
        document.body.classList.add("notes-open");
        notesToggle.classList.add("is-active");
      }}
    }}
    function toggleNotes(force) {{
      const next = typeof force === "boolean" ? force : !document.body.classList.contains("notes-open");
      document.body.classList.toggle("notes-open", next);
      notesToggle.classList.toggle("is-active", next);
    }}
    function toggleChrome() {{
      const next = !document.body.classList.contains("chrome-hidden");
      document.body.classList.toggle("chrome-hidden", next);
      chromeToggle.classList.toggle("is-active", next);
    }}
    function syncDots() {{
      const index = currentIndex();
      dots.forEach((dot, dotIndex) => dot.classList.toggle("active", dotIndex === index));
      progressBar.style.width = `${{((index + 1) / Math.max(dots.length, 1)) * 100}}%`;
      deckStatusIndex.textContent = `${{String(index + 1).padStart(2, "0")}} / ${{String(dots.length).padStart(2, "0")}}`;
      deckStatusLabel.textContent = speakerNotes[index]?.title || "{web_text(plan.get('deck_title') or 'Web Deck')}";
      renderNotes(index);
      renderNextSlide(index);
      fullscreenToggle.classList.toggle("is-active", Boolean(document.fullscreenElement));
    }}
    dots.forEach((dot) => {{
      dot.addEventListener("click", () => goTo(Number(dot.dataset.index)));
    }});
    presenterToggle.addEventListener("click", () => togglePresenter());
    presenterWindowToggle.addEventListener("click", () => openPresenterWindow());
    notesToggle.addEventListener("click", () => toggleNotes());
    chromeToggle.addEventListener("click", () => toggleChrome());
    fullscreenToggle.addEventListener("click", () => toggleFullscreen());
    timerStartToggle.addEventListener("click", () => toggleTimer());
    timerReset.addEventListener("click", () => resetTimerValue());
    presenterConsolePrev.addEventListener("click", () => requestDeckCommand({{ type: "prev" }}));
    presenterConsoleNext.addEventListener("click", () => requestDeckCommand({{ type: "next" }}));
    presenterConsoleTimerToggle.addEventListener("click", () => requestDeckCommand({{ type: "toggle-timer" }}));
    presenterConsoleTimerReset.addEventListener("click", () => requestDeckCommand({{ type: "reset-timer" }}));
    presenterConsoleFullscreen.addEventListener("click", () => requestDeckCommand({{ type: "fullscreen" }}));
    window.addEventListener("message", (event) => {{
      if (event.data?.type === "pptmaster-deck-state") {{
        renderPresenterConsole(event.data.state);
        syncPresenterConsoleConnection();
      }}
    }});
    window.addEventListener("keydown", (event) => {{
      if (isPresenterConsole) {{
        if (event.key === "ArrowRight" || event.key === "PageDown") requestDeckCommand({{ type: "next" }});
        if (event.key === "ArrowLeft" || event.key === "PageUp") requestDeckCommand({{ type: "prev" }});
        if (event.key === "f" || event.key === "F") requestDeckCommand({{ type: "fullscreen" }});
        if (event.key === "t" || event.key === "T") requestDeckCommand({{ type: "toggle-timer" }});
        if (event.key === "r" || event.key === "R") requestDeckCommand({{ type: "reset-timer" }});
        return;
      }}
      const index = currentIndex();
      if (event.key === "ArrowRight" || event.key === "PageDown") goTo(Math.min(index + 1, dots.length - 1));
      if (event.key === "ArrowLeft" || event.key === "PageUp") goTo(Math.max(index - 1, 0));
      if (event.key === "Home") goTo(0);
      if (event.key === "End") goTo(dots.length - 1);
      if (event.key === "f" || event.key === "F") toggleFullscreen();
      if (event.key === "p" || event.key === "P") togglePresenter();
      if (event.key === "o" || event.key === "O") openPresenterWindow();
      if (event.key === "n" || event.key === "N") toggleNotes();
      if (event.key === "h" || event.key === "H") toggleChrome();
      if (event.key === "t" || event.key === "T") toggleTimer();
      if (event.key === "r" || event.key === "R") resetTimerValue();
    }});
    deck.addEventListener("scroll", syncDots, {{ passive: true }});
    window.addEventListener("resize", syncDots);
    document.addEventListener("fullscreenchange", syncDots);
    if (isPresenterConsole) {{
      document.body.classList.add("presenter-console-window");
      pullPresenterStateFromOpener();
      window.setInterval(() => pullPresenterStateFromOpener(), 1000);
    }}
    updateClock();
    syncTimerUi();
    window.setInterval(() => {{
      updateClock();
      syncTimerUi();
      publishPresenterState();
    }}, 1000);
    renderNotes(0);
    renderNextSlide(0);
    syncDots();
    renderPresenterConsole(buildPresenterState());
    publishPresenterState();
  </script>
</body>
</html>"""


def stage_export_web(project_dir: Path, plan: dict, theme: dict, style_key: str, log: LogFn) -> Path:
    exports_dir = project_dir / "exports"
    exports_dir.mkdir(parents=True, exist_ok=True)
    deck_file = web_deck_path(project_dir)
    deck_file.write_text(build_web_deck_html(plan, theme, style_key), encoding="utf-8")
    log(f"Web deck exported: {deck_file.name}")
    return deck_file


def create_draft(request: dict, log: LogFn) -> DraftArtifacts:
    style_key = request.get("styleKey") or "business"
    theme = choose_theme(style_key)
    project_dir = create_project(request, log)
    markdown_text = ingest_source(project_dir, request, log)
    log(f"Source text ready ({len(markdown_text)} chars)")

    planner = request.get("planner") or "heuristic"
    try:
        if planner == "openai":
            plan = build_openai_plan(markdown_text, request, log)
        else:
            plan = build_heuristic_plan(markdown_text, request, log)
    except Exception as exc:
        if planner == "openai":
            log(f"OpenAI planner failed, falling back to heuristic: {exc}")
            plan = build_heuristic_plan(markdown_text, request, log)
        else:
            raise

    plan = apply_style_direction(plan, style_key)
    plan = ensure_slide_ids(plan)
    write_design_files(project_dir, plan, theme)
    slide_svgs = save_slides(project_dir, plan, theme, log)
    save_draft_state(
        project_dir,
        plan,
        theme,
        source_meta={
            "planner": planner,
            "styleKey": style_key,
            "language": request.get("language") or "auto",
            "deckTitle": request.get("deckTitle") or "",
            "exportMode": normalize_export_mode(str(request.get("exportMode") or "pptx")),
        },
    )
    stage_quality_check(project_dir, log)
    return DraftArtifacts(
        project_dir=project_dir,
        slide_svgs=slide_svgs,
        plan=plan,
        theme=theme,
    )


def apply_draft_edits(project_dir: Path, edits: dict, log: LogFn) -> DraftArtifacts:
    state = load_draft_state(project_dir)
    theme = state["theme"]
    plan = ensure_slide_ids(state["plan"])
    slide_map = {slide["id"]: dict(slide) for slide in plan.get("slides") or []}
    requested_slides = edits.get("slides") or []

    if requested_slides:
        seen_ids = set()
        reordered = []
        for item in requested_slides:
            slide_id = item.get("id")
            if slide_id not in slide_map or slide_id in seen_ids:
                continue
            seen_ids.add(slide_id)
            slide = slide_map[slide_id]
            if "title" in item:
                proposed_title = str(item.get("title") or "").strip()
                if proposed_title:
                    slide["title"] = proposed_title
            reordered.append(slide)
        for slide in plan["slides"]:
            if slide["id"] not in seen_ids:
                reordered.append(dict(slide))
    else:
        reordered = [dict(slide) for slide in plan["slides"]]

    for index, slide in enumerate(reordered, start=1):
        slide["index"] = index

    updated_plan = {
        **plan,
        "deck_title": str(edits.get("deckTitle") or plan.get("deck_title") or "").strip() or plan.get("deck_title"),
        "deck_subtitle": str(edits.get("deckSubtitle") or plan.get("deck_subtitle") or "").strip() or plan.get("deck_subtitle"),
        "slides": reordered,
    }
    updated_plan = ensure_slide_ids(updated_plan)
    write_design_files(project_dir, updated_plan, theme)
    slide_svgs = save_slides(project_dir, updated_plan, theme, log)
    save_draft_state(project_dir, updated_plan, theme, state.get("sourceMeta") or {})
    stage_quality_check(project_dir, log)
    return DraftArtifacts(
        project_dir=project_dir,
        slide_svgs=slide_svgs,
        plan=updated_plan,
        theme=theme,
    )


def export_draft(project_dir: Path, log: LogFn, export_mode: str = "pptx") -> GenerationArtifacts:
    state = load_draft_state(project_dir)
    plan = ensure_slide_ids(state["plan"])
    theme = state["theme"]
    source_meta = dict(state.get("sourceMeta") or {})
    style_key = str(source_meta.get("styleKey") or "business")
    export_mode = normalize_export_mode(export_mode)
    if source_meta.get("exportMode") != export_mode:
        source_meta["exportMode"] = export_mode
        save_draft_state(project_dir, plan, theme, source_meta)
    native: Path | None = None
    legacy: Path | None = None
    web_deck: Path | None = None

    if export_mode in {"pptx", "both"}:
        stage_finalize(project_dir, log)
        native, legacy = stage_export(project_dir, log)

    if export_mode in {"web", "both"}:
        web_deck = stage_export_web(project_dir, plan, theme, style_key, log)

    exported = [path.name for path in [native, legacy, web_deck] if path is not None]
    log(f"Export complete: {', '.join(exported)}")
    return GenerationArtifacts(
        project_dir=project_dir,
        native_pptx=native,
        legacy_pptx=legacy,
        web_deck=web_deck,
        slide_svgs=list_slide_svgs(project_dir, "svg_output"),
        plan=plan,
    )


def retry_step(project_dir: Path, step: str, log: LogFn, export_mode: str | None = None) -> dict:
    state = load_draft_state(project_dir)
    plan = ensure_slide_ids(state["plan"])
    theme = state["theme"]

    if step == "render":
        write_design_files(project_dir, plan, theme)
        slide_svgs = save_slides(project_dir, plan, theme, log)
        stage_quality_check(project_dir, log)
        return {
            "projectDir": project_dir,
            "slides": slide_svgs,
            "plan": plan,
        }

    if step == "export":
        selected_mode = export_mode or str((state.get("sourceMeta") or {}).get("exportMode") or "pptx")
        artifacts = export_draft(project_dir, log, export_mode=selected_mode)
        return {
            "projectDir": artifacts.project_dir,
            "nativePptx": artifacts.native_pptx,
            "legacyPptx": artifacts.legacy_pptx,
            "webDeck": artifacts.web_deck,
            "slides": artifacts.slide_svgs,
            "plan": artifacts.plan,
        }

    raise GenerationError(f"Unsupported retry step: {step}")


def generate_deck(request: dict, log: LogFn) -> GenerationArtifacts:
    draft = create_draft(request, log)
    return export_draft(draft.project_dir, log)
