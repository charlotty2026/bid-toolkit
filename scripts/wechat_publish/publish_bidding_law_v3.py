#!/usr/bin/env python3
"""WeChat draft publisher v3.0 - Silicon Design style with column template support.

v3.0 (2026-08-12): Added three-layer pipeline support.
  - Added --column parameter for template binding
  - Added --output-html mode for pipeline integration
  - Wraps sections in card containers for card-based columns (jianghuyehua/baiguiyexing etc.)
  - Backward compatible with v2.3 behavior

v2.3 (2026-08-05): Fixed full-width parens/semicolons being incorrectly
  converted to half-width in Chinese body text.
v2.2 (2026-08-05): Fixed quote direction in normalize_quotes().
"""

import json
import re
import requests
import sys
import os
import argparse
from pathlib import Path


# ============================================================
# Configuration Loading
# ============================================================
def _load_env(path="/opt/zongmen/.env"):
    creds = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    creds[k.strip()] = v.strip()
    return creds


_creds = _load_env()
APP_ID = _creds.get("APP_ID", os.environ.get("WECHAT_APP_ID", ""))
APP_SECRET = _creds.get("APP_SECRET", os.environ.get("WECHAT_APP_SECRET", ""))

# Column template configuration
SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = SCRIPT_DIR / "column_templates.json"


def _load_column_template(column_name):
    """Load column template configuration."""
    if not CONFIG_PATH.exists():
        return None
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("templates", {}).get(column_name)
    except Exception:
        return None


def _get_template_style_config(column_name):
    """Get template style configuration for HTML generation."""
    template = _load_column_template(column_name)
    if not template:
        return None
    return {
        "color_scheme": template.get("color_scheme", {}),
        "layout_rules": template.get("layout_rules", {}),
        "typography": template.get("typography", {}),
        "style": template.get("style", "border_card"),
    }


# Color palette (default, overridden by template)
BLUE = "#4A90E2"
LIGHT_BLUE_BG = "#F0F7FF"
RED = "#E8534F"
GRAY_HEADER = "#f0f0f0"
GRAY_TEXT = "#666666"
DARK_TEXT = "#333333"


# ============================================================
# Character Hygiene: Full-width <-> Half-width normalization
# ============================================================

# Full-width punctuation that is ALWAYS wrong in markdown syntax
_FULLWIDTH_ALWAYS = {
    "\uff5b": "{",  # Fullwidth left brace
    "\uff5d": "}",  # Fullwidth right brace
    "\uff3b": "[",  # Fullwidth left bracket
    "\uff3d": "]",  # Fullwidth right bracket
    "\uff5c": "|",  # Fullwidth vertical bar
    "\uff5e": "~",  # Fullwidth tilde
    "\uff0b": "+",  # Fullwidth plus
    "\uff0d": "-",  # Fullwidth hyphen-minus
    "\uff0f": "/",  # Fullwidth solidus
    "\uff1d": "=",  # Fullwidth equals sign
    "\uff05": "%",  # Fullwidth percent sign
    "\uff06": "&",  # Fullwidth ampersand
    "\uff0a": "*",  # Fullwidth asterisk
    "\uff03": "#",  # Fullwidth number sign
    "\uff40": "`",  # Fullwidth grave accent
    "\uff3c": "\\",  # Fullwidth reverse solidus
    "\uff3e": "^",  # Fullwidth circumflex
    "\uff3f": "_",  # Fullwidth low line
    "\uffe5": "&yen;",  # Fullwidth yen sign -> HTML entity
    "\uff04": "$",  # Fullwidth dollar sign
    "\uff20": "@",  # Fullwidth commercial at
    "\uffe0": "&cent;",  # Fullwidth cent sign
    "\uffe1": "&pound;",  # Fullwidth pound sign
}

# Full-width punctuation that is correct in Chinese body text but must be
# forced to half-width in code blocks, inline code, and URLs.
_FULLWIDTH_CODE_ONLY = {
    "\uff08": "(",  # Fullwidth left parenthesis
    "\uff09": ")",  # Fullwidth right parenthesis
    "\uff1b": ";",  # Fullwidth semicolon
}

# Full-width space
_FULLWIDTH_SPACE = "\u3000"


def _normalize_straight_quotes(text):
    result = []
    double_open = True
    single_open = True
    for ch in text:
        if ch == '"':
            if double_open:
                result.append("\u201c")
            else:
                result.append("\u201d")
            double_open = not double_open
        elif ch == "'":
            if single_open:
                result.append("\u2018")
            else:
                result.append("\u2019")
            single_open = not single_open
        else:
            result.append(ch)
    return "".join(result)


def _force_halfwidth_quotes(text):
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\uff02", '"').replace("\uff07", "'")
    return text


def normalize_quotes(text):
    if not text:
        return text

    text = text.replace(_FULLWIDTH_SPACE, " ")

    for i in range(10):
        text = text.replace(chr(0xFF10 + i), str(i))

    code_block_pattern = re.compile(r"(```[\s\S]*?```)", re.MULTILINE)
    parts = code_block_pattern.split(text)

    result_parts = []
    for part in parts:
        if part.startswith("```"):
            for fw, hw in _FULLWIDTH_ALWAYS.items():
                part = part.replace(fw, hw)
            for fw, hw in _FULLWIDTH_CODE_ONLY.items():
                part = part.replace(fw, hw)
            part = _force_halfwidth_quotes(part)
            part = part.replace("\uff0c", ",").replace("\uff1a", ":")
            part = part.replace("\uff1f", "?").replace("\uff01", "!")
            result_parts.append(part)
        else:
            for fw, hw in _FULLWIDTH_ALWAYS.items():
                part = part.replace(fw, hw)

            inline_code_pattern = re.compile(r"(`[^`]+`)")
            sub_parts = inline_code_pattern.split(part)
            processed_sub = []
            for sp in sub_parts:
                if sp.startswith("`") and sp.endswith("`") and len(sp) > 1:
                    sp = _force_halfwidth_quotes(sp)
                    for fw, hw in _FULLWIDTH_CODE_ONLY.items():
                        sp = sp.replace(fw, hw)
                    processed_sub.append(sp)
                else:

                    def fix_link(m):
                        prefix = m.group(1)
                        text_part = m.group(2)
                        url_part = m.group(3)
                        url_part = _force_halfwidth_quotes(url_part)
                        for fw, hw in _FULLWIDTH_CODE_ONLY.items():
                            url_part = url_part.replace(fw, hw)
                        return f"{prefix}[{text_part}]({url_part})"

                    sp = re.sub(
                        r"(!?)\[([^\]]+)\]\(([^)]+)\)",
                        fix_link,
                        sp,
                    )
                    sp = _normalize_straight_quotes(sp)
                    processed_sub.append(sp)
            part = "".join(processed_sub)
            result_parts.append(part)

    text = "".join(result_parts)
    text = re.sub(r" {3,}", "  ", text)
    text = text.replace("\u200b", "")
    text = text.replace("\u200c", "")
    text = text.replace("\u200d", "")
    text = text.replace("\ufeff", "")

    return text


# ============================================================
# HTML Rendering Functions
# ============================================================
def render_code_block(code, lang=""):
    lines = code.strip().split("\n")
    rendered_lines = []
    for line in lines:
        line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        rendered_lines.append(
            f'<span style="display:block;line-height:1.6;">{line if line.strip() else "&nbsp;"}</span>'
        )
    return (
        f'<section style="background:#1e1e1e;color:#d4d4d4;padding:12px 16px;'
        f"border-radius:6px;margin:16px 0;font-family:Consolas,Monaco,monospace;"
        f'font-size:13px;overflow-x:auto;">'
        f"{''.join(rendered_lines)}"
        f"</section>"
    )


def render_table(header_row, separator_row, data_rows):
    def parse_cells(row):
        cells = [c.strip() for c in row.strip().strip("|").split("|")]
        return cells

    headers = parse_cells(header_row)
    body_rows = [parse_cells(r) for r in data_rows if r.strip()]

    html = f'<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:14px;">'

    html += "<thead><tr>"
    for h in headers:
        h = process_inline(h)
        html += (
            f'<th style="border:1px solid {BLUE};'
            f"background:{BLUE};color:#ffffff;"
            f'padding:8px 12px;text-align:left;font-weight:bold;">'
            f"{h}</th>"
        )
    html += "</tr></thead>"

    html += "<tbody>"
    for row in body_rows:
        html += "<tr>"
        for cell in row:
            cell = process_inline(cell)
            html += (
                f'<td style="border:1px solid #e0e0e0;padding:8px 12px;'
                f'color:{DARK_TEXT};line-height:1.6;">{cell}</td>'
            )
        html += "</tr>"
    html += "</tbody></table>"
    return html


def process_inline(text):
    text = re.sub(
        r"\*\*([^*]+)\*\*",
        rf'<strong style="color:{RED};font-weight:bold;">\1</strong>',
        text,
    )
    text = re.sub(
        r"`([^`]+)`",
        r'<code style="background:#f0f0f0;padding:2px 6px;border-radius:3px;'
        r'font-family:Consolas,Monaco,monospace;font-size:13px;color:#c7254e;">\1</code>',
        text,
    )
    text = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        rf'<a href="\2" style="color:{BLUE};text-decoration:none;">\1</a>',
        text,
    )
    text = re.sub(
        r"\*([^*]+)\*", rf'<em style="color:{GRAY_TEXT};font-size:13px;">\1</em>', text
    )
    return text


# ============================================================
# Card Container Wrapping (Template-aware)
# ============================================================
def _get_card_style(template_config):
    """Generate card container style from template config."""
    if not template_config:
        return ""
    layout = template_config.get("layout_rules", {})
    colors = template_config.get("color_scheme", {})

    if not layout.get("card_border", False):
        return None  # No cards for this template

    radius = layout.get("card_radius", "8px")
    padding = layout.get("card_padding", "16px")
    margin = layout.get("card_margin", "12px 0")
    shadow = layout.get("card_shadow", "0 2px 8px rgba(0,0,0,0.06)")
    bg = colors.get("card_bg", "#FFFFFF")
    border_color = colors.get("border", "#E0E0E0")

    return (
        f"border:1px solid {border_color};"
        f"border-radius:{radius};"
        f"padding:{padding};"
        f"margin:{margin};"
        f"background:{bg};"
        f"box-shadow:{shadow};"
    )


def wrap_sections_in_cards(html_parts, column):
    """Wrap content sections in card containers for card-based columns."""
    template_config = _get_template_style_config(column)
    card_style = _get_card_style(template_config)

    if not card_style:
        # No card wrapping needed
        return html_parts

    # Group content between h2 sections into cards
    result = []
    current_card = []
    in_card = False

    for part in html_parts:
        # Check if this is an h2 heading (section boundary)
        is_h2 = part.startswith("<h2 ")

        if is_h2:
            # Close previous card if exists
            if current_card and in_card:
                card_html = f'<div style="{card_style}">{"".join(current_card)}</div>'
                result.append(card_html)
                current_card = []
            # Add h2 directly (not inside card, or inside its own card)
            # Option: put h2 + following content in the same card
            current_card.append(part)
            in_card = True
        else:
            if in_card:
                current_card.append(part)
            else:
                # Before first h2 or non-card content
                result.append(part)

    # Close final card
    if current_card and in_card:
        card_html = f'<div style="{card_style}">{"".join(current_card)}</div>'
        result.append(card_html)

    return result


def md_to_html(md_text, column=None):
    """Convert markdown to WeChat-compatible HTML with column template support."""
    md_text = normalize_quotes(md_text)

    # Load template for styling overrides
    template_config = _get_template_style_config(column) if column else None

    lines = md_text.split("\n")
    html_parts = []
    in_code_block = False
    code_buffer = []
    code_lang = ""

    i = 0
    while i < len(lines):
        line = lines[i]

        if line.strip().startswith("```"):
            if in_code_block:
                html_parts.append(render_code_block("\n".join(code_buffer), code_lang))
                code_buffer = []
                in_code_block = False
            else:
                in_code_block = True
                code_lang = line.strip()[3:].strip()
            i += 1
            continue

        if in_code_block:
            code_buffer.append(line)
            i += 1
            continue

        # H1
        if line.startswith("# ") and not line.startswith("## "):
            title = line[2:].strip()
            html_parts.append(
                f'<h1 style="font-size:20px;font-weight:bold;color:{DARK_TEXT};'
                f'text-align:center;margin:24px 0 16px 0;line-height:1.4;">'
                f"{title}</h1>"
            )
            i += 1
            continue

        # H2
        if line.startswith("## "):
            title = line[3:].strip()
            html_parts.append(
                f'<h2 style="font-size:17px;font-weight:bold;color:{DARK_TEXT};'
                f"border-left:4px solid {BLUE};padding-left:12px;"
                f'margin:28px 0 14px 0;line-height:1.4;">'
                f"{title}</h2>"
            )
            i += 1
            continue

        # H3
        if line.startswith("### "):
            title = line[4:].strip()
            html_parts.append(
                f'<h3 style="font-size:15px;font-weight:bold;color:{DARK_TEXT};'
                f"border-left:3px solid {BLUE};padding-left:10px;"
                f'margin:20px 0 10px 0;line-height:1.4;">'
                f"{title}</h3>"
            )
            i += 1
            continue

        # Horizontal rule
        if line.strip() == "---":
            html_parts.append(
                f'<p style="text-align:center;color:{BLUE};font-size:14px;'
                f'letter-spacing:8px;margin:20px 0;">\u00b7 \u00b7 \u00b7</p>'
            )
            i += 1
            continue

        # Blockquote
        if line.strip().startswith(">"):
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                content = lines[i].strip()[1:].strip()
                if content:
                    quote_lines.append(content)
                i += 1
            quote_html = ""
            for ql in quote_lines:
                ql = process_inline(ql)
                quote_html += f'<p style="margin:4px 0;padding:0;">{ql}</p>'
            html_parts.append(
                f'<blockquote style="background:{LIGHT_BLUE_BG};'
                f"border-left:4px solid {BLUE};padding:12px 16px;"
                f"margin:16px 0;font-size:14px;color:{GRAY_TEXT};"
                f'line-height:1.6;border-radius:0 4px 4px 0;">'
                f"{quote_html}</blockquote>"
            )
            continue

        # Markdown table
        if "|" in line and line.strip().startswith("|"):
            table_lines = []
            while (
                i < len(lines) and "|" in lines[i] and lines[i].strip().startswith("|")
            ):
                table_lines.append(lines[i])
                i += 1
            if len(table_lines) >= 2:
                header = table_lines[0]
                separator = table_lines[1]
                data = table_lines[2:] if len(table_lines) > 2 else []
                html_parts.append(render_table(header, separator, data))
                continue
            else:
                for tl in table_lines:
                    html_parts.append(
                        f'<p style="font-size:15px;color:{DARK_TEXT};'
                        f'line-height:1.8;margin:10px 0;">{process_inline(tl.strip())}</p>'
                    )
                continue

        # List items
        if line.strip().startswith("- "):
            item = line.strip()[2:].strip()
            item = process_inline(item)
            html_parts.append(
                f'<p style="font-size:15px;color:{DARK_TEXT};'
                f'line-height:1.8;padding-left:20px;margin:6px 0;">'
                f'<span style="color:{BLUE};">\u25b8</span> {item}</p>'
            )
            i += 1
            continue

        # Empty line
        if not line.strip():
            i += 1
            continue

        # Regular paragraph
        content = line.strip()
        content = process_inline(content)
        html_parts.append(
            f'<p style="font-size:15px;color:{DARK_TEXT};line-height:1.8;'
            f'margin:10px 0;text-align:justify;">{content}</p>'
        )
        i += 1

    # Apply card wrapping for card-based templates
    if column:
        html_parts = wrap_sections_in_cards(html_parts, column)

    return "".join(html_parts)


# ============================================================
# WeChat API Functions
# ============================================================
def get_access_token():
    url = f"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={APP_ID}&secret={APP_SECRET}"
    resp = requests.get(url, timeout=10)
    data = resp.json()
    if "access_token" not in data:
        print(f"ERROR getting token: {data}")
        sys.exit(1)
    return data["access_token"]


def upload_thumb(token, image_path):
    url = f"https://api.weixin.qq.com/cgi-bin/material/add_material?access_token={token}&type=thumb"
    with open(image_path, "rb") as f:
        files = {"media": f}
        resp = requests.post(url, files=files, timeout=30)
    data = resp.json()
    if "media_id" not in data:
        print(f"ERROR uploading thumb: {data}")
        sys.exit(1)
    print(f"Thumb uploaded: {data['media_id']}")
    return data["media_id"]


def create_draft(token, title, content_html, thumb_media_id):
    url = f"https://api.weixin.qq.com/cgi-bin/draft/add?access_token={token}"

    article = {
        "title": title,
        "author": "",
        "digest": "",
        "content": content_html,
        "thumb_media_id": thumb_media_id,
        "need_open_comment": 0,
        "only_fans_can_comment": 0,
    }

    payload = {"articles": [article]}

    resp = requests.post(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=utf-8"},
        timeout=30,
    )
    data = resp.json()

    if "media_id" not in data:
        print(f"ERROR creating draft: {data}")
        sys.exit(1)

    return data["media_id"]


# ============================================================
# Main Entry Point (with argparse for pipeline support)
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="WeChat draft publisher v3.0")
    parser.add_argument(
        "md_path",
        nargs="?",
        default="/tmp/2026-08-03_ZHUAN_XIANG_GONG_LUE_SHEN_HE_BAN.md",
        help="Path to markdown file",
    )
    parser.add_argument(
        "cover_path",
        nargs="?",
        default="/tmp/cover_bidding_law_compressed.jpg",
        help="Path to cover image",
    )
    parser.add_argument(
        "--column", default="江湖夜话", help="WeChat column name for template binding"
    )
    parser.add_argument(
        "--output-html",
        default=None,
        help="Output HTML to file instead of pushing to WeChat (for pipeline)",
    )
    parser.add_argument("--title-override", default=None, help="Override article title")
    args = parser.parse_args()

    if not os.path.exists(args.md_path):
        print(f"ERROR: Markdown file not found: {args.md_path}")
        sys.exit(1)

    with open(args.md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    # Extract title (first H1)
    if args.title_override:
        title = args.title_override
    else:
        title_match = re.match(r"^#\s+(.+)$", md_text, re.MULTILINE)
        title = title_match.group(1).strip() if title_match else ""
        if not title:
            title = "2026_ZHAO_BIAO_FA_DA_XIU_LUO_DI"

    print(f"Title: {title}")
    print(f"Column: {args.column}")

    # Convert to HTML with column template
    html_content = md_to_html(md_text, args.column)
    print(f"HTML length: {len(html_content)} chars")

    # If output-html mode, write to file and exit
    if args.output_html:
        with open(args.output_html, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"HTML written to: {args.output_html}")
        sys.exit(0)

    # Normal mode: push to WeChat
    if not os.path.exists(args.cover_path):
        print(f"ERROR: Cover image not found: {args.cover_path}")
        sys.exit(1)

    # Get access token
    token = get_access_token()
    print(f"Token obtained: {token[:20]}...")

    # Upload cover as thumb
    thumb_id = upload_thumb(token, args.cover_path)

    # Create draft
    draft_id = create_draft(token, title, html_content, thumb_id)
    print(f"\n=== SUCCESS ===")
    print(f"Draft media_id: {draft_id}")
    print(f"Title: {title}")
    print(f"Column: {args.column}")


if __name__ == "__main__":
    main()
