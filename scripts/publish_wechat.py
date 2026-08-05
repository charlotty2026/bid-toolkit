#!/usr/bin/env python3
"""WeChat draft publisher - Silicon Design style for bidding law article.

v2.3 (2026-08-05): Fixed full-width parens/semicolons being incorrectly
  converted to half-width in Chinese body text.
  - Moved （）； from _FULLWIDTH_ALWAYS to new _FULLWIDTH_CODE_ONLY map.
  - _FULLWIDTH_ALWAYS: applies everywhere (braces, brackets, etc.)
  - _FULLWIDTH_CODE_ONLY: applies only in code blocks, inline code, URLs.
  - Chinese body text now preserves （）； as correct typography.

v2.2 (2026-08-05): Fixed quote direction in normalize_quotes().
  - Body text: ASCII straight quotes " ' -> Chinese full-width curly quotes ""''
  - Code blocks/inline code/URLs: Chinese full-width quotes -> ASCII straight quotes
  - File moved from /tmp/ to /home/ubuntu/internal/tools/
"""

import json
import re
import requests
import sys
import os

# 微信公众号凭据必须从环境变量读取，禁止硬编码（开源红线）
APP_ID = os.environ.get("WECHAT_APP_ID", "")
APP_SECRET = os.environ.get("WECHAT_APP_SECRET", "")

# Color palette
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
# and should ALWAYS be forced to half-width in ALL contexts.
# NOTE: Full-width QUOTES, PARENS, and SEMICOLON are EXCLUDED because
# they are correct in Chinese body text and should only be converted
# in code blocks, inline code, and URLs.
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
    """Pair-matching: convert ASCII straight quotes to Chinese curly quotes.

    Handles nested quotes correctly by tracking open/close state.
    Operates on raw text that has already been stripped of smart quotes.
    """
    result = []
    double_open = True  # True -> next " is opening; False -> next is closing
    single_open = True

    for ch in text:
        if ch == '"':
            if double_open:
                result.append("\u201c")  # "
            else:
                result.append("\u201d")  # "
            double_open = not double_open
        elif ch == "'":
            if single_open:
                result.append("\u2018")  # '
            else:
                result.append("\u2019")  # '
            single_open = not single_open
        else:
            result.append(ch)
    return "".join(result)


def _force_halfwidth_quotes(text):
    """Force all Chinese curly quotes and full-width quotes to straight ASCII."""
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\uff02", '"').replace("\uff07", "'")
    return text


def normalize_quotes(text):
    """Normalize punctuation for Chinese WeChat article typography.

    Rules (v2.2 - fixed quote direction):

    1. Smart/curly quotes (Word auto-converted) -> straight quotes first.
       This strips accidental smart-quote artifacts before applying
       the Chinese-curly replacement.

    2. Body text (outside code blocks, inline code, and markdown links):
       - Full-width spaces -> regular spaces
       - Full-width digits -> half-width digits
       - Straight ASCII quotes " ' -> Chinese curly quotes ""''
         (paired: opening/closing tracked automatically)
       - Other full-width ASCII punctuation -> half-width
       - Chinese punctuation (、。，：；？！） -> PRESERVED
       - Full-width parens （） and semicolon ； -> PRESERVED
         (correct in Chinese body text)
       - Zero-width chars/BOM -> removed

    3. Code blocks and inline code:
       - ALL full-width punctuation (including quotes, parens, semicolons)
         -> half-width
       - This ensures Python/JS/Shell code executes correctly

    4. Markdown link URLs:
       - Full-width quotes, parens, semicolons inside URL -> half-width
       (URLs cannot contain full-width characters)
    """
    if not text:
        return text

    # Step 1: Full-width spaces -> regular spaces (always)
    text = text.replace(_FULLWIDTH_SPACE, " ")

    # Step 3: Full-width digits -> half-width (always)
    for i in range(10):
        text = text.replace(chr(0xFF10 + i), str(i))

    # Step 4: Process code blocks separately
    code_block_pattern = re.compile(r"(```[\s\S]*?```)", re.MULTILINE)
    parts = code_block_pattern.split(text)

    result_parts = []
    for part in parts:
        if part.startswith("```"):
            # Inside code block: force ALL full-width to half-width
            for fw, hw in _FULLWIDTH_ALWAYS.items():
                part = part.replace(fw, hw)
            # Also force code-only full-width punct (parens, semicolons) to half-width
            for fw, hw in _FULLWIDTH_CODE_ONLY.items():
                part = part.replace(fw, hw)
            # Also force quotes to half-width inside code
            part = _force_halfwidth_quotes(part)
            # Force full-width comma/colon/question/exclamation to half-width in code
            part = part.replace("\uff0c", ",").replace("\uff1a", ":")
            part = part.replace("\uff1f", "?").replace("\uff01", "!")
            result_parts.append(part)
        else:
            # Outside code blocks: apply always-normalize map (no quotes)
            for fw, hw in _FULLWIDTH_ALWAYS.items():
                part = part.replace(fw, hw)

            # Protect inline code segments before quote conversion
            inline_code_pattern = re.compile(r"(`[^`]+`)")
            sub_parts = inline_code_pattern.split(part)
            processed_sub = []
            for sp in sub_parts:
                if sp.startswith("`") and sp.endswith("`") and len(sp) > 1:
                    # Inside inline code: force quotes and code-only punct to half-width
                    sp = _force_halfwidth_quotes(sp)
                    for fw, hw in _FULLWIDTH_CODE_ONLY.items():
                        sp = sp.replace(fw, hw)
                    processed_sub.append(sp)
                else:
                    # Regular text: normalize quotes in markdown link URLs
                    def fix_link(m):
                        prefix = m.group(1)
                        text_part = m.group(2)
                        url_part = m.group(3)
                        # Force half-width quotes and code-only punct inside URL
                        url_part = _force_halfwidth_quotes(url_part)
                        for fw, hw in _FULLWIDTH_CODE_ONLY.items():
                            url_part = url_part.replace(fw, hw)
                        return f"{prefix}[{text_part}]({url_part})"

                    sp = re.sub(
                        r"(!?)\[([^\]]+)\]\(([^)]+)\)",
                        fix_link,
                        sp,
                    )
                    # Now convert straight quotes to Chinese curly quotes
                    sp = _normalize_straight_quotes(sp)
                    processed_sub.append(sp)
            part = "".join(processed_sub)
            result_parts.append(part)

    text = "".join(result_parts)

    # Step 5: Collapse 3+ consecutive spaces
    text = re.sub(r" {3,}", "  ", text)

    # Step 6: Remove zero-width characters
    text = text.replace("\u200b", "")  # Zero-width space
    text = text.replace("\u200c", "")  # Zero-width non-joiner
    text = text.replace("\u200d", "")  # Zero-width joiner
    text = text.replace("\ufeff", "")  # Zero-width no-break space (BOM)

    return text


def get_access_token():
    url = f"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={APP_ID}&secret={APP_SECRET}"
    resp = requests.get(url, timeout=10)
    data = resp.json()
    if "access_token" not in data:
        print(f"ERROR getting token: {data}")
        sys.exit(1)
    return data["access_token"]


def upload_thumb(token, image_path):
    """Upload image as thumb media (max 64KB)."""
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


def render_code_block(code, lang=""):
    """Render code block with WeChat-compatible line wrapping."""
    lines = code.strip().split("\n")
    rendered_lines = []
    for line in lines:
        # Escape HTML
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
    """Render markdown table as WeChat-compatible HTML table."""

    def parse_cells(row):
        cells = [c.strip() for c in row.strip().strip("|").split("|")]
        return cells

    headers = parse_cells(header_row)
    body_rows = [parse_cells(r) for r in data_rows if r.strip()]

    # Build table
    html = f'<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:14px;">'

    # Header
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

    # Body
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


def md_to_html(md_text):
    """Convert markdown to WeChat-compatible HTML with Silicon Design style."""
    # v2.2: Normalize full-width/half-width punctuation BEFORE conversion
    md_text = normalize_quotes(md_text)

    lines = md_text.split("\n")
    html_parts = []
    in_code_block = False
    code_buffer = []
    code_lang = ""

    i = 0
    while i < len(lines):
        line = lines[i]

        # Code block start/end
        if line.strip().startswith("```"):
            if in_code_block:
                # End of code block
                html_parts.append(render_code_block("\n".join(code_buffer), code_lang))
                code_buffer = []
                in_code_block = False
            else:
                # Start of code block
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

        # Horizontal rule -> centered dots
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
            # Collect all consecutive table rows
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
                # Not a real table, fall through
                for tl in table_lines:
                    html_parts.append(
                        f'<p style="font-size:15px;color:{DARK_TEXT};'
                        f'line-height:1.8;margin:10px 0;">{process_inline(tl.strip())}</p>'
                    )
                continue

        # List items
        if line.strip().startswith("- "):
            item = line.strip()[2:].strip()
            # Process inline code in list items
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

    return "".join(html_parts)


def process_inline(text):
    """Process inline markdown: bold, code, links."""
    # Bold -> red highlight
    text = re.sub(
        r"\*\*([^*]+)\*\*",
        rf'<strong style="color:{RED};font-weight:bold;">\1</strong>',
        text,
    )
    # Inline code
    text = re.sub(
        r"`([^`]+)`",
        r'<code style="background:#f0f0f0;padding:2px 6px;border-radius:3px;'
        r'font-family:Consolas,Monaco,monospace;font-size:13px;color:#c7254e;">\1</code>',
        text,
    )
    # Links
    text = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        rf'<a href="\2" style="color:{BLUE};text-decoration:none;">\1</a>',
        text,
    )
    # Italic (italic text at end of article)
    text = re.sub(
        r"\*([^*]+)\*", rf'<em style="color:{GRAY_TEXT};font-size:13px;">\1</em>', text
    )
    return text


def create_draft(token, title, content_html, thumb_media_id):
    """Create WeChat draft."""
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

    # CRITICAL: ensure_ascii=False to prevent \uXXXX encoding
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


def main():
    md_path = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "/tmp/2026-08-03_\u4e13\u9879\u653b\u7565_\u5ba1\u6838\u7248.md"
    )
    cover_path = (
        sys.argv[2] if len(sys.argv) > 2 else "/tmp/cover_bidding_law_compressed.jpg"
    )

    # Read markdown
    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    # Extract title (first H1)
    title_match = re.match(r"^#\s+(.+)$", md_text, re.MULTILINE)
    title = (
        title_match.group(1).strip()
        if title_match
        else "2026\u62db\u6807\u6cd5\u5927\u4fee\u843d\u5730\uff0c\u4f60\u7684\u6807\u4e66\u6392\u7248\u8fd8\u5728\u7528\u624b\u5de5\uff1f"
    )

    print(f"Title: {title}")

    # Convert to HTML
    html_content = md_to_html(md_text)
    print(f"HTML length: {len(html_content)} chars")

    # Get access token
    token = get_access_token()
    print(f"Token obtained: {token[:20]}...")

    # Upload cover as thumb
    thumb_id = upload_thumb(token, cover_path)

    # Create draft
    draft_id = create_draft(token, title, html_content, thumb_id)
    print(f"\n=== SUCCESS ===")
    print(f"Draft media_id: {draft_id}")
    print(f"Title: {title}")


if __name__ == "__main__":
    main()
