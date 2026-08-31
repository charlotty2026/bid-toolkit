import os
#!/usr/bin/env python3
"""Test normalize_quotes() function - v2.2 fixed quote direction."""

import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.publish_wechat import normalize_quotes

tests = [
    # (input, expected_output, description)
    # === Quote direction: ASCII straight -> Chinese curly ===
    # 1. Straight double quotes in body text -> Chinese curly quotes
    (
        '"\u4f60\u597d"',
        "\u201c\u4f60\u597d\u201d",
        "Straight double quotes -> Chinese curly in body",
    ),
    # 2. Straight single quotes in body text -> Chinese curly quotes
    (
        "'\u4f60\u597d'",
        "\u2018\u4f60\u597d\u2019",
        "Straight single quotes -> Chinese curly in body",
    ),
    # 3. Chinese curly quotes already correct -> unchanged (PRESERVED)
    (
        "\u201c\u4f60\u597d\u201d",
        "\u201c\u4f60\u597d\u201d",
        "Already correct Chinese curly -> unchanged",
    ),
    # 4. Multiple sentence pairs (even number of quotes)
    (
        '"\u7b2c\u4e00\u53e5\u3002"\u7b2c\u4e8c\u53e5\u3002"\u7b2c\u4e09\u53e5\u3002"',
        "\u201c\u7b2c\u4e00\u53e5\u3002\u201d\u7b2c\u4e8c\u53e5\u3002\u201c\u7b2c\u4e09\u53e5\u3002\u201d",
        "Multiple sentence pairs: straight -> curly (even count)",
    ),
    # === Code blocks: full-width/smart -> straight ASCII ===
    # 5. Straight double quotes in code block -> stay straight (for Python strings)
    (
        '```python\nprint("hello")\n```',
        '```python\nprint("hello")\n```',
        "Straight quotes in code block -> unchanged",
    ),
    # 6. Full-width double quotes in code block -> straight
    (
        "```python\nprint(\uff02hello\uff02)\n```",
        '```python\nprint("hello")\n```',
        "Full-width quotes in code block -> straight",
    ),
    # 7. Full-width single quotes in code block -> straight
    (
        "```python\nprint(\uff07hello\uff07)\n```",
        "```python\nprint('hello')\n```",
        "Full-width single quotes in code block -> straight",
    ),
    # 8. Smart double quotes in code block -> straight (Word artifact)
    (
        "```python\nprint(\u201chello\u201d)\n```",
        '```python\nprint("hello")\n```',
        "Smart quotes in code block -> straight",
    ),
    # === Inline code: full-width -> straight ASCII ===
    # 9. Full-width double quotes in inline code -> straight
    (
        "`print(\uff02hello\uff02)`",
        '`print("hello")`',
        "Full-width quotes in inline code -> straight",
    ),
    # 10. Straight double quotes in inline code -> stay straight
    (
        '`print("hello")`',
        '`print("hello")`',
        "Straight quotes in inline code -> unchanged",
    ),
    # 11. Smart double quotes in inline code -> straight
    (
        "`print(\u201chello\u201d)`",
        '`print("hello")`',
        "Smart quotes in inline code -> straight",
    ),
    # === Other punctuation ===
    # 12. Full-width parens preserved in body text (v2.3 fix)
    (
        "\uff08\u6d4b\u8bd5\uff09",
        "\uff08\u6d4b\u8bd5\uff09",
        "Full-width parens preserved in body text",
    ),
    # 13. Full-width digits
    ("\uff10\uff11\uff12", "012", "Full-width digits -> half-width"),
    # 14. Full-width space
    ("hello\u3000world", "hello world", "Full-width space -> regular space"),
    # 15. Chinese punctuation preserved
    (
        "\u4f60\u597d\u3001\u4e16\u754c\u3002",
        "\u4f60\u597d\u3001\u4e16\u754c\u3002",
        "Chinese ideographic comma/period preserved",
    ),
    # 16. Full-width comma in body text preserved
    (
        "\u4f60\u597d\uff0c\u4e16\u754c\u3002",
        "\u4f60\u597d\uff0c\u4e16\u754c\u3002",
        "Full-width comma preserved in body text",
    ),
    # 17. Full-width colon in body text preserved
    (
        "\u63d0\u793a\uff1a\u6ce8\u610f\u3002",
        "\u63d0\u793a\uff1a\u6ce8\u610f\u3002",
        "Full-width colon preserved in body text",
    ),
    # 18. Full-width question/exclamation in body preserved
    (
        "\u600e\u4e48\uff1f\u597d\u7684\uff01",
        "\u600e\u4e48\uff1f\u597d\u7684\uff01",
        "Full-width ?/! preserved in body text",
    ),
    # === Zero-width chars ===
    # 19. Zero-width characters removed
    ("hello\u200bworld\u200c\u200d", "helloworld", "Zero-width characters removed"),
    # 20. BOM removed
    ("\ufeffhello", "hello", "BOM removed"),
    # === Already correct: Chinese curly -> unchanged ===
    # 21. Already correct single curly quotes -> stay as is
    (
        "\u2018\u4f60\u597d\u2019",
        "\u2018\u4f60\u597d\u2019",
        "Already correct single curly -> unchanged",
    ),
    # === Code block with Chinese punctuation ===
    # 22. Full-width colon/comma in code block -> half-width
    (
        "```\u63d0\u793a\uff1a\u6ce8\u610f\uff0c\u7ed3\u675f\uff01\uff1f\n```",
        "```\u63d0\u793a:\u6ce8\u610f,\u7ed3\u675f!?\n```",
        "Full-width punct in code block -> half-width",
    ),
    # === Complex: mixed in body text ===
    # 23. Mixed scenario: header + body + code
    # Body text: full-width parens PRESERVED (v2.3 fix)
    # Code block: full-width parens -> half-width
    (
        "\uff03\uff03 \u6807\u9898\n\n\u8fd9\u662f\u6d4b\u8bd5\uff08\uff08\u62ec\u53f7\uff09\uff09\u3002\n\n```\nimport\uff08\uff09\n```",
        "## \u6807\u9898\n\n\u8fd9\u662f\u6d4b\u8bd5\uff08\uff08\u62ec\u53f7\uff09\uff09\u3002\n\n```\nimport()\n```",
        "Mixed scenario: header, body text, code block",
    ),
    # 24. Straight quotes inside inline code with Chinese curly outside
    (
        '\u201c\u4f60\u597d\u201d\uff0c\u8fd9\u662f\u4e00\u53e5`print("hello")`\u3002',
        '\u201c\u4f60\u597d\u201d\uff0c\u8fd9\u662f\u4e00\u53e5`print("hello")`\u3002',
        "Curly outside + straight in inline code -> preserve both",
    ),
    # 25. Straight quotes across paragraphs (restart pairing at each paragraph)
    (
        '"\u7b2c\u4e00\u6bb5\u3002"\n\n"\u7b2c\u4e8c\u6bb5\u3002"',
        "\u201c\u7b2c\u4e00\u6bb5\u3002\u201d\n\n\u201c\u7b2c\u4e8c\u6bb5\u3002\u201d",
        "Straight quotes across paragraphs -> curly each paragraph",
    ),
    # 26. Full-width brackets
    ("\uff3b\u6d4b\u8bd5\uff3d", "[\u6d4b\u8bd5]", "Full-width brackets -> half-width"),
    # 27. Empty string
    ("", "", "Empty string passthrough"),
    # === v2.3: Full-width parens/semicolons preserved in body, converted in code ===
    # 28. Full-width semicolon preserved in body text
    (
        "\u7b2c\u4e00\u53e5\uff1b\u7b2c\u4e8c\u53e5\u3002",
        "\u7b2c\u4e00\u53e5\uff1b\u7b2c\u4e8c\u53e5\u3002",
        "Full-width semicolon preserved in body text",
    ),
    # 29. Full-width parens in inline code -> half-width
    (
        "`func\uff08x\uff09`",
        "`func(x)`",
        "Full-width parens in inline code -> half-width",
    ),
    # 30. Full-width semicolon in code block -> half-width
    (
        "```js\nlet x = 1\uff1b\n```",
        "```js\nlet x = 1;\n```",
        "Full-width semicolon in code block -> half-width",
    ),
    # 31. Full-width parens in markdown link URL -> half-width
    (
        "[\u94fe\u63a5](https://example.com/path\uff08test\uff09)",
        "[\u94fe\u63a5](https://example.com/path(test))",
        "Full-width parens in URL -> half-width",
    ),
    # 32. Real-world sentence: parens and semicolons in body text preserved
    (
        "\u82e6\u529b\u6d3b\uff08\u6293\u8d44\u6599\u3001\u6392\u683c\u5f0f\uff09\u4ea4\u7ed9\u514d\u8d39\u6a21\u578b\uff1b\u786c\u4e8b\u4ea4\u7ed9\u4ed8\u8d39\u6a21\u578b\u3002",
        "\u82e6\u529b\u6d3b\uff08\u6293\u8d44\u6599\u3001\u6392\u683c\u5f0f\uff09\u4ea4\u7ed9\u514d\u8d39\u6a21\u578b\uff1b\u786c\u4e8b\u4ea4\u7ed9\u4ed8\u8d39\u6a21\u578b\u3002",
        "Real-world sentence: parens and semicolons preserved in body",
    ),
    # 33. Full-width parens inside code block -> half-width
    (
        "```python\nprint\uff08\uff09\n```",
        "```python\nprint()\n```",
        "Full-width parens in code block -> half-width",
    ),
]

passed = 0
failed = 0

for i, (inp, expected, desc) in enumerate(tests, 1):
    result = normalize_quotes(inp)
    if result == expected:
        passed += 1
        status = "PASS"
    else:
        failed += 1
        status = "FAIL"
    print(f"[{status}] Test {i}: {desc}")
    if status == "FAIL":
        print(f"  Input:    {repr(inp)}")
        print(f"  Expected: {repr(expected)}")
        print(f"  Got:      {repr(result)}")

print(f"\n=== Results: {passed} passed, {failed} failed ===")
sys.exit(0 if failed == 0 else 1)
