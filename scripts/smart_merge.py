
# ===== 智能合并表格列（bid-toolkit 2026-05-20）=====
def smart_merge_table(table):
    """
    智能合并表格中相邻相同内容的单元格。
    适用场景：技术方案中的层级/组别列自动合并。
    """
    rows = table.rows
    cols = len(table.columns)
    if len(rows) < 2:
        return 0
    total_merges = 0
    for col_idx in range(cols):
        merge_start = None
        prev_text = None
        for row_idx in range(len(rows)):
            cell = table.cell(row_idx, col_idx)
            text = cell.text.strip()
            if text and text == prev_text:
                if merge_start is None:
                    merge_start = row_idx - 1
            else:
                if merge_start is not None:
                    top = table.cell(merge_start, col_idx)
                    bottom = table.cell(row_idx - 1, col_idx)
                    if top is not bottom:
                        top.merge(bottom)
                        total_merges += 1
                    merge_start = None
            prev_text = text
        if merge_start is not None:
            top = table.cell(merge_start, col_idx)
            bottom = table.cell(len(rows) - 1, col_idx)
            if top is not bottom:
                top.merge(bottom)
                total_merges += 1
    return total_merges

def scan_merge_needed(table):
    """扫描表格，返回需要合并的组合数（不实际合并）"""
    rows = table.rows
    cols = len(table.columns)
    if len(rows) < 2:
        return 0
    count = 0
    for col_idx in range(cols):
        found = False
        prev_text = None
        for row_idx in range(len(rows)):
            text = table.cell(row_idx, col_idx).text.strip()
            if text and text == prev_text:
                if not found:
                    count += 1
                    found = True
            else:
                found = False
            prev_text = text
    return count

def fix_table_merge(doc):
    """修复文档中所有表格的合并"""
    total = 0
    for table in doc.tables:
        n = scan_merge_needed(table)
        if n > 0:
            m = smart_merge_table(table)
            total += m
    return total
