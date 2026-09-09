#!/usr/bin/env python3
"""
标书质量防线 - 自动化检查脚本
引用功法：排版定鼎功 + 三层递进13步 + 去AI味 + 废标规避心法 + 内容块库

用法：python3 quality_check.py 投标文件.docx
"""
import sys
import os
import zipfile
import re
import xml.etree.ElementTree as ET


def extract_text_from_docx(docx_path):
    """从docx提取纯文本"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        content = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(content)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    texts = []
    for t in root.iter('{' + ns + '}t'):
        if t.text:
            texts.append(t.text)
    return ' '.join(texts)


def extract_media_info(docx_path):
    """统计图片数量和文件名"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        media_files = [f for f in z.namelist() if f.startswith('word/media/')]
    return media_files


def check_cross_industry(text):
    """检查跨行业模板残留
    依据：废标规避心法 - 实质性响应原则
    """
    keywords = [
        ('建筑施工', '建筑行业残留'),
        ('施工现场', '建筑行业残留'),
        ('安全帽', '建筑行业残留'),
        ('安全绳', '建筑行业残留'),
        ('职业病危害', '建筑/工厂行业残留'),
        ('树叶', '绿化/环卫行业残留'),
        ('取暖', '家庭/物业残留'),
        ('做饭', '家庭/物业残留'),
        ('洗衣粉', '家庭/物业残留'),
        ('路灯', '市政行业残留'),
        ('太阳能照明', '市政行业残留'),
        ('消声器', '工程行业残留'),
        ('蓝天救援队', '救援组织残留'),
        ('卫健委', '政府公文残留'),
        ('销售岗', '非外包岗位残留'),
        ('行政岗', '非外包岗位残留'),
        ('技术岗', '非外包岗位残留'),
    ]
    findings = []
    for kw, desc in keywords:
        count = text.count(kw)
        if count > 0:
            # 找上下文
            idx = text.find(kw)
            context_start = max(0, idx - 20)
            context_end = min(len(text), idx + len(kw) + 30)
            context = text[context_start:context_end].replace('\n', ' ')
            findings.append(f'  🔴 "{kw}" x{count}处（{desc}）\n    位置: ...{context}...')
    return findings


def check_placeholder(text):
    """检查占位符残留
    依据：排版定鼎功 - 格式零错误原则
    """
    patterns = [
        ('XX', '通用占位符'),
        ('____', '下划线占位'),
        ('采购人', '招标方占位符'),
        ('胜任本项目主管', '人员占位符'),
        ('待补充', '内容占位'),
        ('TODO', '开发占位'),
        ('（待填写）', '表格占位'),
        ('<待定>', '尖括号占位'),
        ('AAA', '通用占位符'),
        ('BBB', '通用占位符'),
    ]
    findings = []
    for p, desc in patterns:
        count = text.count(p)
        if count > 0:
            findings.append(f'  🔴 "{p}" x{count}处（{desc}）')
    return findings


def check_legal_traps(text):
    """检查法律自坑表述
    依据：废标规避心法 + 法律合规功法
    """
    trap_patterns = [
        (r'解释权[归由]乙方', '"解释权归乙方"民法典认定无效'),
        (r'零[中断差错事故]', '绝对承诺表述，改"平稳过渡，尽量避免"'),
        (r'赔偿.*封顶', '赔偿上限封顶，改协商方案'),
        (r'零事故', '绝对承诺，建议删除'),
        (r'确保安全', '绝对化表述，改为"加强安全管理"'),
        (r'绝[对不][会发.生]', '绝对承诺'),
        (r'100[%%].*[满完]', '绝对化表述'),
        (r'无[任]何风险', '绝对承诺'),
    ]
    findings = []
    for pattern, desc in trap_patterns:
        matches = re.findall(pattern, text)
        if matches:
            findings.append(f'  🔴 x{len(matches)}处（{desc}）: "{pattern}"')
    return findings


def check_markdown_residue(text):
    """检查markdown标记残留
    依据：排版定鼎功 - 纯Word格式
    """
    patterns = [
        (r'\*\*[^*]+\*\*', '**加粗** 标记'),
        (r'#+\s', '标题 # 标记'),
        (r'`[^`]+`', '代码块 ` 标记'),
        (r'~~[^~]+~~', '删除线 ~~ 标记'),
        (r'\[.*\]\(.*\)', 'Markdown链接 [](url)'),
    ]
    findings = []
    for pattern, desc in patterns:
        count = len(re.findall(pattern, text))
        if count > 0:
            findings.append(f'  🟡 x{count}处 {desc} 残留')
    return findings


def check_ai_tone(text):
    """检查AI八股文风
    依据：去AI味功法 - 绝对禁止清单 + 谨慎使用清单
    """
    # 绝对禁止（从去AI味功法提取）
    forbidden = [
        ('综上所述', 'AI八股 - 直接说结论'),
        ('值得注意的是', 'AI八股 - 直接说内容'),
        ('至关重要', 'AI八股 - 说具体多重要'),
        ('不可或缺', 'AI八股 - 说缺了会怎样'),
        ('赋能', 'AI八股 - 说具体作用'),
        ('提供全方位', 'AI八股 - 列具体项目'),
    ]
    # 谨慎使用
    cautious = [
        ('此外', '建议用"还有/另外"'),
        ('值得一提的是', 'AI八股'),
        ('毋庸置疑', 'AI八股'),
        ('不言而喻', 'AI八股'),
        ('众所周知', 'AI八股'),
    ]
    findings = []
    for c, desc in forbidden:
        count = text.count(c)
        if count > 0:
            findings.append(f'  🔴 x{count}处（{desc}）: "{c}"')
    for c, desc in cautious:
        count = text.count(c)
        if count > 0:
            findings.append(f'  🟡 x{count}处（{desc}）: "{c}"')
    return findings


def check_data_sources(text):
    """检查数据是否有来源标注
    依据：标书数字理解五重验证法
    """
    # 统计数字
    number_matches = re.findall(r'\d+[.%]', text)
    source_sentences = re.findall(r'[^。]*?[来源数据统计根据][^。]*。', text)
    findings = []
    total_numbers = len(number_matches)
    source_count = len(source_sentences)
    if total_numbers > 0:
        ratio = source_count / total_numbers if total_numbers > 0 else 0
        findings.append(f'  🟡 共 {total_numbers} 个数字/百分比，{source_count} 处有来源标注（覆盖率 {ratio:.0%}）')
        if ratio < 0.1 and total_numbers > 10:
            findings.append(f'  🔴 数据来源覆盖率极低（{ratio:.0%}），建议逐条标注出处')
    return findings


def check_table_structure(text):
    """检查表格一致性（简化版）
    依据：排版定鼎功 - 表格规范
    """
    # 检查是否有明显表格结构问题
    findings = []
    # 发现类似表格行但不规范的写法
    table_lines = re.findall(r'^[|].*[|]$', text, re.MULTILINE)
    if len(table_lines) > 0:
        findings.append(f'  🟡 发现 {len(table_lines)} 行markdown表格语法（需确认是否已转为Word表格）')
    return findings


def run_checks(docx_path):
    """执行全部检查"""
    print(f'\n🛡️  标书质量防线检查报告')
    print(f'━━━━━━━━━━━━━━━━━━━━')
    print(f'📄 文件: {os.path.basename(docx_path)}')
    print(f'📚 引用功法: 排版定鼎功/三层递进13步/去AI味/废标规避心法')
    print()

    # 提取文本
    try:
        text = extract_text_from_docx(docx_path)
    except Exception as e:
        print(f'❌ 无法读取文件: {e}')
        return

    print(f'📊 文档字数: {len(text):,} 字')

    # 图片统计
    media = extract_media_info(docx_path)
    print(f'🖼️  嵌入图片: {len(media)} 张')
    print()

    all_high = 0
    all_mid = 0

    # ===== 高危检查 =====
    print('【🔴 高危检查 - 不处理=废标风险】')
    print('━━━━━━━━━━━━━━━━━━━━')

    findings = check_cross_industry(text)
    if findings:
        print('❌ 跨行业模板残留:')
        for f in findings:
            print(f)
        all_high += len(findings)
    else:
        print('✅ 跨行业模板残留检查通过')

    findings = check_placeholder(text)
    if findings:
        print('\n❌ 占位符残留:')
        for f in findings:
            print(f'  {f}')
        all_high += len(findings)
    else:
        print('✅ 占位符残留检查通过')

    findings = check_legal_traps(text)
    if findings:
        print('\n❌ 法律自坑表述:')
        for f in findings:
            print(f'  {f}')
        all_high += len(findings)
    else:
        print('✅ 法律自坑表述检查通过')

    findings = check_ai_tone(text)
    ai_reds = [f for f in findings if '🔴' in f]
    if ai_reds:
        print('\n❌ AI八股禁用词:')
        for f in ai_reds:
            print(f'  {f}')
        all_high += len(ai_reds)
    else:
        print('✅ 去AI味检查通过')

    # ===== 中危检查 =====
    print()
    print('【🟡 中危检查 - 处理=加分，不处理=不扣分】')
    print('━━━━━━━━━━━━━━━━━━━━')

    findings = check_markdown_residue(text)
    if findings:
        print('⚠️  Markdown标记残留:')
        for f in findings:
            print(f'  {f}')
        all_mid += len(findings)
    else:
        print('✅ Markdown残留检查通过')

    ai_yellows = [f for f in findings if '🟡' in f]
    findings = check_data_sources(text)
    if findings:
        print('\n⚠️  数据来源:')
        for f in findings:
            print(f'  {f}')
        all_mid += len(findings)
    else:
        print('✅ 数据来源检查通过')

    findings = check_table_structure(text)
    if findings:
        print('\n⚠️  表格结构:')
        for f in findings:
            print(f'  {f}')
        all_mid += len(findings)

    # ===== 人工审阅 =====
    print()
    print('【📋 人工审阅项 - AI无法自动判断】')
    print('━━━━━━━━━━━━━━━━━━━━')
    print('  1️⃣ 角色定位：逐条核对外包公司口诀')
    print('  2️⃣ 跨章矛盾：检查时间线/考核标准/人数口径一致性')
    print('  3️⃣ 法规验证：确认所有引用法规名称是否准确存在')
    print('  4️⃣ 图片完整性：打开docx原件肉眼核查')
    print(f'  5️⃣ 内容块匹配：确认招标章节名称与{17}个内容块的映射关系')
    print('  6️⃣ 人员资质排布：核对拟派人员名单和证书对应关系')
    print()

    # ===== 总结 =====
    print('━━━━━━━━━━━━━━━━━━━━')
    print(f'📊 检查总结')
    print(f'🔴 高危: {all_high} 项')
    print(f'🟡 中危: {all_mid} 项')
    print(f'📋 人工: 6 项')
    print('━━━━━━━━━━━━━━━━━━━━')
    if all_high > 0:
        print('🔴 存在高危项，必须处理后才能输出！')
        print('   处理策略：按功法「三层递进13步」的优先级修改')
        print('   修改完成后重新运行检查')
    elif all_mid > 0:
        print('🟡 无高危项，但有中危项，建议处理后再输出')
    else:
        print('✅ 全部检查通过，可以输出！')
    print()


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python3 quality_check.py 投标文件.docx')
        sys.exit(1)
    docx_path = sys.argv[1]
    if not os.path.exists(docx_path):
        print(f'❌ 文件不存在: {docx_path}')
        sys.exit(1)
    run_checks(docx_path)
