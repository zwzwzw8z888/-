#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中建四局公文格式化核心逻辑
从 app.py 抽取，供 Flask 后端调用
"""

import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ────────────────────────── 格式常量 ──────────────────────────
FONT_FANGSONG     = "仿宋_GB2312"
FONT_HEITI        = "黑体"
FONT_KAITI        = "楷体_GB2312"
FONT_XIAOBIAOSONG = "方正小标宋简体"
FONT_TIMES_NEW_ROMAN = "Times New Roman"

SIZE_CHUHAO  = Pt(42)
SIZE_ERHAO   = Pt(22)
SIZE_SANHAO  = Pt(16)
SIZE_XIAOSI  = Pt(12)

LINE_SPACING_TWIPS = 579  # 28.9磅 = 579 twips ⚠️ 单位是twips！

MARGIN_TOP    = Cm(3.7)
MARGIN_BOTTOM = Cm(3.5)
MARGIN_LEFT   = Cm(2.8)
MARGIN_RIGHT  = Cm(2.6)

CN_NUMBERS = ['一','二','三','四','五','六','七','八','九','十',
              '十一','十二','十三','十四','十五','十六','十七','十八','十九','二十']
CNUM = {str(i+1): s for i, s in enumerate(CN_NUMBERS)}
CNUM_TO_INT = {s: i+1 for i, s in enumerate(CN_NUMBERS)}
CIRCLE_NUMBERS = ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩',
                  '⑪','⑫','⑬','⑭','⑮','⑯','⑰','⑱','⑲','⑳']


def set_run_font(run, cn_font, size_pt, bold=False, color=None):
    run.font.name = FONT_TIMES_NEW_ROMAN
    run.font.size = size_pt
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), FONT_TIMES_NEW_ROMAN)
    rFonts.set(qn('w:hAnsi'), FONT_TIMES_NEW_ROMAN)


def set_para_spacing(para, twips=LINE_SPACING_TWIPS):
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), str(twips))
    spacing.set(qn('w:lineRule'), 'exact')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')


def set_para_indent(para, first_line_chars=2, char_size_pt=16):
    dxa = int(first_line_chars * char_size_pt * 20)
    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLine'), str(dxa))


def clean_text(text):
    text = re.sub(r'^#{1,6}\s*', '', text)
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'[`~_>|\\^]', '', text)
    text = re.sub(r'  +', ' ', text).strip()
    return text


def detect_level(text):
    t = text.strip()
    if re.match(r'^[一二三四五六七八九十]+、', t):
        return 'h1'
    if re.match(r'^（[一二三四五六七八九十]+）', t):
        return 'h2'
    if re.match(r'^\d+[.、．](?!\d)\s*', t):
        return 'h3'
    if re.match(r'^（\d+）', t):
        return 'h4'
    if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', t):
        return 'h5'
    return 'body'


def is_main_title(text):
    t = text.strip()
    if not t:
        return False
    # 排除时间行（如 "2026年4月"、"2026年4月28日"）
    if re.match(r'^\d{4}年\d{1,2}月\d{0,2}日?\s*$', t):
        return False
    # 排除带括号的日期行（如 "（2026年4月28日）"、"(2026年4月28日)"）
    if re.match(r'^[（(]\d{4}年\d{1,2}月\d{1,2}日[）)]\s*$', t):
        return False
    # 排除"月底"数据行（如 "4月底节点：差异≤40%..."）
    if re.search(r'月底', t):
        return False
    for pat in [
        r'^[一二三四五六七八九十]+、',
        r'^（[一二三四五六七八九十]+）',
        r'^\d+[.、．]\s*',
        r'^（\d+）',
        r'^[①②③④⑤⑥⑦⑧⑨⑩]',
        r'^[\d,.\-+/：:；，。、]+$',
    ]:
        if re.match(pat, t):
            return False
    if '。' in t or '；' in t:
        return False
    # 排除问候语（含冒号且含称呼关键词）
    if re.search(r'[：:]$', t) and re.search(r'领导|同事|各位|尊敬|您好|下午好|上午好|你好', t):
        return False
    # 主标题长度放宽至40字（汇报类文件标题通常较长）
    if len(t) > 40:
        return False
    return True


class HeadingCounter:
    def __init__(self):
        self.h1 = self.h2 = self.h3 = self.h4 = self.h5 = 0

    def next(self, level):
        if level == 'h1':
            self.h1 += 1; self.h2 = self.h3 = self.h4 = self.h5 = 0
            idx = self.h1 - 1
            return (CN_NUMBERS[idx] if idx < len(CN_NUMBERS) else str(self.h1)) + '、'
        elif level == 'h2':
            self.h2 += 1; self.h3 = self.h4 = self.h5 = 0
            idx = self.h2 - 1
            return f'（{CN_NUMBERS[idx] if idx < len(CN_NUMBERS) else str(self.h2)}）'
        elif level == 'h3':
            self.h3 += 1; self.h4 = self.h5 = 0
            return f'{self.h3}.'
        elif level == 'h4':
            self.h4 += 1; self.h5 = 0
            return f'（{self.h4}）'
        elif level == 'h5':
            self.h5 += 1
            idx = self.h5 - 1
            return CIRCLE_NUMBERS[idx] if idx < len(CIRCLE_NUMBERS) else f'({self.h5})'
        return ''


def apply_heading_format(para, level, text, prefix='', no_indent=False, preserve_bold=False):
    para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(para)
    # 全局规则：所有段落（含标题）默认首行缩进2字符，除非明确 no_indent（如问候语/主标题）
    if not no_indent:
        set_para_indent(para, 2)
    display = prefix + text
    font_map = {
        'h1': FONT_HEITI,
        'h2': FONT_KAITI,
        'title': FONT_XIAOBIAOSONG,  # 主标题：方正小标宋简体
    }
    cn_font = font_map.get(level, FONT_FANGSONG)
    
    run = para.add_run(display)
    
    # 主标题特殊处理：居中、二号字体、无首行缩进
    if level == 'title':
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.first_line_indent = None
        set_run_font(run, cn_font, SIZE_ERHAO, bold=False)
    else:
        set_run_font(run, cn_font, SIZE_SANHAO, bold=preserve_bold)


def _calc_smart_col_widths(rows_data, num_cols):
    col_max_chars = [0] * num_cols
    for row_data in rows_data:
        for j in range(min(len(row_data), num_cols)):
            col_max_chars[j] = max(col_max_chars[j], len(row_data[j].strip()))
    weights = []
    for j in range(num_cols):
        max_len = col_max_chars[j]
        is_seq_col = True
        for row_data in rows_data:
            cell = row_data[j].strip() if j < len(row_data) else ''
            if cell and not re.match(
                r'^(\d{1,3}[\.\-]?\d{0,2}|第[一二三四五六七八九十\d]+项?|序号|编号|No\.?|ID|项)$', cell
            ):
                is_seq_col = False
                break
        if is_seq_col:
            weights.append(800)
        elif max_len <= 6:
            weights.append(1000)
        elif max_len <= 12:
            weights.append(1500)
        elif max_len <= 20:
            weights.append(2200)
        elif max_len <= 35:
            weights.append(3000)
        else:
            weights.append(4000)
    return weights


def _split_run_at(target_elem, run, split_pos):
    """在run内部指定位置拆分为两个run，返回 (前半run, 后半run)。
    split_pos: 在run文本中的字符位置。
    如果 split_pos <= 0 或 >= 文本长度，不拆分，返回 (None, run) 或 (run, None)。
    """
    run_text = ''.join(t.text or '' for t in run.iter(qn('w:t')))
    if split_pos <= 0:
        return None, run
    if split_pos >= len(run_text):
        return run, None

    # 复制 rPr
    rPr = run.find(qn('w:rPr'))

    # 前半 run
    before_run = OxmlElement('w:r')
    if rPr is not None:
        before_run.append(rPr.__copy__())
    before_t = OxmlElement('w:t')
    before_t.text = run_text[:split_pos]
    before_t.set(qn('xml:space'), 'preserve')
    before_run.append(before_t)

    # 后半 run
    after_run = OxmlElement('w:r')
    if rPr is not None:
        after_run.append(rPr.__copy__())
    after_t = OxmlElement('w:t')
    after_t.text = run_text[split_pos:]
    after_t.set(qn('xml:space'), 'preserve')
    after_run.append(after_t)

    # 替换原 run
    run_idx = list(target_elem).index(run)
    target_elem.remove(run)
    target_elem.insert(run_idx, before_run)
    target_elem.insert(run_idx + 1, after_run)

    return before_run, after_run


def _add_highlight_to_run(run, color='yellow'):
    """给指定run添加高亮底色（当前暂不启用，保留接口）。"""
    return
    rPr = run.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run.insert(0, rPr)
    hl = rPr.find(qn('w:highlight'))
    if hl is None:
        hl = OxmlElement('w:highlight')
        rPr.append(hl)
    hl.set(qn('w:val'), color)


def _apply_comments_to_doc(doc, comment_list):
    """在文档中统一添加批注。
    comment_list: [(text_prefix, comment_text, anchor_type), ...]
    - text_prefix: 用于匹配段落的文本前缀
    - comment_text: 批注内容
    - anchor_type: 锚定方式，决定批注覆盖范围和高亮颜色
        'trailing_punct'  - 末尾标点错误，只标末尾标点字符，红色高亮
        'missing_end_punct' - 缺少句末标点，标段落末尾几个字，黄色高亮
        'number_prefix'   - 编号格式问题，只标编号部分，黄色高亮
        'heading_skip'    - 标题跳级，只标编号部分，黄色高亮
        'multi_level_num' - 多级编号，只标多级编号部分，黄色高亮
        'verb_after_num'  - 编号后跟动词，标编号+动词，黄色高亮
        'full_para'       - 整段问题，标整段，黄色高亮
    """
    if not comment_list:
        return
    import datetime
    from lxml import etree

    body = doc.element.body

    # 收集所有段落及其文本
    para_map = []  # [(child_index, para_element, text), ...]
    child_idx = 0
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            text = ''.join(t.text for t in child.iter(qn('w:t')) if t.text)
            para_map.append((child_idx, child, text))
        child_idx += 1

    # 创建 comments XML
    comments_xml = (
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '</w:comments>'
    )
    comments_element = etree.fromstring(comments_xml.encode('utf-8'))
    next_id = 0
    now_str = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
    author = '格式化审查'

    # 已匹配的段落+批注组合集合（允许同一段落添加不同批注，但避免完全相同的批注重复）
    matched_keys = set()

    for item in comment_list:
        # 兼容旧格式 (text_prefix, comment_text) 和新格式 (text_prefix, comment_text, anchor_type)
        if len(item) == 3:
            text_prefix, comment_text, anchor_type = item
        else:
            text_prefix, comment_text = item
            anchor_type = 'full_para'

        # 在输出段落中找到匹配的段落
        target_elem = None
        for ci, pelem, ptext in para_map:
            key = (ci, comment_text)
            if key in matched_keys:
                continue
            if ptext.startswith(text_prefix):
                target_elem = pelem
                matched_keys.add(key)
                break
        if target_elem is None:
            continue

        comment_id = str(next_id)

        # 1. 创建 w:comment 元素
        comment_elem = OxmlElement('w:comment')
        comment_elem.set(qn('w:id'), comment_id)
        comment_elem.set(qn('w:author'), author)
        comment_elem.set(qn('w:date'), now_str)
        comment_elem.set(qn('w:initials'), 'GS')

        p_comment = OxmlElement('w:p')
        r_comment = OxmlElement('w:r')
        t_comment = OxmlElement('w:t')
        t_comment.text = comment_text
        t_comment.set(qn('xml:space'), 'preserve')
        r_comment.append(t_comment)
        p_comment.append(r_comment)
        comment_elem.append(p_comment)
        comments_element.append(comment_elem)

        # 2. 根据 anchor_type 决定锚点位置和高亮范围
        runs = target_elem.findall(qn('w:r'))

        # 高亮颜色：错误型用红色，提醒型用黄色
        hl_color = 'red' if anchor_type == 'trailing_punct' else 'yellow'

        if anchor_type == 'trailing_punct' and runs:
            # 标题末尾标点：只标最后一个字符（冒号/句号等标点）
            last_run = runs[-1]
            last_text = ''.join(t.text or '' for t in last_run.iter(qn('w:t')))
            if last_text and not last_text[-1].isalnum() and not last_text[-1].isspace():
                # 末尾是标点，拆分run，只高亮标点
                before, after = _split_run_at(target_elem, last_run, len(last_text) - 1)
                if after is not None:
                    _add_highlight_to_run(after, hl_color)
                    # 批注锚点只包裹最后一个标点字符的run
                    after_idx = list(target_elem).index(after)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(after_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(after_idx + 2, ce)
                else:
                    # 拆分失败，高亮整个last_run
                    _add_highlight_to_run(last_run, hl_color)
                    run_idx = list(target_elem).index(last_run)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx + 2, ce)
            else:
                # 末尾无标点（不应到达此分支），高亮最后一个run
                _add_highlight_to_run(last_run, hl_color)
                run_idx = list(target_elem).index(last_run)
                cs = OxmlElement('w:commentRangeStart')
                cs.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx, cs)
                ce = OxmlElement('w:commentRangeEnd')
                ce.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx + 2, ce)

        elif anchor_type == 'missing_end_punct' and runs:
            # 缺少句末标点：高亮段落末尾几个字（最后run的文本，取最后几个字符）
            last_run = runs[-1]
            last_text = ''.join(t.text or '' for t in last_run.iter(qn('w:t')))
            # 取末尾最多3个字符高亮
            highlight_len = min(3, len(last_text))
            if highlight_len > 0 and len(last_text) > highlight_len:
                before, after = _split_run_at(target_elem, last_run, len(last_text) - highlight_len)
                if after is not None:
                    _add_highlight_to_run(after, hl_color)
                    after_idx = list(target_elem).index(after)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(after_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(after_idx + 2, ce)
                else:
                    _add_highlight_to_run(last_run, hl_color)
                    run_idx = list(target_elem).index(last_run)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx + 2, ce)
            else:
                _add_highlight_to_run(last_run, hl_color)
                run_idx = list(target_elem).index(last_run)
                cs = OxmlElement('w:commentRangeStart')
                cs.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx, cs)
                ce = OxmlElement('w:commentRangeEnd')
                ce.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx + 2, ce)

        elif anchor_type in ('number_prefix', 'heading_skip') and runs:
            # 编号格式/标题跳级：只标编号部分（如"1."）
            first_run = runs[0]
            first_text = ''.join(t.text or '' for t in first_run.iter(qn('w:t')))
            # 匹配编号前缀
            num_match = re.match(r'^(\d+[.、．]\s*|[一二三四五六七八九十]+[、]\s*|（[一二三四五六七八九十]+）\s*|（\d+）\s*)', first_text)
            if num_match and len(num_match.group(0)) < len(first_text):
                # 编号和正文在同一个run中，拆分
                before, after = _split_run_at(target_elem, first_run, len(num_match.group(0)))
                if before is not None:
                    _add_highlight_to_run(before, hl_color)
                    before_idx = list(target_elem).index(before)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(before_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(before_idx + 2, ce)
                else:
                    _add_highlight_to_run(first_run, hl_color)
                    run_idx = list(target_elem).index(first_run)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx + 2, ce)
            elif num_match:
                # 编号在独立run中
                _add_highlight_to_run(first_run, hl_color)
                run_idx = list(target_elem).index(first_run)
                cs = OxmlElement('w:commentRangeStart')
                cs.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx, cs)
                ce = OxmlElement('w:commentRangeEnd')
                ce.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx + 2, ce)
            else:
                # 无编号前缀，高亮第一个run
                _add_highlight_to_run(first_run, hl_color)
                run_idx = list(target_elem).index(first_run)
                cs = OxmlElement('w:commentRangeStart')
                cs.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx, cs)
                ce = OxmlElement('w:commentRangeEnd')
                ce.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx + 2, ce)

        elif anchor_type == 'multi_level_num' and runs:
            # 多级编号：只标多级编号部分（如"1.2"、"2.1"）
            first_run = runs[0]
            first_text = ''.join(t.text or '' for t in first_run.iter(qn('w:t')))
            multi_match = re.match(r'^(\d+[.、．]\d+[.、．]?\s*)', first_text)
            if multi_match and len(multi_match.group(0)) < len(first_text):
                before, after = _split_run_at(target_elem, first_run, len(multi_match.group(0)))
                if before is not None:
                    _add_highlight_to_run(before, hl_color)
                    before_idx = list(target_elem).index(before)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(before_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(before_idx + 2, ce)
                else:
                    _add_highlight_to_run(first_run, hl_color)
                    run_idx = list(target_elem).index(first_run)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx + 2, ce)
            elif multi_match:
                _add_highlight_to_run(first_run, hl_color)
                run_idx = list(target_elem).index(first_run)
                cs = OxmlElement('w:commentRangeStart')
                cs.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx, cs)
                ce = OxmlElement('w:commentRangeEnd')
                ce.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx + 2, ce)
            else:
                _add_highlight_to_run(first_run, hl_color)
                run_idx = list(target_elem).index(first_run)
                cs = OxmlElement('w:commentRangeStart')
                cs.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx, cs)
                ce = OxmlElement('w:commentRangeEnd')
                ce.set(qn('w:id'), comment_id)
                target_elem.insert(run_idx + 2, ce)

        elif anchor_type == 'verb_after_num' and runs:
            # 编号后跟动词：标编号+动词（如"1.是"）
            first_run = runs[0]
            first_text = ''.join(t.text or '' for t in first_run.iter(qn('w:t')))
            # 匹配编号+动词（编号 + 1-2个字的动词）
            verb_match = re.match(r'^(\d+[.、．]\s*[\u662f\u8981\u6709\u80fd\u5e94\u4f1a\u53ef\u5c06\u5f97\u5fc5\u987b]{1,2})', first_text)
            if verb_match and len(verb_match.group(0)) < len(first_text):
                before, after = _split_run_at(target_elem, first_run, len(verb_match.group(0)))
                if before is not None:
                    _add_highlight_to_run(before, hl_color)
                    before_idx = list(target_elem).index(before)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(before_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(before_idx + 2, ce)
                else:
                    _add_highlight_to_run(first_run, hl_color)
                    run_idx = list(target_elem).index(first_run)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx + 2, ce)
            else:
                # 退化为只标编号
                num_match = re.match(r'^(\d+[.、．]\s*)', first_text)
                end_pos = len(num_match.group(0)) if num_match else min(len(first_text), 5)
                if end_pos < len(first_text):
                    before, after = _split_run_at(target_elem, first_run, end_pos)
                    if before is not None:
                        _add_highlight_to_run(before, hl_color)
                        before_idx = list(target_elem).index(before)
                        cs = OxmlElement('w:commentRangeStart')
                        cs.set(qn('w:id'), comment_id)
                        target_elem.insert(before_idx, cs)
                        ce = OxmlElement('w:commentRangeEnd')
                        ce.set(qn('w:id'), comment_id)
                        target_elem.insert(before_idx + 2, ce)
                    else:
                        _add_highlight_to_run(first_run, hl_color)
                        run_idx = list(target_elem).index(first_run)
                        cs = OxmlElement('w:commentRangeStart')
                        cs.set(qn('w:id'), comment_id)
                        target_elem.insert(run_idx, cs)
                        ce = OxmlElement('w:commentRangeEnd')
                        ce.set(qn('w:id'), comment_id)
                        target_elem.insert(run_idx + 2, ce)
                else:
                    _add_highlight_to_run(first_run, hl_color)
                    run_idx = list(target_elem).index(first_run)
                    cs = OxmlElement('w:commentRangeStart')
                    cs.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx, cs)
                    ce = OxmlElement('w:commentRangeEnd')
                    ce.set(qn('w:id'), comment_id)
                    target_elem.insert(run_idx + 2, ce)

        else:
            # full_para 或其他：高亮整段
            for r in runs:
                _add_highlight_to_run(r, hl_color)
            pPr = target_elem.find(qn('w:pPr'))
            if pPr is not None:
                insert_idx = list(target_elem).index(pPr) + 1
            else:
                insert_idx = 0
            cs = OxmlElement('w:commentRangeStart')
            cs.set(qn('w:id'), comment_id)
            target_elem.insert(insert_idx, cs)
            ce = OxmlElement('w:commentRangeEnd')
            ce.set(qn('w:id'), comment_id)
            target_elem.append(ce)

        # 插入 commentReference run（统一处理）
        ref_run = OxmlElement('w:r')
        ref_rPr = OxmlElement('w:rPr')
        ref_rStyle = OxmlElement('w:rStyle')
        ref_rStyle.set(qn('w:val'), 'CommentReference')
        ref_rPr.append(ref_rStyle)
        ref_run.append(ref_rPr)
        ref_cr = OxmlElement('w:commentReference')
        ref_cr.set(qn('w:id'), comment_id)
        ref_run.append(ref_cr)
        # 在 commentRangeEnd 之后插入
        ce_elem = target_elem.find(f'.//w:commentRangeEnd[@w:id="{comment_id}"]',
                                    namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if ce_elem is not None:
            ce_idx = list(target_elem).index(ce_elem)
            target_elem.insert(ce_idx + 1, ref_run)
        else:
            target_elem.append(ref_run)

        next_id += 1

    # 将 comments 保存为 Part
    comments_bytes = etree.tostring(comments_element, xml_declaration=True, encoding='UTF-8', standalone=True)
    doc_part = doc.part

    # 尝试获取已有的 comments part
    for rel in doc_part.rels.values():
        if 'comments' in rel.reltype:
            comments_part = rel.target_part
            comments_part._blob = comments_bytes
            return

    # 创建新的 Part（使用正确的 PackURI）
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    comments_partname = PackURI('/word/comments.xml')
    comments_part = Part(
        partname=comments_partname,
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        blob=comments_bytes,
        package=doc_part.package
    )
    doc_part.relate_to(comments_part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')


def _add_page_number(doc):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in footer.paragraphs:
        p.clear()
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = fp._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')

    def make_run(text_or_field, is_field=False):
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), FONT_TIMES_NEW_ROMAN)
        rFonts.set(qn('w:eastAsia'), '宋体')
        rFonts.set(qn('w:hAnsi'), FONT_TIMES_NEW_ROMAN)
        rPr.append(rFonts)
        for tag in ('w:sz', 'w:szCs'):
            sz = OxmlElement(tag)
            sz.set(qn('w:val'), '28')
            rPr.append(sz)
        r.append(rPr)
        if is_field:
            fld = OxmlElement('w:fldChar')
            fld.set(qn('w:fldCharType'), text_or_field)
            r.append(fld)
        else:
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = text_or_field
            r.append(t)
        return r

    def make_instrText(instr):
        r = OxmlElement('w:r')
        instrT = OxmlElement('w:instrText')
        instrT.set(qn('xml:space'), 'preserve')
        instrT.text = instr
        r.append(instrT)
        return r

    p_elem = fp._p
    p_elem.append(make_run('— ', False))
    p_elem.append(make_run('begin', True))
    p_elem.append(make_instrText(' PAGE '))
    p_elem.append(make_run('separate', True))
    p_elem.append(make_run('end', True))
    p_elem.append(make_run(' —', False))


def _check_punctuation_issues(paragraphs_text):
    """句末标点检测：找出未以句号/问号/叹号结尾的正文段落"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text or len(text) <= 10:
            continue
        # 跳过主标题（主标题不应有句末标点，不应被检测）
        if is_main_title(text):
            continue
        # 标记是否为编号型段落（用于后续判断）
        is_numbered = bool(re.match(r'^\d+[.、．]\s*', text))
        # 跳过标题型文本
        if re.match(r'^[一二三四五六七八九十]+、', text):
            continue
        if re.match(r'^（[一二三四五六七八九十]+）', text):
            continue
        if re.match(r'^（\d+）', text):
            continue
        if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text):
            continue
        # 跳过纯编号+短标题行（如 "1.科技部"，标题不应有句末标点）
        # 但数字编号的正文（含逗号等内部标点）应检查句末标点
        if is_numbered and len(text) <= 25 and not re.search(r'[，、；]', text):
            continue
        # 跳过纯数字/百分比/短数据行
        if re.match(r'^[\d,.\-+%：:（）()]+$', text):
            continue
        # 跳过包含冒号结尾的引导句（如 "科技部："、"商务部："）
        if re.search(r'[：:]$', text):
            continue
        # 跳过纯短数据行（无中文内容且<=15字）
        if len(text) <= 15 and not re.search(r'[\u4e00-\u9fff]', text):
            continue
        # 跳过无标点的短标题行（<=20字且不含任何句内标点、不含中文字数>5）
        if len(text) <= 20 and not re.search(r'[，。；！？]', text):
            continue
        # 检测句末标点
        last_char = text[-1]
        if last_char not in ('。', '？', '！', '…', '"', '"', ')', '）', '；'):
            issues.append((i, text[:60]))
    return issues


def _check_subheading_issues(paragraphs_text):
    """子标题序号混乱检测：原文含 X.Y 格式但被当作普通段落处理"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text:
            continue
        # 检测 X.Y 格式开头（如 1.1, 2.3）
        m = re.match(r'^(\d+)\.(\d+)[.、．]?\s*(.*)', text)
        if m:
            major = int(m.group(1))
            minor = int(m.group(2))
            content = m.group(3)
            # 排除时间数字（如 "1.2025年" 中的 2025 是年份，不是多级编号）
            # 更精确的判断：minor是年份(2000-2030)且content以"年"开头，才是时间数字
            if 2000 <= minor <= 2030 and content.startswith('年'):
                continue
            if minor > 0:
                issues.append((i, text[:60], major, minor, content))
    return issues


def _check_h3_numbering_issues(paragraphs_text):
    """三级标题编号不规范检测：如 '1.是' '2.是' 应为 '一是' '二是'"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text:
            continue
        # 检测 X.是/且/但/将/要 等不规范的三级标题编号
        m = re.match(r'^(\d+)[.、．]\s*(是|且|但|将|要|在|已|以|对|为|从|按|于)\s*(.*)', text)
        if m:
            num = int(m.group(1))
            word = m.group(2)
            rest = m.group(3)
            issues.append((i, text[:70], num, word))
    return issues


def _check_word_numbering_format(paragraphs_text, num_to_abstract, abstract_num_defs):
    """检测Word自动编号使用阿拉伯数字（1.2.3.）且文本像标题的情况，建议改为中文二级标题格式（（一）（二））"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text:
            continue
        # 只检查未识别为标题的正文段落（level=None）
        level = item[3] if len(item) > 3 else None
        if level is not None:
            continue
        orig_num_id = item[4] if len(item) > 4 else None
        if not orig_num_id or orig_num_id == '0':
            continue
        # 检查是否为十进制编号（阿拉伯数字1.2.3.）
        an_id = num_to_abstract.get(orig_num_id)
        if not an_id:
            continue
        levels = abstract_num_defs.get(an_id, {})
        nilvl = item[5] if len(item) > 5 else '0'
        fmt, lvl_txt = levels.get(nilvl, (None, None))
        if fmt != 'decimal':
            continue
        # 是十进制编号，检查文本是否像标题（短文本、含冒号、无句末标点）
        is_like_heading = (
            len(text) <= 40
            and not re.search(r'[。；！？]', text)
        )
        if is_like_heading:
            issues.append((i, text[:50]))
    return issues


def _check_missing_h2(paragraphs_text):
    """检测一级标题下直接使用三级标题的情况，建议补充二级标题。
    
    注意区分：
    - 短文本数字编号（如"1.科技部"）→ 是标题，一级跳三级应提示缺少二级标题
    - 长文本数字编号（如"1.本年度节后新开项目1个..."）→ 是正文，不提示跳级
    """
    issues = []
    last_h1_index = None
    
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        level = item[3] if len(item) > 3 else None
        num_id = item[4] if len(item) > 4 else None
        text = item[1].strip() if len(item) > 1 else ''
        
        # 补充 detect_level：文本编号前缀的段落 item[3] 可能是 None 或 PENDING，
        # 但文本上匹配 h1/h3 模式
        detected = detect_level(text)
        # 综合判断：优先用 item[3] 中的最终 level，否则用 detect_level
        effective_level = level if level in ('h1', 'h2', 'h3', 'h4', 'h5') else detected
        
        # 判断数字编号段落是"标题"还是"正文"
        # 标准：短文本(≤25字)且不含句号/分号 → 标题；否则 → 正文
        is_digit_prefix = bool(re.match(r'^\d+[.、．]', text))
        is_likely_body = is_digit_prefix and (
            len(text) > 25 or '。' in text or '；' in text
        )
        
        if effective_level == 'h1':
            last_h1_index = i
        elif effective_level == 'h3' and last_h1_index is not None:
            # 只有"标题型"三级标题才检查跳级，正文型不检查
            if is_likely_body:
                # 正文列表项，不提示跳级
                continue
            
            # 检查 last_h1_index 和 i 之间是否有 h2
            has_h2_between = False
            for j in range(last_h1_index + 1, i):
                if paragraphs_text[j][0] == 'p':
                    between_level = paragraphs_text[j][3] if len(paragraphs_text[j]) > 3 else None
                    between_text = paragraphs_text[j][1].strip() if len(paragraphs_text[j]) > 1 else ''
                    between_detected = detect_level(between_text)
                    between_effective = between_level if between_level in ('h1', 'h2', 'h3', 'h4', 'h5') else between_detected
                    if between_effective == 'h2':
                        has_h2_between = True
                        break
            
            if not has_h2_between:
                issues.append((i, text[:30]))
                # 不重置 last_h1_index，同一一级标题下可能有多个跳级标题
        elif last_h1_index is not None and num_id and num_id != '0':
            # 扩展检测：一级标题后跟的段落使用了Word自动编号，且看起来像标题
            looks_like_title = (
                len(text) <= 30
                and not re.search(r'[。；！？]', text)
            )
            if looks_like_title:
                issues.append((i, text[:30]))
                # 不重置 last_h1_index，同一一级标题下可能有多个跳级标题
        elif effective_level == 'h2':
            last_h1_index = None  # 有 h2，重置
    
    return issues


def _check_title_punctuation(paragraphs_text):
    """检测标题编号后的标点是否符合规范：
    - 一级标题：编号后接顿号（、），如"一、"
    - 二级标题：编号后无标点，如"（一）"
    - 三级标题：编号后接点号（.），如"1."
    返回问题列表：[(段落索引, 标题文本, 错误类型, 建议), ...]
    """
    issues = []
    
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        
        level = item[3] if len(item) > 3 else None
        text = item[1].strip() if len(item) > 1 else ''
        
        # 如果 level 是 None，尝试使用正则检测标题层级
        if level is None and text:
            # 尝试检测标题层级
            if re.match(r'^[一二三四五六七八九十]+[、】]', text):
                level = 'h1'
            elif re.match(r'^（[一二三四五六七八九十]+）', text):
                level = 'h2'
            elif re.match(r'^\d+[.、．](?!\d)\s*\S', text):
                level = 'h3'
        
        # 只检查已识别为标题的段落
        if level not in ('h1', 'h2', 'h3'):
            continue
        
        # 跳过空文本
        if not text:
            continue
        
        # 检查标题编号后的标点
        if level == 'h1':
            # 一级标题：应该是一、二、三、...（中文数字+顿号）
            # 匹配模式：中文数字开头的标题
            match = re.match(r'^([一二三四五六七八九十]+)([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                
                if punct != '、':
                    if punct:
                        issues.append((i, text[:40], 'h1_wrong_punct',
                            f'一级标题编号"{num_part}"后应为顿号"、"，实为"{punct}"'))
                    else:
                        issues.append((i, text[:40], 'h1_missing_punct',
                            f'一级标题编号"{num_part}"后缺少顿号"、"'))
        
        elif level == 'h2':
            # 二级标题：应该是（一）（二）（三）...（括号后无标点）
            # 匹配模式：（中文数字）开头的标题
            match = re.match(r'^（([一二三四五六七八九十]+)）([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                
                if punct:
                    issues.append((i, text[:40], 'h2_extra_punct',
                        f'二级标题"（{num_part}）"后不应有标点，检测到"{punct}"'))
        
        elif level == 'h3':
            # 三级标题：应该是1. 2. 3.（阿拉伯数字+点号）
            # 匹配模式：阿拉伯数字开头的标题
            match = re.match(r'^(\d+)([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                
                if punct != '.':
                    if punct:
                        issues.append((i, text[:40], 'h3_wrong_punct',
                            f'三级标题编号"{num_part}"后应为点号"."，实为"{punct}"'))
                    else:
                        issues.append((i, text[:40], 'h3_missing_punct',
                            f'三级标题编号"{num_part}"后缺少点号"."'))
    
    return issues


def format_document(src_path: str, dst_path: str):
    """主转换函数：读取 → 应用公文格式 → 保存。
    
    返回格式：(dst_path, warnings_list)
    warnings_list 中每项为 dict，包含 type 和 detail 字段。
    """
    ext = Path(src_path).suffix.lower()
    paragraphs_text = []
    warnings = []

    if ext == '.docx':
        src_doc = Document(src_path)
        body = src_doc.element.body
        ordered_elements = []
        for child in body:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'p':
                ordered_elements.append(('p', child))
            elif tag == 'tbl':
                ordered_elements.append(('tbl', child))

        abstract_num_defs = {}
        num_to_abstract = {}
        try:
            numbering_part = src_doc.part.numbering_part
            if numbering_part:
                num_root = numbering_part._element
                for abstract_num in num_root.findall(qn('w:abstractNum')):
                    an_id = abstract_num.get(qn('w:abstractNumId'))
                    levels = {}
                    for lvl in abstract_num.findall(qn('w:lvl')):
                        ilvl_val = lvl.get(qn('w:ilvl'), '0')
                        num_fmt = lvl.find(qn('w:numFmt'))
                        lvl_text = lvl.find(qn('w:lvlText'))
                        if num_fmt is not None and lvl_text is not None:
                            levels[ilvl_val] = (
                                num_fmt.get(qn('w:val'), ''),
                                lvl_text.get(qn('w:val'), '')
                            )
                    abstract_num_defs[an_id] = levels
                for num in num_root.findall(qn('w:num')):
                    numId = num.get(qn('w:numId'))
                    aRef = num.find(qn('w:abstractNumId'))
                    if aRef is not None:
                        num_to_abstract[numId] = aRef.get(qn('w:val'))
        except Exception:
            pass

        for etype, elem in ordered_elements:
            if etype == 'p':
                text = ''.join(t.text for t in elem.iter(qn('w:t')) if t.text)
                # Bug1修复：只有整段所有文本run都加粗时才标记is_bold=True
                # 避免"部分加粗扩散为全段加粗"的问题
                _runs_with_text = []
                for _r in elem.findall(qn('w:r')):
                    _t = ''.join(t.text for t in _r.iter(qn('w:t')) if t.text)
                    if _t.strip():
                        _runs_with_text.append(_r)
                if _runs_with_text:
                    _all_bold = True
                    for _r in _runs_with_text:
                        _rPr = _r.find(qn('w:rPr'))
                        if _rPr is None:
                            _all_bold = False
                            break
                        _b = _rPr.find(qn('w:b'))
                        if _b is None or _b.get(qn('w:val'), 'true') in ('false', '0'):
                            _all_bold = False
                            break
                    is_bold = _all_bold
                else:
                    is_bold = False
                num_id = num_ilvl = None
                pPr = elem.find(qn('w:pPr'))
                if pPr is not None:
                    numPr = pPr.find(qn('w:numPr'))
                    if numPr is not None:
                        nid_el = numPr.find(qn('w:numId'))
                        ilvl_el = numPr.find(qn('w:ilvl'))
                        if nid_el is not None:
                            num_id = nid_el.get(qn('w:val'))
                        if ilvl_el is not None:
                            num_ilvl = ilvl_el.get(qn('w:val'))
                word_num_level = 'PENDING' if (num_id and num_id != '0') else None
                paragraphs_text.append(('p', text, is_bold, word_num_level, num_id, num_ilvl))
            elif etype == 'tbl':
                rows_data = []
                rows_shading = []  # 保存每个cell的原始底色
                rows_font_color = []  # 保存每个cell的原始字体颜色
                for tr in elem.iter(qn('w:tr')):
                    cells = []
                    shadings = []
                    font_colors = []
                    for tc in tr.iter(qn('w:tc')):
                        cell_text = ''.join(t.text for t in tc.iter(qn('w:t')) if t.text)
                        cells.append(cell_text.strip())
                        # 提取原始单元格底色
                        tcPr = tc.find(qn('w:tcPr'))
                        cell_shading = None
                        if tcPr is not None:
                            shd = tcPr.find(qn('w:shd'))
                            if shd is not None:
                                cell_shading = shd.get(qn('w:fill'))
                        shadings.append(cell_shading)
                        # 提取原始单元格字体颜色（取第一个run的颜色）
                        cell_font_color = None
                        for p_elem in tc.iter(qn('w:p')):
                            for r_elem in p_elem.iter(qn('w:r')):
                                rPr = r_elem.find(qn('w:rPr'))
                                if rPr is not None:
                                    color_elem = rPr.find(qn('w:color'))
                                    if color_elem is not None:
                                        val = color_elem.get(qn('w:val'))
                                        if val:
                                            cell_font_color = val
                                            break
                            if cell_font_color:
                                break
                        font_colors.append(cell_font_color)
                    if cells:
                        rows_data.append(cells)
                        rows_shading.append(shadings)
                        rows_font_color.append(font_colors)
                if rows_data:
                    paragraphs_text.append(('tbl', rows_data, rows_shading, rows_font_color))

        # ──── 预合并：合并被Word拆分的标题碎片 ────
        # 临时禁用：此逻辑导致标题段落被错误合并，编号错乱
        # 直接跳过预合并，使用原始段落列表
        pass  # 下面的while循环已被禁用
        while False and i < len(paragraphs_text):
            if merged[i]:
                i += 1
                continue
            item = paragraphs_text[i]
            if item[0] != 'p':
                i += 1
                continue
            text = item[1]
            # 碎片判断：短文本+无句内标点+无编号前缀+非纯数字+非问候语+不以逗号/顿号/冒号结尾
            # 重要：标题末尾的冒号不应被合并（如"（一）项目上线与业财集成："）
            greeting_kw = '领导|同事|各位|尊敬|您好|下午好|上午好|你好'
            is_greeting_text = (
                re.match(r'^.{2,30}[：:]$', text.strip())
                and re.search(greeting_kw, text.strip())
            )
            # 排除以逗号/顿号/冒号结尾的碎片
            # - 逗号/顿号结尾：句子中间的分隔，不是标题续行
            # - 冒号结尾：标题（如"（一）xxx："），不是正文碎片
            ends_with_sep = bool(re.search(r'[，、：:]\s*$', text))
            is_fragment = (
                1 < len(text) <= 25
                and not re.search(r'[。；！？]', text)
                and not ends_with_sep
                and not re.match(r'^[一二三四五六七八九十]+、', text)
                and not re.match(r'^\d+[.、．]', text)
                and not re.match(r'^（\d+）', text)
                and not re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
                and not re.match(r'^[\d,.\-+%：:（）()]+$', text)
                and not is_greeting_text
            )
            if is_fragment and i > 0 and not merged[i - 1]:
                prev = paragraphs_text[i - 1]
                prev_text = prev[1] if prev[0] == 'p' else ''
                # 跳过空段落，向前继续找
                if not prev_text.strip():
                    i += 1
                    continue
                # 前段也是短碎片（≤25字）且不以标点结尾 → 合并
                # 但如果前段明显较长（>20字），不合并
                prev_is_fragment = (prev[0] == 'p' and 1 < len(prev_text) <= 25
                                    and not re.search(r'[。；！？]', prev_text))
                # 保护：如果前段明显较长（>20字），不合并
                prev_is_long_fragment = (prev[0] == 'p' and len(prev_text) > 20)
                # 临时修复：禁用merge_to_main_title，防止标题段落被错误合并
                merge_to_main_title = False
                # 原代码（暂时禁用）：
                # prev_is_bold_main = (prev[0] == 'p' and prev[2] is True)
                # prev_is_main = (prev[0] == 'p' and is_main_title(prev_text))
                # merge_to_main_title = (prev_is_bold_main or prev_is_main) and len(text) <= 15
                # 禁止合并：当前段或前段看起来像独立条目（含冒号的数据行、月底节点等）
                # 关键修复：禁止合并包含编号前缀的段落（如"1." "2." "（一）"等）
                has_num_prefix_current = bool(
                    re.match(r'^[一二三四五六七八九十]+、', text)
                    or re.match(r'^（[一二三四五六七八九十]+）', text)
                    or re.match(r'^\d+[.、．]\s*', text)
                    or re.match(r'^（\d+）', text)
                    or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
                )
                has_num_prefix_prev = bool(
                    re.match(r'^[一二三四五六七八九十]+、', prev_text)
                    or re.match(r'^（[一二三四五六七八九十]+）', prev_text)
                    or re.match(r'^\d+[.、．]\s*', prev_text)
                    or re.match(r'^（\d+）', prev_text)
                    or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', prev_text)
                )
                looks_like_list_item = (
                    bool(re.search(r'[：:]', text))
                    or bool(re.search(r'[：:]', prev_text))
                    or has_num_prefix_current
                    or has_num_prefix_prev
                )
                if (prev_is_fragment and not prev_is_long_fragment and not looks_like_list_item) or merge_to_main_title:
                    # 合并到前一段
                    merged_text = prev_text + text
                    paragraphs_text[i - 1] = (prev[0], merged_text, prev[2], prev[3], prev[4], prev[5])
                    merged[i] = True
                    i += 1
                    continue
            i += 1
        # 预合并已被禁用，直接设置merged为全False
        merged = [False] * len(paragraphs_text)
        # 过滤掉已合并的碎片
        paragraphs_text = [item for i, item in enumerate(paragraphs_text) if not merged[i]]

    elif ext in ('.txt', '.md'):
        with open(src_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f.read().splitlines():
                paragraphs_text.append(('p', line, False, None, None, None))
    else:
        raise ValueError(f"不支持的文件格式：{ext}（仅支持 .docx .txt .md）")

    # 建立 numId → ilvl → 公文层级映射
    numid_ilvl_level_map = {}
    used_num_ids = {
        item[4] for item in paragraphs_text
        if item[0] == 'p' and len(item) > 4 and item[4] and item[4] != '0'
    }
    sorted_num_ids = sorted(used_num_ids, key=lambda x: int(x) if x.isdigit() else 0)

    for nid in sorted_num_ids:
        an_id = num_to_abstract.get(nid)
        if an_id is None:
            continue
        levels = abstract_num_defs.get(an_id, {})
        for ilvl, (fmt, txt) in levels.items():
            # decimal 格式（1. 2. 3.）通常是正文编号列表，不应映射为标题层级
            if fmt == 'decimal':
                continue
            if fmt in ('chineseCounting', 'chineseCountingThousand',
                       'upperLetter', 'lowerLetter',
                       'ideographDigital', 'ideographEnclosedCircle'):
                mapping = {'0': 'h1', '1': 'h2', '2': 'h3'}
                if ilvl in mapping:
                    numid_ilvl_level_map[(nid, ilvl)] = mapping[ilvl]
            else:
                if txt.startswith('（') or txt.startswith('('):
                    mapping = {'0': 'h2', '1': 'h3', '2': 'h4'}
                elif txt.endswith('.') or txt.endswith('、') or txt.endswith('．'):
                    mapping = {'0': 'h3', '1': 'h4'}
                else:
                    mapping = {'0': 'h1'}
                if ilvl in mapping:
                    numid_ilvl_level_map[(nid, ilvl)] = mapping[ilvl]

    if not numid_ilvl_level_map:
        # 仅当 numId 对应的格式不是 decimal 时才创建回退映射
        # decimal 格式的编号列表通常是正文，不应映射为标题
        for order, nid in enumerate(sorted_num_ids):
            # 检查该 numId 对应的格式是否为 decimal
            an_id = num_to_abstract.get(nid)
            is_decimal_only = False
            if an_id:
                levels = abstract_num_defs.get(an_id, {})
                fmts = [fmt for fmt, _ in levels.values()]
                if fmts and all(f == 'decimal' for f in fmts):
                    is_decimal_only = True
            if is_decimal_only:
                continue  # decimal 格式不映射为标题
            level_map = {0: 'h1', 1: 'h2', 2: 'h3'}
            if order in level_map:
                numid_ilvl_level_map[(nid, '0')] = level_map[order]

    # 回填 PENDING
    # 注意：对于十进制编号（numId对应decimal格式）的段落，应谨慎识别为标题
    # Word的十进制编号常用于正文中的编号列表，不应轻易识别为标题层级
    for idx, item in enumerate(paragraphs_text):
        if item[0] == 'p' and item[3] == 'PENDING':
            nid, nilvl = item[4], item[5] or '0'
            resolved_level = numid_ilvl_level_map.get((nid, nilvl))
            text_content = item[1] if len(item) > 1 else ''
            
            # 如果numId对应的是十进制编号（decimal），应谨慎处理
            # 只有当文本看起来像标题（短且无句末标点）时才识别为标题
            # 否则应识别为正文
            an_id = num_to_abstract.get(nid)
            if an_id:
                levels = abstract_num_defs.get(an_id, {})
                fmt, _ = levels.get(nilvl, (None, None))
                if fmt == 'decimal':
                    # 判断文本是否像标题（短、无句末标点、无数字前缀）
                    looks_like_title = (
                        len(text_content) <= 30
                        and not re.search(r'[。；！？]', text_content)
                        and not re.match(r'^\d+[.、．]', text_content)
                    )
                    # 即使看起来像标题，也要验证文本是否有对应格式的编号前缀
                    # 例如：chineseCounting 格式应有"一、"等前缀，decimal 格式应有"1."等前缀
                    if looks_like_title:
                        # 检查文本是否有对应格式的编号前缀
                        has_correct_prefix = False
                        if fmt == 'chineseCounting':
                            # 检查是否有中文数字前缀（一、 二、 三、）
                            has_correct_prefix = bool(re.match(r'^[一二三四五六七八九十]+[、）]', text_content))
                        elif fmt == 'decimal':
                            # 检查是否有阿拉伯数字前缀（1. 2. 3.）
                            has_correct_prefix = bool(re.match(r'^\d+[.、．]', text_content))
                        
                        if not has_correct_prefix:
                            resolved_level = None  # 无编号前缀，不是标题
                    else:
                        resolved_level = None  # 不像标题，肯定是正文
            
            paragraphs_text[idx] = (
                item[0], item[1], item[2],
                resolved_level,
                item[4], item[5]
            )

    # 运行审查检测
    punct_issues = _check_punctuation_issues(paragraphs_text)
    if punct_issues:
        warnings.append({
            'type': '句末标点缺失',
            'detail': f'共 {len(punct_issues)} 处段落可能缺少句末标点（。）：\n'
                      + '\n'.join(f'  段落{idx+1}: "{txt}…"' for idx, txt in punct_issues)
        })

    subhead_issues = _check_subheading_issues(paragraphs_text)
    if subhead_issues:
        detail_lines = []
        for idx, txt, major, minor, content in subhead_issues:
            detail_lines.append(f'  段落{idx+1}: "{txt}" — 原文使用 {major}.{minor} 多级编号')
        detail_lines.append('  建议：可将多级编号改为四级标题①②③格式，请人工确认。')
        warnings.append({
            'type': '子标题序号格式',
            'detail': f'检测到 {len(subhead_issues)} 处多级编号（X.Y格式）：\n' + '\n'.join(detail_lines)
        })

    h3_issues = _check_h3_numbering_issues(paragraphs_text)
    if h3_issues:
        detail_lines = []
        for idx, txt, num, word in h3_issues:
            detail_lines.append(f'  段落{idx+1}: "{txt}" — 编号 "{num}." 后直接跟"{word}"')
        detail_lines.append('  建议：此类编号建议改为 "一是…""二是…" 格式，请人工确认。')
        warnings.append({
            'type': '三级标题编号不规范',
            'detail': f'检测到 {len(h3_issues)} 处编号后直接跟动词：\n' + '\n'.join(detail_lines)
        })

    # 检查Word自动编号格式（阿拉伯数字1.2.3.）— Bug 1修复
    word_num_issues = _check_word_numbering_format(paragraphs_text, num_to_abstract, abstract_num_defs)
    word_num_indices = {i for i, _ in word_num_issues}
    
    # 检查一级标题下直接使用三级标题的情况
    missing_h2_issues = _check_missing_h2(paragraphs_text)
    missing_h2_indices = {i for i, _ in missing_h2_issues}  # 用于批注
    if word_num_issues:
        detail_lines = []
        for idx, txt in word_num_issues:
            detail_lines.append(f'  段落{idx+1}: "{txt}" — 使用Word自动编号（1. 2. 3.）')
        detail_lines.append('  建议：可改为二级标题（（一）（二））格式，请人工确认。')
        warnings.append({
            'type': 'Word自动编号格式',
            'detail': f'检测到 {len(word_num_issues)} 处Word自动编号（阿拉伯数字格式）：\n' + '\n'.join(detail_lines)
        })
    
    # 检查一级标题下直接使用三级标题的情况
    if missing_h2_issues:
        detail_lines = []
        for idx, txt in missing_h2_issues:
            detail_lines.append(f'  段落{idx+1}: "{txt}" — 一级标题下直接使用三级标题')
        detail_lines.append('  建议：在一级标题和三级标题之间补充二级标题（一）（二）')
        warnings.append({
            'type': '缺少二级标题',
            'detail': f'检测到 {len(missing_h2_issues)} 处一级标题下直接使用三级标题：\n' + '\n'.join(detail_lines)
        })

    # 检查标题编号标点是否符合规范
    title_punct_check_results = _check_title_punctuation(paragraphs_text)
    title_punct_indices = {i for i, _, _, _ in title_punct_check_results}
    
    if title_punct_check_results:
        detail_lines = []
        for idx, txt, issue_type, suggestion in title_punct_check_results:
            detail_lines.append(f'  段落{idx+1}: "{txt}" — {suggestion}')
        warnings.append({
            'type': '标题编号标点不规范',
            'detail': f'检测到 {len(title_punct_check_results)} 处标题编号标点不符合规范：\n' + '\n'.join(detail_lines)
        })
    
    # ──── 标题句末标点检测 ────
    # 规则：只有真正的多级标题（一、二、三、及其子标题）末尾有标点才需要提示
    # 正文编号列表（如"1. xxx 2. xxx"）和Word原生编号段落不算标题
    title_punct_issues = []
    for idx, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text or len(text) <= 5:
            continue
        # 检测各级标题（文本编号前缀）
        level = detect_level(text)
        wnl = item[3] if len(item) > 3 else None
        # 有文本编号前缀才算标题候选
        has_text_prefix = bool(
            re.match(r'^[一二三四五六七八九十]+、', text)
            or re.match(r'^（[一二三四五六七八九十]+）', text)
            or re.match(r'^\d+[.、．](?!\d)\s*\S', text)
            or re.match(r'^（\d+）', text)
            or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
        )
        if not has_text_prefix:
            continue

        # ── 区分"数字编号的标题"和"数字编号的正文" ──
        # 判断标准：短文本（≤25字）且不含句号/分号 → 标题，应检查句末标点
        #           长文本（>25字）或含句号/分号 → 正文列表项，跳过
        is_digit_prefix = bool(re.match(r'^\d+[.、．]', text))
        if is_digit_prefix:
            # 去掉编号前缀后看纯内容长度
            content_after_num = re.sub(r'^\d+[.、．]\s*', '', text)
            is_likely_body = (
                len(text) > 25
                or '。' in text
                or '；' in text
                or (len(content_after_num) > 15 and text.rstrip()[-1] == '。')
            )
            if is_likely_body:
                continue  # 正文列表项，跳过标题句末标点检查

        # Word编号段落（有wnl）通常不是真正的多级标题（而是正文列表）
        # 排除：1) 有Word原生编号的段落，2) 无文本前缀的长段落
        is_word_num_body = (wnl is not None and not has_text_prefix and len(text) > 25)
        if is_word_num_body:
            continue
        # 有编号前缀但内容超长（>30字且有句号或分号）的是正文不是标题
        if len(text) > 30 and ('。' in text or '；' in text):
            continue
        # 关键排除：Word原生十进制编号的正文列表（如"需协同解决的事项"下的1. 2.）
        # 这些段落有numId且对应decimal格式，是正文编号列表，不是标题
        orig_num_id = item[4] if len(item) > 4 else None
        if orig_num_id and orig_num_id != '0':
            # 检查是否是对应decimal格式的编号列表
            an_id = num_to_abstract.get(orig_num_id)
            if an_id:
                levels = abstract_num_defs.get(an_id, {})
                nilvl = item[5] if len(item) > 5 else '0'
                fmt, _ = levels.get(nilvl, (None, None))
                # decimal格式的Word原生编号通常是正文列表，排除
                if fmt == 'decimal':
                    continue
        # 标题应以非句号结尾，如果以句号、分号、逗号、冒号结尾则标记
        last_char = text.rstrip()[-1]
        if last_char in ('。', '；', '，', '：', ':'):
            title_punct_issues.append((idx, text, last_char))

    # 建立审查高亮标记集合
    punct_para_indices = {idx for idx, _ in punct_issues}          # 句末标点缺失
    subhead_para_indices = {idx for idx, _, _, _, _ in subhead_issues}  # X.Y多级编号
    h3_para_indices = {idx for idx, _, _, _ in h3_issues}          # X.是编号
    title_punct_para_indices = {idx for idx, _, _ in title_punct_issues}  # 标题句末标点

    # 标题句末标点检测
    if title_punct_issues:
        detail_lines = []
        for idx, txt, punct in title_punct_issues:
            detail_lines.append(f'  段落{idx+1}: "{txt}" — 标题末尾不应有"{punct}"')
        warnings.append({
            'type': '标题句末标点',
            'detail': f'检测到 {len(title_punct_issues)} 处标题包含句末标点（标题末尾不应有标点符号）：\n'
                      + '\n'.join(detail_lines)
        })

    # 新建文档
    doc = Document()
    comment_list = []  # 收集批注: [(匹配文本前缀, comment_text), ...]
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin    = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM
    section.left_margin   = MARGIN_LEFT
    section.right_margin  = MARGIN_RIGHT

    normal_style = doc.styles['Normal']
    normal_style.font.name = FONT_TIMES_NEW_ROMAN
    normal_style.font.size = SIZE_SANHAO
    nf = normal_style.element.find('.//' + qn('w:rFonts'))
    if nf is None:
        rPr = normal_style.element.find('.//' + qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            normal_style.element.append(rPr)
        nf = OxmlElement('w:rFonts')
        rPr.insert(0, nf)
    nf.set(qn('w:eastAsia'), FONT_FANGSONG)

    # ──── 预扫描：检测文档是否用 X. 作为顶层编号（无中文一、二、三、） ────
    has_cn_h1 = False
    for item in paragraphs_text:
        if item[0] == 'p':
            t = item[1].strip()
            if re.match(r'^[一二三四五六七八九十]+、', t):
                has_cn_h1 = True
                break
    # 如果没有中文一级编号，但存在 X、或 X.（非X.Y）编号段落，
    # 且 X. 后面的子标题是 （X）或 (X)，则提升所有 X. 为 h1
    promote_x_to_h1 = False
    promote_body_indices = set()  # 需要提升为 h1 的 body 级段落索引
    if not has_cn_h1:
        x_prefix_paras = []  # X. 或 X、 开头的段落索引
        for idx, item in enumerate(paragraphs_text):
            if item[0] == 'p':
                t = item[1].strip()
                if re.match(r'^\d+[.、．]\s*\S', t) and not re.match(r'^\d+\.\d+', t):
                    x_prefix_paras.append(idx)
        if x_prefix_paras:
            # 检查任意一个 X. 段落后是否有 （X）或 (X) 子标题
            has_sub_level = False
            for xpi in x_prefix_paras[:5]:
                for j in range(xpi + 1, min(xpi + 6, len(paragraphs_text))):
                    if paragraphs_text[j][0] == 'p':
                        sub_t = paragraphs_text[j][1].strip()
                        if re.match(r'^（\d+）', sub_t) or re.match(r'^[（(]\d+[)）]', sub_t):
                            has_sub_level = True
                            break
                if has_sub_level:
                    break
            # 只要有子标题结构模式，就提升所有 X.（非X.Y）为 h1
            if has_sub_level and len(x_prefix_paras) >= 2:
                promote_x_to_h1 = True
                # 确定标题区域结束位置（首个 X. 前缀段落之前）
                title_end = min(x_prefix_paras) if x_prefix_paras else 0
                # 收集需要提升的 body 段落：仅限无 Word 编号、无编号前缀、
                # 长度极短（<=15字）且夹在两个 X. 段落之间的纯标题行
                for idx, item in enumerate(paragraphs_text):
                    if idx < title_end:
                        continue  # 跳过标题区域
                    if item[0] == 'p' and item[1].strip():
                        t = item[1].strip()
                        # 有 Word 编号的段落一概不提升
                        is_word_num = (len(item) > 4 and item[4] and item[4] != '0')
                        if is_word_num:
                            continue
                        # 只提升极短的无编号标题（<=15字，排除长句）
                        is_short = (len(t) <= 15
                                    and not re.search(r'[。；]', t)
                                    and not re.search(r'^（\d+）', t)
                                    and not t.startswith('附件'))
                        if not is_short:
                            continue
                        # 检查前后是否有 X. 前缀段落（严格只看 X. 段落，不看 Word 编号段落）
                        has_x_neighbor = False
                        for offset in (-1, -2, 1, 2):
                            ni = idx + offset
                            if 0 <= ni < len(paragraphs_text) and paragraphs_text[ni][0] == 'p':
                                n_t = paragraphs_text[ni][1].strip()
                                if re.match(r'^\d+[.、．]\s*\S', n_t) and not re.match(r'^\d+\.\d+', n_t):
                                    has_x_neighbor = True
                                    break
                        if has_x_neighbor:
                            promote_body_indices.add(idx)

    title_mode = True
    title_count = 0  # 连续标题段计数，防止将正文标题误判为主标题
    title_ended = False  # 标记标题区是否已结束
    counter = HeadingCounter()
    merged_titles = set()  # 已合并到主标题的段落索引集合

    # 预计算：标记哪些段落索引最终是标题（用于空行过滤）
    is_heading_index = set()

    def _precompute_heading(idx, item):
        """预计算单个段落的最终层级（不含 counter），返回 level"""
        if item[0] != 'p':
            return None
        raw_text = item[1]
        text = clean_text(raw_text)
        if not text:
            return None
        if is_main_title(text):
            return 'title'
        level = detect_level(text)
        wnl = item[3] if len(item) > 3 else None
        # Bug2修复：优先使用detect_level的文本编号模式识别结果
        # 只有detect_level返回body（无法从文本识别标题）时，才用wnl作回退
        # 避免Word编号映射错误覆盖正确的文本识别（如"四、"被wnl='h4'覆盖）
        if wnl is not None and level == 'body':
            level = wnl
        if promote_x_to_h1 and level == 'h3':
            is_long_sentence = '。' in text and len(text) > 30
            if (re.match(r'^\d+[.、．]\s*\S', text)
                and not re.match(r'^\d+\.\d+', text)
                and not re.match(r'^\d+[.、．]\s*[是是以要为将把让使被]', text)
                and not is_long_sentence):
                level = 'h1'
        if promote_x_to_h1 and level == 'body' and wnl is None and idx in promote_body_indices:
            level = 'h1'
        # Word 编号隐藏了编号的短段落（无文本前缀 ≤15字），promote为 h1
        if promote_x_to_h1 and wnl is not None and wnl in ('h1', 'h2', 'h3', 'h4'):
            has_text_prefix = bool(
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
            )
            if not has_text_prefix and len(text) <= 15 and not re.search(r'[。；]', text):
                level = 'h1'
        # Word 编号段落中，文本无编号前缀且长度 >25 字的，不算标题
        # Bug2修复：但如果Word编号格式是标题型格式（chineseCounting等），不降级
        if wnl is not None and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            has_text_prefix = bool(
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
            )
            if not has_text_prefix and len(text) > 25:
                # 检查Word编号格式
                _orig_nid = item[4] if len(item) > 4 else None
                _orig_nilvl = item[5] if len(item) > 5 else '0'
                _is_heading_num_format = False
                if _orig_nid and _orig_nid != '0':
                    _an_id = num_to_abstract.get(_orig_nid)
                    if _an_id:
                        _levels = abstract_num_defs.get(_an_id, {})
                        _fmt, _ = _levels.get(_orig_nilvl or '0', (None, None))
                        if _fmt in ('chineseCounting', 'chineseCountingThousand',
                                    'ideographDigital', 'ideographEnclosedCircle'):
                            _is_heading_num_format = True
                if not _is_heading_num_format:
                    return None
        return level

    for idx, item in enumerate(paragraphs_text):
        level = _precompute_heading(idx, item)
        if level and level in ('title', 'h1', 'h2', 'h3', 'h4', 'h5'):
            is_heading_index.add(idx)

    # ──── 计算正文编号列表的序号（用于保留Word原生十进制编号） ────
    # 注意：num_seq记录每个段落在其numId序列中的位置，而不是最终累计值
    num_seq = {}  # idx -> 序号
    num_seq_count = {}  # (numId, ilvl) -> 当前计数
    for pi, pitem in enumerate(paragraphs_text):
        if pitem[0] != 'p':
            continue
        orig_num_id = pitem[4] if len(pitem) > 4 else None
        orig_num_ilvl = pitem[5] if len(pitem) > 5 else None
        if orig_num_id and orig_num_id != '0':
            key = (orig_num_id, orig_num_ilvl)
            if key not in num_seq_count:
                num_seq_count[key] = 0
            num_seq_count[key] += 1
            # 记录当前段落在序列中的位置
            num_seq[pi] = num_seq_count[key]

    # ─── 检查正文编号序号是否合理 ───
    # 规则：同一(numId, ilvl)组内，相邻两段之间如果隔了一个 h1 标题，
    # 说明序号跨大节连续，可能不正确，加入批注警告
    # 注意：h1 标题可能是：
    #  1) 文本前缀"一、二、三、"（如手工录入的文档）
    #  2) Word原生编号(numId=1)的段落（如本模板文档）
    # 两种情况都要检测
    discontinuous_seq_warnings = {}
    # 建立每个(numId, ilvl)组的段落索引列表
    num_group_indices = {}
    for pi, pitem in enumerate(paragraphs_text):
        if pitem[0] != 'p':
            continue
        orig_num_id = pitem[4] if len(pitem) > 4 else None
        orig_num_ilvl = pitem[5] if len(pitem) > 5 else None
        if not orig_num_id or orig_num_id == '0':
            continue
        key = (orig_num_id, orig_num_ilvl)
        if key not in num_group_indices:
            num_group_indices[key] = []
        num_group_indices[key].append(pi)
    # 检查：同一组内，相邻两段之间是否有 h1 标题
    for key, indices in num_group_indices.items():
        for j in range(1, len(indices)):
            prev_idx = indices[j - 1]
            curr_idx = indices[j]
            # 检查 prev_idx 和 curr_idx 之间是否有 h1 标题
            has_h1_between = False
            for k in range(prev_idx + 1, curr_idx):
                if k < len(paragraphs_text) and paragraphs_text[k][0] == 'p':
                    t = paragraphs_text[k][1].strip() if len(paragraphs_text[k]) > 1 else ''
                    # 情况1：文本前缀中文编号（一、二、三、）
                    if re.match(r'^[一二三四五六七八九十]+、', t):
                        has_h1_between = True
                        break
                    # 情况2：Word原生编号 numId=1 的段落（h1标题段落）
                    h1_num_id = paragraphs_text[k][4] if len(paragraphs_text[k]) > 4 else None
                    if h1_num_id == '1':
                        has_h1_between = True
                        break
            if has_h1_between:
                prev_seq = num_seq.get(prev_idx)
                curr_seq = num_seq.get(curr_idx)
                if prev_seq is not None and curr_seq is not None and curr_idx in is_heading_index:
                    # 只有当前段落是标题时才触发序号不连续警告
                    # 正文列表（如"需协同解决的事项"下的1. 2.）不触发此警告
                    discontinuous_seq_warnings[curr_idx] = (
                        f'序号 {curr_seq}. 与前一项（序号 {prev_seq}.）'
                        f'之间隔有大节标题，序号可能不连续，建议确认原文'
                    )

    for i, item in enumerate(paragraphs_text):
        # 跳过已合并到主标题的段落
        if i in merged_titles:
            continue
        
        etype = item[0]
        if etype == 'tbl':
            title_mode = False
            rows_data = item[1]
            rows_shading = item[2] if len(item) > 2 else None
            rows_font_color = item[3] if len(item) > 3 else None
            num_cols = max(len(row) for row in rows_data)
            table = doc.add_table(rows=len(rows_data), cols=num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = 'Table Grid'

            tbl_elem = table._tbl
            tblPr = tbl_elem.find(qn('w:tblPr'))
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl_elem.insert(0, tblPr)
            tblW = tblPr.find(qn('w:tblW'))
            if tblW is None:
                tblW = OxmlElement('w:tblW')
                tblPr.insert(0, tblW)
            tblW.set(qn('w:w'), '5000')
            tblW.set(qn('w:type'), 'pct')

            col_widths = _calc_smart_col_widths(rows_data, num_cols)
            total_page_width_dxa = 12818
            total_weight = sum(col_widths)
            for j, weight in enumerate(col_widths):
                width_dxa = int(total_page_width_dxa * weight / total_weight)
                for row in table.rows:
                    cell = row.cells[j]
                    tc_elem = cell._tc
                    tcPr = tc_elem.find(qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = OxmlElement('w:tcPr')
                        tc_elem.insert(0, tcPr)
                    tcW = tcPr.find(qn('w:tcW'))
                    if tcW is None:
                        tcW = OxmlElement('w:tcW')
                        tcPr.append(tcW)
                    tcW.set(qn('w:w'), str(width_dxa))
                    tcW.set(qn('w:type'), 'dxa')

            for ri, row_data in enumerate(rows_data):
                is_header = (ri == 0)
                tr = table.rows[ri]._tr
                trPr = tr.find(qn('w:trPr'))
                if trPr is None:
                    trPr = OxmlElement('w:trPr')
                    tr.insert(0, trPr)
                trHeight = trPr.find(qn('w:trHeight'))
                if trHeight is None:
                    trHeight = OxmlElement('w:trHeight')
                    trPr.append(trHeight)
                trHeight.set(qn('w:val'), '397')
                trHeight.set(qn('w:hRule'), 'atLeast')

                for j, cell_text in enumerate(row_data):
                    cell = table.cell(ri, j)
                    cell.text = ''
                    p = cell.paragraphs[0]
                    set_para_spacing(p, twips=312)
                    cleaned = clean_text(cell_text)
                    run = p.add_run(cleaned)
                    # 全局规则：保持原文表格字体颜色不变
                    orig_font_color = rows_font_color[ri][j] if rows_font_color and ri < len(rows_font_color) and j < len(rows_font_color[ri]) else None
                    if orig_font_color:
                        set_run_font(run, FONT_FANGSONG, SIZE_XIAOSI, bold=is_header, color=RGBColor.from_string(orig_font_color))
                    else:
                        set_run_font(run, FONT_FANGSONG, SIZE_XIAOSI, bold=is_header)
                    if is_header:
                        tc_elem = cell._tc
                        tcPr2 = tc_elem.find(qn('w:tcPr'))
                        if tcPr2 is None:
                            tcPr2 = OxmlElement('w:tcPr')
                            tc_elem.insert(0, tcPr2)
                        # 全局规则：保持原文表格底稿颜色不变
                        orig_shading = rows_shading[ri][j] if rows_shading and ri < len(rows_shading) and j < len(rows_shading[ri]) else None
                        if orig_shading:
                            # 使用原始底色
                            shading = tcPr2.find(qn('w:shd'))
                            if shading is None:
                                shading = OxmlElement('w:shd')
                                tcPr2.append(shading)
                            shading.set(qn('w:fill'), orig_shading)
                            shading.set(qn('w:val'), 'clear')
                        # 如果原文没有底色，也不强制添加
                    else:
                        # 非表头行也保留原始底色
                        orig_shading = rows_shading[ri][j] if rows_shading and ri < len(rows_shading) and j < len(rows_shading[ri]) else None
                        if orig_shading:
                            tc_elem = cell._tc
                            tcPr2 = tc_elem.find(qn('w:tcPr'))
                            if tcPr2 is None:
                                tcPr2 = OxmlElement('w:tcPr')
                                tc_elem.insert(0, tcPr2)
                            shading = tcPr2.find(qn('w:shd'))
                            if shading is None:
                                shading = OxmlElement('w:shd')
                                tcPr2.append(shading)
                            shading.set(qn('w:fill'), orig_shading)
                            shading.set(qn('w:val'), 'clear')
                    header_text = rows_data[0][j].strip() if rows_data and len(rows_data[0]) > j else ''
                    is_seq_col = header_text in ('序号', '编号')
                    stripped = cleaned.strip()
                    if is_header or is_seq_col or len(stripped) <= 20:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif re.match(r'^[\d,.\-+%]+$', stripped) and len(stripped) >= 3:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            continue

        raw = item[1]
        is_bold = item[2] if len(item) > 2 else False
        word_num_level = item[3] if len(item) > 3 else None

        text = clean_text(raw)
        if not text:
            # 空行：如果下一个非空段落是标题，跳过此空行
            skip_empty = False
            for j in range(i + 1, len(paragraphs_text)):
                if paragraphs_text[j][0] == 'p' and paragraphs_text[j][1].strip():
                    if j in is_heading_index:
                        skip_empty = True
                    break
                if paragraphs_text[j][0] == 'tbl':
                    break
            if skip_empty:
                continue
            p = doc.add_paragraph()
            set_para_spacing(p)
            continue

        if title_mode and is_main_title(text):
            title_count += 1
            
            # 收集连续的主标题段落，合并为完整标题
            main_title_parts = [text]
            j = i + 1
            while j < len(paragraphs_text) and paragraphs_text[j][0] == 'p':
                next_text = clean_text(paragraphs_text[j][1])
                if next_text and is_main_title(next_text):
                    # 检查是否紧跟其后（中间无空行或其他内容）
                    if not any(clean_text(paragraphs_text[k][1]) for k in range(i+1, j)):
                        main_title_parts.append(next_text)
                        j += 1
                    else:
                        break
                else:
                    break
            
            # 合并所有主标题部分（用换行符分隔）
            combined_title = '\n'.join(main_title_parts)
            
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_para_spacing(p)
            run = p.add_run(combined_title)
            set_run_font(run, FONT_XIAOBIAOSONG, SIZE_ERHAO, bold=False)
            
            # 跳过已合并的主标题段落（通过continue外部for循环的逻辑）
            # 记录需要跳过的起始索引，外部循环会处理
            for skip_i in range(i + 1, j):
                merged_titles.add(skip_i)
            
            # 主标题后始终插入空行（与正文区隔）
            title_mode = False
            blank = doc.add_paragraph()
            set_para_spacing(blank)
            continue

        title_mode = False
        level = detect_level(text)
        # 先应用 Word 编号层级（PENDING → 实际层级）
        # 但需谨慎：word_num_level 可能不准确（如"协调解决事项"被误判为h1）
        # 仅当 text 看起来像对应层级的标题时才应用
        if word_num_level is not None and level == 'body':
            # Bug2修复：只有detect_level返回body时才用word_num_level回退
            # 避免Word编号映射错误覆盖正确的文本编号识别
            # 保守策略：仅当文本无编号前缀且较短（≤30字）时才应用word_num_level
            # 如果文本看起来像正文（有句号、逗号、长度>30），保持原level判断
            looks_like_title = (
                len(text) <= 30
                and not re.search(r'[。；]', text)
            )
            if looks_like_title:
                level = word_num_level
            # 否则保持原level（可能是body）
        # 调试：查看应用word_num_level后的最终level
        if 18 <= i <= 22:
            print(f'[DEBUG-FINAL-LEVEL] i={i}, text="{text[:30]}", detect_level={detect_level(text)}, wnl={word_num_level}, final_level={level}, promote={promote_x_to_h1}', flush=True)
            print(f'                   promote_body_indices包含i? {i in promote_body_indices}, word_num_indices包含i? {i in word_num_indices}', flush=True)
        # Word 编号段落中，文本无编号前缀且长度 >25 字的，降级为 body（是正文不是标题）
        # Bug2修复：但如果Word编号格式是chineseCounting（一、二、三）等标题型格式，
        # 说明该段落是Word编号隐藏了标题编号，不应仅因文本长度降级
        if word_num_level is not None and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            has_text_prefix = bool(
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
            )
            if not has_text_prefix and len(text) > 25:
                # 检查Word编号格式：如果是标题型编号格式，不降级
                _orig_nid = item[4] if len(item) > 4 else None
                _orig_nilvl = item[5] if len(item) > 5 else '0'
                _is_heading_num_format = False
                if _orig_nid and _orig_nid != '0':
                    _an_id = num_to_abstract.get(_orig_nid)
                    if _an_id:
                        _levels = abstract_num_defs.get(_an_id, {})
                        _fmt, _ = _levels.get(_orig_nilvl or '0', (None, None))
                        if _fmt in ('chineseCounting', 'chineseCountingThousand',
                                    'ideographDigital', 'ideographEnclosedCircle'):
                            _is_heading_num_format = True
                if not _is_heading_num_format:
                    level = 'body'
                    word_num_level = None  # 重置，后续不按标题处理
        # 提升 X. → h1（当文档没有中文一级编号时）
        # 排除：X.Y 多级编号、X.是/要 等动词前缀、正文长句
        if promote_x_to_h1 and level == 'h3':
            is_long_sentence = '。' in text and len(text) > 30
            if (re.match(r'^\d+[.、．]\s*\S', text)
                and not re.match(r'^\d+\.\d+', text)
                and not re.match(r'^\d+[.、．]\s*[是是以要为将把让使被]', text)
                and not is_long_sentence):
                level = 'h1'
        # 提升 promote 模式下夹在编号组之间的无编号短段落为 h1
        if promote_x_to_h1 and level == 'body' and word_num_level is None and i in promote_body_indices:
            level = 'h1'
        # promote 模式下：Word 编号隐藏了编号的短段落，可能需要提升为标题
        # 但仅当 word_num_level 是 h2/h3/h4 时才提升（h1已是最高级，不应再提升）
        # 且需要更严格的判断：文本应像标题（不含句号，长度适中）
        if promote_x_to_h1 and word_num_level is not None and word_num_level in ('h2', 'h3', 'h4'):
            has_text_prefix = bool(
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
            )
            # 更严格的标题判断：无句末标点、长度适中、不包含逗号顿号等正文特征
            looks_like_heading = (
                not has_text_prefix
                and len(text) <= 15
                and not re.search(r'[。；！？]', text)  # 无句末标点
                and not re.match(r'^\d+[.、．]', text)  # 无数字编号前缀
            )
            if looks_like_heading:
                level = 'h1'

        if level == 'body' and word_num_level is None:
            prev_etype = paragraphs_text[i - 1][0] if i > 0 else None
            prev_info = paragraphs_text[i - 1] if i > 0 else None
            is_after_table = (prev_etype == 'tbl')
            # 检查前一段是否是加粗（主标题）
            prev_is_bold = (prev_info[0] == 'p' and len(prev_info) > 2 and prev_info[2] is True) if prev_etype == 'p' and prev_info else False
            # 表格标题（表1、表2…、表3-1…）不算标题
            is_table_title = bool(re.match(r'^表\s*\d+', text))
            # 日期行（如"（2026年4月28日）"）不算标题
            is_date_line = bool(re.match(r'^[（(]\d{4}年\d{1,2}月\d{1,2}日[）)]$', text.strip()))
            is_short_title = (
                len(text) <= 25
                and not re.search(r'[。；]', text)
                and not text.startswith('附件')
                and not re.match(r'^[\d,.\-+%：:（）()]+$', text)
                and not is_table_title
                and not is_date_line
            )
            # 放宽短标题检测：主标题后面的短文本也应该识别为h1
            # 例如："新开项目上线及业财集成情况" 紧跟主标题，应为"一、新开项目..."
            is_after_main_title = (prev_is_bold and len(text) <= 20)
            if is_bold and is_short_title:
                level = 'h1'
            elif is_after_table and is_short_title and len(text) <= 20 and not is_table_title:
                level = 'h1'
            elif is_after_main_title and is_short_title and not is_date_line:
                # 主标题后面的短文本提升为一级标题（日期行除外）
                level = 'h1'

        clean_heading = text
        std_prefix_match = None
        is_multilevel = bool(re.match(r'^\d+\.\d+', text))

        # 检测 X.是 / X.要 / X.以 等编号+动词的不规范格式
        is_verb_prefix = bool(re.match(r'^\d+[.、．]\s*[是是以要为将把让使被]', text))

        if not is_multilevel:
            std_prefix_match = (
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
            )
        if std_prefix_match and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            clean_heading = text[std_prefix_match.end():].lstrip()

        # 全局规则：标题原文加粗的，格式化后也保持加粗
        # 正文段落的加粗保留逻辑不变（需要编号前缀+短文本/结构词）
        has_num_prefix = bool(
            re.match(r'^[一二三四五六七八九十]+、', text)
            or re.match(r'^\d+[.、．]\s*', text)
            or re.match(r'^（\d+）', text)
        )
        has_struct_marker = bool(re.search(r'[：:]', text)) or bool(re.search(r'层面|板块|线条', text))
        if level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            preserve_bold = is_bold
        else:
            preserve_bold = is_bold and has_num_prefix and (len(text) <= 25 or has_struct_marker)

        if is_multilevel and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            # X.Y 多级编号：保留原文编号，不自动重编，添加批注提醒
            p = doc.add_paragraph()
            apply_heading_format(p, level, text, preserve_bold=preserve_bold)
            if i in subhead_para_indices:
                num_m = re.match(r'^(\d+\.\d+[.、．]?)', text)
                num_str = num_m.group(1) if num_m else ''
                comment_list.append((text[:20],
                    f'此处使用"{num_str}"多级编号，建议改为四级标题①②③格式',
                    'multi_level_num'))
        elif is_verb_prefix and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            # X.是/要 等不规范编号：保留原文编号不重编，添加批注提醒
            verb_m = re.match(r'^(\d+[.、．])', text)
            prefix_text = verb_m.group(1) if verb_m else ''
            p = doc.add_paragraph()
            apply_heading_format(p, 'body', text, preserve_bold=preserve_bold)
            if i in h3_para_indices:
                comment_list.append((text[:20],
                    f'编号"{prefix_text}"后直接跟动词，建议改为"一是…""二是…"格式',
                    'verb_after_num'))
        elif level in ('title', 'h1', 'h2', 'h3', 'h4', 'h5'):
            # 调试：如果是编号段落，打印详细信息
            if text and re.match(r'^\d+[.、．]', text):
                print(f'[DEBUG-NUM] i={i}, level={level}, counter状态=h1={counter.h1},h2={counter.h2},h3={counter.h3},h4={counter.h4}, text="{text[:40]}"')
            
            # 优先保留原文编号前缀，而不是强制重编
            # 扩展正则，匹配所有标题前缀格式：h1=一、 h2=（一） h3=1. h4=（1） h5=①
            orig_num_match = re.match(
                r'^([一二三四五六七八九十]+[、]|[0-9]+[.、．]|（[一二三四五六七八九十]+）|（[0-9]+）|[①②③④⑤⑥⑦⑧⑨⑩])\s*',
                text
            )
            if orig_num_match:
                # 原文有编号前缀，提取并保留
                orig_prefix = orig_num_match.group(0)
                # 提取编号后的正文内容
                remaining_text = text[orig_num_match.end():]
                # 对于h3级别，如果原文是"2."格式，保持原样不重编
                display = orig_prefix + remaining_text
                # 不调用 counter.next() 来递增，保持counter状态不变（用于其他未编号的段落）
                prefix_used = orig_prefix
                print(f'[DEBUG]   -> 保留原文编号="{prefix_used.strip()}", counter保持不变')
                # FIX: 更新counter状态，使其与原文编号一致
                # 解析orig_prefix中的编号值
                num_val = None
                prefix_level = level  # 使用当前detect到的level
                # h1格式："一、" "二、"
                m_h1 = re.match(r'[一二三四五六七八九十]+', orig_prefix)
                if m_h1:
                    num_val = CNUM_TO_INT.get(m_h1.group())
                # h3格式："1." "2."
                if num_val is None:
                    m_h3 = re.match(r'[0-9]+', orig_prefix)
                    if m_h3 and re.match(r'[.、．]', orig_prefix[len(m_h3.group()):]):
                        num_val = int(m_h3.group())
                # h2格式："（一）"
                if num_val is None:
                    m_h2 = re.match(r'（([一二三四五六七八九十]+)）', orig_prefix)
                    if m_h2:
                        num_val = CNUM_TO_INT.get(m_h2.group(1))
                # h4格式："（1）"
                if num_val is None:
                    m_h4 = re.match(r'（([0-9]+)）', orig_prefix)
                    if m_h4:
                        num_val = int(m_h4.group(1))
                # h5格式："①"
                if num_val is None:
                    m_h5 = re.match(r'[①②③④⑤⑥⑦⑧⑨⑩]', orig_prefix)
                    if m_h5:
                        try:
                            num_val = CIRCLE_NUMBERS.index(m_h5.group()) + 1
                        except ValueError:
                            pass
                # 更新counter
                if num_val is not None and prefix_level is not None:
                    if prefix_level == 'h1':
                        counter.h1 = num_val
                    elif prefix_level == 'h2':
                        counter.h2 = num_val
                    elif prefix_level == 'h3':
                        counter.h3 = num_val
                    elif prefix_level == 'h4':
                        counter.h4 = num_val
                    elif prefix_level == 'h5':
                        counter.h5 = num_val
                    print(f'[DEBUG-COUNTER] 更新counter: {prefix_level}={num_val}, 新状态=h1={counter.h1},h2={counter.h2},h3={counter.h3},h4={counter.h4},h5={counter.h5}')
            else:
                # 原文无编号前缀，使用counter生成
                prefix = counter.next(level)
                display = prefix + clean_heading
                prefix_used = prefix
                print(f'[DEBUG] i={i} -> prefix="{prefix}", 新counter状态=h1={counter.h1},h2={counter.h2},h3={counter.h3},h4={counter.h4}')
            
            p = doc.add_paragraph()
            apply_heading_format(p, level, display, preserve_bold=preserve_bold)
            if i in punct_para_indices:
                comment_list.append((display[:20],
                    '此标题/段落可能缺少句末标点，请人工确认',
                    'missing_end_punct'))
            elif i in title_punct_para_indices:
                comment_list.append((display[:20],
                    '标题末尾不应有标点符号',
                    'trailing_punct'))
            # 检测阿拉伯数字编号格式（如"1. xxx"），判断是否需要提示改中文格式
            # 公文规范：h1=一、  h2=（一）  h3=1.  h4=（1）
            # 如果文档已按规范使用（一）作为h2，则h3使用"1."是正确格式，不需提示
            # 仅当层级编号格式不匹配规范时才提示
            arabic_num_match = re.match(r'^(\d+)([.、．])', text)
            if arabic_num_match and level in ('h1', 'h2', 'h3'):
                # 判断文档的二级标题格式是否已经是中文（（一）格式）
                # 如果是，则三级标题用"1."是规范写法，不提示
                # 只有在一级标题直接用"1."（无中文一、二、三）时才需要提示
                if level == 'h1':
                    # 一级标题用了"1、"格式，建议改为"一、"
                    num_str = arabic_num_match.group(1)
                    sep = arabic_num_match.group(2)
                    orig_num_prefix = num_str + sep
                    chinese_num = CNUM.get(num_str, num_str)
                    comment_list.append((orig_num_prefix,
                        f'建议将"{orig_num_prefix}"改为"{chinese_num}、"',
                        'number_prefix'))
                elif level == 'h2' and not has_cn_h1:
                    # 二级标题用了"1、"格式，且文档无中文一级标题，建议改为"（一）"
                    num_str = arabic_num_match.group(1)
                    sep = arabic_num_match.group(2)
                    orig_num_prefix = num_str + sep
                    chinese_num = CNUM.get(num_str, num_str)
                    comment_list.append((orig_num_prefix,
                        f'建议将"{orig_num_prefix}"改为"（{chinese_num}）"',
                        'number_prefix'))
            # 检查是否缺少二级标题
            if i in missing_h2_indices:
                comment_list.append((text[:30],
                    '一级标题下直接使用三级标题，建议补充二级标题（一）（二）',
                    'heading_skip'))
            # 检查标题编号标点是否符合规范
            if i in title_punct_indices:
                # 找到对应的问题，获取建议文本
                        for idx, txt, issue_type, suggestion in title_punct_check_results:
                            if idx == i:
                                comment_list.append((text[:30], suggestion, 'number_prefix'))
                                break
        else:
            # 正文段落
            p = doc.add_paragraph()
            # 问候语（含称呼关键词+以：结尾）不缩进
            greeting_kw = '领导|同事|各位|尊敬|您好|下午好|上午好|上午好'
            is_greeting = bool(re.match(r'^.{2,30}[：:]$', text.strip()) and re.search(greeting_kw, text.strip()))
            # 全局规则：表格标题（表1、表2、表3等）居中
            is_table_title = bool(re.match(r'^表\s*\d+', text.strip()))
            
            # 对于有numId的正文段落（Word原生编号列表），保留原始编号格式
            orig_num_id = item[4] if len(item) > 4 else None
            orig_num_ilvl = item[5] if len(item) > 5 else None
            display_text = text
            if orig_num_id and orig_num_id != '0':
                an_id = num_to_abstract.get(orig_num_id)
                if an_id:
                    levels = abstract_num_defs.get(an_id, {})
                    fmt, lvl_txt = levels.get(orig_num_ilvl or '0', (None, None))
                    # lvl_txt 是 Word 的编号模板，如 '（%1）'、'%1.'、'①' 等
                    is_non_decimal = fmt and fmt not in ('decimal', 'none', None, '')
                    is_decimal_no_prefix = (
                        fmt == 'decimal'
                        and not bool(re.match(r'^\d+[.、．]\s*\S', text))
                    )
                    if is_non_decimal:
                        # 非十进制格式（如①②③④），添加编号前缀到文本
                        seq_val = num_seq.get(i, 1)
                        circle_num = CIRCLE_NUMBERS[seq_val - 1] if 1 <= seq_val <= len(CIRCLE_NUMBERS) else str(seq_val)
                        display_text = circle_num + text
                        fmt_name = {
                            'ideographEnclosedCircle': '①②③④',
                            'decimalEnclosedCircleChinese': '①②③④',
                            'chineseCountingThousand': '中文千位',
                            'chineseCounting': '中文计数',
                            'lowerLetter': 'a.b.c.',
                            'upperLetter': 'A.B.C.',
                        }.get(fmt, fmt)
                        if i not in discontinuous_seq_warnings:
                            comment_list.append((circle_num,
                                f'原文使用编号{fmt_name}，是Word自带编号。建议改为（1）（2）或保留原格式',
                                'number_prefix'))
                    elif is_decimal_no_prefix:
                        # 文本无编号前缀但有Word原生编号（如 numId=1 -> abstractNumId=0 -> lvlText='（%1）'）
                        # 根据 lvl_txt 模板生成对应的编号前缀
                        seq_val = num_seq.get(i, 1)
                        # 模板替换：'（%1）' -> '（1）', '%1.' -> '1.', '①' -> '①'
                        if lvl_txt and lvl_txt != '':
                            prefix = lvl_txt.replace('%1', str(seq_val))
                        else:
                            prefix = f'{seq_val}.'
                        display_text = prefix + text
            
            apply_heading_format(p, level, display_text, no_indent=is_greeting or is_table_title)
            # 表格标题居中且不加粗（保持原文加粗状态）
            if is_table_title:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i in punct_para_indices:
                comment_list.append((text[:20],
                    '此段落可能缺少句末标点，请人工确认',
                    'missing_end_punct'))
            if i in discontinuous_seq_warnings:
                # 用 text[:30] 而不是 display_text[:30]，因为输出文档的纯文本不含编号前缀
                comment_list.append((text[:30],
                    discontinuous_seq_warnings[i],
                    'number_prefix'))
            # Bug 1修复：Word自动编号格式批注
            if i in word_num_indices:
                comment_list.append((text[:30],
                    '原文使用Word自动编号（阿拉伯数字1.2.3.），建议改为二级标题（（一）（二））格式',
                    'number_prefix'))

    # 统一添加批注
    _apply_comments_to_doc(doc, comment_list)

    _add_page_number(doc)
    doc.save(dst_path)
    return dst_path, warnings
