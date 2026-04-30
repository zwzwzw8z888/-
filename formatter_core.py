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

# ─────────────────────────── 格式常量 ───────────────────────────
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
    if re.match(r'^\d+[.、．]\s*', t):
        return 'h3'
    if re.match(r'^（\d+）', t):
        return 'h4'
    if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', t):
        return 'h5'
    return 'body'


def is_main_title(text):
    t = text.strip()
    if not t or len(t) > 25:
        return False
    # 排除时间行（如 "2026年4月"、"2026年4月28日"）
    if re.match(r'^\d{4}年\d{1,2}月\d{0,2}日?\s*$', t):
        return False
    for pat in [
        r'^[一二三四五六七八九十]+、',
        r'^（[一二三四五六七八九十]+）',
        r'^\d+[.、．]\s*',
        r'^（\d+）',
        r'^[①②③④⑤⑥⑦⑧⑨⑩]',
        r'^[\d,.\-+%/：:；;，。、]+$',
    ]:
        if re.match(pat, t):
            return False
    if '。' in t or '；' in t:
        return False
    # 排除问候语（含冒号且含称呼关键词）
    if re.search(r'[：:]$', t) and re.search(r'领导|同事|各位|尊敬|您好|下午好|上午好|你好', t):
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


def apply_heading_format(para, level, text, prefix='', no_indent=False):
    para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(para)
    # 全局规则：所有段落（含标题）默认首行缩进2字符，除非明确 no_indent（如问候语）
    if not no_indent:
        set_para_indent(para, 2)
    display = prefix + text
    font_map = {
        'h1': FONT_HEITI,
        'h2': FONT_KAITI,
    }
    cn_font = font_map.get(level, FONT_FANGSONG)
    run = para.add_run(display)
    set_run_font(run, cn_font, SIZE_SANHAO, bold=False)


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


def _apply_comments_to_doc(doc, comment_list):
    """在文档中统一添加批注。
    comment_list: [(text_prefix, comment_text), ...]
    通过文本前缀匹配输出文档中的段落来定位批注位置。
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

    # 已匹配的段落集合（避免重复匹配）
    matched_indices = set()

    for text_prefix, comment_text in comment_list:
        # 在输出段落中找到匹配的段落
        target_elem = None
        for ci, pelem, ptext in para_map:
            if ci in matched_indices:
                continue
            if ptext.startswith(text_prefix) or text_prefix in ptext:
                target_elem = pelem
                matched_indices.add(ci)
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

        # 2. 在段落开头插入 commentRangeStart
        cs = OxmlElement('w:commentRangeStart')
        cs.set(qn('w:id'), comment_id)
        target_elem.insert(0, cs)

        # 3. 在段落末尾插入 commentRangeEnd + commentReference
        ce = OxmlElement('w:commentRangeEnd')
        ce.set(qn('w:id'), comment_id)
        target_elem.append(ce)

        ref_run = OxmlElement('w:r')
        ref_rPr = OxmlElement('w:rPr')
        ref_rStyle = OxmlElement('w:rStyle')
        ref_rStyle.set(qn('w:val'), 'CommentReference')
        ref_rPr.append(ref_rStyle)
        ref_run.append(ref_rPr)
        ref_cr = OxmlElement('w:commentReference')
        ref_cr.set(qn('w:id'), comment_id)
        ref_run.append(ref_cr)
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
        # 纯编号+标题行跳过标点检测（如 "1.科技部："）
        if is_numbered and len(text) <= 12:
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
                is_bold = any(
                    rPr.find(qn('w:b')) is not None
                    and rPr.find(qn('w:b')).get(qn('w:val'), 'true') not in ('false', '0')
                    for rPr in elem.iter(qn('w:rPr'))
                )
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
                for tr in elem.iter(qn('w:tr')):
                    cells = []
                    for tc in tr.iter(qn('w:tc')):
                        cell_text = ''.join(t.text for t in tc.iter(qn('w:t')) if t.text)
                        cells.append(cell_text.strip())
                    if cells:
                        rows_data.append(cells)
                if rows_data:
                    paragraphs_text.append(('tbl', rows_data))

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
        for order, nid in enumerate(sorted_num_ids):
            level_map = {0: 'h1', 1: 'h2', 2: 'h3'}
            if order in level_map:
                numid_ilvl_level_map[(nid, '0')] = level_map[order]

    # 回填 PENDING
    for idx, item in enumerate(paragraphs_text):
        if item[0] == 'p' and item[3] == 'PENDING':
            nid, nilvl = item[4], item[5] or '0'
            paragraphs_text[idx] = (
                item[0], item[1], item[2],
                numid_ilvl_level_map.get((nid, nilvl)),
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

    # ──── 标题句末标点检测 ────
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
        # 有文本编号前缀的才算标题候选
        has_text_prefix = bool(
            re.match(r'^[一二三四五六七八九十]+、', text)
            or re.match(r'^（[一二三四五六七八九十]+）', text)
            or re.match(r'^\d+[.、．]\s*\S', text)
            or re.match(r'^（\d+）', text)
            or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
        )
        if not has_text_prefix:
            continue
        # Word编号但无文本前缀的长段落不算标题
        is_word_num_body = (wnl is not None and not has_text_prefix and len(text) > 25)
        if is_word_num_body:
            continue
        # 有编号前缀但内容超长（>30字且有句号）的是正文不是标题
        if len(text) > 30 and '。' in text:
            continue
        # 标题应以非句号结尾，如果以句号结尾则标记
        if text.rstrip()[-1] in ('。', '；', '，'):
            title_punct_issues.append((idx, text, text.rstrip()[-1]))

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
    # 且 X 后面的子标题是 （X）或 (X)，则提升所有 X. 为 h1
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
        if wnl is not None:
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
        if wnl is not None and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            has_text_prefix = bool(
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
            )
            if not has_text_prefix and len(text) > 25:
                return None
        return level

    for idx, item in enumerate(paragraphs_text):
        level = _precompute_heading(idx, item)
        if level and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            is_heading_index.add(idx)

    for i, item in enumerate(paragraphs_text):
        etype = item[0]
        if etype == 'tbl':
            title_mode = False
            rows_data = item[1]
            num_cols = max(len(row) for row in rows_data)
            table = doc.add_table(rows=len(rows_data), cols=num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = 'Table Grid'

            tbl_elem = table._tbl
            tblPr = tbl_elem.find(qn('w:tblPr')) or OxmlElement('w:tblPr')
            if not tbl_elem.find(qn('w:tblPr')):
                tbl_elem.insert(0, tblPr)
            tblW = tblPr.find(qn('w:tblW')) or OxmlElement('w:tblW')
            if not tblPr.find(qn('w:tblW')):
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
                    set_run_font(run, FONT_FANGSONG, SIZE_XIAOSI, bold=is_header)
                    if is_header:
                        tc_elem = cell._tc
                        tcPr2 = tc_elem.find(qn('w:tcPr'))
                        if tcPr2 is None:
                            tcPr2 = OxmlElement('w:tcPr')
                            tc_elem.insert(0, tcPr2)
                        shading = tcPr2.find(qn('w:shd'))
                        if shading is None:
                            shading = OxmlElement('w:shd')
                            tcPr2.append(shading)
                        shading.set(qn('w:fill'), 'F0F0F0')
                        shading.set(qn('w:val'), 'clear')
                    header_text = rows_data[0][j].strip() if rows_data[0] else ''
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
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_para_spacing(p)
            run = p.add_run(text)
            set_run_font(run, FONT_XIAOBIAOSONG, SIZE_ERHAO, bold=False)
            # 超过2段连续标题则退出标题模式并插入空行
            if title_count >= 2:
                title_mode = False
                title_ended = True
                blank = doc.add_paragraph()
                set_para_spacing(blank)
            continue

        title_mode = False
        level = detect_level(text)
        # 先应用 Word 编号层级（PENDING → 实际层级）
        if word_num_level is not None:
            level = word_num_level
        # Word 编号段落中，文本无编号前缀且长度 >25 字的，降级为 body（是正文不是标题）
        if word_num_level is not None and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            has_text_prefix = bool(
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
            )
            if not has_text_prefix and len(text) > 25:
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
        # promote 模式下：Word 编号隐藏了编号的短段落（无文本前缀 ≤15字），提升为 h1
        if promote_x_to_h1 and word_num_level is not None and word_num_level in ('h1', 'h2', 'h3', 'h4'):
            has_text_prefix = bool(
                re.match(r'^[一二三四五六七八九十]+、', text)
                or re.match(r'^（[一二三四五六七八九十]+）', text)
                or re.match(r'^\d+[.、．]\s*', text)
                or re.match(r'^（\d+）', text)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
            )
            if not has_text_prefix and len(text) <= 15 and not re.search(r'[。；]', text):
                level = 'h1'

        if level == 'body' and word_num_level is None:
            prev_etype = paragraphs_text[i - 1][0] if i > 0 else None
            is_after_table = (prev_etype == 'tbl')
            # 表格标题（表1、表2…、表3-1…）不算标题
            is_table_title = bool(re.match(r'^表\s*\d+', text))
            is_short_title = (
                len(text) <= 25
                and not re.search(r'[。；]', text)
                and not text.startswith('附件')
                and not re.match(r'^[\d,.\-+%：:（]+', text)
                and not is_table_title
            )
            if is_bold and is_short_title:
                level = 'h1'
            elif is_after_table and is_short_title and len(text) <= 20 and not is_table_title:
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

        if is_multilevel and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            # X.Y 多级编号：保留原文编号，不自动重编，添加批注提醒
            p = doc.add_paragraph()
            apply_heading_format(p, level, text)
            if i in subhead_para_indices:
                num_m = re.match(r'^(\d+\.\d+[.、．]?)', text)
                num_str = num_m.group(1) if num_m else ''
                comment_list.append((text[:20],
                    f'此处使用"{num_str}"多级编号，建议改为四级标题①②③格式'))
        elif is_verb_prefix and level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            # X.是/要 等不规范编号：保留原文编号不重编，添加批注提醒
            verb_m = re.match(r'^(\d+[.、．])', text)
            prefix_text = verb_m.group(1) if verb_m else ''
            p = doc.add_paragraph()
            apply_heading_format(p, 'body', text)
            if i in h3_para_indices:
                comment_list.append((text[:20],
                    f'编号"{prefix_text}"后直接跟动词，建议改为"一是…""二是…"格式'))
        elif level in ('h1', 'h2', 'h3', 'h4', 'h5'):
            prefix = counter.next(level)
            display = prefix + clean_heading
            p = doc.add_paragraph()
            apply_heading_format(p, level, display)
            if i in punct_para_indices:
                comment_list.append((display[:20],
                    '此标题/段落可能缺少句末标点，请人工确认'))
            elif i in title_punct_para_indices:
                comment_list.append((display[:20],
                    '标题末尾不应有标点符号'))
        else:
            # 正文段落
            p = doc.add_paragraph()
            # 问候语（含称呼关键词+以：结尾）不缩进
            greeting_kw = '领导|同事|各位|尊敬|您好|下午好|上午好|你好'
            is_greeting = bool(re.match(r'^.{2,30}[：:]$', text.strip())
                              and re.search(greeting_kw, text.strip()))
            apply_heading_format(p, level, text, no_indent=is_greeting)
            if i in punct_para_indices:
                comment_list.append((text[:20],
                    '此段落可能缺少句末标点，请人工确认'))

    # 统一添加批注
    _apply_comments_to_doc(doc, comment_list)

    _add_page_number(doc)
    doc.save(dst_path)
    return dst_path, warnings
