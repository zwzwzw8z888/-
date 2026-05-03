#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中建四局公文格式化 — 审查逻辑

包含 6 个 _check_* 函数，用于检测公文格式问题并返回问题列表。
"""

import re
import os
import json
from detect import detect_level, is_main_title
from constants import has_text_number_prefix, CN_NUMBERS, CNUM_TO_INT


# ──── AI 语义判断（可选，配置 DMP_AI_KEY 环境变量启用）────
def ai_is_body(text):
    """调 DeepSeek API 判断编号段落是正文还是标题
    返回 True=正文，False=标题，None=未配置/出错（走原有规则）
    """
    api_key = os.environ.get('CSCEC_AI_KEY', '')
    if not api_key:
        return None
    try:
        import urllib.request
        req = urllib.request.Request(
            'https://api.deepseek.com/v1/chat/completions',
            data=json.dumps({
                'model': 'deepseek-chat',
                'messages': [{
                    'role': 'system',
                    'content': '判断编号段落是【正文】还是【标题】。正文：描述动作/目标/措施的完整句（如"提升…能力""加强…工作"）、具体说明。标题：名词性短语、纲要要点、时间地点项（如"培训时间""培训地点"）、联系信息。只答"正文"或"标题"。'
                }, {
                    'role': 'user',
                    'content': f'"{text}"'
                }],
                'max_tokens': 5,
                'temperature': 0
            }).encode(),
            headers={'Content-Type': 'application/json', 'Authorization': f'Bearer {api_key}'}
        )
        resp = json.loads(urllib.request.urlopen(req, timeout=5).read())
        return '正文' in resp['choices'][0]['message']['content']
    except Exception:
        return None


# ──── 审查函数 ────
def check_punctuation_issues(paragraphs_text):
    """句末标点检测：找出未以句号/问号/叹号结尾的正文段落"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text or len(text) <= 10:
            continue
        if is_main_title(text):
            continue
        if re.match(r'^[一二三四五六七八九十]+、', text):
            continue
        if re.match(r'^（[一二三四五六七八九十]+）', text):
            continue
        if re.match(r'^（\d+）', text):
            continue
        if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text):
            continue
        # 有编号前缀 + 短文本 + 无句末标点 → 是标题，跳过句末标点检查
        if has_text_number_prefix(text) and len(text) <= 15 and not re.search(r'[。；]', text):
            continue
        # 有编号前缀 + 含冒号引导 + 冒号后内容短 → 标题式引导，跳过
        if has_text_number_prefix(text) and not re.search(r'[。；]', text):
            colon_m = re.search(r'[：:]', text)
            if colon_m:
                remaining = len(text) - colon_m.end()
                if remaining <= 20:
                    continue
        if re.match(r'^[\d,.\-+%：:（）()]+$', text):
            continue
        if re.search(r'[：:]$', text):
            continue
        if len(text) <= 15 and not re.search(r'[\u4e00-\u9fff]', text):
            continue
        if len(text) <= 20 and not re.search(r'[，。；！？]', text):
            continue
        # 含冒号 + 以数字或右括号结尾 → 列表项/联系信息，不查句号
        if re.search(r'[：:]', text) and re.search(r'[\d）\)\]】]$', text):
            continue
        # AI 兜底：15-30字无句号的编号段落（>30字的标题少见，直接当正文查句号）
        if has_text_number_prefix(text) and 15 < len(text) <= 30 and not re.search(r'[。；]', text):
            if ai_is_body(text) is False:
                continue  # AI 判为标题，跳过句末标点检查
        last_char = text[-1]
        if last_char not in ('。', '？', '！', '…', '"', '"', ')', '）', '；'):
            issues.append((i, text[:60]))
    return issues


def check_subheading_issues(paragraphs_text):
    """子标题序号混乱检测：原文含 X.Y 格式但被当作普通段落处理"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text:
            continue
        m = re.match(r'^(\d+)\.(\d+)[.、．]?\s*(.*)', text)
        if m:
            major = int(m.group(1))
            minor = int(m.group(2))
            content = m.group(3)
            if 2000 <= minor <= 2030 and content.startswith('年'):
                continue
            if minor > 0:
                issues.append((i, text[:60], major, minor, content))
    return issues


def check_h3_numbering_issues(paragraphs_text):
    """三级标题编号不规范检测：如 '1.是' '2.是' 应为 '一是' '二是'"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text:
            continue
        m = re.match(r'^(\d+)[.、．]\s*(是|且|但|将|要|在|已|以|对|为|从|按|于)\s*(.*)', text)
        if m:
            num = int(m.group(1))
            word = m.group(2)
            issues.append((i, text[:70], num, word))
    return issues


def check_word_numbering_format(paragraphs_text, num_to_abstract, abstract_num_defs):
    """检测Word自动编号使用阿拉伯数字（1.2.3.）且文本像标题的情况"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text:
            continue
        level = item[3] if len(item) > 3 else None
        if level is not None:
            continue
        orig_num_id = item[4] if len(item) > 4 else None
        if not orig_num_id or orig_num_id == '0':
            continue
        an_id = num_to_abstract.get(orig_num_id)
        if not an_id:
            continue
        levels = abstract_num_defs.get(an_id, {})
        nilvl = item[5] if len(item) > 5 else '0'
        fmt, lvl_txt = levels.get(nilvl, (None, None))
        if fmt != 'decimal':
            continue
        is_like_heading = (
            len(text) <= 40
            and not re.search(r'[。；！？]', text)
        )
        if is_like_heading:
            issues.append((i, text[:50]))
    return issues


def check_numbering_separator(paragraphs_text, num_to_abstract, abstract_num_defs):
    """检测Word自动编号分隔符不规范（、或．应改为.）"""
    issues = []
    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text:
            continue
        level = item[3] if len(item) > 3 else None
        if level is not None:
            continue  # 已有标题层级，由 check_title_punctuation 处理
        orig_num_id = item[4] if len(item) > 4 else None
        if not orig_num_id or orig_num_id == '0':
            continue
        an_id = num_to_abstract.get(orig_num_id)
        if not an_id:
            continue
        levels = abstract_num_defs.get(an_id, {})
        nilvl = item[5] if len(item) > 5 else '0'
        fmt, lvl_txt = levels.get(nilvl, (None, None))
        if not lvl_txt:
            continue
        if '、' in lvl_txt or '．' in lvl_txt:
            correct = lvl_txt.replace('、', '.').replace('．', '.')
            issues.append((i, text[:40], lvl_txt.strip(), correct.strip()))
    return issues


def check_missing_h2(paragraphs_text):
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

        detected = detect_level(text)
        effective_level = level if level in ('h1', 'h2', 'h3', 'h4', 'h5') else detected

        is_digit_prefix = bool(re.match(r'^\d+[.、．]', text))
        # 是否为正文：有句号（真正文结束标记）或很长（>40字）才算正文
        # 仅有分号但结构为"标题：内容"的仍算标题
        is_likely_body = is_digit_prefix and (
            '。' in text or len(text) > 40
            or bool(re.search(r'\d{11}', text))
            or ('@' in text and '.' in text)
            or bool(re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text))
            or bool(re.search(r'\d{1,2}:\d{2}', text))
        )
        # AI 兜底：规则拿不准时（15-40字、无句号、无联系信息），调 AI 判断
        if is_digit_prefix and not is_likely_body:
            ai_result = ai_is_body(text)
            if ai_result is True:
                is_likely_body = True  # AI 判为正文

        if effective_level == 'h1':
            last_h1_index = i
        elif effective_level == 'h3' and last_h1_index is not None:
            if is_likely_body:
                continue
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
        elif last_h1_index is not None and num_id and num_id != '0':
            looks_like_title = (
                len(text) <= 30
                and not re.search(r'[。；！？]', text)
            )
            if looks_like_title:
                issues.append((i, text[:30]))
        elif effective_level == 'h2':
            last_h1_index = None

    return issues


def check_title_punctuation(paragraphs_text):
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

        if level is None and text:
            if re.match(r'^[一二三四五六七八九十]+[、】]', text):
                level = 'h1'
            elif re.match(r'^（[一二三四五六七八九十]+）', text):
                level = 'h2'
            elif re.match(r'^\d+[.、．](?!\d)\s*\S', text):
                level = 'h3'

        if level not in ('h1', 'h2', 'h3'):
            continue

        if not text:
            continue

        if level == 'h1':
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
            match = re.match(r'^（([一二三四五六七八九十]+)）([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                if punct:
                    issues.append((i, text[:40], 'h2_extra_punct',
                        f'二级标题"（{num_part}）"后不应有标点，检测到"{punct}"'))

        elif level == 'h3':
            match = re.match(r'^(\d+)([、．.：:；;]?)', text)
            if match:
                num_part = match.group(1)
                punct = match.group(2)
                if punct != '.':
                    if punct:
                        issues.append((i, text[:40], 'h3_wrong_punct',
                            f'三级编号"{num_part}"后应为点号"."，实为"{punct}"'))
                    else:
                        issues.append((i, text[:40], 'h3_missing_punct',
                            f'三级编号"{num_part}"后缺少点号"."'))

    return issues


def check_title_trailing_punct(paragraphs_text, num_to_abstract, abstract_num_defs):
    """标题句末标点检测（从 format_document 中提取）

    规则：只有真正的多级标题末尾有标点才需要提示
    正文编号列表和Word原生编号段落不算标题
    返回：[(idx, text, punct_char), ...]
    """
    issues = []
    for idx, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        if not text or len(text) <= 5:
            continue

        level = detect_level(text)
        wnl = item[3] if len(item) > 3 else None

        # 有文本编号前缀才算标题候选
        if not has_text_number_prefix(text):
            continue

        # 区分"数字编号的标题"和"数字编号的正文"
        is_digit_prefix = bool(re.match(r'^\d+[.、．]', text))
        if is_digit_prefix:
            content_after_num = re.sub(r'^\d+[.、．]\s*', '', text)
            is_likely_body = (
                len(text) > 25
                or '。' in text
                or '；' in text
                or (len(content_after_num) > 15 and text.rstrip()[-1] == '。')
            )
            if is_likely_body:
                continue

        # Word编号段落排除
        is_word_num_body = (wnl is not None and not has_text_number_prefix(text) and len(text) > 25)
        if is_word_num_body:
            continue

        # 有编号前缀但内容超长的是正文
        if len(text) > 30 and ('。' in text or '；' in text):
            continue

        # 排除Word原生十进制编号的正文列表
        orig_num_id = item[4] if len(item) > 4 else None
        if orig_num_id and orig_num_id != '0':
            an_id = num_to_abstract.get(orig_num_id)
            if an_id:
                levels = abstract_num_defs.get(an_id, {})
                nilvl = item[5] if len(item) > 5 else '0'
                fmt, _ = levels.get(nilvl, (None, None))
                if fmt == 'decimal':
                    continue

        # 标题应以非句号结尾
        last_char = text.rstrip()[-1]
        if last_char in ('。', '；', '，', '：', ':'):
            issues.append((idx, text, last_char))

    return issues


def check_list_numbering_restart(paragraphs_text, num_seq=None):
    """检测一级标题下编号列表的起始编号是否正确。

    每个一级标题下的编号列表应从 1 开始重新编序。
    num_seq: {idx: 序号}，用于获取 Word 自带给编号的实际序号。
    返回：[(text_prefix, suggestion, anchor_type), ...]
    """
    if num_seq is None:
        num_seq = {}
    issues = []
    current_h1_idx = None
    expected_counter = 0  # h1/h2 区间内的期望计数（h2 边界重置）
    last_text_h1_num = 0  # 上一级标题的编号值（一→1, 二→2, ...）

    for i, item in enumerate(paragraphs_text):
        if item[0] != 'p':
            continue
        text = item[1].strip()
        wnl = item[3] if len(item) > 3 else None
        nid = item[4] if len(item) > 4 else None

        # 检测一级标题
        is_text_h1 = bool(re.match(r'^[一二三四五六七八九十]+、', text))
        is_word_h1 = (wnl == 'h1')
        if is_text_h1 or is_word_h1:
            # 检测文字编号的一级标题序号是否连续
            if is_text_h1:
                m = re.match(r'^([一二三四五六七八九十]+)、', text)
                if m:
                    cur_num = CNUM_TO_INT.get(m.group(1), 0)
                    if last_text_h1_num > 0 and cur_num != last_text_h1_num + 1:
                        expected_cn = CN_NUMBERS[last_text_h1_num] if last_text_h1_num < len(CN_NUMBERS) else str(last_text_h1_num + 1)
                        issues.append((
                            text[:20],
                            f'一级标题编号不连续，上一标题为"{CN_NUMBERS[last_text_h1_num - 1] if last_text_h1_num <= len(CN_NUMBERS) else ""}、"，当前为"{m.group(1)}、"，建议改为"{CN_NUMBERS[last_text_h1_num] if last_text_h1_num < len(CN_NUMBERS) else ""}、"',
                            'number_prefix'
                        ))
                    last_text_h1_num = cur_num
            current_h1_idx = i
            expected_counter = 0  # 新 h1 下计数器重置
            continue

        # h2 标题：重置编号计数（每个 h2 下都应从 1 开始）
        is_text_h2 = bool(re.match(r'^（[一二三四五六七八九十]+）', text))
        is_word_h2 = (wnl == 'h2')
        if is_text_h2 or is_word_h2:
            expected_counter = 0
            continue

        if current_h1_idx is None:
            continue

        # 提取当前段落的实际编号值
        num_val = None
        num_prefix = None

        # 方式1：文本中可见的数字编号
        txt_match = re.match(r'^(\d+)([.、．])', text)
        if txt_match:
            num_val = int(txt_match.group(1))
            num_prefix = txt_match.group(0)
            # 子编号（如 1.1 2.1.3）不计入顶层编号序列
            if re.match(r'^\d+\.\d{1,2}(?!\d)', text.lstrip()):
                continue

        # 方式2：Word 自带编号（用 num_seq 获取实际序号）
        if num_val is None and nid and nid != '0':
            # Word编号项：跳过，不参与文本编号连续性检查
            continue
        if num_val is None:
            continue

        expected_counter += 1
        expected = expected_counter

        if num_val != expected:
            # Word编号：输出段落带编号前缀（如"3.看板上..."），锚点需含编号
            is_word = (nid and nid != '0')
            if is_word:
                anchor = f'{num_val}.{text[:15]}'  # 输出为"3.看板上数据..."
            else:
                anchor = num_prefix or text[:20]
            issues.append((
                anchor,
                f'一级标题下的编号应从"{expected}"开始，当前为"{num_val}"，建议将"{num_val}"改为"{expected}"',
                'number_prefix'
            ))

    return issues
