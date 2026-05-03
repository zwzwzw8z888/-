#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中建四局公文格式化 — 格式常量 + 全局规则集中声明

所有格式常量和全局规则函数统一在此定义，
其他模块通过 from constants import * 使用。
"""

import re

# ────────────────────────── 格式常量 ──────────────────────────
FONT_FANGSONG     = "仿宋_GB2312"
FONT_HEITI        = "黑体"
FONT_KAITI        = "楷体_GB2312"
FONT_XIAOBIAOSONG = "方正小标宋简体"
FONT_TIMES_NEW_ROMAN = "仿宋_GB2312"

SIZE_CHUHAO  = 42  # Pt(42) — 调用处自行包装
SIZE_ERHAO   = 22
SIZE_SANHAO  = 16
SIZE_XIAOSI  = 12

LINE_SPACING_TWIPS = 579  # 28.9磅 = 579 twips

MARGIN_TOP    = 3.7   # cm
MARGIN_BOTTOM = 3.5
MARGIN_LEFT   = 2.8
MARGIN_RIGHT  = 2.6

CN_NUMBERS = ['一','二','三','四','五','六','七','八','九','十',
              '十一','十二','十三','十四','十五','十六','十七','十八','十九','二十']
CNUM = {str(i+1): s for i, s in enumerate(CN_NUMBERS)}
CNUM_TO_INT = {s: i+1 for i, s in enumerate(CN_NUMBERS)}
CIRCLE_NUMBERS = ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩',
                  '⑪','⑫','⑬','⑭','⑮','⑯','⑰','⑱','⑲','⑳']


# ────────────────────────── 全局规则函数 ──────────────────────────
# 集中管理所有全局规则，单一来源，避免散落导致冲突

def is_date_line(text):
    """全局规则：日期行排除（带括号格式）
    匹配格式：（2026年4月28日）或 (2026年4月28日)
    用途：is_main_title / is_short_title / is_after_main_title 三处统一调用
    """
    return bool(re.match(r'^[（(]\d{4}年\d{1,2}月\d{1,2}日[）)]$', text.strip()))


def is_signature_date(text):
    """全局规则：落款日期行检测（不带括号）
    匹配格式：2026年4月23日
    用途：落款格式处理（右对齐、右空4字）
    """
    return bool(re.match(r'^\d{4}年\d{1,2}月\d{1,2}日\s*$', text.strip()))


def is_table_title(text):
    """全局规则：表格标题识别
    匹配格式：表1、表2、表3-1 等
    用途：正文段落居中 + 不缩进
    """
    return bool(re.match(r'^表\s*\d+', text.strip()))


def is_pure_data_line(text):
    """纯数字/数据行排除
    匹配格式：纯数字、标点、运算符组成的行
    用途：is_main_title / is_short_title 排除
    """
    return bool(re.match(r'^[\d,.\-+/：:；，。、]+$', text.strip()))


def should_preserve_bold(level, is_bold, text):
    """全局规则1：原文加粗的，格式化后也保持加粗。

    标题(h1-h5)及正文均适用。不加过滤条件——对"一是/二是"、
    "1.xxx"、"商务部"等主观加粗一视同仁，原文怎样就怎样。
    """
    return bool(is_bold)


def is_greeting(text):
    """问候语/主送机关识别
    匹配格式：含称呼关键词 + 以冒号结尾的短文本
    包含：问候语（尊敬、各位领导）、主送机关（局属各单位、总部各部门）
    用途：正文段落不缩进
    """
    greeting_kw = '领导|同事|各位|尊敬|您好|下午好|上午好|你好'
    recipient_kw = '单位|部门|公司|分公司|事业部'
    t = text.strip()
    return bool(
        re.match(r'^.{2,30}[：:]$', t) and (
            re.search(greeting_kw, t) or re.search(recipient_kw, t)
        )
    )


def has_text_number_prefix(text):
    """文本编号前缀检测
    匹配：一、 / （一） / 1. / （1） / ① 等格式
    用途：多处层级判断共用
    """
    return bool(
        re.match(r'^[一二三四五六七八九十]+、', text)
        or re.match(r'^（[一二三四五六七八九十]+）', text)
        or re.match(r'^\d+[.、．]\s*', text)
        or re.match(r'^（\d+）', text)
        or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
    )


def is_verb_after_number(text):
    """编号+动词格式检测
    匹配：1.是 / 2.要 / 3.以 等不规范格式
    """
    return bool(re.match(r'^\d+[.、．]\s*[是是以要为将把让使被]', text))
