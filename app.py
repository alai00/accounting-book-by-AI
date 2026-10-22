# -*- coding: utf-8 -*-
"""
记账本 - 本地桌面客户端后端
架构：Python 标准库 HTTP 服务器 + pywebview 窗口 + Win32 原生文件对话框
前端用 fetch 调用 /api/* 端点，完全不依赖 pywebview JS bridge（解决 JS bridge 注入不可靠问题）
文件对话框用 ctypes 调用 comdlg32（tkinter 在子线程会死锁）
"""
import os
import re
import sys
import json
import socket
import threading
import datetime
import ctypes
from ctypes import wintypes
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

import webview
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# 路径工具（兼容 PyInstaller 打包）
# ============================================================
def resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


DATA_FILE = os.path.join(app_dir(), '记账数据.json')


# ============================================================
# 默认值
# ============================================================
def default_categories():
    return {
        'expense': [
            {'name': '饮食', 'subs': []},
            {'name': '购物', 'subs': []},
            {'name': '交通', 'subs': []},
            {'name': '生活消费', 'subs': []},
            {'name': '特殊开支', 'subs': []},
        ],
        'income': [
            {'name': '工资', 'subs': []},
            {'name': '理财', 'subs': []},
            {'name': '其他', 'subs': []},
        ],
    }


def default_prefs():
    return {
        'filterType': 'all', 'filterCategory': '', 'filterFrom': '', 'filterTo': '',
        'search': '', 'statsPeriod': 'month', 'periodOffset': 0,
        'statsFlow': 'expense', 'catLevel': '1', 'view': 'A',
    }


def _num(v):
    try:
        x = float(v)
        return x if x == x else 0
    except Exception:
        return 0


def _round2(n):
    return round(_num(n), 2)


def _fmt_md(n):
    v = round(_num(n), 2)
    return '\u00a5' + format(abs(v), ',.2f')


# ============================================================
# 数据读写
# ============================================================
def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                d = json.load(f)
            if not isinstance(d, dict):
                raise ValueError('bad format')
            d.setdefault('records', [])
            d.setdefault('categories', default_categories())
            d.setdefault('prefs', default_prefs())
            for r in d['records']:
                if 'category' not in r:
                    r['category'] = '\u5176\u4ed6' if r.get('income', 0) > 0 else '\u672a\u5206\u7c7b'
                if 'subcategory' not in r:
                    r['subcategory'] = ''
                if 'type' not in r:
                    r['type'] = 'income' if r.get('income', 0) > 0 else 'expense'
            return d
        except Exception as e:
            print('\u52a0\u8f7d\u6570\u636e\u5931\u8d25\uff1a', e, file=sys.stderr)
    return {'records': [], 'categories': default_categories(), 'prefs': default_prefs()}


def save_data_full(records, categories, prefs):
    d = {'records': records, 'categories': categories, 'prefs': prefs}
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print('\u4fdd\u5b58\u5931\u8d25\uff1a', e, file=sys.stderr)
        return False


def save_data_partial(records=None, categories=None, prefs=None):
    d = load_data()
    if records is not None:
        d['records'] = records
    if categories is not None:
        d['categories'] = categories
    if prefs is not None:
        d['prefs'] = prefs
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print('\u4fdd\u5b58\u5931\u8d25\uff1a', e, file=sys.stderr)
        return False


# ============================================================
# 文件对话框（Windows 原生 API，不依赖 tkinter/pywebview）
# tkinter filedialog 在子线程调用会死锁（Tk() 必须在主线程），
# HTTP 请求处理在子线程，改用 ctypes 调用 comdlg32 的原生对话框
# ============================================================
OFN_OVERWRITEPROMPT = 0x00000002
OFN_FILEMUSTEXIST = 0x00001000
OFN_PATHMUSTEXIST = 0x00000800
OFN_NOCHANGEDIR = 0x00000008


class OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ('lStructSize', wintypes.DWORD),
        ('hwndOwner', wintypes.HWND),
        ('hInstance', wintypes.HINSTANCE),
        ('lpstrFilter', wintypes.LPCWSTR),
        ('lpstrCustomFilter', wintypes.LPWSTR),
        ('nMaxCustFilter', wintypes.DWORD),
        ('nFilterIndex', wintypes.DWORD),
        ('lpstrFile', wintypes.LPWSTR),
        ('nMaxFile', wintypes.DWORD),
        ('lpstrFileTitle', wintypes.LPWSTR),
        ('nMaxFileTitle', wintypes.DWORD),
        ('lpstrInitialDir', wintypes.LPCWSTR),
        ('lpstrTitle', wintypes.LPCWSTR),
        ('Flags', wintypes.DWORD),
        ('nFileExtension', wintypes.WORD),
        ('lpstrDefExt', wintypes.LPCWSTR),
        ('lCustData', wintypes.LPARAM),
        ('lpfnHook', wintypes.LPVOID),
        ('lpTemplateName', wintypes.LPCWSTR),
        ('pvReserved', wintypes.LPVOID),
        ('dwReserved', wintypes.DWORD),
        ('FlagsEx', wintypes.DWORD),
    ]


def save_dialog(default_name, filetypes):
    """Windows 原生保存对话框，可在任意线程安全调用"""
    try:
        # 过滤器字符串：Name1\0*.ext1\0Name2\0*.ext2\0\0
        filter_str = ''.join('{}\0{}\0'.format(name, ext) for name, ext in filetypes) + '\0'
        buf = ctypes.create_unicode_buffer(520)
        buf.value = default_name or ''
        ofn = OPENFILENAMEW()
        ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
        ofn.hwndOwner = None
        ofn.lpstrFilter = filter_str
        ofn.nFilterIndex = 1
        ofn.lpstrFile = buf
        ofn.nMaxFile = 520
        ofn.lpstrTitle = '\u4fdd\u5b58\u6587\u4ef6'
        ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR
        # 自动设置默认扩展名
        if '.' in (default_name or ''):
            ofn.lpstrDefExt = default_name.rsplit('.', 1)[1]
        ok = ctypes.windll.comdlg32.GetSaveFileNameW(ctypes.byref(ofn))
        return buf.value if ok else None
    except Exception as e:
        print('save dialog error:', e, file=sys.stderr)
        return None


def open_dialog(filetypes):
    """Windows 原生打开文件对话框，可在任意线程安全调用"""
    try:
        filter_str = ''.join('{}\0{}\0'.format(name, ext) for name, ext in filetypes) + '\0'
        buf = ctypes.create_unicode_buffer(520)
        ofn = OPENFILENAMEW()
        ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
        ofn.hwndOwner = None
        ofn.lpstrFilter = filter_str
        ofn.nFilterIndex = 1
        ofn.lpstrFile = buf
        ofn.nMaxFile = 520
        ofn.lpstrTitle = '\u6253\u5f00\u6587\u4ef6'
        ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR
        ok = ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn))
        return buf.value if ok else None
    except Exception as e:
        print('open dialog error:', e, file=sys.stderr)
        return None


# ============================================================
# Markdown 导出
# ============================================================
def build_md(records):
    recs = sorted(records, key=lambda r: (r.get('date', ''), r.get('createdAt', '')), reverse=True)
    total_inc = sum(_num(r.get('income', 0)) for r in recs)
    total_exp = sum(_num(r.get('expense', 0)) for r in recs)
    now = datetime.datetime.now()
    L = []
    L.append('# \u8bb0\u8d26\u62a5\u544a')
    L.append('')
    L.append('> \u5bfc\u51fa\u65f6\u95f4\uff1a' + now.strftime('%Y-%m-%d %H:%M'))
    L.append('> \u8bb0\u5f55\u603b\u6570\uff1a' + str(len(recs)) + ' \u6761')
    L.append('')
    L.append('## \u603b\u89c8')
    L.append('')
    L.append('| \u6307\u6807 | \u91d1\u989d |')
    L.append('|------|------:|')
    L.append('| \u603b\u6536\u5165 | ' + _fmt_md(total_inc) + ' |')
    L.append('| \u603b\u652f\u51fa | ' + _fmt_md(total_exp) + ' |')
    L.append('| \u7ed3\u4f59 | ' + _fmt_md(total_inc - total_exp) + ' |')
    L.append('')

    # 支出分类
    exp_map, exp_cnt = {}, {}
    for r in recs:
        e = _num(r.get('expense', 0))
        if e > 0:
            k = r.get('category', '\u672a\u5206\u7c7b')
            exp_map[k] = exp_map.get(k, 0) + e
            exp_cnt[k] = exp_cnt.get(k, 0) + 1
    if exp_map:
        L.append('### \u652f\u51fa')
        L.append('')
        L.append('| \u5206\u7c7b | \u7b14\u6570 | \u91d1\u989d | \u5360\u6bd4 |')
        L.append('|------|------:|------:|------:|')
        et = sum(exp_map.values())
        for k in sorted(exp_map, key=lambda x: -exp_map[x]):
            v = exp_map[k]
            pct = (v / et * 100) if et else 0
            L.append('| ' + k + ' | ' + str(exp_cnt[k]) + ' | ' + _fmt_md(v) + ' | ' + ('%.1f' % pct) + '% |')
        L.append('| **\u5408\u8ba1** | **' + str(sum(exp_cnt.values())) + '** | **' + _fmt_md(et) + '** | |')
        L.append('')

    # 收入分类
    inc_map, inc_cnt = {}, {}
    for r in recs:
        i = _num(r.get('income', 0))
        if i > 0:
            k = r.get('category', '\u672a\u5206\u7c7b')
            inc_map[k] = inc_map.get(k, 0) + i
            inc_cnt[k] = inc_cnt.get(k, 0) + 1
    if inc_map:
        L.append('### \u6536\u5165')
        L.append('')
        L.append('| \u5206\u7c7b | \u7b14\u6570 | \u91d1\u989d | \u5360\u6bd4 |')
        L.append('|------|------:|------:|------:|')
        it = sum(inc_map.values())
        for k in sorted(inc_map, key=lambda x: -inc_map[x]):
            v = inc_map[k]
            pct = (v / it * 100) if it else 0
            L.append('| ' + k + ' | ' + str(inc_cnt[k]) + ' | ' + _fmt_md(v) + ' | ' + ('%.1f' % pct) + '% |')
        L.append('| **\u5408\u8ba1** | **' + str(sum(inc_cnt.values())) + '** | **' + _fmt_md(it) + '** | |')
        L.append('')

    # 月度汇总
    m_map = {}
    for r in recs:
        ym = (r.get('date', '') or '')[:7]
        if not ym:
            continue
        if ym not in m_map:
            m_map[ym] = [0, 0]
        m_map[ym][0] += _num(r.get('income', 0))
        m_map[ym][1] += _num(r.get('expense', 0))
    if m_map:
        L.append('## \u6708\u5ea6\u6c47\u603b')
        L.append('')
        L.append('| \u6708\u4efd | \u6536\u5165 | \u652f\u51fa | \u7ed3\u4f59 |')
        L.append('|------|------:|------:|------:|')
        for ym in sorted(m_map, reverse=True):
            inc, exp = m_map[ym]
            L.append('| ' + ym + ' | ' + _fmt_md(inc) + ' | ' + _fmt_md(exp) + ' | ' + _fmt_md(inc - exp) + ' |')
        L.append('')

    # 明细
    L.append('## \u660e\u7ec6\u8bb0\u5f55')
    L.append('')
    L.append('| \u65e5\u671f | \u6536\u652f | \u5206\u7c7b | \u9879\u76ee\u5185\u5bb9 | \u6536\u5165 | \u652f\u51fa | \u5907\u6ce8 |')
    L.append('|------|------|------|------|------:|------:|------|')
    for r in recs:
        is_inc = _num(r.get('income', 0)) > 0
        cat = r.get('category', '')
        sub = r.get('subcategory', '')
        if sub:
            cat = cat + '-' + sub
        content = (r.get('content', '') or '').replace('|', '\\|').replace('\n', ' ')
        note = (r.get('note', '') or '').replace('|', '\\|').replace('\n', ' ')
        inc = _num(r.get('income', 0))
        exp = _num(r.get('expense', 0))
        L.append('| ' + (r.get('date', '') or '') + ' | ' + ('\u6536\u5165' if is_inc else '\u652f\u51fa') + ' | ' + cat +
                 ' | ' + content + ' | ' + (_fmt_md(inc) if inc else '') + ' | ' + (_fmt_md(exp) if exp else '') + ' | ' + note + ' |')
    L.append('')
    L.append('---')
    L.append('')
    L.append('*\u7531\u300c\u8bb0\u8d26\u672c\u300d\u5bfc\u51fa \u00b7 ' + now.strftime('%Y-%m-%d %H:%M') + '*')
    return '\n'.join(L)


# ============================================================
# Excel 导出
# ============================================================
def build_xlsx(path, records):
    recs = sorted(records, key=lambda r: (r.get('date', ''), r.get('createdAt', '')), reverse=True)
    wb = Workbook()

    header_font = Font(bold=True, color='FFFFFF', name='\u5fae\u8f6f\u96c5\u9ed1')
    header_fill = PatternFill('solid', fgColor='3370FF')
    inc_font = Font(color='16A34A', name='\u5fae\u8f6f\u96c5\u9ed1')
    exp_font = Font(color='E5484D', name='\u5fae\u8f6f\u96c5\u9ed1')
    bold = Font(bold=True, name='\u5fae\u8f6f\u96c5\u9ed1')
    normal = Font(name='\u5fae\u8f6f\u96c5\u9ed1')
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    thin = Side(style='thin', color='ECECEC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money_fmt = '\u00a5#,##0.00'

    # Sheet1 明细
    ws = wb.active
    ws.title = '\u660e\u7ec6'
    headers = ['\u65e5\u671f', '\u6536\u652f', '\u4e00\u7ea7\u5206\u7c7b', '\u4e8c\u7ea7\u5206\u7c7b', '\u9879\u76ee\u5185\u5bb9', '\u6536\u5165', '\u652f\u51fa', '\u5907\u6ce8']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = header_font; cell.fill = header_fill; cell.alignment = center; cell.border = border
    for i, r in enumerate(recs, 2):
        inc = _round2(r.get('income', 0))
        exp = _round2(r.get('expense', 0))
        is_inc = inc > 0
        row = [r.get('date', ''), '\u6536\u5165' if is_inc else '\u652f\u51fa', r.get('category', ''),
               r.get('subcategory', ''), r.get('content', ''), inc if inc else None, exp if exp else None, r.get('note', '')]
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = border; cell.font = normal
            if c == 2:
                cell.font = inc_font if is_inc else exp_font; cell.alignment = center
            elif c in (6, 7):
                cell.number_format = money_fmt; cell.alignment = right
                if v: cell.font = inc_font if c == 6 else exp_font
            elif c == 1: cell.alignment = center
            else: cell.alignment = left
    widths = [12, 8, 12, 12, 26, 14, 14, 26]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    # Sheet2 分类汇总
    ws2 = wb.create_sheet('\u5206\u7c7b\u6c47\u603b')
    exp_map, exp_cnt = {}, {}
    for r in recs:
        e = _num(r.get('expense', 0))
        if e > 0:
            k = r.get('category', '\u672a\u5206\u7c7b')
            exp_map[k] = exp_map.get(k, 0) + e
            exp_cnt[k] = exp_cnt.get(k, 0) + 1
    ws2.cell(row=1, column=1, value='\u652f\u51fa\u5206\u7c7b\u6c47\u603b').font = Font(bold=True, size=13, color='E5484D', name='\u5fae\u8f6f\u96c5\u9ed1')
    hdr = ['\u5206\u7c7b', '\u7b14\u6570', '\u91d1\u989d', '\u5360\u6bd4']
    for c, h in enumerate(hdr, 1):
        cell = ws2.cell(row=2, column=c, value=h)
        cell.font = header_font; cell.fill = header_fill; cell.alignment = center; cell.border = border
    et = sum(exp_map.values())
    row = 3
    for k in sorted(exp_map, key=lambda x: -exp_map[x]):
        v = exp_map[k]
        ws2.cell(row=row, column=1, value=k).border = border; ws2.cell(row=row, column=1).font = normal
        ws2.cell(row=row, column=2, value=exp_cnt[k]).border = border; ws2.cell(row=row, column=2).alignment = center; ws2.cell(row=row, column=2).font = normal
        c3 = ws2.cell(row=row, column=3, value=round(v, 2)); c3.number_format = money_fmt; c3.border = border; c3.font = exp_font
        c4 = ws2.cell(row=row, column=4, value=(v / et) if et else 0); c4.number_format = '0.0%'; c4.border = border; c4.alignment = center
        row += 1
    if exp_map:
        ws2.cell(row=row, column=1, value='\u5408\u8ba1').font = bold
        c = ws2.cell(row=row, column=2, value=sum(exp_cnt.values())); c.font = bold; c.border = border; c.alignment = center
        c = ws2.cell(row=row, column=3, value=round(et, 2)); c.number_format = money_fmt; c.font = bold; c.border = border
        c = ws2.cell(row=row, column=4, value=1); c.number_format = '0.0%'; c.font = bold; c.border = border; c.alignment = center
        row += 1

    start = row + 2
    ws2.cell(row=start, column=1, value='\u6536\u5165\u5206\u7c7b\u6c47\u603b').font = Font(bold=True, size=13, color='16A34A', name='\u5fae\u8f6f\u96c5\u9ed1')
    for c, h in enumerate(hdr, 1):
        cell = ws2.cell(row=start + 1, column=c, value=h)
        cell.font = header_font; cell.fill = header_fill; cell.alignment = center; cell.border = border
    inc_map, inc_cnt = {}, {}
    for r in recs:
        i = _num(r.get('income', 0))
        if i > 0:
            k = r.get('category', '\u672a\u5206\u7c7b')
            inc_map[k] = inc_map.get(k, 0) + i
            inc_cnt[k] = inc_cnt.get(k, 0) + 1
    it = sum(inc_map.values())
    rr = start + 2
    for k in sorted(inc_map, key=lambda x: -inc_map[x]):
        v = inc_map[k]
        ws2.cell(row=rr, column=1, value=k).border = border; ws2.cell(row=rr, column=1).font = normal
        ws2.cell(row=rr, column=2, value=inc_cnt[k]).border = border; ws2.cell(row=rr, column=2).alignment = center; ws2.cell(row=rr, column=2).font = normal
        c3 = ws2.cell(row=rr, column=3, value=round(v, 2)); c3.number_format = money_fmt; c3.border = border; c3.font = inc_font
        c4 = ws2.cell(row=rr, column=4, value=(v / it) if it else 0); c4.number_format = '0.0%'; c4.border = border; c4.alignment = center
        rr += 1
    if inc_map:
        ws2.cell(row=rr, column=1, value='\u5408\u8ba1').font = bold
        c = ws2.cell(row=rr, column=2, value=sum(inc_cnt.values())); c.font = bold; c.border = border; c.alignment = center
        c = ws2.cell(row=rr, column=3, value=round(it, 2)); c.number_format = money_fmt; c.font = bold; c.border = border
        c = ws2.cell(row=rr, column=4, value=1); c.number_format = '0.0%'; c.font = bold; c.border = border; c.alignment = center
    for i, w in enumerate([16, 8, 16, 10], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # Sheet3 月度汇总
    ws3 = wb.create_sheet('\u6708\u5ea6\u6c47\u603b')
    hdr3 = ['\u6708\u4efd', '\u6536\u5165', '\u652f\u51fa', '\u7ed3\u4f59']
    for c, h in enumerate(hdr3, 1):
        cell = ws3.cell(row=1, column=c, value=h)
        cell.font = header_font; cell.fill = header_fill; cell.alignment = center; cell.border = border
    m_map = {}
    for r in recs:
        ym = (r.get('date', '') or '')[:7]
        if not ym:
            continue
        if ym not in m_map:
            m_map[ym] = [0, 0]
        m_map[ym][0] += _num(r.get('income', 0))
        m_map[ym][1] += _num(r.get('expense', 0))
    rr = 2
    for ym in sorted(m_map, reverse=True):
        inc, exp = m_map[ym]
        ws3.cell(row=rr, column=1, value=ym).border = border; ws3.cell(row=rr, column=1).alignment = center; ws3.cell(row=rr, column=1).font = normal
        c = ws3.cell(row=rr, column=2, value=round(inc, 2)); c.number_format = money_fmt; c.border = border; c.font = inc_font if inc else normal
        c = ws3.cell(row=rr, column=3, value=round(exp, 2)); c.number_format = money_fmt; c.border = border; c.font = exp_font if exp else normal
        c = ws3.cell(row=rr, column=4, value=round(inc - exp, 2)); c.number_format = money_fmt; c.border = border; c.font = bold
        rr += 1
    ti = sum(v[0] for v in m_map.values()); te = sum(v[1] for v in m_map.values())
    if m_map:
        ws3.cell(row=rr, column=1, value='\u5408\u8ba1').font = bold
        c = ws3.cell(row=rr, column=2, value=round(ti, 2)); c.number_format = money_fmt; c.font = bold; c.border = border
        c = ws3.cell(row=rr, column=3, value=round(te, 2)); c.number_format = money_fmt; c.font = bold; c.border = border
        c = ws3.cell(row=rr, column=4, value=round(ti - te, 2)); c.number_format = money_fmt; c.font = bold; c.border = border
    for i, w in enumerate([12, 14, 14, 14], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)


# ============================================================
# 导入校验
# ============================================================
def validate_record(obj):
    if not isinstance(obj, dict):
        return None
    if not isinstance(obj.get('id'), str) or not obj['id']:
        return None
    if not isinstance(obj.get('date'), str) or not re.match(r'^\d{4}-\d{2}-\d{2}$', obj['date']):
        return None
    if not isinstance(obj.get('content'), str) or not obj['content'].strip():
        return None
    income = _num(obj.get('income', 0))
    expense = _num(obj.get('expense', 0))
    if income < 0 or expense < 0:
        return None
    if income == 0 and expense == 0:
        return None
    t = obj.get('type')
    if t not in ('income', 'expense'):
        t = 'income' if income > 0 else 'expense'
    cat = obj.get('category')
    if not isinstance(cat, str) or not cat.strip():
        cat = '\u5176\u4ed6' if income > 0 else '\u672a\u5206\u7c7b'
    sub = obj.get('subcategory', '')
    if not isinstance(sub, str):
        sub = ''
    note = obj.get('note', '')
    if not isinstance(note, str):
        note = ''
    created = obj.get('createdAt') or (obj['date'] + 'T00:00:00.000Z')
    updated = obj.get('updatedAt') or (obj['date'] + 'T00:00:00.000Z')
    return {
        'id': obj['id'], 'date': obj['date'], 'type': t,
        'category': cat.strip(), 'subcategory': sub,
        'content': obj['content'].strip(),
        'income': round(income, 2), 'expense': round(expense, 2),
        'note': note, 'createdAt': created, 'updatedAt': updated,
    }


# ============================================================
# HTTP 服务器
# ============================================================
class RequestHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/api/load':
            self._send_json(load_data())
        elif self.path == '/' or self.path == '/index.html':
            self._serve_file('index.html', 'text/html; charset=utf-8')
        elif self.path.startswith('/lib/'):
            mime = 'application/javascript'
            self._serve_file(self.path.lstrip('/'), mime)
        else:
            self._send_json({'error': 'not found'}, 404)

    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length).decode('utf-8') if content_length else '{}'
        try:
            data = json.loads(body)
        except Exception:
            data = {}

        if self.path == '/api/save':
            records = data.get('records')
            categories = data.get('categories')
            prefs = data.get('prefs')
            ok = save_data_partial(records, categories, prefs)
            self._send_json({'ok': ok})

        elif self.path == '/api/export/md':
            records = data.get('records', [])
            md = build_md(records)
            path = self._save_export('\u8bb0\u8d26\u62a5\u544a_' + datetime.date.today().isoformat() + '.md')
            if path is None:
                self._send_json({'ok': False, 'msg': '\u521b\u5efa\u5bfc\u51fa\u76ee\u5f55\u5931\u8d25'})
                return
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(md)
                self._reveal_file(path)
                self._send_json({'ok': True, 'path': path})
            except Exception as e:
                self._send_json({'ok': False, 'msg': str(e)})

        elif self.path == '/api/export/xlsx':
            records = data.get('records', [])
            path = self._save_export('\u8bb0\u8d26\u660e\u7ec6_' + datetime.date.today().isoformat() + '.xlsx')
            if path is None:
                self._send_json({'ok': False, 'msg': '\u521b\u5efa\u5bfc\u51fa\u76ee\u5f55\u5931\u8d25'})
                return
            try:
                build_xlsx(path, records)
                self._reveal_file(path)
                self._send_json({'ok': True, 'path': path})
            except Exception as e:
                self._send_json({'ok': False, 'msg': str(e)})

        elif self.path == '/api/export/json':
            records = data.get('records', [])
            categories = data.get('categories', default_categories())
            payload = {
                'app': 'income_outcome', 'version': 2,
                'exportedAt': datetime.datetime.now().isoformat(),
                'records': records, 'categories': categories,
            }
            path = self._save_export('\u8bb0\u8d26\u5907\u4efd_' + datetime.date.today().isoformat() + '.json')
            if path is None:
                self._send_json({'ok': False, 'msg': '\u521b\u5efa\u5bfc\u51fa\u76ee\u5f55\u5931\u8d25'})
                return
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
                self._reveal_file(path)
                self._send_json({'ok': True, 'path': path})
            except Exception as e:
                self._send_json({'ok': False, 'msg': str(e)})

        elif self.path == '/api/import':
            # 前端用 <input type="file"> 选文件并上传内容，后端不弹对话框
            content = data.get('content', '')
            if not content:
                self._send_json({'ok': False, 'msg': '\u672a\u63d0\u4f9b\u6587\u4ef6\u5185\u5bb9'})
                return
            try:
                payload = json.loads(content)
            except Exception:
                self._send_json({'ok': False, 'msg': '\u6587\u4ef6\u683c\u5f0f\u9519\u8bef\uff0c\u4e0d\u662f\u6709\u6548\u7684 JSON'})
                return
            if not isinstance(payload, (list, dict)):
                self._send_json({'ok': False, 'msg': '\u6587\u4ef6\u5185\u5bb9\u4e0d\u662f\u6709\u6548\u7684\u8bb0\u8d26\u5907\u4efd'})
                return
            raw = payload if isinstance(payload, list) else payload.get('records')
            if not isinstance(raw, list):
                self._send_json({'ok': False, 'msg': '\u6587\u4ef6\u4e2d\u672a\u627e\u5230 records \u6570\u636e'})
                return
            valid = []
            invalid = 0
            for item in raw:
                rec = validate_record(item)
                if rec:
                    valid.append(rec)
                else:
                    invalid += 1
            imported_cats = payload.get('categories') if isinstance(payload, dict) else None
            self._send_json({'ok': True, 'records': valid, 'categories': imported_cats, 'invalid': invalid})

        elif self.path == '/api/reveal':
            folder = os.path.dirname(DATA_FILE) or '.'
            if sys.platform == 'win32':
                os.system('explorer /select,"' + DATA_FILE.replace('/', '\\') + '"')
            self._send_json({'ok': True, 'path': DATA_FILE, 'folder': folder})

        else:
            self._send_json({'error': 'not found'}, 404)

    def _send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _serve_file(self, rel, mime):
        path = resource_path(rel)
        if not os.path.exists(path):
            self.send_error(404)
            return
        try:
            with open(path, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', mime)
            self.send_header('Content-Length', str(len(content)))
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_error(500, str(e))

    def _save_export(self, filename):
        """保存到 exe 同目录下的'导出'文件夹，返回完整路径"""
        try:
            export_dir = os.path.join(app_dir(), '\u5bfc\u51fa')
            os.makedirs(export_dir, exist_ok=True)
            return os.path.join(export_dir, filename)
        except Exception as e:
            print('export dir error:', e, file=sys.stderr)
            return None

    def _reveal_file(self, path):
        """在资源管理器中选中指定文件（与 /api/reveal 同机制，已验证可用）"""
        try:
            if sys.platform == 'win32':
                os.system('explorer /select,"' + path.replace('/', '\\') + '"')
        except Exception:
            pass

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, *args):
        pass


def get_random_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('localhost', 0))
        return s.getsockname()[1]


# ============================================================
# 启动
# ============================================================
def main():
    port = get_random_port()
    server = ThreadingHTTPServer(('localhost', port), RequestHandler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

    webview.create_window(
        '\u8bb0\u8d26\u672c',
        'http://localhost:{}/'.format(port),
        width=1280, height=860,
        min_size=(900, 600),
        text_select=True,
    )
    webview.start()
    server.shutdown()


if __name__ == '__main__':
    main()
