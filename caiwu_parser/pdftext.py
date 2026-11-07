# -*- coding: utf-8 -*-
"""
pdftext.py —— 整本 PDF 的快速文字层（pypdfium2，约为 pdfplumber 的 20–40 倍速）

只用来「找页」：快速定位报表标题页与全局金额单位声明。
真正取表格数字仍然用 pdfplumber（只在少数几页报表页上运行）。
同一进程内按文件路径缓存，避免重复提取。
"""
import os
import re

_CACHE = {}


def pages(path):
	"""→ [(字符数, 原始文字)]，按页序；失败返回 None（调用方退回 pdfplumber）"""
	key = os.path.abspath(path)
	if key in _CACHE:
		return _CACHE[key]
	try:
		import pypdfium2 as pdfium
		doc = pdfium.PdfDocument(path)
		out = []
		try:
			for i in range(len(doc)):
				pg = doc[i]
				tp = pg.get_textpage()
				try:
					out.append((tp.count_chars(), tp.get_text_range() or ""))
				finally:
					tp.close()
					pg.close()
		finally:
			doc.close()
	except Exception:
		out = None
	_CACHE[key] = out
	return out


def text(path, i):
	"""第 i 页（0 起）的原始文字；没有快速文字层时返回 None"""
	ps = pages(path)
	if ps is None or i >= len(ps):
		return None
	return ps[i][1]


def flat(path, i):
	"""第 i 页去掉全部空白后的文字"""
	t = text(path, i)
	return None if t is None else re.sub(r"\s+", "", t)


def lines(path, i):
	t = text(path, i)
	return None if t is None else t.replace("\r\n", "\n").replace("\r", "\n").split("\n")
