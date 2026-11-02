import sys
import io
import os
import json
import copy
import zipfile
import argparse
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from layout_rules import LayoutPlan, strip_math_number  # noqa: E402

# Ensure stdout uses utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DRAWING_A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

ET.register_namespace('w', WORD_NS)
ET.register_namespace('m', OMML_NS)

# Office 常用命名空间的标准前缀；用于修复被旧版本写成 ns1 的文件，并补全 mc:Ignorable 引用的前缀
KNOWN_NS = {
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': DRAWING_A_NS if 'DRAWING_A_NS' in globals() else 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml', 'o': 'urn:schemas-microsoft-com:office:office',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
    'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
    'w16sdtdh': 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
    'w16du': 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
}
for _p, _u in KNOWN_NS.items():
    ET.register_namespace(_p, _u)

MATH_FUNCS = (
    'arcsin', 'arccos', 'arctan', 'arcsec', 'arccsc', 'arccot',
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'ln', 'log', 'lg', 'exp',
    'lim', 'max', 'min', 'sup', 'inf', 'det', 'tr', 'sgn',
    'sinh', 'cosh', 'tanh', 'coth', 'arg', 'gcd', 'dim', 'ker', 'deg'
)

XML_NS = 'http://www.w3.org/XML/1998/namespace'

# 所有积分类 n 元运算符 (∫∬∭∮∯∰∱∲∳⨌)
INTEGRAL_CHARS = ('∫', '∬', '∭', '∮', '∯', '∰', '∱', '∲', '∳', '⨌')
# 所有求和/乘积/并交类大型运算符
BIGOP_CHARS = ('∑', '∏', '∐', '⋃', '⋂', '⋁', '⋀', '⨁', '⨂', '⨀', '⨄', '⨆', 'sum', 'prod')

# OMML 属性元素的 schema 顺序 (ECMA-376 shared-math.xsd)，插入新子元素时必须遵守，否则新版 Word 会报“内容有问题”
NARY_PR_ORDER = ('chr', 'limLoc', 'grow', 'subHide', 'supHide', 'ctrlPr')
D_PR_ORDER = ('begChr', 'sepChr', 'endChr', 'grow', 'shp', 'ctrlPr')

# 包含公式的所有 Word 文档部件 (正文、页眉页脚、脚注尾注、批注)
import re as _re
MATH_PART_RE = _re.compile(r'^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$')
PART_LABELS = {'header': '页眉', 'footer': '页脚', 'footnotes': '脚注', 'endnotes': '尾注', 'comments': '批注'}

# 段落遍历时不进入的元素 (嵌套段落由 findall('.//w:p') 单独处理；删除的修订不可见)
WALK_SKIP_TAGS = ('p', 'txbxContent', 'del', 'moveFrom', 'pPr', 'rPr', 'sectPr')


def mtag(name):
    return f'{{{OMML_NS}}}{name}'


def local(tag):
    return tag.split('}')[-1] if isinstance(tag, str) else ''


def find_parent(root, node):
    for p in root.iter():
        for c in p:
            if c is node:
                return p
    return None


def set_ordered_prop(pr, name, val, order):
    """在 xxxPr 中设置 <m:name m:val=val/>，并按 schema 顺序插入。返回是否有改动。"""
    node = pr.find(mtag(name))
    if node is not None:
        cur = node.attrib.get(mtag('val'), 'on')
        same = {'on': ('on', '1', 'true'), 'off': ('off', '0', 'false')}.get(val, (val,))
        if cur in same:
            return False
        node.set(mtag('val'), val)
        return True
    node = ET.Element(mtag(name))
    node.set(mtag('val'), val)
    my_rank = order.index(name)
    insert_at = len(pr)
    for i, child in enumerate(list(pr)):
        n = local(child.tag)
        if n in order and order.index(n) > my_rank:
            insert_at = i
            break
    pr.insert(insert_at, node)
    return True


def ensure_m_rpr(r):
    """m:r 的子元素顺序必须是 m:rPr → w:rPr → m:t。"""
    rPr = r.find(mtag('rPr'))
    if rPr is None:
        rPr = ET.Element(mtag('rPr'))
        r.insert(0, rPr)
    return rPr


def ensure_w_rpr(r):
    w_rPr = r.find(f'{{{WORD_NS}}}rPr')
    if w_rPr is None:
        w_rPr = ET.Element(f'{{{WORD_NS}}}rPr')
        idx = 1 if r.find(mtag('rPr')) is not None else 0
        r.insert(idx, w_rPr)
    return w_rPr


def normalize_schema_order(omath):
    """修正旧版本或其他工具写出的错误子元素顺序：m:r 内 m:rPr → w:rPr → 内容；m:e 内 ctrlPr 置于末尾。"""
    rank_r = {mtag('rPr'): 0, f'{{{WORD_NS}}}rPr': 1}
    for node in omath.iter():
        kids = list(node)
        if not kids:
            continue
        if node.tag == mtag('r'):
            ordered = sorted(kids, key=lambda c: rank_r.get(c.tag, 2))
        elif local(node.tag) in ('e', 'num', 'den', 'sub', 'sup', 'lim', 'deg', 'fName', 'oMath'):
            ordered = sorted(kids, key=lambda c: 1 if c.tag == mtag('ctrlPr') else 0)
            ordered = sorted(ordered, key=lambda c: 0 if c.tag == mtag('argPr') else 1) \
                if node.tag != mtag('oMath') else ordered
        else:
            continue
        if ordered != kids:
            for c in kids:
                node.remove(c)
            node.extend(ordered)


def get_sty(r):
    rPr = r.find(mtag('rPr'))
    sty = rPr.find(mtag('sty')) if rPr is not None else None
    return sty.attrib.get(mtag('val')) if sty is not None else None


def set_upright(r):
    """设为正体，保留加粗 (bi→b)。返回是否有改动。"""
    cur = get_sty(r)
    target = 'b' if cur in ('b', 'bi') else 'p'
    if cur == target:
        return False
    rPr = ensure_m_rpr(r)
    sty = rPr.find(mtag('sty'))
    if sty is None:
        sty = ET.Element(mtag('sty'))
        # m:rPr 子顺序: lit, nor, scr, sty, brk, aln
        idx = sum(1 for c in rPr if local(c.tag) in ('lit', 'nor', 'scr'))
        rPr.insert(idx, sty)
    sty.set(mtag('val'), target)
    return True


def is_normal_text(r):
    """m:nor 表示“普通文本”，是作者刻意设置的非数学文本，不应拆分或改正斜体。"""
    rPr = r.find(mtag('rPr'))
    if rPr is None:
        return False
    nor = rPr.find(mtag('nor'))
    return nor is not None and nor.attrib.get(mtag('val'), 'on') not in ('off', '0', 'false')


def run_text(r):
    return ''.join(t.text or '' for t in r.findall(mtag('t')))


def set_t_text(t, text):
    t.text = text
    if text != text.strip():
        t.set(f'{{{XML_NS}}}space', 'preserve')


def clone_run(orig, text, upright=None):
    """按原 run 的字体/字号/加粗等属性生成新 run。upright: True 正体，False 斜体，None 保持原样。"""
    r = ET.Element(mtag('r'))
    orig_sty = get_sty(orig) if orig is not None else None
    if orig is not None:
        m_rPr = orig.find(mtag('rPr'))
        if m_rPr is not None:
            m_copy = copy.deepcopy(m_rPr)
            if upright is not None:
                for s in m_copy.findall(mtag('sty')):
                    m_copy.remove(s)
            if len(m_copy):
                r.append(m_copy)
        w_rPr = orig.find(f'{{{WORD_NS}}}rPr')
        if w_rPr is not None:
            r.append(copy.deepcopy(w_rPr))
    bold = orig_sty in ('b', 'bi')
    if upright is True:
        set_upright(r)
        if bold:
            r.find(mtag('rPr')).find(mtag('sty')).set(mtag('val'), 'b')
    elif upright is False and bold:
        set_upright(r)
        r.find(mtag('rPr')).find(mtag('sty')).set(mtag('val'), 'bi')
    t = ET.SubElement(r, mtag('t'))
    set_t_text(t, text)
    return r


def is_empty_arg(node):
    """参数为空：没有文字，也没有任何数学对象 (忽略 ctrlPr/argPr 等属性节点)。"""
    if node is None:
        return True
    if ''.join(node.itertext()).strip():
        return False
    return all(local(c.tag) in ('ctrlPr', 'argPr', 'rPr') or (local(c.tag) == 'r' and not run_text(c).strip())
               for c in node)

def escape_latex_text(text):
    if not text:
        return ""
    # 先转义 LaTeX 特殊字符，再替换 Unicode 符号 (替换结果本身含 \ 和 {})
    text = text.replace('\\', '\\backslash ').replace('{', '\\{').replace('}', '\\}')
    text = text.replace('^', '\\hat{}').replace('~', '\\sim ')
    replacements = [
        ('∑', '\\sum '), ('∏', '\\prod '), ('∐', '\\coprod '), ('∫', '\\int '), ('∬', '\\iint '),
        ('∭', '\\iiint '), ('∮', '\\oint '), ('ℝ', '\\mathbb{R}'), ('ℕ', '\\mathbb{N}'),
        ('ℤ', '\\mathbb{Z}'), ('ℚ', '\\mathbb{Q}'), ('ℂ', '\\mathbb{C}'), ('∀', '\\forall '),
        ('∃', '\\exists '), ('¬', '\\neg '), ('∧', '\\land '), ('∨', '\\lor '), ('⊥', '\\perp '),
        ('∥', '\\parallel '), ('∠', '\\angle '), ('△', '\\triangle '), ('≈', '\\approx '),
        ('≡', '\\equiv '), ('∼', '\\sim '), ('≅', '\\cong '), ('∝', '\\propto '), ('⊇', '\\supseteq '),
        ('⊃', '\\supset '), ('⊊', '\\subsetneq '), ('⊄', '\\not\\subset '), ('←', '\\leftarrow '),
        ('↔', '\\leftrightarrow '), ('⟶', '\\longrightarrow '), ('⇐', '\\Leftarrow '), ('↦', '\\mapsto '),
        ('∖', '\\setminus '), ('∘', '\\circ '), ('⋅', '\\cdot '), ('∗', '\\ast '), ('ℓ', '\\ell '),
        ('∣', '\\mid '), ('‖', '\\| '), ('（', '('), ('）', ')'), ('，', ',\\,'), ('：', ':'), ('；', ';'),
        ('⩽', '\\le '), ('⩾', '\\ge '),
        ('≤', '\\le '), ('≥', '\\ge '), ('≠', '\\ne '), ('±', '\\pm '),
        ('×', '\\times '), ('·', '\\cdot '), ('÷', '\\div '), ('∈', '\\in '),
        ('∉', '\\notin '), ('⊂', '\\subset '), ('⊆', '\\subseteq '),
        ('∪', '\\cup '), ('∩', '\\cap '), ('∅', '\\emptyset '),
        ('→', '\\to '), ('⇒', '\\implies '), ('⇔', '\\iff '),
        ('∞', '\\infty '), ('∂', '\\partial '), ('∇', '\\nabla '),
        ('⋯', '\\cdots '), ('…', '\\dots '), ('⋮', '\\vdots '), ('⋱', '\\ddots '),
        ('−', '-'), ('′', "'"), ('″', "''"), ('°', "^\\circ "),
        ('%', '\\%'), ('$', '\\$'), ('#', '\\#'), ('&', '\\&'), ('_', '\\_')
    ]
    for old, new in replacements:
        text = text.replace(old, new)
    return text


def escape_text_mode(text):
    """\\text{} 内部是文本模式，不能出现 \\le 之类数学命令，只转义特殊字符。"""
    if not text:
        return ""
    out = []
    for ch in text:
        if ch in '\\':
            out.append('\\textbackslash ')
        elif ch in '{}$%#&_':
            out.append('\\' + ch)
        elif ch == '^':
            out.append('\\textasciicircum ')
        elif ch == '~':
            out.append('\\textasciitilde ')
        else:
            out.append(ch)
    return ''.join(out)


SCR_CMDS = {'double-struck': '\\mathbb', 'script': '\\mathcal', 'fraktur': '\\mathfrak',
            'sans-serif': '\\mathsf', 'monospace': '\\mathtt'}

def omml_to_latex(elem):
    """Convert an OMML XML element tree to LaTeX representation for KaTeX preview."""
    if elem is None:
        return ''
    tag = elem.tag.split('}')[-1]
    
    if tag == 't':
        t_text = elem.text or ''
        return escape_latex_text(t_text)
        
    elif tag == 'r':
        t_text = ''.join([t.text or '' for t in elem.findall(f'{{{OMML_NS}}}t')])
        if not t_text:
            return ''
        rPr = elem.find(f'{{{OMML_NS}}}rPr')
        sty = rPr.find(f'{{{OMML_NS}}}sty') if rPr is not None else None
        val = sty.attrib.get(f'{{{OMML_NS}}}val') if sty is not None else ''

        # 普通文本 (m:nor)
        if is_normal_text(elem):
            return f"\\text{{{escape_text_mode(t_text)}}}"

        # 特殊字体 (双线体 ℝ、花体、哥特体等)
        scr = rPr.find(mtag('scr')) if rPr is not None else None
        scr_val = scr.attrib.get(mtag('val'), 'roman') if scr is not None else 'roman'
        if scr_val in SCR_CMDS and t_text.isascii() and t_text.isalnum():
            return f"{SCR_CMDS[scr_val]}{{{t_text}}}"

        if val in ('b', 'bi') and t_text.isascii() and t_text.strip().isalnum():
            cmd = '\\mathbf' if val == 'b' else '\\boldsymbol'
            return f"{cmd}{{{t_text}}}"

        # Plain upright text or function
        if val == 'p':
            if t_text in MATH_FUNCS:
                if t_text in ('tr', 'sgn', 'arcsec', 'arccsc', 'arccot'):
                    return f"\\operatorname{{{t_text}}} "
                return f"\\{t_text} "
            elif t_text in ('d', 'e', 'T', 'r', 'o', 'E', 'O'):
                return f"\\mathrm{{{t_text}}}"
            elif t_text in ('=', '+', '-', '−', '<', '>', '≤', '≥', '≠', '±', '×', '·', '→', '∈'):
                return f" {escape_latex_text(t_text)} "
            elif t_text.isdigit():
                return t_text
            elif t_text in ('(', ')', '[', ']', '{', '}', ',', '.', ';', ':', '|', '⋯', '∞', '′', '″'):
                return escape_latex_text(t_text)
            elif t_text.isascii() and t_text.isalpha():
                return f"\\mathrm{{{t_text}}}"
            elif not any('\u4e00' <= ch <= '\u9fff' for ch in t_text):
                # 正体的符号/数字混合，如 "2π"、"≤"
                return escape_latex_text(t_text)
            return f"\\text{{{escape_text_mode(t_text)}}}"

        # If contains Chinese characters, wrap in \text{} for KaTeX stability
        if any('\u4e00' <= char <= '\u9fff' for char in t_text):
            return f"\\text{{{escape_text_mode(t_text)}}}"
            
        return escape_latex_text(t_text)
        
    elif tag == 'sSup':
        e = elem.find(f'{{{OMML_NS}}}e')
        sup = elem.find(f'{{{OMML_NS}}}sup')
        e_str = omml_to_latex(e) if e is not None else ''
        sup_str = omml_to_latex(sup) if sup is not None else ''
        return f"{{{e_str}}}^{{{sup_str}}}"
        
    elif tag == 'sSub':
        e = elem.find(f'{{{OMML_NS}}}e')
        sub = elem.find(f'{{{OMML_NS}}}sub')
        e_str = omml_to_latex(e) if e is not None else ''
        sub_str = omml_to_latex(sub) if sub is not None else ''
        return f"{{{e_str}}}_{{{sub_str}}}"
        
    elif tag == 'sSubSup':
        e = elem.find(f'{{{OMML_NS}}}e')
        sub = elem.find(f'{{{OMML_NS}}}sub')
        sup = elem.find(f'{{{OMML_NS}}}sup')
        e_str = omml_to_latex(e) if e is not None else ''
        sub_str = omml_to_latex(sub) if sub is not None else ''
        sup_str = omml_to_latex(sup) if sup is not None else ''
        return f"{{{e_str}}}_{{{sub_str}}}^{{{sup_str}}}"
        
    elif tag == 'sPre':
        sub = elem.find(f'{{{OMML_NS}}}sub')
        sup = elem.find(f'{{{OMML_NS}}}sup')
        e = elem.find(f'{{{OMML_NS}}}e')
        sub_str = omml_to_latex(sub) if sub is not None else ''
        sup_str = omml_to_latex(sup) if sup is not None else ''
        e_str = omml_to_latex(e) if e is not None else ''
        return f"{{}}_{{{sub_str}}}^{{{sup_str}}}{{{e_str}}}"
        
    elif tag == 'f':
        fPr = elem.find(f'{{{OMML_NS}}}fPr')
        fType = fPr.find(f'{{{OMML_NS}}}type') if fPr is not None else None
        fType_val = fType.attrib.get(f'{{{OMML_NS}}}val', 'bar') if fType is not None else 'bar'
        num = elem.find(f'{{{OMML_NS}}}num')
        den = elem.find(f'{{{OMML_NS}}}den')
        num_str = omml_to_latex(num) if num is not None else ''
        den_str = omml_to_latex(den) if den is not None else ''
        if fType_val == 'noBar':
            # 无分数线的上下堆叠 (组合数外层括号由 m:d 单独提供)
            return f"\\genfrac{{}}{{}}{{0pt}}{{}}{{{num_str}}}{{{den_str}}}"
        elif fType_val in ('skw', 'lin'):
            return f"{{{num_str}}}/{{{den_str}}}"
        return f"\\frac{{{num_str}}}{{{den_str}}}"
        
    elif tag == 'rad':
        deg = elem.find(f'{{{OMML_NS}}}deg')
        e = elem.find(f'{{{OMML_NS}}}e')
        deg_str = omml_to_latex(deg) if deg is not None else ''
        e_str = omml_to_latex(e) if e is not None else ''
        if deg_str.strip():
            return f"\\sqrt[{deg_str}]{{{e_str}}}"
        return f"\\sqrt{{{e_str}}}"
        
    elif tag == 'd':
        dPr = elem.find(f'{{{OMML_NS}}}dPr')
        beg = '('
        end = ')'
        sep = None
        if dPr is not None:
            beg_elem = dPr.find(f'{{{OMML_NS}}}begChr')
            end_elem = dPr.find(f'{{{OMML_NS}}}endChr')
            sep_elem = dPr.find(f'{{{OMML_NS}}}sepChr')
            if beg_elem is not None:
                beg = beg_elem.attrib.get(f'{{{OMML_NS}}}val', '')
            if end_elem is not None:
                end = end_elem.attrib.get(f'{{{OMML_NS}}}val', '')
            if sep_elem is not None:
                sep = sep_elem.attrib.get(f'{{{OMML_NS}}}val', None)
                
        e_list = [omml_to_latex(e) for e in elem.findall(f'{{{OMML_NS}}}e')]
        
        def map_delim(c):
            if not c: return '.'
            if c in ('{', '｛'): return '\\{'
            if c in ('}', '｝'): return '\\}'
            if c in ('(', '（'): return '('
            if c in (')', '）'): return ')'
            if c in ('[', '【'): return '['
            if c in (']', '】'): return ']'
            if c in ('|', '∣'): return '|'
            if c in ('‖', '∥'): return '\\|'
            if c in ('⟨', '<', '〈'): return '\\langle '
            if c in ('⟩', '>', '〉'): return '\\rangle '
            if c == '⌊': return '\\lfloor '
            if c == '⌋': return '\\rfloor '
            if c == '⌈': return '\\lceil '
            if c == '⌉': return '\\rceil '
            if c == '⟦': return '\\llbracket '
            if c == '⟧': return '\\rrbracket '
            if c == '/': return '/'
            if c == '\\': return '\\backslash '
            return '.'

        beg_latex = map_delim(beg)
        end_latex = map_delim(end)

        # ECMA-376：省略 sepChr 时默认分隔符为 "|"
        sep_raw = sep if sep is not None else '|'
        if len(e_list) <= 1 or sep_raw == '':
            inner = ' '.join(e_list)
        else:
            sep_tex = '\\mid ' if sep_raw in ('|', '∣') else escape_latex_text(sep_raw)
            inner = f' {sep_tex} '.join(e_list)
            
        return f"\\left{beg_latex} {inner} \\right{end_latex}"
        
    elif tag == 'm':
        rows = []
        for mr in elem.findall(f'{{{OMML_NS}}}mr'):
            cells = [omml_to_latex(e).strip() for e in mr.findall(f'{{{OMML_NS}}}e')]
            rows.append(' & '.join(cells))
        return "\\begin{matrix} " + ' \\\\ '.join(rows) + " \\end{matrix}"
        
    elif tag == 'eqArr':
        # eqArr 本身没有大括号 (分段函数的大括号来自外层 m:d)，用左对齐 array 表示
        lines = [omml_to_latex(e).strip() for e in elem.findall(f'{{{OMML_NS}}}e')]
        return "\\begin{array}{l} " + ' \\\\ '.join(lines) + " \\end{array}"
        
    elif tag == 'nary':
        naryPr = elem.find(f'{{{OMML_NS}}}naryPr')
        chr_val = '∫'
        if naryPr is not None:
            chr_elem = naryPr.find(f'{{{OMML_NS}}}chr')
            if chr_elem is not None:
                chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '∫')
        sub = elem.find(f'{{{OMML_NS}}}sub')
        sup = elem.find(f'{{{OMML_NS}}}sup')
        e = elem.find(f'{{{OMML_NS}}}e')
        
        NARY_CMDS = {
            '∑': '\\sum', 'sum': '\\sum', '∏': '\\prod', 'prod': '\\prod', '∐': '\\coprod',
            '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint', '∯': '\\oiint', '∰': '\\oiiint',
            '⋃': '\\bigcup', '⋂': '\\bigcap', '⋁': '\\bigvee', '⋀': '\\bigwedge',
            '⨁': '\\bigoplus', '⨂': '\\bigotimes', '⨀': '\\bigodot', '⨄': '\\biguplus', '⨆': '\\bigsqcup',
        }
        cmd = NARY_CMDS.get(chr_val, '\\int' if not chr_val else f'\\mathop{{{escape_latex_text(chr_val)}}}')

        def hidden(name):
            node = naryPr.find(mtag(name)) if naryPr is not None else None
            return node is not None and node.attrib.get(mtag('val'), 'on') in ('on', '1', 'true')

        lim_loc = naryPr.find(mtag('limLoc')) if naryPr is not None else None
        if lim_loc is not None and lim_loc.attrib.get(mtag('val')) == 'undOvr':
            cmd += '\\limits'
        sub_tex = omml_to_latex(sub).strip() if sub is not None and not hidden('subHide') else ''
        sup_tex = omml_to_latex(sup).strip() if sup is not None and not hidden('supHide') else ''
        sub_str = f"_{{{sub_tex}}}" if sub_tex else ""
        sup_str = f"^{{{sup_tex}}}" if sup_tex else ""
        e_str = omml_to_latex(e) if e is not None else ""
        return f"{cmd}{sub_str}{sup_str} {e_str}"
        
    elif tag == 'limLow':
        e = elem.find(f'{{{OMML_NS}}}e')
        lim = elem.find(f'{{{OMML_NS}}}lim')
        e_str = omml_to_latex(e).strip() if e is not None else ''
        lim_str = omml_to_latex(lim).strip() if lim is not None else ''
        if e_str in ('lim', '\\lim', '\\text{lim}', '\\mathrm{lim}'):
            return f"\\lim_{{{lim_str}}} "
        elif e_str in ('∬', '\\iint', '\\text{∬}'):
            return f"\\iint_{{{lim_str}}} "
        elif e_str in ('∭', '\\iiint', '\\text{∭}'):
            return f"\\iiint_{{{lim_str}}} "
        elif e_str in ('∑', '\\sum', '\\text{∑}'):
            return f"\\sum_{{{lim_str}}} "
        elif e_str in ('∏', '\\prod', '\\text{∏}'):
            return f"\\prod_{{{lim_str}}} "
        base_plain = ''.join(e.itertext()).strip() if e is not None else ''
        if base_plain in ('max', 'min', 'sup', 'inf', 'limsup', 'liminf', 'lim sup', 'lim inf'):
            return f"\\operatorname*{{{base_plain}}}_{{{lim_str}}} "
        if base_plain in BIGOP_CHARS or base_plain in INTEGRAL_CHARS:
            return f"\\mathop{{{e_str}}}\\limits_{{{lim_str}}} "
        return f"\\underset{{{lim_str}}}{{{e_str}}} "
        
    elif tag == 'limUpp':
        e = elem.find(f'{{{OMML_NS}}}e')
        lim = elem.find(f'{{{OMML_NS}}}lim')
        sup_str = omml_to_latex(lim).strip() if lim is not None else ''
        
        if e is not None:
            inner_limLow = e.find(f'{{{OMML_NS}}}limLow')
            if inner_limLow is not None:
                inner_e = inner_limLow.find(f'{{{OMML_NS}}}e')
                inner_lim = inner_limLow.find(f'{{{OMML_NS}}}lim')
                sub_str = omml_to_latex(inner_lim).strip() if inner_lim is not None else ''
                base_text = ''.join(inner_e.itertext()).strip() if inner_e is not None else ''
                if base_text in ('∑', 'sum', '\\sum'):
                    return f"\\sum_{{{sub_str}}}^{{{sup_str}}} "
                elif base_text in ('∏', 'prod', '\\prod'):
                    return f"\\prod_{{{sub_str}}}^{{{sup_str}}} "
                    
        e_str = omml_to_latex(e).strip() if e is not None else ''
        if e_str in ('lim', '\\lim', '\\text{lim}'):
            return f"\\lim^{{{sup_str}}} "
        base_plain = ''.join(e.itertext()).strip() if e is not None else ''
        if base_plain in BIGOP_CHARS or base_plain in INTEGRAL_CHARS:
            return f"\\mathop{{{e_str}}}\\limits^{{{sup_str}}} "
        return f"\\overset{{{sup_str}}}{{{e_str}}} "
        
    elif tag == 'bar':
        e = elem.find(f'{{{OMML_NS}}}e')
        barPr = elem.find(f'{{{OMML_NS}}}barPr')
        pos = 'bot'  # ECMA-376：m:bar 省略 pos 时横线在下方
        if barPr is not None:
            pos_elem = barPr.find(f'{{{OMML_NS}}}pos')
            if pos_elem is not None:
                pos = pos_elem.attrib.get(f'{{{OMML_NS}}}val', 'bot')
        e_str = omml_to_latex(e) if e is not None else ''
        if pos == 'bot':
            return f"\\underline{{{e_str}}}"
        return f"\\overline{{{e_str}}}"
        
    elif tag == 'acc':
        e = elem.find(f'{{{OMML_NS}}}e')
        accPr = elem.find(f'{{{OMML_NS}}}accPr')
        chr_val = '^'
        if accPr is not None:
            chr_elem = accPr.find(f'{{{OMML_NS}}}chr')
            if chr_elem is not None:
                chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '^')
        e_str = omml_to_latex(e) if e is not None else ''
        ACC_MAP = {
            '̇': '\\dot', '̈': '\\ddot', '̂': '\\hat', '̃': '\\tilde',
            '̄': '\\bar', '⃗': '\\vec', '^': '\\hat', 'ˇ': '\\check',
            '́': '\\acute', '̀': '\\grave', '̆': '\\breve', '̌': '\\check', '̊': '\\mathring',
            '⃖': '\\overleftarrow', '⃡': '\\overleftrightarrow', '⃑': '\\overrightharpoon',
            '⃐': '\\overleftharpoon', '̅': '\\overline', '̿': '\\overline', '⃛': '\\dddot',
            '→': '\\vec', '¯': '\\bar', '~': '\\tilde', '˙': '\\dot', '¨': '\\ddot', '˘': '\\breve',
        }
        cmd = ACC_MAP.get(chr_val, '\\hat')
        return f"{cmd}{{{e_str}}}"
        
    elif tag == 'groupChr':
        e = elem.find(f'{{{OMML_NS}}}e')
        chr_pr = elem.find(f'{{{OMML_NS}}}groupChrPr')
        pos = 'bot'
        chr_val = '⏟'
        if chr_pr is not None:
            pos_elem = chr_pr.find(f'{{{OMML_NS}}}pos')
            if pos_elem is not None:
                pos = pos_elem.attrib.get(f'{{{OMML_NS}}}val', 'bot')
            chr_elem = chr_pr.find(f'{{{OMML_NS}}}chr')
            if chr_elem is not None:
                chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '⏟')
        e_str = omml_to_latex(e) if e is not None else ''
        if chr_val == '⏞':
            return f"\\overbrace{{{e_str}}}"
        if chr_val == '⏟':
            return f"\\underbrace{{{e_str}}}"
        GROUP_ARROWS = {'→': 'rightarrow', '←': 'leftarrow', '↔': 'leftrightarrow',
                        '⇒': 'Rightarrow', '⇐': 'Leftarrow', '⇔': 'Leftrightarrow'}
        if chr_val in GROUP_ARROWS:
            return f"\\{'over' if pos == 'top' else 'under'}{GROUP_ARROWS[chr_val]}{{{e_str}}}"
        sym = escape_latex_text(chr_val) or '\\_'
        if pos == 'top':
            return f"\\overset{{{sym}}}{{{e_str}}}"
        return f"\\underset{{{sym}}}{{{e_str}}}"
        
    elif tag == 'borderBox':
        e = elem.find(f'{{{OMML_NS}}}e')
        e_str = omml_to_latex(e) if e is not None else ''
        return f"\\boxed{{{e_str}}}"
        
    elif tag == 'func':
        fName = elem.find(f'{{{OMML_NS}}}fName')
        e = elem.find(f'{{{OMML_NS}}}e')
        fn_str = omml_to_latex(fName).strip() if fName is not None else ''
        e_str = omml_to_latex(e).strip() if e is not None else ''
        return f"{fn_str} {e_str}"
        
    elif tag == 'box':
        e = elem.find(f'{{{OMML_NS}}}e')
        return omml_to_latex(e) if e is not None else ''

    elif tag == 'phant':
        e = elem.find(mtag('e'))
        e_str = omml_to_latex(e) if e is not None else ''
        show = elem.find(f'{mtag("phantPr")}/{mtag("show")}')
        if show is not None and show.attrib.get(mtag('val'), 'on') in ('off', '0', 'false'):
            return f"\\phantom{{{e_str}}}"
        return e_str

    elif tag == 'oMathPara':
        parts = [omml_to_latex(om).strip() for om in elem.findall(mtag('oMath'))]
        return ' \\\\ '.join(p for p in parts if p)

    elif tag in ('del', 'moveFrom') or tag.endswith('Pr'):
        # 已删除的修订和属性节点不参与显示
        return ''

    else:
        return ''.join(omml_to_latex(c) for c in elem)


class DocxMathRepairEngine:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.xml_content = None
        self.root = None
        self.load()

    @staticmethod
    def parse_xml_keep_ns(content):
        """解析 XML 并记录根节点声明的全部命名空间前缀。
        ElementTree 默认会把 w14/wp14 等改名为 ns1 并丢弃未使用的声明，
        而 mc:Ignorable="w14 wp14" 引用的前缀一旦缺失，新版 Word 会提示文件损坏。"""
        root_ns = []
        depth = 0
        for event, item in ET.iterparse(io.BytesIO(content), events=('start-ns', 'start')):
            if event == 'start-ns':
                prefix, uri = item
                if depth == 0:
                    root_ns.append((prefix, uri))
                if prefix and not _re.match(r'ns\d+$', prefix):
                    try:
                        ET.register_namespace(prefix, uri)
                    except ValueError:
                        pass
            else:
                depth += 1
                if depth > 1:
                    break
        return ET.fromstring(content), root_ns

    @staticmethod
    def serialize_xml_keep_ns(root, root_ns):
        xml = ET.tostring(root, encoding='unicode')
        end = xml.index('>')
        head = xml[:end]
        wanted = [(p, u) for p, u in root_ns if p and not _re.match(r'ns\d+$', p)]
        # mc:Ignorable / mc:ProcessContent 中引用的前缀必须在根节点声明
        ign = _re.search(r'Ignorable="([^"]*)"', head)
        if ign:
            wanted += [(p, KNOWN_NS[p]) for p in ign.group(1).split() if p in KNOWN_NS]
        extra = ''
        for p, u in wanted:
            if f'xmlns:{p}=' not in head and f'xmlns:{p}=' not in extra:
                extra += f' xmlns:{p}="{u}"'
        xml = head + extra + xml[end:]
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml).encode('utf-8')

    def load(self):
        self.parts = {}
        with zipfile.ZipFile(self.docx_path, 'r') as z:
            names = z.namelist()
            ordered = ['word/document.xml'] + sorted(n for n in names if MATH_PART_RE.match(n) and n != 'word/document.xml')
            for name in ordered:
                if name in names:
                    self.parts[name] = self.parse_xml_keep_ns(z.read(name))
            # 样式表只读，用于判断段落的有效行距
            self.styles_root = ET.fromstring(z.read('word/styles.xml')) if 'word/styles.xml' in names else None
        self.root = self.parts['word/document.xml'][0]

    def save(self, output_path, unify_fonts=True):
        """Save the updated document back into a new docx archive with safe XML tree writing."""
        parts = getattr(self, 'parts', None) or {'word/document.xml': (self.root, [])}
        new_parts = {name: self.serialize_xml_keep_ns(root, ns) for name, (root, ns) in parts.items()}
        tmp_output = output_path + '.tmp_writing.docx'

        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp_output, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    if item.filename in new_parts:
                        zout.writestr(item, new_parts[item.filename])
                    elif not unify_fonts:
                        zout.writestr(item, content)
                    elif item.filename == 'word/theme/theme1.xml':
                        try:
                            theme_root, theme_ns = self.parse_xml_keep_ns(content)
                            for latin in theme_root.iter(f'{{{DRAWING_A_NS}}}latin'):
                                latin.set('typeface', 'Times New Roman')
                            content = self.serialize_xml_keep_ns(theme_root, theme_ns)
                        except Exception:
                            pass
                        zout.writestr(item, content)
                    elif item.filename in ('word/styles.xml', 'word/stylesWithEffects.xml'):
                        try:
                            styles_root, styles_ns = self.parse_xml_keep_ns(content)
                            for rFonts in styles_root.iter(f'{{{WORD_NS}}}rFonts'):
                                rFonts.set(f'{{{WORD_NS}}}ascii', 'Times New Roman')
                                rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Times New Roman')
                                rFonts.set(f'{{{WORD_NS}}}cs', 'Times New Roman')
                            content = self.serialize_xml_keep_ns(styles_root, styles_ns)
                        except Exception:
                            pass
                        zout.writestr(item, content)
                    elif item.filename == 'word/settings.xml':
                        try:
                            settings_root, settings_ns = self.parse_xml_keep_ns(content)
                            mathPr = settings_root.find(f'{{{OMML_NS}}}mathPr')
                            if mathPr is None:
                                # CT_Settings 中 m:mathPr 必须位于这些元素之前
                                after = ('attachedSchema', 'themeFontLang', 'clrSchemeMapping',
                                         'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures',
                                         'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType',
                                         'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags',
                                         'decimalSymbol', 'listSeparator')
                                pos = len(settings_root)
                                for i, c in enumerate(list(settings_root)):
                                    if local(c.tag) in after or not c.tag.startswith(f'{{{WORD_NS}}}'):
                                        pos = i
                                        break
                                mathPr = ET.Element(f'{{{OMML_NS}}}mathPr')
                                settings_root.insert(pos, mathPr)
                            mathFont = mathPr.find(f'{{{OMML_NS}}}mathFont')
                            if mathFont is None:
                                mathFont = ET.Element(f'{{{OMML_NS}}}mathFont')
                                mathPr.insert(0, mathFont)  # mathFont 是 mathPr 的第一个子元素
                            mathFont.set(f'{{{OMML_NS}}}val', 'Times New Roman')
                            content = self.serialize_xml_keep_ns(settings_root, settings_ns)
                        except Exception:
                            pass
                        zout.writestr(item, content)
                    else:
                        zout.writestr(item, content)

        # Atomic and safe replacement
        try:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except PermissionError:
                    # File is currently opened by Word or another program
                    pass
            os.replace(tmp_output, output_path)
        except Exception:
            # If replacing failed because target is locked, try saving with unique counter
            counter = 1
            base, ext = os.path.splitext(output_path)
            alt_path = f"{base}_{counter}{ext}"
            while os.path.exists(alt_path):
                counter += 1
                alt_path = f"{base}_{counter}{ext}"
            os.replace(tmp_output, alt_path)
            return alt_path
        return output_path

    # ================== Repair Rules ==================

    def apply_rule_split_tokens(self, omath):
        """拆分连写函数 (如 lncos -> ln + cos)"""
        modified = False
        # 任意函数名的连写组合：lncosx → ln|cos|x，sin2x → sin|2x，arctant → arctan|t
        func_alt = '|'.join(sorted(MATH_FUNCS, key=lambda x: -len(x)))
        glued_re = _re.compile(rf'^((?:{func_alt})+)(.*)$', _re.S)
        func_re = _re.compile(func_alt)

        for r in list(omath.iter(mtag('r'))):
            if is_normal_text(r) or len(r.findall(mtag('t'))) != 1:
                continue
            text = run_text(r)
            m = glued_re.match(text)
            if not m or text in MATH_FUNCS:
                continue
            funcs = func_re.findall(m.group(1))
            rest = m.group(2)
            if len(funcs) < 2 and not rest:
                continue
            # 剩余部分须以变量/数字/括号开头，避免把普通单词误拆
            if rest and not _re.match(r'[A-Za-z0-9(（\[{α-ωΑ-Ω]', rest):
                continue
            parent = find_parent(omath, r)
            if parent is None:
                continue
            idx = list(parent).index(r)
            parent.remove(r)
            pieces = [(f, True) for f in funcs] + ([(rest, None)] if rest else [])
            for piece, upright in pieces:
                new_r = clone_run(r, piece, upright=upright if upright else None)
                if upright is None and get_sty(new_r) == 'p':
                    # 原 run 整体是正体时，剩余变量部分恢复斜体
                    new_r.find(mtag('rPr')).remove(new_r.find(mtag('rPr')).find(mtag('sty')))
                parent.insert(idx, new_r)
                idx += 1
            modified = True
        return modified

    def apply_rule_isolated_delimiters(self, omath):
        """修复括号上标结构 (如 (式子)^n)"""
        modified = False
        DELIM_PAIRS = {
            ')': '(',
            ']': '[',
            '}': '{'
        }
        DELIM_PAIRS.update({'）': '（', '】': '【'})
        scripts = [n for n in omath.iter() if local(n.tag) in ('sSup', 'sSub', 'sSubSup')]
        for sSup in scripts:
            e = sSup.find(f'{{{OMML_NS}}}e')
            if e is None:
                continue
            # 底座必须只由普通 run 组成，且以闭括号结尾，如 ")" 或 "b+c)"
            e_runs = [c for c in e if local(c.tag) not in ('ctrlPr', 'argPr')]
            if not e_runs or any(local(c.tag) != 'r' for c in e_runs):
                continue
            e_text = ''.join(run_text(c) for c in e_runs).rstrip()
            if not e_text or e_text[-1] not in DELIM_PAIRS:
                continue
            close_char = e_text[-1]
            beg_char = DELIM_PAIRS[close_char]
            base_tail = e_text[:-1]
            # 底座内部若自身括号已配对，例如 "(x)"，说明不是被拆散的结构
            if base_tail.count(beg_char) >= base_tail.count(close_char) + 1:
                continue
            e_text = close_char

            parent = find_parent(omath, sSup)
            if parent is None:
                continue
                
            siblings = list(parent)
            sSup_pos = siblings.index(sSup)
            
            # 使用括号平衡深度计数器（Nesting Depth Balance）从右向左寻找与之配对的开括号
            match_idx = -1
            match_pos = -1
            # 闭括号嵌套计数 (底座内多余的闭括号也要计入)
            depth = max(0, base_tail.count(close_char) - base_tail.count(beg_char))
            found = False
            
            for i in range(sSup_pos - 1, -1, -1):
                t = ''.join(siblings[i].itertext())
                # 从右往左逐字符扫描
                for c_pos in range(len(t) - 1, -1, -1):
                    ch = t[c_pos]
                    if ch == close_char:
                        depth += 1
                    elif ch == beg_char:
                        if depth == 0:
                            # 找到了真正与当前闭括号匹配的最外层配对开括号
                            match_idx = i
                            match_pos = c_pos
                            found = True
                            break
                        else:
                            depth -= 1
                if found:
                    break
                    
            if not found or match_idx == -1:
                continue

            match_elem = siblings[match_idx]
            # 开括号必须位于普通 run 中；若在嵌套结构 (分式、根式) 内部则无法安全移动，跳过
            if local(match_elem.tag) != 'r' or len(match_elem.findall(mtag('t'))) != 1:
                continue
            inner_items = []

            t_elem = match_elem.find(f'{{{OMML_NS}}}t')
            prefix_exists = False
            if t_elem is not None and t_elem.text:
                orig_text = t_elem.text
                prefix_text = orig_text[:match_pos]
                suffix_text = orig_text[match_pos + 1:]
                
                if prefix_text:
                    set_t_text(t_elem, prefix_text)
                    prefix_exists = True
                else:
                    parent.remove(match_elem)

                if suffix_text:
                    inner_items.append(clone_run(match_elem, suffix_text))
            else:
                parent.remove(match_elem)
                
            current_siblings = list(parent)
            current_sSup_pos = current_siblings.index(sSup)
            start_idx = (match_idx + 1) if prefix_exists else match_idx
            middle_items = current_siblings[start_idx:current_sSup_pos]
            
            for item in middle_items:
                parent.remove(item)
                inner_items.append(item)

            # 底座中闭括号前的内容 (如 "b+c)" 中的 "b+c") 也属于括号内部
            if base_tail:
                last = [run for run in e_runs if run_text(run).strip()][-1]
                last_t = last.find(mtag('t'))
                set_t_text(last_t, run_text(last).rstrip()[:-1])
                for run in e_runs:
                    if run_text(run):
                        inner_items.append(run)

            d_elem = ET.Element(f'{{{OMML_NS}}}d')
            dPr = ET.SubElement(d_elem, f'{{{OMML_NS}}}dPr')
            beg_node = ET.SubElement(dPr, f'{{{OMML_NS}}}begChr')
            beg_node.set(f'{{{OMML_NS}}}val', beg_char)
            end_node = ET.SubElement(dPr, f'{{{OMML_NS}}}endChr')
            end_node.set(f'{{{OMML_NS}}}val', e_text)
            grow_node = ET.SubElement(dPr, f'{{{OMML_NS}}}grow')
            grow_node.set(f'{{{OMML_NS}}}val', 'on')
            
            d_e = ET.SubElement(d_elem, f'{{{OMML_NS}}}e')
            for item in inner_items:
                d_e.append(item)
                
            e.clear()
            e.append(d_elem)
            modified = True
            
        return modified

    def apply_rule_nary_empty_operand(self, omath):
        """消除积分虚线框 (补全被积项)"""
        modified = False
        for nary in list(omath.iter(f'{{{OMML_NS}}}nary')):
            chr_elem = nary.find(f'{mtag("naryPr")}/{mtag("chr")}')
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '∫') if chr_elem is not None else '∫'
            if chr_val not in INTEGRAL_CHARS:
                continue

            e = nary.find(f'{{{OMML_NS}}}e')
            if e is None:
                e = ET.SubElement(nary, f'{{{OMML_NS}}}e')
                
            if is_empty_arg(e):
                parent = find_parent(omath, nary)
                if parent is None:
                    continue
                    
                siblings = list(parent)
                pos = siblings.index(nary)
                following_siblings = siblings[pos + 1:]
                
                if not following_siblings:
                    continue
                    
                # 被积项吸收到：微分 dx/dxdy 结束为止，或遇到括号外的关系符/逗号为止
                state = {'depth': 0, 'pending_d': False}
                operand_items = []
                for sib in following_siblings:
                    if local(sib.tag) in ('ctrlPr', 'argPr'):
                        continue
                    if local(sib.tag) == 'r' and len(sib.findall(mtag('t'))) == 1:
                        text = run_text(sib)
                        cut = self._integrand_stop(text, state)
                        if cut is not None:
                            if cut > 0:
                                operand_items.append(clone_run(sib, text[:cut]))
                            if text[cut:]:
                                set_t_text(sib.find(mtag('t')), text[cut:])
                            else:
                                parent.remove(sib)
                            break
                    else:
                        state['pending_d'] = False
                    parent.remove(sib)
                    operand_items.append(sib)
                    
                if operand_items:
                    # m:e 的 ctrlPr 必须是最后一个子元素
                    ctrl = e.find(mtag('ctrlPr'))
                    pos_in_e = list(e).index(ctrl) if ctrl is not None else len(e)
                    for item in operand_items:
                        e.insert(pos_in_e, item)
                        pos_in_e += 1
                    modified = True

        return modified

    @staticmethod
    def _integrand_stop(text, state):
        """扫描被积项文本，返回截断位置 (该位置之前属于被积项)；None 表示整段都属于被积项。"""
        LETTER = _re.compile(r'[A-Za-zα-ωθφρσ]')
        RELATIONS = '=<>≤≥≠≈⩽⩾,，;；'
        if state['pending_d']:
            state['pending_d'] = False
            if text and LETTER.match(text[0]):
                i = 1
                while text[i:i + 1] == 'd' and LETTER.match(text[i + 1:i + 2] or ' '):
                    i += 2
                return i
        i = 0
        while i < len(text):
            ch = text[i]
            if ch in '([{（':
                state['depth'] += 1
            elif ch in ')]}）':
                state['depth'] = max(0, state['depth'] - 1)
            elif state['depth'] == 0 and ch in RELATIONS:
                return i
            elif state['depth'] == 0 and ch == 'd' and not text.startswith('det', i):
                nxt = text[i + 1:i + 2]
                if nxt and LETTER.match(nxt):
                    j = i + 2
                    while text[j:j + 1] == 'd' and LETTER.match(text[j + 1:j + 2] or ' '):
                        j += 2
                    return j
                if not nxt:
                    state['pending_d'] = True
            i += 1
        return None

    def apply_rule_integral_size(self, omath):
        """固定积分号高度 (防止拉伸)"""
        modified = False
        for nary in omath.iter(f'{{{OMML_NS}}}nary'):
            naryPr = nary.find(f'{{{OMML_NS}}}naryPr')
            chr_elem = naryPr.find(f'{{{OMML_NS}}}chr') if naryPr is not None else None
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '∫') if chr_elem is not None else '∫'
            
            if chr_val in INTEGRAL_CHARS:
                if naryPr is None:
                    naryPr = ET.Element(mtag('naryPr'))
                    nary.insert(0, naryPr)
                grow = naryPr.find(mtag('grow'))
                if grow is None:
                    continue  # 省略 grow 时积分号本来就不随内容拉伸
                cur = grow.attrib.get(mtag('val'), 'on')
                if cur in ('off', '0', 'false'):
                    continue  # 已经固定，不重复修改
                # OMML 的 ST_OnOff 只允许 on/off，并按 schema 顺序插入
                if set_ordered_prop(naryPr, 'grow', 'off', NARY_PR_ORDER):
                    modified = True
        return modified

    def apply_rule_fix_limits(self, omath):
        """极限下标居中 (lim)"""
        modified = False
        for sSub in list(omath.iter(f'{{{OMML_NS}}}sSub')):
            e = sSub.find(f'{{{OMML_NS}}}e')
            sub = sSub.find(f'{{{OMML_NS}}}sub')
            if e is not None and sub is not None:
                e_text = ''.join(e.itertext()).strip()
                # 只处理底座恰好是极限类算子的情况；旧逻辑 'lim' in e_text 会把 limsup 等改成 lim
                LIMIT_OPS = ('lim', 'limsup', 'liminf', 'lim sup', 'lim inf', 'max', 'min', 'sup', 'inf')
                if e_text in LIMIT_OPS and all(local(c.tag) in ('r', 'ctrlPr') for c in e):
                    parent = find_parent(omath, sSub)
                    if parent is not None:
                        idx = list(parent).index(sSub)
                        parent.remove(sSub)
                        
                        limLow = ET.Element(f'{{{OMML_NS}}}limLow')
                        new_e = ET.SubElement(limLow, f'{{{OMML_NS}}}e')
                        new_r = ET.SubElement(new_e, f'{{{OMML_NS}}}r')
                        rPr = ET.SubElement(new_r, f'{{{OMML_NS}}}rPr')
                        sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                        sty.set(f'{{{OMML_NS}}}val', 'p')
                        src_r = e.find(mtag('r'))
                        if src_r is not None and src_r.find(f'{{{WORD_NS}}}rPr') is not None:
                            new_r.insert(1, copy.deepcopy(src_r.find(f'{{{WORD_NS}}}rPr')))
                        new_t = ET.SubElement(new_r, f'{{{OMML_NS}}}t')
                        set_t_text(new_t, e_text)
                        
                        new_lim = ET.SubElement(limLow, f'{{{OMML_NS}}}lim')
                        for child in list(sub):
                            new_lim.append(child)
                        
                        parent.insert(idx, limLow)
                        modified = True
        return modified

    def apply_rule_nary_lim_loc(self, omath):
        """求和与乘积上下标居中 (∑/∏)"""
        modified = False
        for nary in list(omath.iter(f'{{{OMML_NS}}}nary')):
            chr_elem = nary.find(f'{mtag("naryPr")}/{mtag("chr")}')
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '') if chr_elem is not None else ''

            if chr_val in BIGOP_CHARS:
                parent = find_parent(omath, nary)
                if parent is None:
                    continue
                    
                idx = list(parent).index(nary)
                sub = nary.find(f'{{{OMML_NS}}}sub')
                sup = nary.find(f'{{{OMML_NS}}}sup')
                e = nary.find(f'{{{OMML_NS}}}e')
                
                has_sub = sub is not None and bool(''.join(sub.itertext()).strip())
                has_sup = sup is not None and bool(''.join(sup.itertext()).strip())
                
                if not has_sub and not has_sup:
                    continue
                    
                sym_char = {'sum': '∑', 'prod': '∏'}.get(chr_val, chr_val)

                def create_sym_r():
                    # m:r 子元素顺序必须是 m:rPr → w:rPr → m:t
                    r_sym = ET.Element(f'{{{OMML_NS}}}r')
                    rPr_m = ET.SubElement(r_sym, f'{{{OMML_NS}}}rPr')
                    sty = ET.SubElement(rPr_m, f'{{{OMML_NS}}}sty')
                    sty.set(f'{{{OMML_NS}}}val', 'p')
                    rPr_w = ET.SubElement(r_sym, f'{{{WORD_NS}}}rPr')
                    sz = ET.SubElement(rPr_w, f'{{{WORD_NS}}}sz')
                    sz.set(f'{{{WORD_NS}}}val', '28')
                    szCs = ET.SubElement(rPr_w, f'{{{WORD_NS}}}szCs')
                    szCs.set(f'{{{WORD_NS}}}val', '28')
                    t_sym = ET.SubElement(r_sym, f'{{{OMML_NS}}}t')
                    t_sym.text = sym_char
                    return r_sym
                
                top_elem = None
                if has_sub and has_sup:
                    limUpp = ET.Element(f'{{{OMML_NS}}}limUpp')
                    e_outer = ET.SubElement(limUpp, f'{{{OMML_NS}}}e')
                    limLow = ET.SubElement(e_outer, f'{{{OMML_NS}}}limLow')
                    e_inner = ET.SubElement(limLow, f'{{{OMML_NS}}}e')
                    e_inner.append(create_sym_r())
                    
                    lim_sub = ET.SubElement(limLow, f'{{{OMML_NS}}}lim')
                    for child in list(sub):
                        lim_sub.append(child)
                        
                    lim_sup = ET.SubElement(limUpp, f'{{{OMML_NS}}}lim')
                    for child in list(sup):
                        lim_sup.append(child)
                    top_elem = limUpp
                elif has_sub:
                    limLow = ET.Element(f'{{{OMML_NS}}}limLow')
                    e_inner = ET.SubElement(limLow, f'{{{OMML_NS}}}e')
                    e_inner.append(create_sym_r())
                    
                    lim_sub = ET.SubElement(limLow, f'{{{OMML_NS}}}lim')
                    for child in list(sub):
                        lim_sub.append(child)
                    top_elem = limLow
                else:
                    limUpp = ET.Element(f'{{{OMML_NS}}}limUpp')
                    e_inner = ET.SubElement(limUpp, f'{{{OMML_NS}}}e')
                    e_inner.append(create_sym_r())
                    
                    lim_sup = ET.SubElement(limUpp, f'{{{OMML_NS}}}lim')
                    for child in list(sup):
                        lim_sup.append(child)
                    top_elem = limUpp
                
                parent.remove(nary)
                parent.insert(idx, top_elem)
                
                if e is not None and list(e):
                    insert_offset = 1
                    for child in list(e):
                        if local(child.tag) in ('ctrlPr', 'argPr'):
                            continue
                        if local(child.tag) == 'r' and not run_text(child).strip():
                            continue
                        parent.insert(idx + insert_offset, child)
                        insert_offset += 1
                        
                modified = True
        return modified

    def apply_rule_double_integral_region(self, omath):
        """二重积分区域居中 (∬)"""
        modified = False
        for nary in list(omath.iter(f'{{{OMML_NS}}}nary')):
            chr_elem = nary.find(f'{mtag("naryPr")}/{mtag("chr")}')
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '∫') if chr_elem is not None else '∫'

            if chr_val in ('∬', '∭', '∯', '∰'):
                sub = nary.find(f'{{{OMML_NS}}}sub')
                sup = nary.find(f'{{{OMML_NS}}}sup')
                
                has_sup = sup is not None and bool(''.join(sup.itertext()).strip())
                has_sub = sub is not None and bool(''.join(sub.itertext()).strip())
                if not has_sup and has_sub:
                    parent = None
                    for p in omath.iter():
                        if nary in list(p):
                            parent = p
                            break
                    if parent is None:
                        continue
                        
                    idx = list(parent).index(nary)
                    e = nary.find(f'{{{OMML_NS}}}e')
                    e_children = list(e) if e is not None else []
                    
                    limLow = ET.Element(f'{{{OMML_NS}}}limLow')
                    new_e = ET.SubElement(limLow, f'{{{OMML_NS}}}e')
                    r_sym = ET.SubElement(new_e, f'{{{OMML_NS}}}r')
                    rPr_m = ET.SubElement(r_sym, f'{{{OMML_NS}}}rPr')
                    sty = ET.SubElement(rPr_m, f'{{{OMML_NS}}}sty')
                    sty.set(f'{{{OMML_NS}}}val', 'p')
                    t_sym = ET.SubElement(r_sym, f'{{{OMML_NS}}}t')
                    t_sym.text = chr_val
                    
                    lim_sub = ET.SubElement(limLow, f'{{{OMML_NS}}}lim')
                    for child in list(sub):
                        lim_sub.append(child)
                        
                    parent.remove(nary)
                    parent.insert(idx, limLow)
                    
                    insert_offset = 1
                    for c in e_children:
                        if local(c.tag) in ('ctrlPr', 'argPr'):
                            continue
                        if local(c.tag) == 'r' and not run_text(c).strip():
                            continue
                        parent.insert(idx + insert_offset, c)
                        insert_offset += 1
                        
                    modified = True
        return modified

    def apply_rule_piecewise_alignment(self, omath):
        """分段函数左对齐"""
        modified = False
        for d in omath.iter(f'{{{OMML_NS}}}d'):
            dPr = d.find(f'{{{OMML_NS}}}dPr')
            beg = dPr.find(f'{{{OMML_NS}}}begChr') if dPr is not None else None
            end = dPr.find(f'{{{OMML_NS}}}endChr') if dPr is not None else None
            beg_val = beg.attrib.get(f'{{{OMML_NS}}}val', '') if beg is not None else ''
            end_val = end.attrib.get(f'{{{OMML_NS}}}val', '') if end is not None else ''
            
            if beg_val == '{' and end_val in ('', None):
                for m in d.iter(f'{{{OMML_NS}}}m'):
                    mPr = m.find(f'{{{OMML_NS}}}mPr')
                    if mPr is not None:
                        for mcJc in mPr.iter(f'{{{OMML_NS}}}mcJc'):
                            if mcJc.attrib.get(f'{{{OMML_NS}}}val') != 'left':
                                mcJc.set(f'{{{OMML_NS}}}val', 'left')
                                modified = True
        return modified

    def apply_rule_eval_bar(self, omath):
        """修复代入线高度 (|)"""
        modified = False
        targets = list(omath.iter(f'{{{OMML_NS}}}sSub')) + list(omath.iter(f'{{{OMML_NS}}}sSubSup'))
        for item_node in targets:
            e = item_node.find(f'{{{OMML_NS}}}e')
            if e is not None:
                # 竖线必须是底座最后一个直接子 run，避免误改嵌套结构内部的 |
                content = [c for c in e if local(c.tag) not in ('ctrlPr', 'argPr')]
                runs = content[-1:] if content and local(content[-1].tag) == 'r' else []
                if runs and is_normal_text(runs[0]):
                    runs = []
                if runs:
                    last_t = runs[-1].find(f'{{{OMML_NS}}}t')
                    if last_t is not None and last_t.text and last_t.text.endswith('|'):
                        e_all_text = ''.join(e.itertext()).strip()
                        if e_all_text == '|':
                            continue
                            
                        # 底座内部若有成对的 |，如 |x| 本身是绝对值，不是代入线
                        if e_all_text.count('|') % 2 == 0:
                            continue
                        if last_t.text == '|':
                            e.remove(runs[-1])
                        else:
                            set_t_text(last_t, last_t.text[:-1])

                        inner_children = [c for c in e if local(c.tag) not in ('ctrlPr', 'argPr')]
                        if not inner_children:
                            continue
                            
                        e.clear()
                        
                        d_elem = ET.SubElement(e, f'{{{OMML_NS}}}d')
                        dPr = ET.SubElement(d_elem, f'{{{OMML_NS}}}dPr')
                        beg_node = ET.SubElement(dPr, f'{{{OMML_NS}}}begChr')
                        beg_node.set(f'{{{OMML_NS}}}val', '')
                        end_node = ET.SubElement(dPr, f'{{{OMML_NS}}}endChr')
                        end_node.set(f'{{{OMML_NS}}}val', '|')
                        grow_node = ET.SubElement(dPr, f'{{{OMML_NS}}}grow')
                        grow_node.set(f'{{{OMML_NS}}}val', 'on')
                        
                        d_e = ET.SubElement(d_elem, f'{{{OMML_NS}}}e')
                        for c in inner_children:
                            d_e.append(c)
                        modified = True
        return modified

    def apply_rule_typography(self, omath):
        """正斜体规范 (数字/微分d/常数e/转置T正体，变量保持斜体)"""
        modified = False
        import re

        TOKEN_REGEX = re.compile(
            # 特征数 Re、Ma、We 等不能拆开 (否则 e 会被当成自然常数改成正体)
            r'(?P<dimless>(?<![A-Za-z])(?:' + '|'.join(self.DIMLESS) + r')(?![a-z]))|' +
            r'(?P<func>' + '|'.join(sorted(MATH_FUNCS, key=lambda x: -len(x))) + r')|' +
            r'(?P<diff>d[xytzuvwrθσρφ])|' +
            r'(?P<cjk>[\u4e00-\u9fff\u3000-\u303f\uff0c\uff1a\uff1b\uff01\uff1f]+)|' +
            r'(?P<rank>r\()|' +
            r'(?P<little_o>\+?o\()|' +
            r'(?P<digit>\d+(?:\.\d+)?)|' +
            r'(?P<op>[\=\+\-\−\±\×\·\<\>\≤\≥\≠\→\∈\∉\∪\∩])|' +
            r'(?P<punct>[\,\.\;\:\|\/\(\)\[\]\{\}])|' +
            r'(?P<symbol>[∞ℝ′″⋯∣∂∫\*（）])|' +
            r'(?P<other>.)',
            re.UNICODE
        )

        def make_r(text, is_plain=False, orig_r=None):
            # 保留原 run 的 m:rPr (字体样式/对齐) 与 w:rPr (字号/颜色)，并保证元素顺序正确
            return clone_run(orig_r, text, upright=bool(is_plain))

        def next_starts_with_letter(node):
            parent = find_parent(omath, node)
            if parent is None:
                return False
            kids = list(parent)
            i = kids.index(node) + 1
            while i < len(kids) and local(kids[i].tag) in ('ctrlPr', 'argPr'):
                i += 1
            if i >= len(kids) or local(kids[i].tag) != 'r':
                return False
            nxt = run_text(kids[i])
            return bool(nxt) and (nxt[0].isascii() and nxt[0].isalpha() or 'α' <= nxt[0] <= 'ω')

        # 转置上标 T
        for sSup in omath.iter(f'{{{OMML_NS}}}sSup'):
            sup = sSup.find(f'{{{OMML_NS}}}sup')
            if sup is not None:
                for r in sup.findall(f'.//{{{OMML_NS}}}r'):
                    t = r.find(f'{{{OMML_NS}}}t')
                    if t is not None and t.text == 'T':
                        rPr = r.find(f'{{{OMML_NS}}}rPr')
                        if rPr is None:
                            rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                            r.insert(0, rPr)
                        sty = rPr.find(f'{{{OMML_NS}}}sty')
                        if sty is None:
                            sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                            sty.set(f'{{{OMML_NS}}}val', 'p')
                            modified = True
                        elif sty.attrib.get(f'{{{OMML_NS}}}val') != 'p':
                            sty.set(f'{{{OMML_NS}}}val', 'p')
                            modified = True

        # 自然常数 e 底座
        for sSup in omath.iter(f'{{{OMML_NS}}}sSup'):
            e_elem = sSup.find(f'{{{OMML_NS}}}e')
            if e_elem is not None:
                runs = e_elem.findall(f'{{{OMML_NS}}}r')
                if len(runs) == 1:
                    t = runs[0].find(f'{{{OMML_NS}}}t')
                    if t is not None and t.text == 'e':
                        rPr = runs[0].find(f'{{{OMML_NS}}}rPr')
                        if rPr is None:
                            rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                            runs[0].insert(0, rPr)
                        sty = rPr.find(f'{{{OMML_NS}}}sty')
                        if sty is None:
                            sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                            sty.set(f'{{{OMML_NS}}}val', 'p')
                            modified = True
                        elif sty.attrib.get(f'{{{OMML_NS}}}val') != 'p':
                            sty.set(f'{{{OMML_NS}}}val', 'p')
                            modified = True

        # 分式中的微分 d
        for frac in omath.iter(f'{{{OMML_NS}}}f'):
            num = frac.find(f'{{{OMML_NS}}}num')
            den = frac.find(f'{{{OMML_NS}}}den')
            for part in (num, den):
                if part is not None:
                    for r in part.findall(f'.//{{{OMML_NS}}}r'):
                        t = r.find(f'{{{OMML_NS}}}t')
                        sole_d = ''.join(part.itertext()).strip() == 'd' and part is num and \
                            ''.join(den.itertext()).strip().startswith('d') if den is not None else False
                        if t is not None and t.text == 'd' and (sole_d or next_starts_with_letter(r)):
                            rPr = r.find(f'{{{OMML_NS}}}rPr')
                            if rPr is None:
                                rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                                r.insert(0, rPr)
                            sty = rPr.find(f'{{{OMML_NS}}}sty')
                            if sty is None:
                                sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                                sty.set(f'{{{OMML_NS}}}val', 'p')
                                modified = True
                            elif sty.attrib.get(f'{{{OMML_NS}}}val') != 'p':
                                sty.set(f'{{{OMML_NS}}}val', 'p')
                                modified = True

        # 全公式词元拆分与正斜体处理
        runs = list(omath.iter(f'{{{OMML_NS}}}r'))
        for r in runs:
            t_elem = r.find(f'{{{OMML_NS}}}t')
            if t_elem is None or not t_elem.text:
                continue
            text = t_elem.text

            if len(r.findall(mtag('t'))) != 1 or is_normal_text(r):
                continue
            rPr = r.find(f'{{{OMML_NS}}}rPr')
            sty = rPr.find(f'{{{OMML_NS}}}sty') if rPr is not None else None
            is_plain = (sty is not None and sty.attrib.get(f'{{{OMML_NS}}}val') in ('p', 'b'))
            scr = rPr.find(mtag('scr')) if rPr is not None else None
            if scr is not None and scr.attrib.get(mtag('val'), 'roman') != 'roman':
                continue  # 双线体/花体等特殊字体保持原样

            # 作者已设为正体的多字符文本 (如单位 cm、算子名 Re) 不再拆成斜体字母
            if is_plain and (len(text) > 1 or text in ('∬', '∭', '∑', '∏', '∫')):
                continue
            if is_plain and text.isdigit():
                continue

            parent = find_parent(omath, r)
            if parent is None:
                continue

            parent_tag = parent.tag.split('}')[-1]
            is_in_sup = (parent_tag == 'sup')

            # 纯数字正体
            if text.isdigit() and not is_plain:
                modified = set_upright(r) or modified
                continue

            needs_split = False
            if len(text) > 1 and text not in MATH_FUNCS:
                needs_split = True
            elif text in ('<', '>', '=', '≤', '≥', '≠', '+', '-', '−', '±', '×', '·', ',', '.', ';', ':', '|', '/', '(', ')', '[', ']', '{', '}', '∞', 'ℝ', '⋯', '∣', '′', '″', '∂', '∫', '*', '（', '）'):
                if not is_plain:
                    modified = set_upright(r) or modified
                    continue
            elif text == 'e' and not is_plain:
                modified = set_upright(r) or modified
                continue
            elif text == 'd' and not is_plain:
                # 只有后面紧跟变量 (d x) 时才是微分算子；单独的 d (如等差数列公差) 保持斜体
                if next_starts_with_letter(r):
                    modified = set_upright(r) or modified
                continue

            if needs_split:
                idx = list(parent).index(r)
                parent.remove(r)

                new_tokens = []
                for m in TOKEN_REGEX.finditer(text):
                    kind = m.lastgroup
                    val = m.group()

                    if kind == 'dimless':
                        new_tokens.append((val, False))
                    elif kind == 'digit':
                        new_tokens.append((val, True))
                    elif kind == 'cjk':
                        new_tokens.append((val, True))
                    elif kind == 'func':
                        new_tokens.append((val, True))
                    elif kind == 'diff':
                        new_tokens.append(('d', True))
                        new_tokens.append((val[1:], False))
                    elif kind == 'rank':
                        new_tokens.append(('r', True))
                        new_tokens.append(('(', True))
                    elif kind == 'little_o':
                        if val.startswith('+'):
                            new_tokens.append(('+', True))
                        new_tokens.append(('o', True))
                        new_tokens.append(('(', True))
                    elif kind in ('op', 'punct', 'symbol'):
                        new_tokens.append((val, True))
                    else:
                        if val == 'e':
                            new_tokens.append(('e', True))
                        elif val == 'T' and is_in_sup:
                            new_tokens.append(('T', True))
                        elif val == 'd' and m.end() == len(text) and next_starts_with_letter(r):
                            new_tokens.append(('d', True))
                        else:
                            new_tokens.append((val, False))

                # 合并相邻同样式的 token，减少碎片 run
                merged = []
                for tok_val, tok_plain in new_tokens:
                    if merged and merged[-1][1] == tok_plain and not tok_plain:
                        merged[-1] = (merged[-1][0] + tok_val, tok_plain)
                    else:
                        merged.append((tok_val, tok_plain))
                if len(merged) == 1 and merged[0][1] == is_plain:
                    parent.insert(idx, r)  # 拆分后样式没有任何变化，放回原 run
                    continue

                for tok_val, tok_plain in merged:
                    parent.insert(idx, make_r(tok_val, is_plain=tok_plain, orig_r=r))
                    idx += 1
                modified = True

        return modified

    def apply_rule_set_braces(self, omath):
        """集合大括号修复：将大括号正体化并规范化自适应高度"""
        modified = False
        runs = list(omath.iter(f'{{{OMML_NS}}}r'))
        for r in runs:
            t = r.find(f'{{{OMML_NS}}}t')
            # 只处理纯括号 run；"{x|x>0}" 这类整段 run 若设为正体会把变量 x 也变成正体
            if t is not None and t.text and t.text.strip() and set(t.text.strip()) <= set('{}｛｝'):
                if not is_normal_text(r) and set_upright(r):
                    modified = True

        for d in omath.iter(f'{{{OMML_NS}}}d'):
            dPr = d.find(f'{{{OMML_NS}}}dPr')
            if dPr is not None:
                beg = dPr.find(f'{{{OMML_NS}}}begChr')
                end = dPr.find(f'{{{OMML_NS}}}endChr')
                beg_val = beg.attrib.get(f'{{{OMML_NS}}}val', '') if beg is not None else ''
                end_val = end.attrib.get(f'{{{OMML_NS}}}val', '') if end is not None else ''
                if beg_val in ('{', '｛') or end_val in ('}', '｝'):
                    if set_ordered_prop(dPr, 'grow', 'on', D_PR_ORDER):
                        modified = True
        return modified

    def apply_rule_font_times_new_roman(self, omath):
        """统一公式英文字体为 Times New Roman"""
        modified = False
        for r in omath.iter(f'{{{OMML_NS}}}r'):
            w_rPr = ensure_w_rpr(r)  # 放在 m:rPr 之后
            rFonts = w_rPr.find(f'{{{WORD_NS}}}rFonts')
            if rFonts is None:
                # rFonts 是 w:rPr 的第一个子元素 (rStyle 除外)
                rFonts = ET.Element(f'{{{WORD_NS}}}rFonts')
                has_style = len(w_rPr) and local(w_rPr[0].tag) == 'rStyle'
                w_rPr.insert(1 if has_style else 0, rFonts)
            if rFonts.attrib.get(f'{{{WORD_NS}}}ascii') != 'Times New Roman':
                rFonts.set(f'{{{WORD_NS}}}ascii', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}cs', 'Times New Roman')
                modified = True
        return modified

    FULLWIDTH_MAP = {**{chr(0xFF10 + i): str(i) for i in range(10)},
                     **{chr(0xFF21 + i): chr(0x41 + i) for i in range(26)},
                     **{chr(0xFF41 + i): chr(0x61 + i) for i in range(26)},
                     '（': '(', '）': ')', '［': '[', '］': ']', '｛': '{', '｝': '}', '，': ',', '：': ':',
                     '；': ';', '＋': '+', '－': '−', '＝': '=', '＜': '<', '＞': '>', '｜': '|', '／': '/',
                     '．': '.', '！': '!', '＇': '′', '　': ' '}

    # 量纲一的特征数 (GB 3102.12 / ISO 80000-11)：是物理量，应为斜体且两字母不可拆开
    DIMLESS = ('Re', 'Ma', 'Pr', 'Nu', 'Fr', 'St', 'Sr', 'Eu', 'We', 'Kn', 'Pe', 'Gr', 'Ra', 'Bi',
               'Fo', 'Sc', 'Sh', 'Ri', 'Ec', 'Le', 'Ca', 'Da', 'Ek', 'Ro', 'Ta', 'Ste', 'Ar', 'Bo')
    # 描述性下标 (单词或缩写，不是量)：应为正体
    DESC_SUBS = ('max', 'min', 'ref', 'eff', 'cr', 'crit', 'in', 'out', 'avg', 'ave', 'mean', 'tot',
                 'total', 'rel', 'abs', 'atm', 'air', 'opt', 'eq', 'st', 'th', 'sat', 'cal', 'exp', 'test',
                 'act', 'nom', 'inc', 'ext', 'int', 'amb', 'fuel', 'ad', 'ex', 'wall', 'inlet', 'outlet',
                 'ideal', 'real', 'stall', 'cruise', 'design', 'loss', 'gas', 'liq', 'vap', 'sol', 'turb',
                 'lam', 'visc', 'dyn', 'stat', 'cond', 'conv', 'rad', 'gen', 'mech', 'elec', 'kin', 'pot')
    # 气动系数 C_L 升力、C_D 阻力、C_M 力矩、C_N 法向力、C_Y 侧力：下标是缩写，应为正体
    COEF_SUBS = ('L', 'D', 'M', 'N', 'Y')
    # 单位符号 (GB 3100/3102)：正体。单字母单位与变量同形，只在与数字间有空格或后接 / · ^ 时才判定
    UNITS_MULTI = ('mm', 'cm', 'km', 'μm', 'nm', 'ms', 'μs', 'ns', 'min', 'kg', 'kN', 'MN', 'Pa', 'kPa',
                   'MPa', 'GPa', 'hPa', 'bar', 'kJ', 'MJ', 'kW', 'MW', 'Hz', 'kHz', 'MHz', 'rad', 'sr',
                   'mol', 'kmol', 'cd', 'lm', 'lx', 'Wb', 'dB', 'mL', 'eV', 'rpm', 'kWh', 'atm', 'mbar')
    UNITS_SINGLE = ('m', 's', 'g', 'N', 'K', 'J', 'W', 'L', 'h', 'A', 'V', 'C', 'T', 'F', 'Ω')

    def apply_rule_engineering(self, omath):
        """工程/物理符号正斜体规范 (GB 3101-93、ISO 80000)。"""
        modified = False

        def set_italic(r):
            cur = get_sty(r)
            if cur not in ('p', 'b'):
                return False
            sty = r.find(mtag('rPr')).find(mtag('sty'))
            if cur == 'b':
                sty.set(mtag('val'), 'bi')
            else:
                r.find(mtag('rPr')).remove(sty)
            return True

        # 1. 特征数：被拆成 "R" + 正体 "e" 的合并回来；被设成正体的改回斜体
        for parent in list(omath.iter()):
            kids = list(parent)
            i = 0
            while i < len(kids) - 1:
                a, b = kids[i], kids[i + 1]
                if a.tag == mtag('r') and b.tag == mtag('r') and run_text(a) + run_text(b) in self.DIMLESS \
                        and len(a.findall(mtag('t'))) == 1 and not is_normal_text(a):
                    set_t_text(a.find(mtag('t')), run_text(a) + run_text(b))
                    parent.remove(b)
                    kids.pop(i + 1)
                    modified = True
                i += 1
        for r in omath.iter(mtag('r')):
            if run_text(r) in self.DIMLESS and not is_normal_text(r) and set_italic(r):
                modified = True
        return self._engineering_part2(omath) or modified

    def _engineering_part2(self, omath):
        modified = False

        # 2. 描述性下标正体：T_max、p_ref、v_in、Re_cr、C_L、C_D
        for node in omath.iter():
            if local(node.tag) not in ('sSub', 'sSubSup'):
                continue
            sub, base = node.find(mtag('sub')), node.find(mtag('e'))
            runs = [r for r in sub.iter(mtag('r'))] if sub is not None else []
            if not runs or any(local(c.tag) not in ('r', 'ctrlPr') for c in sub):
                continue
            text = ''.join(run_text(r) for r in runs).strip()
            base_text = ''.join(base.itertext()).strip() if base is not None else ''
            if text in self.DESC_SUBS or (base_text in ('C', 'c') and text in self.COEF_SUBS):
                for r in runs:
                    if not is_normal_text(r) and set_upright(r):
                        modified = True

        # 3. 数字后的单位正体：101325 Pa、30 m/s、9.8 m/s²
        for parent in list(omath.iter()):
            kids = list(parent)
            for i, kid in enumerate(kids):
                if kid.tag != mtag('r') or not _re.search(r'\d\s*$', run_text(kid)):
                    continue
                spaced = run_text(kid).endswith(' ')
                j = i + 1
                unit_runs = []
                while j < len(kids):
                    nxt = kids[j]
                    base = nxt.find(mtag('e')) if local(nxt.tag) == 'sSup' else None
                    target = base.find(mtag('r')) if base is not None and len(base.findall(mtag('r'))) == 1 else nxt
                    if target is None or target.tag != mtag('r'):
                        break
                    t = run_text(target).strip()
                    if t in self.UNITS_MULTI or t in ('/', '·', '⋅'):
                        unit_runs.append(target)
                    elif t in self.UNITS_SINGLE:
                        follow = run_text(kids[j + 1]).strip() if j + 1 < len(kids) and kids[j + 1].tag == mtag('r') else ''
                        if spaced or run_text(target).startswith(' ') or unit_runs or base is not None \
                                or follow in ('/', '·', '⋅'):
                            unit_runs.append(target)
                        else:
                            break
                    else:
                        break
                    j += 1
                if any(run_text(u).strip() not in ('/', '·', '⋅') for u in unit_runs):
                    for u in unit_runs:
                        if set_upright(u):
                            modified = True

        # 4. 数学常数 π 正体；作差分算子的 Δ (后接变量，如 Δp、Δt) 正体
        #    先把 "πr"、"Δp" 这类连在一个 run 里的拆开
        for r in list(omath.iter(mtag('r'))):
            text = run_text(r)
            if len(text) < 2 or is_normal_text(r) or len(r.findall(mtag('t'))) != 1:
                continue
            if 'π' not in text and not _re.search(r'Δ[A-Za-zα-ω]', text):
                continue
            if any('\u4e00' <= ch <= '\u9fff' for ch in text):
                continue
            pieces = [s for s in _re.split(r'(π|Δ)', text) if s]
            parent = find_parent(omath, r)
            if parent is None or len(pieces) < 2:
                continue
            idx = list(parent).index(r)
            parent.remove(r)
            for s in pieces:
                parent.insert(idx, clone_run(r, s))
                idx += 1
            modified = True
        for r in omath.iter(mtag('r')):
            t = run_text(r)
            if is_normal_text(r):
                continue
            if t == 'π' or (t == 'Δ' and self._next_is_letter(omath, r)):
                if set_upright(r):
                    modified = True
        return modified

    @staticmethod
    def _next_is_letter(omath, r):
        parent = find_parent(omath, r)
        kids = list(parent) if parent is not None else []
        i = kids.index(r) + 1 if r in kids else len(kids)
        if i < len(kids):
            nxt = ''.join(kids[i].itertext())
            return bool(nxt) and nxt[0].isalpha()
        return False

    def _math_text_runs(self, omath):
        """可安全改字符的数学 run：排除普通文本 (m:nor) 和含中文的 run。"""
        for r in omath.iter(mtag('r')):
            if is_normal_text(r):
                continue
            text = run_text(r)
            if text and not any('\u4e00' <= ch <= '\u9fff' for ch in text):
                yield r

    def apply_rule_fullwidth(self, omath):
        """全角符号转半角：OCR 常把 （ ） ， ＝ １ 等全角字符写进公式，导致间距过宽、字体不统一。"""
        modified = False
        for r in self._math_text_runs(omath):
            for t in r.findall(mtag('t')):
                new = ''.join(self.FULLWIDTH_MAP.get(ch, ch) for ch in (t.text or ''))
                if new != t.text:
                    set_t_text(t, new)
                    modified = True
        return modified

    def apply_rule_minus_sign(self, omath):
        """连字符改为数学减号：ASCII 连字符 - 在公式中显示为短横线，标准减号是 U+2212 −。"""
        modified = False
        for r in self._math_text_runs(omath):
            for t in r.findall(mtag('t')):
                if t.text and '-' in t.text:
                    set_t_text(t, t.text.replace('-', '−'))
                    modified = True
        return modified

    def apply_rule_math_position(self, omath):
        """清除公式字符的上浮/下沉偏移 (w:position)，修复公式与正文基线不对齐。"""
        modified = False
        for r in omath.iter(mtag('r')):
            w_rPr = r.find(f'{{{WORD_NS}}}rPr')
            if w_rPr is None:
                continue
            for pos in w_rPr.findall(f'{{{WORD_NS}}}position'):
                if pos.attrib.get(f'{{{WORD_NS}}}val', '0') != '0':
                    w_rPr.remove(pos)
                    modified = True
        return modified

    # ================== Extraction & Process ==================

    def process(self, options, modify=False):
        """
        Analyze or repair the document.
        Returns document paragraphs (for full text view) and issuesList (for comparison view).
        """
        parts = getattr(self, 'parts', None) or {'word/document.xml': (self.root, [])}
        paragraphs = []      # (段落, 部件标签)
        total_omath = 0
        for part_name, (part_root, _ns) in parts.items():
            kind = _re.sub(r'\d+$', '', part_name[5:-4])
            label = PART_LABELS.get(kind, '')
            paragraphs.extend((p, label) for p in part_root.iter(f'{{{WORD_NS}}}p'))
            # 不计入已删除修订中的公式
            deleted = sum(len(d.findall(f'.//{{{OMML_NS}}}oMath'))
                          for tag in ('del', 'moveFrom') for d in part_root.iter(f'{{{WORD_NS}}}{tag}'))
            total_omath += len(part_root.findall(f'.//{{{OMML_NS}}}oMath')) - deleted
        
        issues_found = {
            'splitGluedTokens': 0,
            'fixIsolatedDelimiters': 0,
            'fixNaryEmptyOperand': 0,
            'fixDoubleIntegralRegion': 0,
            'fixIntegralSize': 0,
            'fixLimits': 0,
            'fixNaryLimLoc': 0,
            'fixPiecewiseAlignment': 0,
            'fixEvalBar': 0,
            'fixTypography': 0,
            'fixSetBraces': 0,
            'fixFontTimesNewRoman': 0,
            'fixEngineeringSymbols': 0,
            'fixFullwidth': 0,
            'fixMinusSign': 0,
            'fixMathPosition': 0,
            'fixLineSpacing': 0,
            'fixSnapToGrid': 0,
            'fixDisplayCenter': 0,
            'fixEquationNumbering': 0,
            'addEquationNumbers': 0,
        }
        DEFAULT_OFF = ('addEquationNumbers',)
        options = {k: options.get(k, k not in DEFAULT_OFF) for k in issues_found}
        plan = LayoutPlan(paragraphs, self.root, getattr(self, 'styles_root', None), options)
        for k, v in plan.counts.items():
            issues_found[k] = v
        
        issues_list = []
        issue_id = 1
        math_counter = 1
        
        document_paragraphs = []

        for p_idx, (p, part_label) in enumerate(paragraphs):
            p_text = ''.join(p.itertext()).strip()
            if part_label:
                p_text = f'【{part_label}】' + p_text
            
            # 遍历提取段落内的视觉元素 (文字、换行与公式)
            elements = []
            
            def walk_paragraph(elem):
                nonlocal issue_id, math_counter
                tag = elem.tag.split('}')[-1]
                
                if tag == 'oMathPara':
                    for om in elem.findall(f'.//{{{OMML_NS}}}oMath'):
                        process_math_node(om, is_display=True)
                elif tag == 'oMath':
                    process_math_node(elem, is_display=False)
                elif tag in WALK_SKIP_TAGS:
                    return
                elif tag == 'AlternateContent':
                    # 只处理 mc:Choice 分支，Fallback 是同一内容的副本
                    choice = next((c for c in elem if local(c.tag) == 'Choice'), None)
                    if choice is not None:
                        for sub in choice:
                            walk_paragraph(sub)
                elif tag == 'r' and elem.tag.startswith(f'{{{WORD_NS}}}'):
                    for sub in elem:
                        sub_tag = sub.tag.split('}')[-1]
                        if sub_tag == 't':
                            if sub.text:
                                elements.append({'type': 'text', 'text': sub.text})
                        elif sub_tag == 'br':
                            elements.append({'type': 'br'})
                        elif sub_tag == 'tab':
                            elements.append({'type': 'text', 'text': '    '})
                    # run 内部也可能嵌套公式 (如 mc:AlternateContent)，继续向下查找
                    for sub in elem:
                        if local(sub.tag) not in ('t', 'br', 'tab'):
                            walk_paragraph(sub)
                else:
                    # hyperlink / sdt / smartTag / fldSimple / ins / moveTo / customXml / dir / bdo 等容器
                    for sub in elem:
                        walk_paragraph(sub)

            def process_math_node(om_node, is_display):
                nonlocal issue_id, math_counter
                om_clone = copy.deepcopy(om_node)
                before_latex = omml_to_latex(om_clone).strip()
                before_text = ''.join(om_clone.itertext()).strip()
                
                changed_rules = []
                target_om = om_node if modify else om_clone
                # 公式内部的编号文字 (#(1) 或末尾的 "  (1)") 移到右侧编号单元格
                if id(om_node) in plan.math_strip:
                    strip_math_number(target_om, plan.math_strip[id(om_node)])
                
                if options.get('splitGluedTokens', True):
                    if self.apply_rule_split_tokens(target_om):
                        issues_found['splitGluedTokens'] += 1
                        changed_rules.append(('splitGluedTokens', '拆分连写函数', '拆分连写的 lncos、sincos 等为独立函数'))
                
                for key, func, name, desc in (
                    ('fixFullwidth', self.apply_rule_fullwidth, '全角符号转半角', '把公式中的全角括号、逗号、数字等改为半角'),
                    ('fixMinusSign', self.apply_rule_minus_sign, '连字符改为减号', '把公式中的短横线 - 改为数学减号 −'),
                    ('fixMathPosition', self.apply_rule_math_position, '清除公式上浮/下沉', '去掉公式字符的位置偏移，与正文基线对齐'),
                ):
                    if options.get(key) and func(target_om):
                        issues_found[key] += 1
                        changed_rules.append((key, name, desc))

                if options.get('fixIsolatedDelimiters', True):
                    if self.apply_rule_isolated_delimiters(target_om):
                        issues_found['fixIsolatedDelimiters'] += 1
                        changed_rules.append(('fixIsolatedDelimiters', '修复括号上标', '修复 (式子)^n 括号上标结构，防止括号被误删'))
                
                if options.get('fixNaryLimLoc', True):
                    if self.apply_rule_nary_lim_loc(target_om):
                        issues_found['fixNaryLimLoc'] += 1
                        changed_rules.append(('fixNaryLimLoc', '求和/乘积上下标居中', '将 ∑ 和 ∏ 的上下标置于正上方和正下方'))
                
                if options.get('fixNaryEmptyOperand', True):
                    if self.apply_rule_nary_empty_operand(target_om):
                        issues_found['fixNaryEmptyOperand'] += 1
                        changed_rules.append(('fixNaryEmptyOperand', '消除积分虚线框', '补全被积项，消除 Word 积分虚线占位框'))
                
                if options.get('fixDoubleIntegralRegion', True):
                    if self.apply_rule_double_integral_region(target_om):
                        issues_found['fixDoubleIntegralRegion'] += 1
                        changed_rules.append(('fixDoubleIntegralRegion', '二重积分区域居中', '将二重/三重积分区域 (如 D) 置于积分符号正下方'))

                if options.get('fixIntegralSize', True):
                    if self.apply_rule_integral_size(target_om):
                        issues_found['fixIntegralSize'] += 1
                        changed_rules.append(('fixIntegralSize', '固定积分号高度', '防止积分符号随内容过度拉伸变长'))
                
                if options.get('fixLimits', True):
                    if self.apply_rule_fix_limits(target_om):
                        issues_found['fixLimits'] += 1
                        changed_rules.append(('fixLimits', '极限下标居中', '将 lim 的下标置于正下方居中'))
                
                if options.get('fixPiecewiseAlignment', True):
                    if self.apply_rule_piecewise_alignment(target_om):
                        issues_found['fixPiecewiseAlignment'] += 1
                        changed_rules.append(('fixPiecewiseAlignment', '分段函数左对齐', '将分段函数条件列改为左对齐'))
                
                if options.get('fixEvalBar', True):
                    if self.apply_rule_eval_bar(target_om):
                        issues_found['fixEvalBar'] += 1
                        changed_rules.append(('fixEvalBar', '修复代入线高度', '将求导/积分代入短线 | 改为自适应高度竖线'))
                
                if options.get('fixTypography', True):
                    if self.apply_rule_typography(target_om):
                        issues_found['fixTypography'] += 1
                        changed_rules.append(('fixTypography', '正斜体规范', '数字、微分d、常数e、转置T设为正体，变量保持斜体'))

                if options.get('fixEngineeringSymbols'):
                    if self.apply_rule_engineering(target_om):
                        issues_found['fixEngineeringSymbols'] += 1
                        changed_rules.append(('fixEngineeringSymbols', '工程符号规范',
                                              '特征数 Re/Ma 斜体、描述性下标 max/ref/C_L 正体、单位 Pa/m/s 正体、π 与差分 Δ 正体'))
                
                if options.get('fixSetBraces', True):
                    if self.apply_rule_set_braces(target_om):
                        issues_found['fixSetBraces'] += 1
                        changed_rules.append(('fixSetBraces', '集合大括号修复', '规范集合大括号定界符'))

                if options.get('fixFontTimesNewRoman', True):
                    if self.apply_rule_font_times_new_roman(target_om):
                        issues_found['fixFontTimesNewRoman'] += 1
                        changed_rules.append(('fixFontTimesNewRoman', '公式字体设为 Times New Roman', '将公式中的英文字母和数字字体设为 Times New Roman'))
                
                normalize_schema_order(target_om)
                is_modified = bool(changed_rules)
                after_latex = omml_to_latex(target_om).strip()
                after_text = ''.join(target_om.itertext()).strip()
                
                curr_math_id = math_counter
                math_counter += 1

                if is_modified:
                    issues_list.append({
                        'id': issue_id,
                        'mathId': curr_math_id,
                        'paragraphIndex': p_idx,
                        'paragraphSnippet': p_text[:80] + ('...' if len(p_text) > 80 else ''),
                        'rulesApplied': [r[0] for r in changed_rules],
                        'ruleNames': [r[1] for r in changed_rules],
                        'descriptions': [r[2] for r in changed_rules],
                        'beforeText': before_text,
                        'afterText': after_text,
                        'beforeLatex': before_latex,
                        'afterLatex': after_latex
                    })
                    issue_id += 1

                elements.append({
                    'type': 'math',
                    'mathId': curr_math_id,
                    'isDisplay': is_display,
                    'isModified': is_modified,
                    'beforeLatex': before_latex,
                    'afterLatex': after_latex,
                    'beforeText': before_text,
                    'afterText': after_text,
                    'rulesApplied': [r[0] for r in changed_rules] if is_modified else [],
                    'ruleNames': [r[1] for r in changed_rules] if is_modified else [],
                    'descriptions': [r[2] for r in changed_rules] if is_modified else []
                })

            for child in p:
                walk_paragraph(child)

            has_math = any(e['type'] == 'math' for e in elements)
            has_modified = any(e.get('isModified') for e in elements if e['type'] == 'math')

            # 段落级修改 (行距、网格、编号) 合并进对比列表
            p_fixes = plan.para_fixes.get(id(p), [])
            eq_info = plan.eq.get(id(p))
            eq_number = eq_info['display_number'] if eq_info else None
            math_elems = [e for e in elements if e['type'] == 'math']
            if p_fixes and math_elems:
                has_modified = True
                me = math_elems[0]
                if eq_number:
                    me['isDisplay'] = True
                after = me['afterLatex'] + (f' \\qquad ({eq_number})' if eq_number else '')
                before = me['beforeLatex']
                if eq_info and eq_info['source'] == 'text':
                    before += f' \\quad \\text{{({eq_info["number"]})}}'
                existing = next((it for it in issues_list if it['mathId'] == me['mathId']), None)
                if existing is None:
                    existing = {'id': issue_id, 'mathId': me['mathId'], 'paragraphIndex': p_idx,
                                'paragraphSnippet': p_text[:80] + ('...' if len(p_text) > 80 else ''),
                                'rulesApplied': [], 'ruleNames': [], 'descriptions': [],
                                'beforeText': me['beforeText'], 'afterText': me['afterText']}
                    issues_list.append(existing)
                    issue_id += 1
                existing['beforeLatex'], existing['afterLatex'] = before, after
                for key, name in p_fixes:
                    existing['rulesApplied'].append(key)
                    existing['ruleNames'].append(name)
                    existing['descriptions'].append(name)
                for target in (me,):
                    target['isModified'] = True
                    target['rulesApplied'] = me.get('rulesApplied', []) + [k for k, _ in p_fixes]
                    target['ruleNames'] = me.get('ruleNames', []) + [n for _, n in p_fixes]
                    target['descriptions'] = me.get('descriptions', []) + [n for _, n in p_fixes]

            document_paragraphs.append({
                'index': p_idx,
                'snippet': p_text[:100] + ('...' if len(p_text) > 100 else ''),
                'hasMath': has_math,
                'hasModified': has_modified,
                'eqNumber': eq_number,
                'eqNumberSource': eq_info['source'] if eq_info else None,
                'elements': elements
            })

        if modify:
            plan.apply(paragraphs, getattr(self, 'styles_root', None))

        # 修复模式时，统一正文英文字体为 Times New Roman
        if modify and options.get('fixFontTimesNewRoman', True):
            all_w_runs = [r for part_root, _ns in parts.values() for r in part_root.iter(f'{{{WORD_NS}}}r')]
            for r in all_w_runs:
                w_rPr = r.find(f'{{{WORD_NS}}}rPr')
                if w_rPr is None:
                    w_rPr = ET.Element(f'{{{WORD_NS}}}rPr')
                    r.insert(0, w_rPr)
                rFonts = w_rPr.find(f'{{{WORD_NS}}}rFonts')
                if rFonts is None:
                    rFonts = ET.Element(f'{{{WORD_NS}}}rFonts')
                    has_style = len(w_rPr) and local(w_rPr[0].tag) == 'rStyle'
                    w_rPr.insert(1 if has_style else 0, rFonts)
                rFonts.set(f'{{{WORD_NS}}}ascii', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}cs', 'Times New Roman')

        return {
            'filePath': self.docx_path,
            'fileName': os.path.basename(self.docx_path),
            'totalParagraphs': len(paragraphs),
            'totalOMath': total_omath,
            'totalIssues': len(issues_list),
            'issuesFound': issues_found,
            'issuesList': issues_list,
            'paragraphs': document_paragraphs
        }


def render_native_pages(docx_path, output_dir, max_pages=12):
    """调用本机 Microsoft Word 导出高保真 PDF 并利用 PyMuPDF 渲染为 Word 原生真实页面图像"""
    import win32com.client
    import fitz
    
    os.makedirs(output_dir, exist_ok=True)
    abs_docx = os.path.abspath(docx_path)
    temp_pdf = os.path.join(output_dir, '_native_render.pdf')
    
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(abs_docx, ReadOnly=True)
        doc.SaveAs2(os.path.abspath(temp_pdf), FileFormat=17) # 17 = wdExportFormatPDF
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass
        
    pdf_doc = fitz.open(temp_pdf)
    image_names = []
    try:
        for i in range(min(max_pages, len(pdf_doc))):
            page = pdf_doc[i]
            pix = page.get_pixmap(dpi=150)
            img_name = f'page_{i+1}.png'
            img_file = os.path.join(output_dir, img_name)
            pix.save(img_file)
            image_names.append(img_name)
        page_count = len(pdf_doc)
    finally:
        pdf_doc.close()
        
    return {
        'success': True,
        'pageCount': page_count,
        'renderedPages': image_names,
        'outputDir': output_dir
    }

def main():
    parser = argparse.ArgumentParser(description="Word Math OMML Repair Engine")
    parser.add_argument("--mode", choices=["analyze", "repair", "render_pages"], default="analyze", help="Mode: analyze, repair, or render_pages")
    parser.add_argument("--input", required=True, help="Path to input .docx file")
    parser.add_argument("--output", help="Path to output repaired .docx file (required for repair mode)")
    parser.add_argument("--options", default="{}", help="JSON string of rule options")
    parser.add_argument("--export-pdf", action="store_true", help="Also export PDF via Word COM")
    parser.add_argument("--render-dir", default="rendered_pages", help="Directory to save rendered page images")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(json.dumps({"error": f"输入文件不存在: {args.input}"}, ensure_ascii=False))
        sys.exit(1)
        
    try:
        options = json.loads(args.options)
    except Exception:
        options = {}
    
    try:
        if args.mode == "render_pages":
            result = render_native_pages(args.input, os.path.abspath(args.render_dir))
            print(json.dumps(result, ensure_ascii=False))
            return

        engine = DocxMathRepairEngine(args.input)
        if args.mode == "analyze":
            result = engine.process(options, modify=False)
            print(json.dumps(result, ensure_ascii=False))
        elif args.mode == "repair":
            if not args.output:
                print(json.dumps({"error": "修复模式必须指定 --output 输出路径"}, ensure_ascii=False))
                sys.exit(1)
            result = engine.process(options, modify=True)
            actual_saved = engine.save(args.output, unify_fonts=options.get('fixFontTimesNewRoman', True))
            result['savedTo'] = actual_saved
            
            if args.export_pdf:
                try:
                    import win32com.client
                    word = win32com.client.Dispatch('Word.Application')
                    word.Visible = False
                    pdf_path = os.path.splitext(os.path.abspath(args.output))[0] + '.pdf'
                    doc = word.Documents.Open(os.path.abspath(args.output))
                    doc.SaveAs2(pdf_path, FileFormat=17)
                    doc.Close(False)
                    word.Quit()
                    result['exportedPdf'] = pdf_path
                except Exception as e_pdf:
                    result['pdfWarning'] = f"PDF 导出提示: {str(e_pdf)}"
                    
            print(json.dumps(result, ensure_ascii=False))
    except Exception as e:
        import traceback
        err_msg = {"error": str(e), "traceback": traceback.format_exc()}
        print(json.dumps(err_msg, ensure_ascii=False))
        sys.exit(1)

if __name__ == "__main__":
    main()
