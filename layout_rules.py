# -*- coding: utf-8 -*-
"""段落级公式排版规则：公式编号、行距截断、文档网格、独立公式居中。"""
import re
import copy
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
XML_NS = 'http://www.w3.org/XML/1998/namespace'


def w(name):
    return f'{{{W}}}{name}'


def m(name):
    return f'{{{M}}}{name}'


def local(tag):
    return tag.split('}')[-1] if isinstance(tag, str) else ''


# CT_PPrBase 子元素顺序
PPR_ORDER = ('pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
             'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
             'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
             'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
             'sectPr', 'pPrChange')

NUM_RE = re.compile(r'^[(（]\s*([0-9]+(?:[.\-－–][0-9]+)*[a-z]?)\s*[)）]$')
HASH_NUM_RE = re.compile(r'#\s*[(（]\s*([^()（）]+?)\s*[)）]\s*$')


def get_ppr(p, create=True):
    pPr = p.find(w('pPr'))
    if pPr is None and create:
        pPr = ET.Element(w('pPr'))
        p.insert(0, pPr)
    return pPr


def ppr_child(pPr, name, create=True):
    node = pPr.find(w(name))
    if node is not None or not create:
        return node
    node = ET.Element(w(name))
    rank = PPR_ORDER.index(name)
    pos = len(pPr)
    for i, c in enumerate(list(pPr)):
        n = local(c.tag)
        if n in PPR_ORDER and PPR_ORDER.index(n) > rank:
            pos = i
            break
    pPr.insert(pos, node)
    return node


def own_content(p):
    """段落自身 (不含嵌套文本框段落) 的公式节点与公式外文字。"""
    maths, texts = [], []

    def walk(node):
        for c in node:
            tag = local(c.tag)
            if tag in ('p', 'txbxContent', 'del', 'moveFrom', 'pPr', 'rPr', 'instrText', 'delText'):
                continue
            if tag == 'oMath':
                maths.append(c)
            elif tag == 't' and c.tag == w('t'):
                texts.append(c)
            else:
                walk(c)
    walk(p)
    return maths, texts


class StyleResolver:
    """解析段落的有效行距 (直接格式 → 段落样式链 → 文档默认)。"""

    def __init__(self, styles_root):
        self.styles = {}
        self.default_para = None
        self.doc_default_spacing = None
        if styles_root is None:
            return
        dd = styles_root.find(f'{w("docDefaults")}/{w("pPrDefault")}/{w("pPr")}/{w("spacing")}')
        self.doc_default_spacing = dd
        for st in styles_root.findall(w('style')):
            if st.get(w('type')) == 'paragraph':
                sid = st.get(w('styleId'))
                self.styles[sid] = st
                if st.get(w('default')) in ('1', 'true', 'on'):
                    self.default_para = sid

    @staticmethod
    def _read(sp):
        """w:spacing 有 line 但省略 lineRule 时表示 auto (多倍行距)。"""
        if sp is None or (sp.get(w('lineRule')) is None and sp.get(w('line')) is None):
            return None
        return sp.get(w('lineRule')) or 'auto', sp.get(w('line'))

    def line_spacing(self, p):
        pPr = p.find(w('pPr'))
        got = self._read(pPr.find(w('spacing')) if pPr is not None else None)
        if got:
            return got
        ps = pPr.find(w('pStyle')) if pPr is not None else None
        sid = ps.get(w('val')) if ps is not None else self.default_para
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            st = self.styles[sid]
            got = self._read(st.find(f'{w("pPr")}/{w("spacing")}'))
            if got:
                return got
            based = st.find(w('basedOn'))
            sid = based.get(w('val')) if based is not None else None
        return self._read(self.doc_default_spacing) or ('auto', '240')


def fix_line_spacing(p, resolver, modify):
    """行距为“固定值”时高公式会被截掉上下部分：改为“最小值”，保留原行距下限。"""
    rule, line = resolver.line_spacing(p)
    if rule != 'exact':
        return False
    if modify:
        sp = ppr_child(get_ppr(p), 'spacing')
        sp.set(w('lineRule'), 'atLeast')
        sp.set(w('line'), line or '240')
    return True


def doc_uses_grid(root):
    for g in root.iter(w('docGrid')):
        if g.get(w('type')) in ('lines', 'linesAndChars', 'snapToChars'):
            return True
    return False


def fix_snap_to_grid(p, modify):
    """文档设置了网格时，含公式的段落会被“对齐到网格”撑成两倍行高：取消该段对齐网格。"""
    pPr = p.find(w('pPr'))
    node = pPr.find(w('snapToGrid')) if pPr is not None else None
    if node is not None and node.get(w('val')) in ('0', 'false', 'off'):
        return False
    if modify:
        ppr_child(get_ppr(p), 'snapToGrid').set(w('val'), '0')
    return True


def fix_display_center(p, modify):
    """独立公式 (oMathPara) 被设为左/右对齐时改为居中。"""
    changed = False
    for para in p.iter(m('oMathPara')):
        jc = para.find(f'{m("oMathParaPr")}/{m("jc")}')
        if jc is not None and jc.get(m('val')) in ('left', 'right'):
            if modify:
                jc.set(m('val'), 'centerGroup')
            changed = True
    return changed


# ---------------- 公式编号 ----------------
TAIL_NUM_RE = re.compile(r'\s[(（]\s*([0-9]+(?:[.\-－–][0-9]+)*[a-z]?)\s*[)）]\s*$')


def _mrun_text(r):
    return ''.join(t.text or '' for t in r.findall(m('t')))


def _set_mt(r, text):
    ts = r.findall(m('t'))
    for extra in ts[1:]:
        r.remove(extra)
    ts[0].text = text
    if text != text.strip():
        ts[0].set(f'{{{XML_NS}}}space', 'preserve')


def strip_math_number(om, source):
    """从公式内部删除编号文字，返回编号；source='hash' 对应 Word 365 的 #(1) 写法，'tail' 对应末尾空格+(1)。"""
    if source == 'hash':
        full = ''.join(om.itertext())
        mm = HASH_NUM_RE.search(full)
        if not mm:
            return None
        runs = [r for r in om.iter(m('r')) if '#' in _mrun_text(r)]
        if not runs:
            return None
        target = runs[-1]
        parent = next((p for p in om.iter() if target in list(p)), None)
        kids = list(parent)
        for later in kids[kids.index(target) + 1:]:
            if local(later.tag) != 'ctrlPr':
                parent.remove(later)
        text = _mrun_text(target)
        before = text[:text.rindex('#')].rstrip()
        if before:
            _set_mt(target, before)
        else:
            parent.remove(target)
        return mm.group(1)
    if source == 'tail':
        kids = [c for c in om if local(c.tag) != 'ctrlPr']
        tail_runs = []
        for c in reversed(kids):
            if local(c.tag) != 'r':
                break
            tail_runs.insert(0, c)
            tail = ''.join(_mrun_text(r) for r in tail_runs)
            mm = TAIL_NUM_RE.search(tail)
            if mm:
                keep = tail[:mm.start()].rstrip()
                for r in tail_runs:
                    om.remove(r)
                if keep:
                    first = tail_runs[0]
                    _set_mt(first, keep)
                    om.insert(list(om).index(kids[len(kids) - len(tail_runs) - 1]) + 1
                              if len(kids) > len(tail_runs) else 0, first)
                return mm.group(1)
            if len(tail_runs) >= 6:
                break
    return None


def detect_equation(p):
    """判断段落是否为独立公式，以及它已有的编号。
    返回 None 或 {'math': oMath, 'number': '1'|None, 'source': 'text'|'hash'|'tail'|None}。"""
    pPr = p.find(w('pPr'))
    if pPr is not None and pPr.find(w('sectPr')) is not None:
        return None
    maths, texts = own_content(p)
    if len(maths) != 1:
        return None
    om = maths[0]
    other = ''.join(t.text or '' for t in texts).strip()
    if other:
        mm = NUM_RE.match(other)
        if mm:
            return {'math': om, 'number': mm.group(1), 'source': 'text'}
        return None  # 公式和正文混排，是行内公式
    full = ''.join(om.itertext())
    mm = HASH_NUM_RE.search(full)
    if mm:
        return {'math': om, 'number': mm.group(1), 'source': 'hash'}
    mm = TAIL_NUM_RE.search(full)
    if mm and all(local(c.tag) == 'r' for c in list(om)[-1:]):
        return {'math': om, 'number': mm.group(1), 'source': 'tail'}
    return {'math': om, 'number': None, 'source': None}


def text_width(root):
    """正文版心宽度 (twips)。"""
    sect = None
    for s in root.iter(w('sectPr')):
        sect = s
    try:
        pg = sect.find(w('pgSz'))
        mar = sect.find(w('pgMar'))
        width = int(pg.get(w('w'))) - int(mar.get(w('left'))) - int(mar.get(w('right')))
        if width > 2000:
            return width
    except Exception:
        pass
    return 8306  # A4，左右页边距 3.17 cm


def _el(parent, name, **attrs):
    node = ET.SubElement(parent, w(name))
    for k, v in attrs.items():
        node.set(w(k), str(v))
    return node


def build_numbered_table(p, om, number, width):
    """把“公式 + 编号”改成一行三列的无边框表格：中间单元格放独立公式 (保持显示模式)，
    右侧单元格放编号。编号是 SEQ 域，可交叉引用，按 F9 更新。"""
    side = int(width * 0.12)
    mid = width - 2 * side
    tbl = ET.Element(w('tbl'))
    tblPr = _el(tbl, 'tblPr')
    _el(tblPr, 'tblW', w=5000, type='pct')
    _el(tblPr, 'jc', val='center')
    borders = _el(tblPr, 'tblBorders')
    for b in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        _el(borders, b, val='nil')
    _el(tblPr, 'tblLayout', type='fixed')
    mar = _el(tblPr, 'tblCellMar')
    _el(mar, 'left', w=0, type='dxa')
    _el(mar, 'right', w=0, type='dxa')
    _el(tblPr, 'tblLook', val='0000')
    grid = _el(tbl, 'tblGrid')
    for cw in (side, mid, side):
        _el(grid, 'gridCol', w=cw)
    tr = _el(tbl, 'tr')

    orig_pPr = p.find(w('pPr'))

    def cell(cw, jc):
        tc = _el(tr, 'tc')
        tcPr = _el(tc, 'tcPr')
        _el(tcPr, 'tcW', w=cw, type='dxa')
        _el(tcPr, 'vAlign', val='center')
        cp = _el(tc, 'p')
        pPr = copy.deepcopy(orig_pPr) if orig_pPr is not None else ET.Element(w('pPr'))
        for name in ('ind', 'jc', 'numPr', 'tabs', 'framePr'):
            for n in pPr.findall(w(name)):
                pPr.remove(n)
        ppr_child(pPr, 'jc').set(w('val'), jc)
        cp.append(pPr)
        return cp

    cell(side, 'left')
    mid_p = cell(mid, 'center')
    para = next((x for x in p.iter(m('oMathPara')) if om in list(x)), None)
    holder = para if para is not None else ET.Element(m('oMathPara'))
    if para is None:
        parent = next(x for x in p.iter() if om in list(x))
        parent.remove(om)
        holder.append(om)
    else:
        parent = next(x for x in p.iter() if para in list(x))
        parent.remove(para)
    mid_p.append(holder)

    num_p = cell(side, 'right')
    rpr = ET.Element(w('rPr'))
    _el(rpr, 'rFonts', ascii='Times New Roman', hAnsi='Times New Roman', cs='Times New Roman')

    def run(text):
        r = _el(num_p, 'r')
        r.append(copy.deepcopy(rpr))
        _el(r, 't').text = text
        return r

    run('(')
    if number.isdigit():
        fld = _el(num_p, 'fldSimple', instr=' SEQ Equation \\* ARABIC ')
        r = run(number)
        num_p.remove(r)
        fld.append(r)
    else:
        run(number)
    run(')')
    return tbl


class LayoutPlan:
    """分析阶段先算出每段要做的段落级修改；修复阶段 (modify=True) 再统一执行。"""

    def __init__(self, paragraphs, body_root, styles_root, options):
        self.options = options
        self.counts = {'fixLineSpacing': 0, 'fixSnapToGrid': 0, 'fixDisplayCenter': 0,
                       'fixEquationNumbering': 0, 'addEquationNumbers': 0}
        self.para_fixes = {}      # id(p) -> [(key, 名称)]
        self.eq = {}              # id(p) -> {'math','number','source','display_number'}
        self.math_strip = {}      # id(oMath) -> source
        self.body_root = body_root
        resolver = StyleResolver(styles_root)
        use_grid = doc_uses_grid(body_root)
        counter = 0
        for p, label in paragraphs:
            maths, _t = own_content(p)
            if not maths:
                continue
            fixes = []
            if options.get('fixLineSpacing') and fix_line_spacing(p, resolver, False):
                fixes.append(('fixLineSpacing', '公式行距不被截断'))
            if options.get('fixSnapToGrid') and use_grid and fix_snap_to_grid(p, False):
                fixes.append(('fixSnapToGrid', '取消公式段落对齐网格'))
            if options.get('fixDisplayCenter') and fix_display_center(p, False):
                fixes.append(('fixDisplayCenter', '独立公式居中'))
            if not label:  # 只给正文的独立公式编号，页眉页脚脚注不编号
                info = self._plan_number(p)
                if info:
                    counter += 1
                    if options.get('addEquationNumbers'):
                        info['display_number'] = str(counter)
                    fixes.append((info['key'], '公式编号右对齐' if info['key'] == 'fixEquationNumbering'
                                  else '自动添加公式编号'))
                    self.eq[id(p)] = info
                    if info['source'] in ('hash', 'tail'):
                        self.math_strip[id(info['math'])] = info['source']
            for key, _name in fixes:
                self.counts[key] += 1
            if fixes:
                self.para_fixes[id(p)] = fixes

    def _plan_number(self, p):
        info = detect_equation(p)
        if not info:
            return None
        parent_ok = True
        if info['number'] is not None and self.options.get('fixEquationNumbering'):
            info['key'] = 'fixEquationNumbering'
            info['display_number'] = info['number']
            return info if parent_ok else None
        if info['number'] is None and self.options.get('addEquationNumbers'):
            # 只有段落里仅有这一个公式时才算独立公式
            info['key'] = 'addEquationNumbers'
            return info
        if info['number'] is not None and self.options.get('addEquationNumbers'):
            info['key'] = 'addEquationNumbers'
            return info
        return None


    def apply(self, paragraphs, styles_root):
        """修复阶段执行段落级修改。"""
        resolver = StyleResolver(styles_root)
        width = text_width(self.body_root)
        parent_of = {}
        for node in self.body_root.iter():
            for c in node:
                if c.tag == w('p'):
                    parent_of[id(c)] = node
        for p, _label in paragraphs:
            for key, _name in self.para_fixes.get(id(p), []):
                if key == 'fixLineSpacing':
                    fix_line_spacing(p, resolver, True)
                elif key == 'fixSnapToGrid':
                    fix_snap_to_grid(p, True)
                elif key == 'fixDisplayCenter':
                    fix_display_center(p, True)
            info = self.eq.get(id(p))
            parent = parent_of.get(id(p))
            if not info or parent is None:
                continue
            if local(parent.tag) not in ('body', 'tc', 'sdtContent', 'txbxContent'):
                continue
            tbl = build_numbered_table(p, info['math'], info['display_number'], width)
            idx = list(parent).index(p)
            parent.remove(p)
            parent.insert(idx, tbl)
            # 表格单元格不能是容器最后一个元素 (tc 末尾必须是段落)
            if local(parent.tag) == 'tc' and idx == len(parent) - 1:
                parent.append(ET.Element(w('p')))
