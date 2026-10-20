import sys
import io
import os
import json
import copy
import zipfile
import argparse
import xml.etree.ElementTree as ET

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

MATH_FUNCS = (
    'arcsin', 'arccos', 'arctan', 'arcsec', 'arccsc', 'arccot',
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'ln', 'log', 'lg', 'exp',
    'lim', 'max', 'min', 'sup', 'inf', 'det', 'tr', 'sgn'
)

def escape_latex_text(text):
    if not text:
        return ""
    replacements = [
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
        rPr = elem.find(f'{{{OMML_NS}}}rPr')
        sty = rPr.find(f'{{{OMML_NS}}}sty') if rPr is not None else None
        val = sty.attrib.get(f'{{{OMML_NS}}}val') if sty is not None else ''
        
        # Plain upright text or function
        if val == 'p':
            if t_text in MATH_FUNCS:
                return f"\\{t_text} "
            elif t_text in ('d', 'e', 'T', 'r', 'o', 'E', 'O'):
                return f"\\mathrm{{{t_text}}}"
            elif t_text in ('=', '+', '-', '−', '<', '>', '≤', '≥', '≠', '±', '×', '·', '→', '∈'):
                return f" {escape_latex_text(t_text)} "
            elif t_text.isdigit():
                return t_text
            elif t_text in ('(', ')', '[', ']', '{', '}', ',', '.', ';', ':', '|', '⋯', '∞', '′', '″'):
                return escape_latex_text(t_text)
            return f"\\text{{{escape_latex_text(t_text)}}}"
            
        # If contains Chinese characters, wrap in \text{} for KaTeX stability
        if any('\u4e00' <= char <= '\u9fff' for char in t_text):
            return f"\\text{{{escape_latex_text(t_text)}}}"
            
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
        pre = ""
        if sub_str: pre += f"_{{{sub_str}}}"
        if sup_str: pre += f"^{{{sup_str}}}"
        return f"{{{pre}{e_str}}}"
        
    elif tag == 'f':
        fPr = elem.find(f'{{{OMML_NS}}}fPr')
        fType = fPr.find(f'{{{OMML_NS}}}type') if fPr is not None else None
        fType_val = fType.attrib.get(f'{{{OMML_NS}}}val', 'bar') if fType is not None else 'bar'
        num = elem.find(f'{{{OMML_NS}}}num')
        den = elem.find(f'{{{OMML_NS}}}den')
        num_str = omml_to_latex(num) if num is not None else ''
        den_str = omml_to_latex(den) if den is not None else ''
        if fType_val == 'noBar':
            return f"\\binom{{{num_str}}}{{{den_str}}}"
        elif fType_val == 'skw':
            return f"{num_str}/{den_str}"
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
            if c in ('⟨', '<'): return '\\langle '
            if c in ('⟩', '>'): return '\\rangle '
            if c in ('/', '\\', '⌊', '⌋', '⌈', '⌉'): return c
            return '.'
            
        beg_latex = map_delim(beg)
        end_latex = map_delim(end)
        
        sep_str = (sep if sep is not None else ', ')
        if sep_str == '':
            inner = ' '.join(e_list)
        else:
            inner = f'{sep_str} '.join(e_list)
            
        return f"\\left{beg_latex} {inner} \\right{end_latex}"
        
    elif tag == 'm':
        rows = []
        for mr in elem.findall(f'{{{OMML_NS}}}mr'):
            cells = [omml_to_latex(e).strip() for e in mr.findall(f'{{{OMML_NS}}}e')]
            rows.append(' & '.join(cells))
        return "\\begin{matrix} " + ' \\\\ '.join(rows) + " \\end{matrix}"
        
    elif tag == 'eqArr':
        lines = [omml_to_latex(e).strip() for e in elem.findall(f'{{{OMML_NS}}}e')]
        return "\\begin{cases} " + ' \\\\ '.join(lines) + " \\end{cases}"
        
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
        
        cmd = '\\int'
        if chr_val in ('∑', 'sum'): cmd = '\\sum'
        elif chr_val in ('∏', 'prod'): cmd = '\\prod'
        elif chr_val == '∬': cmd = '\\iint'
        elif chr_val == '∭': cmd = '\\iiint'
        elif chr_val == '∮': cmd = '\\oint'
        
        sub_str = f"_{{{omml_to_latex(sub)}}}" if sub is not None and omml_to_latex(sub).strip() else ""
        sup_str = f"^{{{omml_to_latex(sup)}}}" if sup is not None and omml_to_latex(sup).strip() else ""
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
        return f"{{{e_str}}}_{{{lim_str}}} "
        
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
        return f"{{{e_str}}}^{{{sup_str}}} "
        
    elif tag == 'bar':
        e = elem.find(f'{{{OMML_NS}}}e')
        barPr = elem.find(f'{{{OMML_NS}}}barPr')
        pos = 'top'
        if barPr is not None:
            pos_elem = barPr.find(f'{{{OMML_NS}}}pos')
            if pos_elem is not None:
                pos = pos_elem.attrib.get(f'{{{OMML_NS}}}val', 'top')
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
            '́': '\\acute', '̀': '\\grave'
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
        if pos == 'top' or chr_val == '⏞':
            return f"\\overbrace{{{e_str}}}"
        return f"\\underbrace{{{e_str}}}"
        
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
        
    else:
        return ''.join(omml_to_latex(c) for c in elem)


class DocxMathRepairEngine:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.xml_content = None
        self.root = None
        self.load()

    def load(self):
        with zipfile.ZipFile(self.docx_path, 'r') as z:
            self.xml_content = z.read('word/document.xml')
        self.root = ET.fromstring(self.xml_content)

    def save(self, output_path):
        """Save the updated document back into a new docx archive with safe XML tree writing."""
        new_xml = ET.tostring(self.root, encoding='utf-8', xml_declaration=True)
        tmp_output = output_path + '.tmp_writing.docx'
        
        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp_output, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    if item.filename == 'word/document.xml':
                        zout.writestr(item, new_xml)
                    elif item.filename == 'word/theme/theme1.xml':
                        try:
                            theme_root = ET.fromstring(content)
                            for latin in theme_root.iter(f'{{{DRAWING_A_NS}}}latin'):
                                latin.set('typeface', 'Times New Roman')
                            content = ET.tostring(theme_root, encoding='utf-8', xml_declaration=True)
                        except Exception:
                            pass
                        zout.writestr(item, content)
                    elif item.filename in ('word/styles.xml', 'word/stylesWithEffects.xml'):
                        try:
                            styles_root = ET.fromstring(content)
                            for rFonts in styles_root.iter(f'{{{WORD_NS}}}rFonts'):
                                rFonts.set(f'{{{WORD_NS}}}ascii', 'Times New Roman')
                                rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Times New Roman')
                                rFonts.set(f'{{{WORD_NS}}}cs', 'Times New Roman')
                            content = ET.tostring(styles_root, encoding='utf-8', xml_declaration=True)
                        except Exception:
                            pass
                        zout.writestr(item, content)
                    elif item.filename == 'word/settings.xml':
                        try:
                            settings_root = ET.fromstring(content)
                            mathPr = settings_root.find(f'{{{OMML_NS}}}mathPr')
                            if mathPr is None:
                                mathPr = ET.SubElement(settings_root, f'{{{OMML_NS}}}mathPr')
                            mathFont = mathPr.find(f'{{{OMML_NS}}}mathFont')
                            if mathFont is None:
                                mathFont = ET.SubElement(mathPr, f'{{{OMML_NS}}}mathFont')
                            mathFont.set(f'{{{OMML_NS}}}val', 'Times New Roman')
                            content = ET.tostring(settings_root, encoding='utf-8', xml_declaration=True)
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
        glued_pairs = [
            ('lncos', ['ln', 'cos']),
            ('sincos', ['sin', 'cos']),
            ('cossin', ['cos', 'sin']),
            ('sint', ['sin', 't']),
            ('cost', ['cos', 't']),
            ('tant', ['tan', 't']),
            ('arctant', ['arctan', 't']),
            ('lnx', ['ln', 'x']),
            ('sinx', ['sin', 'x']),
            ('cosx', ['cos', 'x']),
            ('tanx', ['tan', 'x']),
            ('secx', ['sec', 'x']),
            ('cotx', ['cot', 'x'])
        ]
        
        runs = list(omath.iter(f'{{{OMML_NS}}}r'))
        for r in runs:
            t_elem = r.find(f'{{{OMML_NS}}}t')
            if t_elem is None or not t_elem.text:
                continue
            text = t_elem.text
            
            for token, parts in glued_pairs:
                if token in text:
                    if text == token:
                        parent = None
                        for p in omath.iter():
                            if r in list(p):
                                parent = p
                                break
                        if parent is not None:
                            idx = list(parent).index(r)
                            parent.remove(r)
                            for part in parts:
                                new_r = ET.Element(f'{{{OMML_NS}}}r')
                                if part in ('sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'ln', 'log', 'arctan', 'exp'):
                                    rPr = ET.SubElement(new_r, f'{{{OMML_NS}}}rPr')
                                    sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                                    sty.set(f'{{{OMML_NS}}}val', 'p')
                                new_t = ET.SubElement(new_r, f'{{{OMML_NS}}}t')
                                new_t.text = part
                                parent.insert(idx, new_r)
                                idx += 1
                            modified = True
                            break
        return modified

    def apply_rule_isolated_delimiters(self, omath):
        """修复括号上标结构 (如 (式子)^n)"""
        modified = False
        DELIM_PAIRS = {
            ')': '(',
            ']': '[',
            '}': '{'
        }
        for sSup in list(omath.iter(f'{{{OMML_NS}}}sSup')):
            e = sSup.find(f'{{{OMML_NS}}}e')
            if e is None:
                continue
            e_text = ''.join(e.itertext()).strip()
            if e_text not in DELIM_PAIRS:
                continue
            
            beg_char = DELIM_PAIRS[e_text]
            close_char = e_text
            
            parent = None
            for p in omath.iter():
                if sSup in list(p):
                    parent = p
                    break
            if parent is None:
                continue
                
            siblings = list(parent)
            sSup_pos = siblings.index(sSup)
            
            # 使用括号平衡深度计数器（Nesting Depth Balance）从右向左寻找与之配对的开括号
            match_idx = -1
            match_pos = -1
            depth = 0  # 闭括号嵌套计数
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
            inner_items = []
            
            t_elem = match_elem.find(f'.//{{{OMML_NS}}}t')
            prefix_exists = False
            if t_elem is not None and t_elem.text:
                orig_text = t_elem.text
                prefix_text = orig_text[:match_pos]
                suffix_text = orig_text[match_pos + 1:]
                
                if prefix_text:
                    t_elem.text = prefix_text
                    prefix_exists = True
                else:
                    parent.remove(match_elem)
                    
                if suffix_text:
                    new_r = ET.Element(f'{{{OMML_NS}}}r')
                    new_t = ET.SubElement(new_r, f'{{{OMML_NS}}}t')
                    new_t.text = suffix_text
                    inner_items.append(new_r)
            else:
                parent.remove(match_elem)
                
            current_siblings = list(parent)
            current_sSup_pos = current_siblings.index(sSup)
            start_idx = (match_idx + 1) if prefix_exists else match_idx
            middle_items = current_siblings[start_idx:current_sSup_pos]
            
            for item in middle_items:
                parent.remove(item)
                inner_items.append(item)
                
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
            chr_elem = nary.find(f'.//{{{OMML_NS}}}chr')
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '∫') if chr_elem is not None else '∫'
            if chr_val not in ('∫', '∬', '∭', '∮'):
                continue
                
            e = nary.find(f'{{{OMML_NS}}}e')
            if e is None:
                e = ET.SubElement(nary, f'{{{OMML_NS}}}e')
                
            e_text = ''.join(e.itertext()).strip()
            e_children = list(e)
            
            if not e_text and len(e_children) == 0:
                parent = None
                for p in omath.iter():
                    if nary in list(p):
                        parent = p
                        break
                if parent is None:
                    continue
                    
                siblings = list(parent)
                pos = siblings.index(nary)
                following_siblings = siblings[pos + 1:]
                
                if not following_siblings:
                    continue
                    
                operand_items = []
                for sib in following_siblings:
                    if sib.tag.endswith('r'):
                        t_elem = sib.find(f'.//{{{OMML_NS}}}t')
                        if t_elem is not None and t_elem.text and '=' in t_elem.text:
                            eq_idx = t_elem.text.index('=')
                            before_eq = t_elem.text[:eq_idx]
                            after_eq = t_elem.text[eq_idx:]
                            if before_eq:
                                new_r = ET.Element(f'{{{OMML_NS}}}r')
                                new_t = ET.SubElement(new_r, f'{{{OMML_NS}}}t')
                                new_t.text = before_eq
                                operand_items.append(new_r)
                            t_elem.text = after_eq
                            break
                    
                    sib_text = ''.join(sib.itertext())
                    if '=' in sib_text:
                        break
                        
                    parent.remove(sib)
                    operand_items.append(sib)
                    
                if operand_items:
                    for item in operand_items:
                        e.append(item)
                    modified = True
                    
        return modified

    def apply_rule_integral_size(self, omath):
        """固定积分号高度 (防止拉伸)"""
        modified = False
        for nary in omath.iter(f'{{{OMML_NS}}}nary'):
            naryPr = nary.find(f'{{{OMML_NS}}}naryPr')
            chr_elem = naryPr.find(f'{{{OMML_NS}}}chr') if naryPr is not None else None
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '∫') if chr_elem is not None else '∫'
            
            if chr_val in ('∫', '∬', '∭', '∮'):
                if naryPr is not None:
                    grow = naryPr.find(f'{{{OMML_NS}}}grow')
                    if grow is not None:
                        if grow.attrib.get(f'{{{OMML_NS}}}val') != '0':
                            grow.set(f'{{{OMML_NS}}}val', '0')
                            modified = True
                    else:
                        grow = ET.SubElement(naryPr, f'{{{OMML_NS}}}grow')
                        grow.set(f'{{{OMML_NS}}}val', '0')
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
                if e_text == 'lim' or 'lim' in e_text:
                    parent = None
                    for p in omath.iter():
                        if sSub in list(p):
                            parent = p
                            break
                    if parent is not None:
                        idx = list(parent).index(sSub)
                        parent.remove(sSub)
                        
                        limLow = ET.Element(f'{{{OMML_NS}}}limLow')
                        new_e = ET.SubElement(limLow, f'{{{OMML_NS}}}e')
                        new_r = ET.SubElement(new_e, f'{{{OMML_NS}}}r')
                        rPr = ET.SubElement(new_r, f'{{{OMML_NS}}}rPr')
                        sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                        sty.set(f'{{{OMML_NS}}}val', 'p')
                        new_t = ET.SubElement(new_r, f'{{{OMML_NS}}}t')
                        new_t.text = 'lim'
                        
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
            chr_elem = nary.find(f'.//{{{OMML_NS}}}chr')
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '') if chr_elem is not None else ''
            
            if chr_val in ('∑', '∏', 'sum', 'prod'):
                parent = None
                for p in omath.iter():
                    if nary in list(p):
                        parent = p
                        break
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
                    
                sym_char = '∑' if chr_val in ('∑', 'sum') else '∏'
                
                def create_sym_r():
                    r_sym = ET.Element(f'{{{OMML_NS}}}r')
                    rPr_w = ET.SubElement(r_sym, f'{{{WORD_NS}}}rPr')
                    sz = ET.SubElement(rPr_w, f'{{{WORD_NS}}}sz')
                    sz.set(f'{{{WORD_NS}}}val', '28')
                    szCs = ET.SubElement(rPr_w, f'{{{WORD_NS}}}szCs')
                    szCs.set(f'{{{WORD_NS}}}val', '28')
                    rPr_m = ET.SubElement(r_sym, f'{{{OMML_NS}}}rPr')
                    sty = ET.SubElement(rPr_m, f'{{{OMML_NS}}}sty')
                    sty.set(f'{{{OMML_NS}}}val', 'p')
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
                        if child.tag.endswith('r') and not (''.join(child.itertext()).strip()):
                            continue
                        parent.insert(idx + insert_offset, child)
                        insert_offset += 1
                        
                modified = True
        return modified

    def apply_rule_double_integral_region(self, omath):
        """二重积分区域居中 (∬)"""
        modified = False
        for nary in list(omath.iter(f'{{{OMML_NS}}}nary')):
            chr_elem = nary.find(f'.//{{{OMML_NS}}}chr')
            chr_val = chr_elem.attrib.get(f'{{{OMML_NS}}}val', '∫') if chr_elem is not None else '∫'
            
            if chr_val in ('∬', '∭'):
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
                        if c.tag.endswith('r') and not (''.join(c.itertext()).strip()):
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
                runs = list(e.iter(f'{{{OMML_NS}}}r'))
                if runs:
                    last_t = runs[-1].find(f'{{{OMML_NS}}}t')
                    if last_t is not None and last_t.text and last_t.text.endswith('|'):
                        e_all_text = ''.join(e.itertext()).strip()
                        if e_all_text == '|':
                            continue
                            
                        if last_t.text == '|':
                            for p in e.iter():
                                if runs[-1] in list(p):
                                    p.remove(runs[-1])
                                    break
                        else:
                            last_t.text = last_t.text[:-1]
                        
                        inner_children = list(e)
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
            r'(?P<func>' + '|'.join(sorted(MATH_FUNCS, key=lambda x: -len(x))) + r')|' +
            r'(?P<diff>d[xytzuvwrθσρφ])|' +
            r'(?P<cjk>[\u4e00-\u9fa5]+)|' +
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
            r_elem = ET.Element(f'{{{OMML_NS}}}r')
            if orig_r is not None:
                w_rPr = orig_r.find(f'{{{WORD_NS}}}rPr')
                if w_rPr is not None:
                    r_elem.append(copy.deepcopy(w_rPr))
            if is_plain:
                rPr = r_elem.find(f'{{{OMML_NS}}}rPr')
                if rPr is None:
                    rPr = ET.SubElement(r_elem, f'{{{OMML_NS}}}rPr')
                sty = rPr.find(f'{{{OMML_NS}}}sty')
                if sty is None:
                    sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                sty.set(f'{{{OMML_NS}}}val', 'p')
            t_elem = ET.SubElement(r_elem, f'{{{OMML_NS}}}t')
            t_elem.text = text
            return r_elem

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
                        if t is not None and t.text == 'd':
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

            rPr = r.find(f'{{{OMML_NS}}}rPr')
            sty = rPr.find(f'{{{OMML_NS}}}sty') if rPr is not None else None
            is_plain = (sty is not None and sty.attrib.get(f'{{{OMML_NS}}}val') == 'p')

            if is_plain and text in MATH_FUNCS:
                continue
            if is_plain and text in ('∬', '∭', '∑', '∏', '∫'):
                continue
            if is_plain and text.isdigit():
                continue

            parent = None
            for p in omath.iter():
                if r in list(p):
                    parent = p
                    break
            if parent is None:
                continue

            parent_tag = parent.tag.split('}')[-1]
            is_in_sup = (parent_tag == 'sup')

            # 纯数字正体
            if text.isdigit() and not is_plain:
                if rPr is None:
                    rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                    r.insert(0, rPr)
                if sty is None:
                    sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                sty.set(f'{{{OMML_NS}}}val', 'p')
                modified = True
                continue

            needs_split = False
            if len(text) > 1 and text not in MATH_FUNCS:
                needs_split = True
            elif text in ('<', '>', '=', '≤', '≥', '≠', '+', '-', '−', '±', '×', '·', ',', '.', ';', ':', '|', '/', '(', ')', '[', ']', '{', '}', '∞', 'ℝ', '⋯', '∣', '′', '″', '∂', '∫', '*', '（', '）'):
                if not is_plain:
                    if rPr is None:
                        rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                        r.insert(0, rPr)
                    if sty is None:
                        sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                    sty.set(f'{{{OMML_NS}}}val', 'p')
                    modified = True
                    continue
            elif text == 'e' and not is_plain:
                if rPr is None:
                    rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                    r.insert(0, rPr)
                if sty is None:
                    sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                sty.set(f'{{{OMML_NS}}}val', 'p')
                modified = True
                continue
            elif text == 'd' and not is_plain:
                if rPr is None:
                    rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                    r.insert(0, rPr)
                if sty is None:
                    sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                sty.set(f'{{{OMML_NS}}}val', 'p')
                modified = True
                continue

            if needs_split:
                idx = list(parent).index(r)
                parent.remove(r)

                new_tokens = []
                for m in TOKEN_REGEX.finditer(text):
                    kind = m.lastgroup
                    val = m.group()

                    if kind == 'digit':
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
                        elif val == 'd':
                            new_tokens.append(('d', True))
                        else:
                            new_tokens.append((val, False))

                for tok_val, tok_plain in new_tokens:
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
            if t is not None and t.text:
                if '{' in t.text or '}' in t.text:
                    rPr = r.find(f'{{{OMML_NS}}}rPr')
                    if rPr is None:
                        rPr = ET.Element(f'{{{OMML_NS}}}rPr')
                        r.insert(0, rPr)
                    sty = rPr.find(f'{{{OMML_NS}}}sty')
                    if sty is None:
                        sty = ET.SubElement(rPr, f'{{{OMML_NS}}}sty')
                    if sty.attrib.get(f'{{{OMML_NS}}}val') != 'p':
                        sty.set(f'{{{OMML_NS}}}val', 'p')
                        modified = True

        for d in omath.iter(f'{{{OMML_NS}}}d'):
            dPr = d.find(f'{{{OMML_NS}}}dPr')
            if dPr is not None:
                beg = dPr.find(f'{{{OMML_NS}}}begChr')
                end = dPr.find(f'{{{OMML_NS}}}endChr')
                beg_val = beg.attrib.get(f'{{{OMML_NS}}}val', '') if beg is not None else ''
                end_val = end.attrib.get(f'{{{OMML_NS}}}val', '') if end is not None else ''
                if beg_val in ('{', '｛') or end_val in ('}', '｝'):
                    grow = dPr.find(f'{{{OMML_NS}}}grow')
                    if grow is None:
                        grow = ET.SubElement(dPr, f'{{{OMML_NS}}}grow')
                        grow.set(f'{{{OMML_NS}}}val', 'on')
                        modified = True
                    elif grow.attrib.get(f'{{{OMML_NS}}}val') != 'on':
                        grow.set(f'{{{OMML_NS}}}val', 'on')
                        modified = True
        return modified

    def apply_rule_font_times_new_roman(self, omath):
        """统一公式英文字体为 Times New Roman"""
        modified = False
        for r in omath.iter(f'{{{OMML_NS}}}r'):
            w_rPr = r.find(f'{{{WORD_NS}}}rPr')
            if w_rPr is None:
                w_rPr = ET.Element(f'{{{WORD_NS}}}rPr')
                r.insert(0, w_rPr)
            rFonts = w_rPr.find(f'{{{WORD_NS}}}rFonts')
            if rFonts is None:
                rFonts = ET.SubElement(w_rPr, f'{{{WORD_NS}}}rFonts')
            if rFonts.attrib.get(f'{{{WORD_NS}}}ascii') != 'Times New Roman':
                rFonts.set(f'{{{WORD_NS}}}ascii', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}cs', 'Times New Roman')
                modified = True
        return modified

    # ================== Extraction & Process ==================

    def process(self, options, modify=False):
        """
        Analyze or repair the document.
        Returns document paragraphs (for full text view) and issuesList (for comparison view).
        """
        paragraphs = self.root.findall(f'.//{{{WORD_NS}}}p')
        total_omath = len(self.root.findall(f'.//{{{OMML_NS}}}oMath'))
        
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
            'fixFontTimesNewRoman': 0
        }
        
        issues_list = []
        issue_id = 1
        math_counter = 1
        
        document_paragraphs = []

        for p_idx, p in enumerate(paragraphs):
            p_text = ''.join(p.itertext()).strip()
            
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
                elif tag == 'r':
                    for sub in elem:
                        sub_tag = sub.tag.split('}')[-1]
                        if sub_tag == 't':
                            if sub.text:
                                elements.append({'type': 'text', 'text': sub.text})
                        elif sub_tag == 'br':
                            elements.append({'type': 'br'})
                        elif sub_tag == 'tab':
                            elements.append({'type': 'text', 'text': '    '})
                elif tag in ('hyperlink', 'sdt', 'sdtContent', 'smartTag', 'fldSimple', 'ins'):
                    for sub in elem:
                        walk_paragraph(sub)

            def process_math_node(om_node, is_display):
                nonlocal issue_id, math_counter
                om_clone = copy.deepcopy(om_node)
                before_latex = omml_to_latex(om_clone).strip()
                before_text = ''.join(om_clone.itertext()).strip()
                
                changed_rules = []
                target_om = om_node if modify else om_clone
                
                if options.get('splitGluedTokens', True):
                    if self.apply_rule_split_tokens(target_om):
                        issues_found['splitGluedTokens'] += 1
                        changed_rules.append(('splitGluedTokens', '拆分连写函数', '拆分连写的 lncos、sincos 等为独立函数'))
                
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
                
                if options.get('fixSetBraces', True):
                    if self.apply_rule_set_braces(target_om):
                        issues_found['fixSetBraces'] += 1
                        changed_rules.append(('fixSetBraces', '集合大括号修复', '规范集合大括号定界符'))

                if options.get('fixFontTimesNewRoman', True):
                    if self.apply_rule_font_times_new_roman(target_om):
                        issues_found['fixFontTimesNewRoman'] += 1
                        changed_rules.append(('fixFontTimesNewRoman', '公式字体设为 Times New Roman', '将公式中的英文字母和数字字体设为 Times New Roman'))
                
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
            
            document_paragraphs.append({
                'index': p_idx,
                'snippet': p_text[:100] + ('...' if len(p_text) > 100 else ''),
                'hasMath': has_math,
                'hasModified': has_modified,
                'elements': elements
            })

        # 修复模式时，统一正文英文字体为 Times New Roman
        if modify and options.get('fixFontTimesNewRoman', True):
            for r in self.root.findall(f'.//{{{WORD_NS}}}r'):
                w_rPr = r.find(f'{{{WORD_NS}}}rPr')
                if w_rPr is None:
                    w_rPr = ET.Element(f'{{{WORD_NS}}}rPr')
                    r.insert(0, w_rPr)
                rFonts = w_rPr.find(f'{{{WORD_NS}}}rFonts')
                if rFonts is None:
                    rFonts = ET.SubElement(w_rPr, f'{{{WORD_NS}}}rFonts')
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
            actual_saved = engine.save(args.output)
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
