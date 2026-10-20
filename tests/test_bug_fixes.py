# -*- coding: utf-8 -*-
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import unittest
import xml.etree.ElementTree as ET
from engine import DocxMathRepairEngine, omml_to_latex, OMML_NS, WORD_NS

class TestBugFixes(unittest.TestCase):
    def setUp(self):
        # 创建一个虚拟最小 engine 实例（不加载实际磁盘文件）
        self.engine = DocxMathRepairEngine.__new__(DocxMathRepairEngine)
        self.engine.docx_path = "dummy.docx"
        self.engine.root = ET.Element(f'{{{WORD_NS}}}document')

    def test_nested_brackets_in_isolated_delimiters(self):
        """测试嵌套括号结构 (a + (b + c))^2 不会被误截断为 (b + c)^2"""
        # 构造 OMML 片段：
        # r: '('
        # r: 'a + '
        # r: '('
        # r: 'b + c'
        # r: ')'
        # sSup: (e: r(')'), sup: r('2'))
        omath = ET.Element(f'{{{OMML_NS}}}oMath')
        
        def add_r(parent, text):
            r = ET.SubElement(parent, f'{{{OMML_NS}}}r')
            t = ET.SubElement(r, f'{{{OMML_NS}}}t')
            t.text = text
            return r

        add_r(omath, '(')
        add_r(omath, 'a + ')
        add_r(omath, '(')
        add_r(omath, 'b + c')
        add_r(omath, ')')
        
        sSup = ET.SubElement(omath, f'{{{OMML_NS}}}sSup')
        e = ET.SubElement(sSup, f'{{{OMML_NS}}}e')
        add_r(e, ')')
        sup = ET.SubElement(sSup, f'{{{OMML_NS}}}sup')
        add_r(sup, '2')

        # 执行修复
        mod = self.engine.apply_rule_isolated_delimiters(omath)
        self.assertTrue(mod, "应该触发括号上标修复")

        # 验证修复后，底座中的 d 元素内应该包含完整的内容（包含 a + (b + c)）
        d_elem = sSup.find(f'.//{{{OMML_NS}}}d')
        self.assertIsNotNone(d_elem, "sSup 的 e 应该包含 d 元素")
        
        d_text = ''.join(d_elem.itertext())
        self.assertIn("a +", d_text, "外层括号内的 'a +' 必须包含在底座中，不能被漏掉")
        self.assertIn("b + c", d_text, "内层括号内的 'b + c' 也必须包含在底座中")

    def test_nary_lim_loc_no_sup_creates_no_empty_lim(self):
        """测试仅有下标的求和符号不会产生空的 limUpp，杜绝虚线框"""
        omath = ET.Element(f'{{{OMML_NS}}}oMath')
        nary = ET.SubElement(omath, f'{{{OMML_NS}}}nary')
        naryPr = ET.SubElement(nary, f'{{{OMML_NS}}}naryPr')
        chr_node = ET.SubElement(naryPr, f'{{{OMML_NS}}}chr')
        chr_node.set(f'{{{OMML_NS}}}val', '∑')
        
        sub = ET.SubElement(nary, f'{{{OMML_NS}}}sub')
        r_sub = ET.SubElement(sub, f'{{{OMML_NS}}}r')
        t_sub = ET.SubElement(r_sub, f'{{{OMML_NS}}}t')
        t_sub.text = 'i=1'
        
        # 没有 sup
        mod = self.engine.apply_rule_nary_lim_loc(omath)
        self.assertTrue(mod)

        # 检查是否只生成了 limLow，而没有空的 limUpp
        limUpp = omath.find(f'.//{{{OMML_NS}}}limUpp')
        self.assertIsNone(limUpp, "只有下标时绝不应该生成 limUpp，防止产生空虚线框")
        
        limLow = omath.find(f'.//{{{OMML_NS}}}limLow')
        self.assertIsNotNone(limLow, "应该生成 limLow")
        
        lim_sub = limLow.find(f'{{{OMML_NS}}}lim')
        self.assertIsNotNone(lim_sub)
        self.assertEqual(''.join(lim_sub.itertext()).strip(), 'i=1')

    def test_chinese_text_in_typography(self):
        """测试公式中包含中文字符时保持正体且不被拆散"""
        omath = ET.Element(f'{{{OMML_NS}}}oMath')
        r = ET.SubElement(omath, f'{{{OMML_NS}}}r')
        t = ET.SubElement(r, f'{{{OMML_NS}}}t')
        t.text = '当x大于0时'
        
        mod = self.engine.apply_rule_typography(omath)
        self.assertTrue(mod)
        
        # 检查所有的 runs
        runs = omath.findall(f'.//{{{OMML_NS}}}r')
        # 中文汉字应该作为词块保持正体 (sty val="p")
        chinese_runs = [run for run in runs if any('\u4e00' <= c <= '\u9fa5' for c in ''.join(run.itertext()))]
        self.assertTrue(len(chinese_runs) > 0)
        for c_run in chinese_runs:
            rPr = c_run.find(f'{{{OMML_NS}}}rPr')
            self.assertIsNotNone(rPr, "中文字符应该有 rPr")
            sty = rPr.find(f'{{{OMML_NS}}}sty')
            self.assertIsNotNone(sty, "中文字符应该设置 sty")
            self.assertEqual(sty.attrib.get(f'{{{OMML_NS}}}val'), 'p', "中文字符必须是正体 (p)")

    def test_set_braces_rule(self):
        """测试集合大括号修复"""
        omath = ET.Element(f'{{{OMML_NS}}}oMath')
        d = ET.SubElement(omath, f'{{{OMML_NS}}}d')
        dPr = ET.SubElement(d, f'{{{OMML_NS}}}dPr')
        beg = ET.SubElement(dPr, f'{{{OMML_NS}}}begChr')
        beg.set(f'{{{OMML_NS}}}val', '{')
        end = ET.SubElement(dPr, f'{{{OMML_NS}}}endChr')
        end.set(f'{{{OMML_NS}}}val', '}')
        
        # 没有 grow 或 grow="0"
        mod = self.engine.apply_rule_set_braces(omath)
        self.assertTrue(mod, "应该修复集合大括号")
        
        grow = dPr.find(f'{{{OMML_NS}}}grow')
        self.assertIsNotNone(grow)
        self.assertEqual(grow.attrib.get(f'{{{OMML_NS}}}val'), 'on')

    def test_omml_to_latex_chinese_and_delims(self):
        """测试 LaTeX 转换中对中文与全角定界符的安全处理"""
        # 测试全角括号定界符
        d = ET.Element(f'{{{OMML_NS}}}d')
        dPr = ET.SubElement(d, f'{{{OMML_NS}}}dPr')
        beg = ET.SubElement(dPr, f'{{{OMML_NS}}}begChr')
        beg.set(f'{{{OMML_NS}}}val', '（')
        end = ET.SubElement(dPr, f'{{{OMML_NS}}}endChr')
        end.set(f'{{{OMML_NS}}}val', '）')
        e = ET.SubElement(d, f'{{{OMML_NS}}}e')
        r = ET.SubElement(e, f'{{{OMML_NS}}}r')
        t = ET.SubElement(r, f'{{{OMML_NS}}}t')
        t.text = 'x+1'
        
        latex = omml_to_latex(d)
        self.assertEqual(latex, r"\left( x+1 \right)")

        # 测试中文文本转义
        r_cn = ET.Element(f'{{{OMML_NS}}}r')
        t_cn = ET.SubElement(r_cn, f'{{{OMML_NS}}}t')
        t_cn.text = '其中x为实数'
        latex_cn = omml_to_latex(r_cn)
        self.assertTrue(latex_cn.startswith(r"\text{") and latex_cn.endswith("}"))

if __name__ == '__main__':
    unittest.main()
