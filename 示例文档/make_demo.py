# -*- coding: utf-8 -*-
"""生成“公式修复演示.docx”：每段一种典型 OMML 排版问题。"""
import os, zipfile

def R(t, sty=None, nor=False):
    pr = ''
    if sty or nor:
        pr = '<m:rPr>' + ('<m:nor/>' if nor else '') + (f'<m:sty m:val="{sty}"/>' if sty else '') + '</m:rPr>'
    return f'<m:r>{pr}<m:t xml:space="preserve">{t}</m:t></m:r>'

def nary(ch, sub, sup, e='', grow='1'):
    return (f'<m:nary><m:naryPr><m:chr m:val="{ch}"/><m:limLoc m:val="subSup"/><m:grow m:val="{grow}"/>'
            f'<m:subHide m:val="off"/><m:supHide m:val="off"/></m:naryPr>'
            f'<m:sub>{sub}</m:sub><m:sup>{sup}</m:sup><m:e>{e}</m:e></m:nary>')

def sSup(e, s): return f'<m:sSup><m:e>{e}</m:e><m:sup>{s}</m:sup></m:sSup>'
def sSub(e, s): return f'<m:sSub><m:e>{e}</m:e><m:sub>{s}</m:sub></m:sSub>'
def frac(n, d): return f'<m:f><m:num>{n}</m:num><m:den>{d}</m:den></m:f>'

def para(label, math=None, display=False):
    p = f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{label}</w:t></w:r></w:p>'
    if math is None:
        return p
    om = f'<m:oMath>{math}</m:oMath>'
    if display:
        om = f'<m:oMathPara>{om}</m:oMathPara>'
    return p + f'<w:p>{om}</w:p>'

piecewise = ('<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr><m:e><m:m><m:mPr>'
             '<m:mcs><m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>'
             '<m:mr><m:e>' + R('x+1,') + '</m:e><m:e>' + R('x≥0') + '</m:e></m:mr>'
             '<m:mr><m:e>' + R('−x,') + '</m:e><m:e>' + R('x&lt;−100') + '</m:e></m:mr></m:m></m:e></m:d>')

CASES = [
    ('1. 积分后出现虚线框（被积项没放进积分号里）',
     nary('∫', R('0'), R('1')) + R('xdx=') + frac(R('1'), R('2')), True),
    ('2. 二重积分区域 D 不在积分号正下方',
     nary('∬', R('D'), '') + R('f(x,y)dxdy'), True),
    ('3. 求和上下标在右侧，没有居中',
     nary('∑', R('i=1'), R('n')) + R('i=') + frac(R('n(n+1)'), R('2')), True),
    ('4. lim 的下标在右下角，没有放到正下方',
     sSub(R('lim', 'p'), R('x→0')) + frac(R('sinx'), R('x')) + R('=1'), True),
    ('5. 括号上标被拆散：(a+b)² 的左括号和右括号不在同一结构里',
     R('(a+b') + sSup(R(')'), R('2')) + R('=a²+2ab+b²'), True),
    ('6. 函数名连写：lncosx、sin2x 中间没有间距',
     R('y=lncosx+sin2x'), True),
    ('7. 正斜体：数字、微分 d、常数 e 应为正体，变量 a、d（公差）保持斜体',
     frac(R('d'), R('dx')) + sSup(R('e'), R('2x')) + R('=2') + sSup(R('e'), R('2x')) + R(',  ') + sSub(R('a'), R('n')) + R('=') + sSub(R('a'), R('1')) + R('+(n−1)d'), True),
    ('8. 代入竖线高度不够',
     f'<m:sSubSup><m:e>{R("x²|")}</m:e><m:sub>{R("0")}</m:sub><m:sup>{R("1")}</m:sup></m:sSubSup>' + R('=1'), True),
    ('9. 积分号随被积函数拉长（grow=1）',
     nary('∫', '', '', frac(R('1'), f'<m:rad><m:radPr><m:degHide m:val="on"/></m:radPr><m:deg/><m:e>{R("1−x²")}</m:e></m:rad>') + R('dx')), True),
    ('10. 分段函数条件列居中，应改为左对齐',
     R('f(x)=') + piecewise, True),
    ('11. 集合大括号不随内容变高',
     R('A=') + '<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val="}"/><m:grow m:val="0"/></m:dPr><m:e>'
     + R('x|') + frac(R('1'), R('x')) + R('&gt;0') + '</m:e></m:d>', True),
    ('12. 保护项：limsup 下标居中但仍是 limsup（旧版会改成 lim）；普通文本“单位 cm”保持原样',
     sSub(R('limsup', 'p'), R('n→∞')) + sSub(R('a'), R('n')) + R('=1,  ') + R('单位 cm', nor=True), True),
]

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"')

def raw_para(math, ppr='', tail=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><m:oMathPara><m:oMath>{math}</m:oMath></m:oMathPara>{tail}</w:p>'


def text_para(math, text, ppr=''):
    # 公式后直接打编号：Word 会把公式降为行内模式
    return (f'<w:p><w:pPr>{ppr}</w:pPr><m:oMath>{math}</m:oMath><w:r><w:tab/></w:r>'
            f'<w:r><w:t>{text}</w:t></w:r></w:p>')


EXTRA = [
    para('13. 编号写在公式后面：公式被降为行内模式，编号位置不对') +
    text_para(R('F=ma'), '(1)'),
    para('14. Word 365 的 #(2) 编号写法：旧版 Word / WPS 显示成 #(2)') +
    raw_para(R('p+') + frac(R('1'), R('2')) + sSup(R('ρv'), R('2')) + R('=const#(2)')),
    para('15. 行距为固定值 12 磅：高公式上下被截掉') +
    raw_para(frac(R('∂u'), R('∂t')) + R('+u') + frac(R('∂u'), R('∂x')) + R('=−') + frac(R('1'), R('ρ')) +
             frac(R('∂p'), R('∂x')), '<w:spacing w:line="240" w:lineRule="exact"/>'),
    para('16. 工程符号：特征数 Re、Ma 应斜体，下标 max/ref、C_L 的 L 应正体，单位 Pa、m/s 应正体') +
    raw_para(R('Re', 'p') + R('=') + frac(R('ρvL'), R('μ')) + R(',  ') + R('Ma', 'p') + R('=') + frac(R('v'), R('a')) +
             R(',  ') + sSub(R('T'), R('max')) + R(',  ') + sSub(R('p'), R('ref')) + R('=101325Pa,  ') +
             sSub(R('C'), R('L')) + R(',  ') + sSub(R('v'), R('in')) + R('=30 m/s')),
    para('17. π 与差分 Δ 应正体；全角括号、逗号和连字符应改为半角与减号') +
    raw_para(R('S=πr²,  Δp=') + sSub(R('p'), R('1')) + R('-') + sSub(R('p'), R('2')) + R('，  f（x）＝x+１')),
]

body = para('Word 公式修复演示文档（每段一种典型问题）')
for label, math, disp in CASES:
    body += para(label, math, disp)
body += ''.join(EXTRA)
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
            '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'
            '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>')
header = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}><w:p>'
          '<w:r><w:t xml:space="preserve">页眉中的公式（旧版不处理）：</w:t></w:r>'
          f'<m:oMath>{R("lnx+cosx")}</m:oMath></w:p></w:hdr>')
content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>')
doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    '</Relationships>')

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), '公式修复演示_修复前.docx')
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/_rels/document.xml.rels', doc_rels)
    z.writestr('word/document.xml', document)
    z.writestr('word/header1.xml', header)
print(out)
