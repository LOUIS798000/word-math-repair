// 规则配置：侧边栏分组、默认开关、场景预设与对比筛选都从这里生成
window.RULE_GROUPS = [
  {
    id: 'structure', title: '公式结构', rules: [
      { key: 'fixNaryEmptyOperand', name: '消除积分虚线框', desc: '补全遗漏的被积项，消除积分号后的虚线方框 <code>[::]</code>' },
      { key: 'fixIsolatedDelimiters', name: '修复括号上标', desc: '修复 <code>(式子)^n</code>、<code>(式子)_n</code> 被拆散的括号结构' },
      { key: 'fixLimits', name: '极限下标居中', desc: '<code>lim</code>、<code>max</code>、<code>limsup</code> 的下标放到正下方' },
      { key: 'fixNaryLimLoc', name: '求和/乘积上下标居中', desc: '<code>∑ ∏ ⋃ ⋂</code> 的上下标放到正上方与正下方' },
      { key: 'fixDoubleIntegralRegion', name: '二重积分区域居中', desc: '积分区域 <code>D</code> 放到 <code>∬</code> 正下方' },
      { key: 'fixIntegralSize', name: '固定积分号高度', desc: '防止积分号随被积函数过度拉长' },
      { key: 'fixEvalBar', name: '修复代入线高度', desc: '代入竖线 <code>|</code> 随内容自适应高度' },
      { key: 'fixSetBraces', name: '集合大括号修复', desc: '集合大括号随内容高度伸缩' },
      { key: 'fixPiecewiseAlignment', name: '分段函数左对齐', desc: '分段函数的条件列改为左对齐' },
    ]
  },
  {
    id: 'symbols', title: '正斜体与符号', rules: [
      { key: 'fixTypography', name: '正斜体规范', desc: '数字、微分 <code>d</code>、常数 <code>e</code>、转置 <code>T</code> 正体，变量斜体' },
      { key: 'fixEngineeringSymbols', name: '工程符号规范', desc: '特征数 <code>Re Ma</code> 斜体；下标 <code>max ref C_L</code>、单位 <code>Pa m/s</code>、<code>π Δ</code> 正体 (GB 3101 / ISO 80000)' },
      { key: 'splitGluedTokens', name: '拆分连写函数', desc: '拆开 <code>lncosx</code>、<code>sin2x</code> 等连写函数' },
      { key: 'fixFullwidth', name: '全角符号转半角', desc: '公式中的 <code>（ ）， ＝ １</code> 改为半角' },
      { key: 'fixMinusSign', name: '连字符改为减号', desc: '短横线 <code>-</code> 改为数学减号 <code>−</code>' },
    ]
  },
  {
    id: 'layout', title: '段落排版', rules: [
      { key: 'fixLineSpacing', name: '公式行距不被截断', desc: '“固定值”行距改为“最小值”，避免高公式被裁掉' },
      { key: 'fixSnapToGrid', name: '取消公式段落对齐网格', desc: '避免文档网格把公式所在行撑成双倍行高' },
      { key: 'fixDisplayCenter', name: '独立公式居中', desc: '左/右对齐的独立公式改为居中' },
      { key: 'fixMathPosition', name: '清除公式上浮/下沉', desc: '去掉字符位置偏移，公式与正文基线对齐' },
    ]
  },
  {
    id: 'numbering', title: '公式编号', rules: [
      { key: 'fixEquationNumbering', name: '编号右对齐（不丢失）', desc: '把公式后的 <code>(1)</code>、公式内的 <code>#(1)</code> 改为“公式居中 + 编号右对齐”，编号为可更新的 SEQ 域' },
      { key: 'addEquationNumbers', name: '自动为独立公式编号', desc: '为所有独立成行的公式按顺序编号 (1)(2)…，适合论文', defaultOff: true },
    ]
  },
  {
    id: 'font', title: '字体', rules: [
      { key: 'fixFontTimesNewRoman', name: '公式字体设为 Times New Roman', desc: '公式及正文的西文、数字统一为 Times New Roman' },
    ]
  },
];

window.RULE_PRESETS = {
  recommended: (r) => !r.defaultOff,
  all: () => true,
  none: () => false,
  // 试卷：不自动编号
  exam: (r) => r.key !== 'addEquationNumbers',
  // 论文：全部开启，包括自动编号
  thesis: () => true,
};
