// State management
let currentFilePath = null;
let currentFileName = null;
let currentAnalysisData = null;
let isProcessing = false;

let activeTab = 'document'; // 'document', 'native', or 'issues'
let docDisplayMode = 'repaired'; // 'repaired' or 'original'
let highlightModified = true;
let viewMode = 'split'; // 'split' or 'after'
let searchDebounceTimer = null;
let nativeRenderData = null;

// DOM Elements
const btnOpenFile = document.getElementById('btnOpenFile');
const btnQuickSample = document.getElementById('btnQuickSample');
const btnBatchRepair = document.getElementById('btnBatchRepair');
const btnOpenInWord = document.getElementById('btnOpenInWord');
const btnRepairAndSave = document.getElementById('btnRepairAndSave');
const chkExportPdf = document.getElementById('chkExportPdf');

const btnSelectAllRules = document.getElementById('btnSelectAllRules');
const btnResetRules = document.getElementById('btnResetRules');
const btnClearRules = document.getElementById('btnClearRules');

const labelFileName = document.getElementById('labelFileName');
const labelFilePath = document.getElementById('labelFilePath');
const statusMessage = document.getElementById('statusMessage');

const appHeaderDocTitle = document.getElementById('appHeaderDocTitle');
const appHeaderSub = document.getElementById('appHeaderSub');

const btnTabDocument = document.getElementById('btnTabDocument');
const btnTabNativeRender = document.getElementById('btnTabNativeRender');
const btnTabIssues = document.getElementById('btnTabIssues');
const badgeIssueCount = document.getElementById('badgeIssueCount');

const inputSearch = document.getElementById('inputSearch');
const docViewControls = document.getElementById('docViewControls');
const nativeViewControls = document.getElementById('nativeViewControls');
const issuesViewControls = document.getElementById('issuesViewControls');

const btnShowRepaired = document.getElementById('btnShowRepaired');
const btnShowOriginal = document.getElementById('btnShowOriginal');
const chkHighlightModified = document.getElementById('chkHighlightModified');
const btnRefreshNative = document.getElementById('btnRefreshNative');
const labelNativePageInfo = document.getElementById('labelNativePageInfo');

const selectRuleFilter = document.getElementById('selectRuleFilter');
const btnViewSplit = document.getElementById('btnViewSplit');
const btnViewAfter = document.getElementById('btnViewAfter');

const documentView = document.getElementById('documentView');
const documentSheet = document.getElementById('documentSheet');
const nativeRenderView = document.getElementById('nativeRenderView');
const nativePagesContainer = document.getElementById('nativePagesContainer');
const issuesView = document.getElementById('issuesView');
const issuesContainer = document.getElementById('issuesContainer');
const emptyState = document.getElementById('emptyState');

const sbParagraphCount = document.getElementById('sbParagraphCount');
const sbFormulaCount = document.getElementById('sbFormulaCount');
const sbIssueCount = document.getElementById('sbIssueCount');
const sbDocStatus = document.getElementById('sbDocStatus');

const mathTooltip = document.getElementById('mathTooltip');
const tooltipId = document.getElementById('tooltipId');
const tooltipRules = document.getElementById('tooltipRules');
const tooltipBefore = document.getElementById('tooltipBefore');
const tooltipAfter = document.getElementById('tooltipAfter');

const loadingModal = document.getElementById('loadingModal');
const loadingTitle = document.getElementById('loadingTitle');
const loadingSubtitle = document.getElementById('loadingSubtitle');

const btnToggleSidebar = document.getElementById('btnToggleSidebar');
const appSidebar = document.getElementById('appSidebar');

if (btnToggleSidebar && appSidebar) {
  btnToggleSidebar.addEventListener('click', () => {
    appSidebar.classList.toggle('collapsed');
  });
}

// ---------- 规则侧边栏：由 rules-config.js 生成 ----------
const ALL_RULES = window.RULE_GROUPS.flatMap(g => g.rules);
const RULE_STORE_KEY = 'wmr.ruleOptions.v2';
const savedRuleOptions = (() => {
  try { return JSON.parse(localStorage.getItem(RULE_STORE_KEY) || '{}'); } catch (_) { return {}; }
})();

(function buildRuleSidebar() {
  const list = document.getElementById('rulesList');
  const collapsed = JSON.parse(localStorage.getItem('wmr.collapsedGroups') || '[]');
  list.innerHTML = window.RULE_GROUPS.map(g => `
    <div class="rule-group ${collapsed.includes(g.id) ? 'collapsed' : ''}" data-group="${g.id}">
      <button class="rule-group-header" type="button">
        <svg class="chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"></polyline></svg>
        <span class="rule-group-title">${g.title}</span>
        <span class="rule-group-count" data-count-for="${g.id}"></span>
      </button>
      <div class="rule-group-body">
        ${g.rules.map(r => {
          const on = (r.key in savedRuleOptions) ? savedRuleOptions[r.key] : !r.defaultOff;
          return `
          <div class="rule-item" data-rule-item="${r.key}">
            <div class="rule-info">
              <label class="rule-title" for="rule_${r.key}">${r.name}<span class="rule-hit" data-hit-for="${r.key}"></span></label>
              <p class="rule-desc">${r.desc}</p>
            </div>
            <label class="switch">
              <input type="checkbox" id="rule_${r.key}" data-rule="${r.key}" ${on ? 'checked' : ''}>
              <span class="slider"></span>
            </label>
          </div>`;
        }).join('')}
      </div>
    </div>`).join('');

  list.querySelectorAll('.rule-group-header').forEach(h => h.addEventListener('click', () => {
    h.parentElement.classList.toggle('collapsed');
    const ids = [...list.querySelectorAll('.rule-group.collapsed')].map(x => x.dataset.group);
    localStorage.setItem('wmr.collapsedGroups', JSON.stringify(ids));
  }));

  const sel = document.getElementById('selectRuleFilter');
  window.RULE_GROUPS.forEach(g => {
    const og = document.createElement('optgroup');
    og.label = g.title;
    g.rules.forEach(r => og.appendChild(new Option(r.name, r.key)));
    sel.appendChild(og);
  });
})();

function updateGroupCounts() {
  window.RULE_GROUPS.forEach(g => {
    const on = g.rules.filter(r => document.getElementById(`rule_${r.key}`).checked).length;
    const el = document.querySelector(`[data-count-for="${g.id}"]`);
    if (el) el.textContent = `${on}/${g.rules.length}`;
  });
}

function updateRuleHits(issuesFound) {
  ALL_RULES.forEach(r => {
    const el = document.querySelector(`[data-hit-for="${r.key}"]`);
    if (!el) return;
    const n = issuesFound ? (issuesFound[r.key] || 0) : 0;
    el.textContent = n ? n : '';
    el.classList.toggle('visible', n > 0);
  });
}

function applyPreset(name) {
  const pick = window.RULE_PRESETS[name];
  ALL_RULES.forEach(r => { document.getElementById(`rule_${r.key}`).checked = pick(r); });
  saveRuleOptions();
  triggerReAnalysis();
}

function saveRuleOptions() {
  localStorage.setItem(RULE_STORE_KEY, JSON.stringify(getRuleOptions()));
  updateGroupCounts();
}

document.querySelectorAll('.preset-chip').forEach(btn => btn.addEventListener('click', () => applyPreset(btn.dataset.preset)));

const ruleCheckboxes = document.querySelectorAll('.rule-item input[type="checkbox"]');
updateGroupCounts();

// Helper to get rule options
function getRuleOptions() {
  const options = {};
  ruleCheckboxes.forEach(cb => {
    const ruleKey = cb.dataset.rule;
    options[ruleKey] = cb.checked;
  });
  return options;
}

// Show/Hide Loading Overlay
function showLoading(title = '正在处理中...', subtitle = '正在解析文档与公式') {
  loadingTitle.textContent = title;
  loadingSubtitle.textContent = subtitle;
  loadingModal.classList.add('active');
}

function hideLoading() {
  loadingModal.classList.remove('active');
}

// KaTeX Render Helper
function renderLatexHtml(latexStr, displayMode = false) {
  if (!latexStr || !latexStr.trim()) {
    return '<span class="empty-math">(空公式)</span>';
  }
  try {
    return katex.renderToString(latexStr, {
      throwOnError: false,
      displayMode: displayMode,
      strict: false
    });
  } catch (err) {
    return `<span class="render-error">[公式错误: ${escapeHtml(latexStr)}]</span>`;
  }
}

// Escape HTML
function escapeHtml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// Load and Analyze Document
async function loadAndAnalyzeDocument(filePath) {
  if (!filePath || isProcessing) return;
  isProcessing = true;
  showLoading('正在解析文档内容...', '正在遍历段落与公式并构建 Word 视图');

  currentFilePath = filePath;
  currentFileName = filePath.split(/[\\/]/).pop();
  nativeRenderData = null; // reset cached native render

  labelFileName.textContent = currentFileName;
  labelFilePath.textContent = currentFilePath;
  labelFilePath.title = currentFilePath;
  appHeaderDocTitle.textContent = currentFileName;
  appHeaderSub.textContent = '正在读取与排版公式...';
  statusMessage.textContent = '解析中...';

  try {
    const options = getRuleOptions();
    const result = await window.api.analyzeDocx(filePath, options);

    if (result.error) {
      alert(`解析失败: ${result.error}`);
      statusMessage.textContent = '解析失败';
      appHeaderSub.textContent = '解析失败';
      return;
    }

    currentAnalysisData = result;
    updateRuleHits(result.issuesFound);

    // Update Status Bar & Header in Claude style
    sbParagraphCount.textContent = `段落: ${result.totalParagraphs}`;
    sbFormulaCount.textContent = `公式: ${result.totalOMath}`;
    sbIssueCount.textContent = `可优化: ${result.totalIssues}`;
    badgeIssueCount.textContent = result.totalIssues;

    appHeaderSub.textContent = `${result.totalOMath} 个公式 · ${result.totalIssues} 处已优化`;
    appHeaderSub.title = currentFilePath;

    statusMessage.textContent = result.totalIssues > 0
      ? `就绪: 发现 ${result.totalIssues} 处可优化项`
      : '就绪: 当前公式排版合规';

    btnRepairAndSave.disabled = false;
    btnOpenInWord.disabled = false;
    emptyState.style.display = 'none';

    // Render active tab view
    renderCurrentActiveView();

  } catch (err) {
    alert(`解析异常: ${err.message}`);
    statusMessage.textContent = '解析异常';
  } finally {
    isProcessing = false;
    hideLoading();
  }
}

// Re-analyze when switches change
let analyzeDebounceTimer = null;
function triggerReAnalysis() {
  if (!currentFilePath) return;
  clearTimeout(analyzeDebounceTimer);
  analyzeDebounceTimer = setTimeout(() => {
    loadAndAnalyzeDocument(currentFilePath);
  }, 200);
}

// Switch active view
function renderCurrentActiveView() {
  if (!currentAnalysisData) return;

  // Toggle active tab buttons
  btnTabDocument.classList.toggle('active', activeTab === 'document');
  btnTabNativeRender.classList.toggle('active', activeTab === 'native');
  btnTabIssues.classList.toggle('active', activeTab === 'issues');

  // Toggle view containers
  documentView.style.display = (activeTab === 'document') ? 'block' : 'none';
  nativeRenderView.style.display = (activeTab === 'native') ? 'block' : 'none';
  issuesView.style.display = (activeTab === 'issues') ? 'block' : 'none';

  // Toggle toolbar controls
  docViewControls.style.display = (activeTab === 'document') ? 'flex' : 'none';
  nativeViewControls.style.display = (activeTab === 'native') ? 'flex' : 'none';
  issuesViewControls.style.display = (activeTab === 'issues') ? 'flex' : 'none';

  if (activeTab === 'document') {
    renderDocumentView();
  } else if (activeTab === 'native') {
    renderNativeView();
  } else if (activeTab === 'issues') {
    renderIssuesList();
  }
}

// Math Element Lookup Map
const mathElementMap = new Map();

// 1. Render Document View (Word A4 Paper Layout)
function renderDocumentView() {
  if (!currentAnalysisData || !currentAnalysisData.paragraphs) return;

  mathElementMap.clear();
  const query = (inputSearch.value || '').trim().toLowerCase();
  const isRepaired = (docDisplayMode === 'repaired');
  const isHighlight = chkHighlightModified.checked;

  let html = '';
  let visibleCount = 0;

  currentAnalysisData.paragraphs.forEach(p => {
    let pSnippet = p.snippet || '';
    if (query) {
      const matchText = pSnippet.toLowerCase();
      const hasMatch = matchText.includes(query) || p.elements.some(e => {
        if (e.type === 'text') return e.text.toLowerCase().includes(query);
        if (e.type === 'math') return (e.beforeLatex + ' ' + e.afterLatex).toLowerCase().includes(query);
        return false;
      });
      if (!hasMatch) return;
    }

    visibleCount++;
    let pContentHtml = '';

    p.elements.forEach(elem => {
      if (elem.type === 'text') {
        pContentHtml += `<span class="doc-text">${escapeHtml(elem.text)}</span>`;
      } else if (elem.type === 'br') {
        pContentHtml += `<br>`;
      } else if (elem.type === 'math') {
        mathElementMap.set(elem.mathId, elem);

        const latexToShow = isRepaired
          ? (elem.isModified ? elem.afterLatex : elem.beforeLatex)
          : elem.beforeLatex;

        const mathHtml = renderLatexHtml(latexToShow, elem.isDisplay);
        const modClass = (elem.isModified && isHighlight) ? 'math-modified' : '';

        if (elem.isDisplay) {
          pContentHtml += `
            <div class="doc-math-block ${modClass}" data-math-id="${elem.mathId}" title="${elem.isModified ? '已优化公式 (鼠标悬停查看对比)' : ''}">
              ${elem.isModified && isHighlight ? '<span class="modified-badge">已优化</span>' : ''}
              ${mathHtml}
            </div>
          `;
        } else {
          pContentHtml += `
            <span class="doc-math-inline ${modClass}" data-math-id="${elem.mathId}" title="${elem.isModified ? '已优化公式 (鼠标悬停查看对比)' : ''}">
              ${mathHtml}
              ${elem.isModified && isHighlight ? '<span class="inline-mod-indicator"></span>' : ''}
            </span>
          `;
        }
      }
    });

    // 公式编号：修复后视图里显示为“公式居中 + 编号右对齐”
    if (isRepaired && p.eqNumber) {
      const mathParts = p.elements.filter(e => e.type === 'math').map(elem => {
        const latex = elem.isModified ? elem.afterLatex : elem.beforeLatex;
        return renderLatexHtml(latex, true);
      }).join('');
      const firstMath = p.elements.find(e => e.type === 'math');
      const modClass = isHighlight ? 'math-modified' : '';
      pContentHtml = `
        <div class="doc-math-numbered ${modClass}" data-math-id="${firstMath ? firstMath.mathId : ''}">
          <div class="eq-body">${mathParts}</div>
          <div class="eq-number">(${escapeHtml(p.eqNumber)})</div>
        </div>`;
    }

    const isEmpty = p.elements.length === 0 || (!pContentHtml.trim());
    if (isEmpty) {
      pContentHtml = '&nbsp;';
    }

    html += `
      <div class="doc-paragraph ${p.hasModified ? 'has-modified-formula' : ''}" data-idx="${p.index}">
        <span class="p-index-tag">[${p.index + 1}]</span>
        <div class="p-body">${pContentHtml}</div>
      </div>
    `;
  });

  if (visibleCount === 0) {
    documentSheet.innerHTML = `
      <div class="no-results-hint">
        <p>未找到包含 “${escapeHtml(query)}” 的文档内容</p>
      </div>
    `;
    return;
  }

  documentSheet.innerHTML = html;
  bindMathTooltipEvents();
}

// 2. Render Native View (High-fidelity Word Engine Rendering)
async function renderNativeView(forceRefresh = false) {
  if (!currentFilePath) return;

  if (nativeRenderData && !forceRefresh) {
    displayNativePages(nativeRenderData);
    return;
  }

  nativePagesContainer.innerHTML = `
    <div class="native-loading-hint">
      <div class="spinner-small"></div>
      <span>正在调用本机 Microsoft Word 引擎生成高保真原生渲染...</span>
    </div>
  `;

  try {
    const res = await window.api.renderPages(currentFilePath);
    if (res && res.success) {
      nativeRenderData = res;
      displayNativePages(res);
    } else {
      nativePagesContainer.innerHTML = `
        <div class="native-error-hint">
          <p>⚠️ 生成 Word 原生渲染遇到问题：${escapeHtml(res ? res.error : '未知错误')}</p>
          <p style="margin-top: 8px; font-size: 12px; color: var(--text-muted);">
            提示：您也可以直接点击右上角的“在 Word 中打开”按钮直接查看。
          </p>
        </div>
      `;
    }
  } catch (err) {
    nativePagesContainer.innerHTML = `
      <div class="native-error-hint">
        <p>⚠️ 调用本地 Word 引擎失败: ${escapeHtml(err.message)}</p>
      </div>
    `;
  }
}

function displayNativePages(data) {
  labelNativePageInfo.textContent = `共 ${data.pageCount} 页 (Word 渲染)`;
  let html = '';
  const timestamp = Date.now();
  data.renderedPages.forEach((imgName, idx) => {
    // 渲染图片在用户数据目录 (见 main.js)，用绝对 file:// 地址加载
    const dirUrl = 'file:///' + encodeURI(String(data.outputDir).replace(/\\/g, '/'));
    const imgSrc = `${dirUrl}/${imgName}?t=${timestamp}`;
    html += `
      <div class="native-page-card" data-page="${idx + 1}">
        <div class="native-page-header">
          <span class="page-num-badge">第 ${idx + 1} 页 / 共 ${data.pageCount} 页</span>
        </div>
        <img src="${imgSrc}" class="native-page-img" alt="第 ${idx + 1} 页 (Word 原生渲染)" loading="lazy"/>
      </div>
    `;
  });
  nativePagesContainer.innerHTML = html;
}

// 3. Render Issues View (Side-by-side Comparison)
function renderIssuesList() {
  if (!currentAnalysisData || !currentAnalysisData.issuesList) return;

  const query = (inputSearch.value || '').trim().toLowerCase();
  const filterRule = selectRuleFilter.value;

  const filtered = currentAnalysisData.issuesList.filter(item => {
    if (filterRule !== 'ALL' && !item.rulesApplied.includes(filterRule)) {
      return false;
    }
    if (query) {
      const matchText = (item.paragraphSnippet + ' ' + item.beforeText + ' ' + item.afterText + ' ' + item.beforeLatex + ' ' + item.afterLatex).toLowerCase();
      if (!matchText.includes(query)) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    issuesContainer.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">✨</div>
        <h3>未发现符合条件的待优化项</h3>
        <p>当前筛选条件下所有公式均合规或未命中规则，可调整上方规则开关或搜索关键词。</p>
      </div>`;
    return;
  }

  let html = '';
  filtered.forEach(item => {
    const beforeRendered = renderLatexHtml(item.beforeLatex, true);
    const afterRendered = renderLatexHtml(item.afterLatex, true);
    const tagsHtml = item.ruleNames.map(rName => `<span class="tag">${escapeHtml(rName)}</span>`).join('');
    const descText = item.descriptions.join('； ');

    if (viewMode === 'split') {
      html += `
        <div class="issue-card" data-id="${item.id}">
          <div class="card-header">
            <div class="header-meta">
              <span class="issue-badge">#${item.id}</span>
              <span class="paragraph-badge">第 ${item.paragraphIndex + 1} 段</span>
              <div class="rule-tags">${tagsHtml}</div>
            </div>
          </div>
          <div class="card-snippet">
            <strong>所在题目/段落：</strong> ${escapeHtml(item.paragraphSnippet)}
          </div>
          <div class="card-comparison">
            <div class="comparison-col col-before">
              <div class="col-header">
                <span class="col-title">修改前</span>
              </div>
              <div class="math-render-box">${beforeRendered}</div>
            </div>
            <div class="comparison-col col-after">
              <div class="col-header">
                <span class="col-title">修改后</span>
              </div>
              <div class="math-render-box">${afterRendered}</div>
            </div>
          </div>
          <div class="card-footer">
            <span class="footer-icon">💡</span>
            <span>${escapeHtml(descText)}</span>
          </div>
        </div>
      `;
    } else {
      html += `
        <div class="issue-card" data-id="${item.id}">
          <div class="card-header">
            <div class="header-meta">
              <span class="issue-badge">#${item.id}</span>
              <span class="paragraph-badge">第 ${item.paragraphIndex + 1} 段</span>
              <div class="rule-tags">${tagsHtml}</div>
            </div>
          </div>
          <div class="card-snippet">
            <strong>所在题目/段落：</strong> ${escapeHtml(item.paragraphSnippet)}
          </div>
          <div class="comparison-col col-after" style="padding:16px;">
            <div class="col-header">
              <span class="col-title">修复后效果</span>
            </div>
            <div class="math-render-box">${afterRendered}</div>
          </div>
          <div class="card-footer">
            <span class="footer-icon">💡</span>
            <span>${escapeHtml(descText)}</span>
          </div>
        </div>
      `;
    }
  });

  issuesContainer.innerHTML = html;
}

// Tooltip positioning
function bindMathTooltipEvents() {
  const modifiedElements = documentSheet.querySelectorAll('.math-modified');

  modifiedElements.forEach(el => {
    el.addEventListener('mouseenter', () => {
      const mathId = parseInt(el.dataset.mathId, 10);
      const mathData = mathElementMap.get(mathId);
      if (!mathData || !mathData.isModified) return;

      tooltipId.textContent = `#${mathData.mathId}`;
      tooltipRules.innerHTML = mathData.ruleNames.map(r => `<span class="tooltip-tag">${escapeHtml(r)}</span>`).join(' ');
      tooltipBefore.innerHTML = renderLatexHtml(mathData.beforeLatex, false);
      tooltipAfter.innerHTML = renderLatexHtml(mathData.afterLatex, false);

      const rect = el.getBoundingClientRect();
      const tooltipWidth = 380;
      let left = rect.left;

      if (left + tooltipWidth > window.innerWidth - 20) {
        left = window.innerWidth - tooltipWidth - 20;
      }
      if (left < 10) left = 10;

      mathTooltip.style.display = 'block';
      const tooltipHeight = mathTooltip.offsetHeight || 190;

      let top = rect.bottom + 8;
      if (top + tooltipHeight > window.innerHeight - 15) {
        top = Math.max(10, rect.top - tooltipHeight - 8);
      }

      mathTooltip.style.left = `${left}px`;
      mathTooltip.style.top = `${top}px`;
    });

    el.addEventListener('mouseleave', () => {
      mathTooltip.style.display = 'none';
    });
  });
}

// 页面滚动时自动隐藏浮窗
if (documentView) {
  documentView.addEventListener('scroll', () => {
    if (mathTooltip) mathTooltip.style.display = 'none';
  }, { passive: true });
}

// Tab Events
btnTabDocument.addEventListener('click', () => {
  activeTab = 'document';
  renderCurrentActiveView();
});

btnTabNativeRender.addEventListener('click', () => {
  activeTab = 'native';
  renderCurrentActiveView();
});

btnTabIssues.addEventListener('click', () => {
  activeTab = 'issues';
  renderCurrentActiveView();
});

// Refresh Native Pages
btnRefreshNative.addEventListener('click', () => {
  renderNativeView(true);
});

// Action: Open in Word
btnOpenInWord.addEventListener('click', async () => {
  if (!currentFilePath) return;

  showLoading('正在准备并在 Word 中打开...', '正在写入优化规则并调用 Microsoft Word');
  try {
    const options = getRuleOptions();
    // Repair into temporary view docx
    const tempWordDocx = currentFilePath.replace(/\.docx$/i, '_Word预览.docx');
    const repairRes = await window.api.repairDocx(currentFilePath, tempWordDocx, options, false);
    const targetToOpen = (repairRes && repairRes.savedTo) ? repairRes.savedTo : tempWordDocx;
    await window.api.openInWord(targetToOpen);
    statusMessage.textContent = '已在 Microsoft Word 中呈现文档';
  } catch (err) {
    alert(`在 Word 中打开失败: ${err.message}`);
  } finally {
    hideLoading();
  }
});

// Document View Controls
btnShowRepaired.addEventListener('click', () => {
  docDisplayMode = 'repaired';
  btnShowRepaired.classList.add('active');
  btnShowOriginal.classList.remove('active');
  renderDocumentView();
});

btnShowOriginal.addEventListener('click', () => {
  docDisplayMode = 'original';
  btnShowOriginal.classList.add('active');
  btnShowRepaired.classList.remove('active');
  renderDocumentView();
});

chkHighlightModified.addEventListener('change', () => {
  highlightModified = chkHighlightModified.checked;
  renderDocumentView();
});

// Search input
inputSearch.addEventListener('input', () => {
  clearTimeout(searchDebounceTimer);
  searchDebounceTimer = setTimeout(() => {
    if (activeTab === 'document') renderDocumentView();
    else if (activeTab === 'issues') renderIssuesList();
  }, 150);
});

// Filter & View Mode for Issues Tab
selectRuleFilter.addEventListener('change', () => {
  renderIssuesList();
});

btnViewSplit.addEventListener('click', () => {
  btnViewSplit.classList.add('active');
  btnViewAfter.classList.remove('active');
  viewMode = 'split';
  renderIssuesList();
});

btnViewAfter.addEventListener('click', () => {
  btnViewAfter.classList.add('active');
  btnViewSplit.classList.remove('active');
  viewMode = 'after';
  renderIssuesList();
});

// Rule switch listeners
ruleCheckboxes.forEach(cb => {
  cb.addEventListener('change', () => {
    saveRuleOptions();
    triggerReAnalysis();
  });
});

btnSelectAllRules.addEventListener('click', () => applyPreset('all'));
btnClearRules.addEventListener('click', () => applyPreset('none'));
// “推荐”恢复默认开关 (自动编号默认关闭)
btnResetRules.addEventListener('click', () => applyPreset('recommended'));

// Button: Open File
btnOpenFile.addEventListener('click', async () => {
  const file = await window.api.openFileDialog(false);
  if (file) {
    loadAndAnalyzeDocument(file);
  }
});

// Button: Quick Sample
btnQuickSample.addEventListener('click', async () => {
  const samplePath = await window.api.getDefaultSamplePath();
  if (samplePath) {
    loadAndAnalyzeDocument(samplePath);
  } else {
    alert('未在工作区找到示例文件，请点击“打开文档”手动选取。');
  }
});

// Button: Batch Repair
btnBatchRepair.addEventListener('click', async () => {
  const files = await window.api.openFileDialog(true);
  if (!files || files.length === 0) return;

  const outputDir = await window.api.selectDirectory();
  if (!outputDir) return;

  const exportPdf = chkExportPdf ? chkExportPdf.checked : false;
  let stopProgress = null;

  showLoading(`正在批量修复 1/${files.length}...`, '正在准备文档');

  try {
    stopProgress = window.api.onBatchProgress((progress) => {
      showLoading(
        `正在批量修复 ${progress.index}/${progress.total}...`,
        progress.fileName
      );
    });

    const { results } = await window.api.repairBatch(files, outputDir, getRuleOptions(), exportPdf);
    const succeeded = results.filter(r => r.ok);
    const failed = results.filter(r => !r.ok);

    let message = `批量修复完成：成功 ${succeeded.length} 个，失败 ${failed.length} 个。\n输出目录：${outputDir}`;
    if (failed.length > 0) {
      const shown = failed.slice(0, 10).map(f => `- ${f.file}: ${f.error}`).join('\n');
      message += `\n\n失败详情：\n${shown}`;
      if (failed.length > 10) message += `\n…另有 ${failed.length - 10} 个失败`;
    }

    statusMessage.textContent = `批量修复完成：成功 ${succeeded.length}/${results.length}，输出至 ${outputDir}`;
    if (confirm(`${message}\n\n是否打开输出文件夹？`)) {
      window.api.openPath(outputDir);
    }
  } catch (err) {
    alert(`批量修复失败: ${err.message}`);
  } finally {
    if (stopProgress) stopProgress();
    hideLoading();
  }
});

// Button: Repair and Save
btnRepairAndSave.addEventListener('click', async () => {
  if (!currentFilePath) return;

  const defaultSaveName = '已修复_' + currentFileName;
  const targetSavePath = await window.api.saveFileDialog(defaultSaveName);

  if (!targetSavePath) return;

  const exportPdf = chkExportPdf ? chkExportPdf.checked : false;
  showLoading('正在保存修复后的文档...', `正在写入至 ${targetSavePath}${exportPdf ? ' (同时导出 PDF)' : ''}`);

  try {
    const options = getRuleOptions();
    const result = await window.api.repairDocx(currentFilePath, targetSavePath, options, exportPdf);

    if (result.error) {
      alert(`保存失败: ${result.error}`);
    } else {
      const finalSavedPath = result.savedTo || targetSavePath;
      let msg = `修复完成！\n已处理 ${result.totalIssues} 处公式。\n文档已保存: ${finalSavedPath}`;
      if (finalSavedPath !== targetSavePath) {
        msg += `\n(注：原目标文件被占用，已自动存为新文件名)`;
      }
      if (result.exportedPdf) {
        msg += `\nPDF 已生成: ${result.exportedPdf}`;
      } else if (result.pdfWarning) {
        msg += `\n(${result.pdfWarning})`;
      }
      msg += `\n\n是否立即在 Microsoft Word 中打开？`;
      statusMessage.textContent = `已保存至: ${finalSavedPath}`;
      const openWord = confirm(msg);
      if (openWord) {
        window.api.openInWord(finalSavedPath);
      }
    }
  } catch (err) {
    alert(`修复出错: ${err.message}`);
  } finally {
    hideLoading();
  }
});

// Drag and drop docx files into the window
window.addEventListener('dragover', (e) => {
  e.preventDefault();
  e.stopPropagation();
});

window.addEventListener('drop', (e) => {
  e.preventDefault();
  e.stopPropagation();
  if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
    const file = e.dataTransfer.files[0];
    const path = (window.api && typeof window.api.getPathForFile === 'function')
      ? window.api.getPathForFile(file)
      : file.path;
    if (path && path.toLowerCase().endsWith('.docx')) {
      loadAndAnalyzeDocument(path);
    } else {
      alert('请拖入有效的 Word (.docx) 文件！');
    }
  }
});

// Open .docx passed via "Open with" / second instance
if (window.api && typeof window.api.onOpenFile === 'function') {
  window.api.onOpenFile((filePath) => {
    if (filePath && filePath.toLowerCase().endsWith('.docx')) {
      loadAndAnalyzeDocument(filePath);
    }
  });
}

// Auto-load the launch file, otherwise the bundled sample, on startup
window.addEventListener('DOMContentLoaded', async () => {
  const launchFile = window.api && typeof window.api.getLaunchFile === 'function'
    ? await window.api.getLaunchFile()
    : null;
  if (launchFile) {
    loadAndAnalyzeDocument(launchFile);
    return;
  }
  const sample = await window.api.getDefaultSamplePath();
  if (sample) {
    loadAndAnalyzeDocument(sample);
  }
});
