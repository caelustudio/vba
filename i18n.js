/* Caelus Studio VBA 生成器 多语言切换：简体中文 / 繁體中文 / English */
(function () {
  var DICT = {
    'VBA · DeepSeek 代码生成器': ['VBA · DeepSeek 程式碼產生器', 'VBA · DeepSeek Code Generator'],
    'PPT 内调用 DeepSeek 的 VBA 代码生成器': ['PPT 內呼叫 DeepSeek 的 VBA 程式碼產生器', 'A VBA code generator for calling DeepSeek inside PowerPoint'],
    '只需 4 项配置。': ['只需 4 項配置。', 'Just 4 settings.'],
    '所有名称必须来自 PowerPoint 的': ['所有名稱必須來自 PowerPoint 的', 'All names must come from an '],
    '（开发工具 → 控件 → 文本框）， 代码通过': ['（開發工具 → 控制項 → 文字方塊），程式碼透過', ' (Developer → Controls → Text Box) on your slides; the code reads and writes them via '],
    '读写它们的值。API Key 从控件动态读取，不会硬编码到 VBA 里。': ['讀寫它們的值。API Key 從控制項動態讀取，不會硬編碼到 VBA 裡。', ' The API key is read from the controls at runtime — never hard-coded into the VBA.'],
    '让': ['讓', 'Let '],
    '直接向': ['直接向', ' talk to '],
    '提问。': ['提問。', ' directly.'],
    '使用步骤': ['使用步驟', 'How to use'],
    '使用提示': ['使用提示', 'Tips'],
    '填写配置': ['填寫配置', 'Configure'],
    '生成结果': ['產生結果', 'Result'],
    '01 提问': ['01 提問', '01 Question'],
    '02 回答': ['02 回答', '02 Answer'],
    '04 页码': ['04 頁碼', '04 Slide #'],
    '填写用于用户提问内容的 ActiveX 文本框名称': ['填寫用於使用者提問內容的 ActiveX 文字方塊名稱', 'Name of the ActiveX text box for the user question'],
    '填写用于 DeepSeek 输出内容的 ActiveX 文本框名称': ['填寫用於 DeepSeek 輸出內容的 ActiveX 文字方塊名稱', 'Name of the ActiveX text box for the DeepSeek output'],
    '填写你的 DeepSeek 的 API Key': ['填寫你的 DeepSeek 的 API Key', 'Your DeepSeek API key'],
    '填写你用于输入和输出的 ActiveX 文本框所在的幻灯片页码': ['填寫你用於輸入和輸出的 ActiveX 文字方塊所在的投影片頁碼', 'Slide number containing the input/output ActiveX text boxes'],
    '例如：txtQuestion': ['例如：txtQuestion', 'e.g. txtQuestion'],
    '例如：txtAnswer': ['例如：txtAnswer', 'e.g. txtAnswer'],
    '例如：txtApiKey': ['例如：txtApiKey', 'e.g. txtApiKey'],
    '例如：1': ['例如：1', 'e.g. 1'],
    '生成': ['產生', 'Generate'],
    '重置': ['重置', 'Reset'],
    'ActiveX 文本框': ['ActiveX 文字方塊', 'ActiveX text box'],
    'VBA 生成器': ['VBA 產生器', 'VBA Generator'],
    '生成器': ['產生器', 'Generator'],
    '项配置': ['項配置', ' settings'],
    '次生成': ['次產生', ' generations'],
    '次复用': ['次複用', ' reuses'],
    '下滑探索': ['下滑探索', 'Scroll to explore'],
    '主站': ['主站', 'Main site'],
    '本站': ['本站', 'This site'],
    '导航': ['導覽', 'Nav'],
    '首页': ['首頁', 'Home'],
    '教育': ['教育', 'Education'],
    '环保': ['環保', 'Environment'],
    '隐私': ['隱私', 'Privacy'],
    '联系': ['聯繫', 'Contact'],
    '关闭': ['關閉', 'Close'],
    '空格': ['空格', 'Space'],
    '跳跃 ·': ['跳躍 ·', 'Jump ·'],
    '重启 ·': ['重啟 ·', 'Restart ·'],
    '网页由人工智能辅助生成': ['網頁由人工智慧輔助生成', 'This website was created with AI assistance']
  };

  /* JS 动态生成的提示文案（多语言版） */
  var MSG = {
    fillPrefix: { 'zh-CN': '请填写：', 'zh-TW': '請填寫：', 'en': 'Please fill in: ' },
    slide: { 'zh-CN': '幻灯片号码必须是大于等于 1 的整数。', 'zh-TW': '投影片號碼必須是大於等於 1 的整數。', 'en': 'The slide number must be an integer greater than or equal to 1.' },
    name: { 'zh-CN': '控件名称只能包含中文、字母、数字和下划线，且不能以数字开头。', 'zh-TW': '控制項名稱只能包含中文、字母、數字和底線，且不能以數字開頭。', 'en': 'Control names may only contain Chinese characters, letters, numbers and underscores, and cannot start with a number.' },
    placeholder: { 'zh-CN': '填写上方四项配置后，点击「生成」。', 'zh-TW': '填寫上方四項配置後，點擊「產生」。', 'en': 'Fill in the four settings above, then click "Generate".' },
    fields: {
      'zh-CN': ['提问输入框名称', '回答存放框名称', 'API 输入框名称', '幻灯片号码'],
      'zh-TW': ['提問輸入框名稱', '回答存放框名稱', 'API 輸入框名稱', '投影片號碼'],
      'en': ['question box name', 'answer box name', 'API box name', 'slide number']
    }
  };

  function cur() { try { return localStorage.getItem('caelus_lang') || 'zh-CN'; } catch (e) { return 'zh-CN'; } }
  function norm(s) { return String(s).replace(/\s+/g, ' ').trim(); }
  function tr(s) { var e = DICT[norm(s)]; return e ? e[cur()] : null; }
  function msg(k) { return MSG[k][cur()]; }

  var obs = null;
  function apply() {
    if (obs) obs.disconnect();
    try {
      if (!document.body) return;
      var w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null), n, nodes = [];
      while ((n = w.nextNode())) nodes.push(n);
      for (var i = 0; i < nodes.length; i++) {
        var key = norm(nodes[i].nodeValue);
        if (!key) continue;
        var t = tr(key);
        if (t !== null && t !== nodes[i].nodeValue) nodes[i].nodeValue = t;
      }
      var els = document.querySelectorAll('input[placeholder], textarea[placeholder]');
      for (var j = 0; j < els.length; j++) {
        var tp = tr(els[j].getAttribute('placeholder'));
        if (tp !== null) els[j].setAttribute('placeholder', tp);
      }
      var tt = tr(document.title);
      if (tt !== null) document.title = tt;
      document.documentElement.setAttribute('lang', HTMLLANG[cur()]);
      var btn = document.getElementById('langToggle');
      if (btn) btn.textContent = LABEL[cur()];
    } finally {
      if (obs) obs.observe(document.body, { childList: true, subtree: true, characterData: true });
    }
  }

  function inject() {
    if (document.getElementById('langToggle')) return;
    var mount = document.querySelector('.nav-right') || document.querySelector('.site-nav') || document.body;
    var btn = document.createElement('button');
    btn.id = 'langToggle';
    btn.className = 'theme-toggle lang-toggle';
    btn.type = 'button';
    btn.title = '简 / 繁 / EN';
    btn.textContent = LABEL[cur()];
    var tt = document.getElementById('themeToggle');
    if (tt && tt.parentNode === mount) mount.insertBefore(btn, tt); else mount.appendChild(btn);
    btn.addEventListener('click', function () {
      var order = ['zh-CN', 'zh-TW', 'en'];
      setLang(order[(order.indexOf(cur()) + 1) % 3]);
      apply();
    });
  }

  function setLang(l) { try { localStorage.setItem('caelus_lang', l); } catch (e) {} }

  var LABEL = { 'zh-CN': '简', 'zh-TW': '繁', 'en': 'EN' };
  var HTMLLANG = { 'zh-CN': 'zh-CN', 'zh-TW': 'zh-Hant', 'en': 'en' };

  var style = document.createElement('style');
  style.textContent = '.lang-toggle{font-size:12px;font-weight:600;line-height:1;display:flex;align-items:center;justify-content:center;letter-spacing:.02em;}';
  document.head.appendChild(style);

  /* 把语言工具暴露给页面自身 JS（错误提示等动态文案） */
  window.caelusLang = { cur: cur, msg: msg, apply: apply };

  inject();
  apply();
  obs = new MutationObserver(function () { clearTimeout(apply._t); apply._t = setTimeout(apply, 120); });
  obs.observe(document.body, { childList: true, subtree: true, characterData: true });
})();
