/* 成绩汇总工作台 · 界面与导出 */
(function () {
  'use strict';

  var XLSX = window.XLSX;
  var Core = window.GradeCore;
  var LS_KEY = 'gradeWorkbench.v1';
  var ITEMS = Core.ITEMS;
  var SUBJECTS = Core.SUBJECTS;

  var DEFAULT_TIERS = [
    { name: '精英班', classes: ['高一1', '高一8', '高一9', '高一16', '高一17'] },
    { name: '火箭班', classes: ['高一2', '高一4', '高一5', '高一7', '高一10', '高一11', '高一12', '高一13', '高一15', '高一19', '高一20'] },
    { name: '实验班', classes: ['高一3', '高一6', '高一14', '高一18'] }
  ];

  var state = {
    cur: null,
    prev: null,
    prevMap: null,
    tiers: JSON.parse(JSON.stringify(DEFAULT_TIERS)),
    teachers: {},
    headTeachers: {},
    bands: [850, 800, 750, 700, 650, 600, 550, 500],
    awardPct: 16.7,
    texts: {},
    info: {},
    zkCutA: 688,
    zkCutB: 699,
    specialMap: {},
    excludes: {},
    exportSel: {},
    exportClasses: {},
    exportClassAnalysis: false,
    exportClassBands: false,
    examName: '',
    page: 'import',
    ui: {
      statTab: 'grade',
      gradePage: 1,
      gradeSize: 100,
      gradeSearch: '',
      gradeClass: '',
      gradeSort: { col: null, dir: -1 },
      classSel: '',
      stuPage: 1,
      stuSearch: '',
      stuClass: '',
      analysisCls: ''
    }
  };

  /* ---------------- helpers ---------------- */
  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function fmt(v) {
    if (v === null || v === undefined || v === '') return '—';
    if (typeof v === 'number') return String(Math.round(v * 10) / 10);
    return esc(v);
  }
  function fmtScore(v) {
    if (v === null || v === undefined || v === '') return '—';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v) : v.toFixed(2);
    return esc(v);
  }
  function num(v) {
    if (v === null || v === undefined) return '—';
    if (v === '') return '';
    if (typeof v === 'number') return String(Math.round(v * 10) / 10);
    return esc(v);
  }
  function progCell(v) {
    if (v === 'new') return '<span class="tag new">新进</span>';
    if (v === null || v === undefined) return '<span class="flat">—</span>';
    if (v > 0) return '<span class="up">+' + v + '</span>';
    if (v < 0) return '<span class="down">' + v + '</span>';
    return '<span class="flat">0</span>';
  }
  function deltaCell(v) {
    if (v === null || v === undefined) return '<span class="flat">—</span>';
    if (v > 0) return '<span class="up">+' + fmtScore(v) + '</span>';
    if (v < 0) return '<span class="down">' + fmtScore(v) + '</span>';
    return '<span class="flat">0</span>';
  }
  function toast(msg, isError) {
    var t = $('toast');
    t.textContent = msg;
    t.className = 'toast' + (isError ? ' error' : '');
    t.hidden = false;
    clearTimeout(toast._timer);
    toast._timer = setTimeout(function () { t.hidden = true; }, 3200);
  }
  function saveCfg() {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify({
        tiers: state.tiers,
        teachers: state.teachers,
        headTeachers: state.headTeachers,
        bands: state.bands,
        awardPct: state.awardPct,
        texts: state.texts,
        zkCutA: state.zkCutA,
        zkCutB: state.zkCutB,
        examName: state.examName
      }));
    } catch (e) { /* ignore quota */ }
  }
  function loadCfg() {
    try {
      var raw = localStorage.getItem(LS_KEY);
      if (!raw) return;
      var o = JSON.parse(raw);
      if (o.tiers && o.tiers.length) state.tiers = o.tiers;
      if (o.teachers) state.teachers = o.teachers;
      if (o.headTeachers) state.headTeachers = o.headTeachers;
      if (o.bands && o.bands.length) state.bands = o.bands;
      if (typeof o.awardPct === 'number') state.awardPct = o.awardPct;
      if (o.texts) state.texts = o.texts;
      if (typeof o.zkCutA === 'number') state.zkCutA = o.zkCutA;
      if (typeof o.zkCutB === 'number') state.zkCutB = o.zkCutB;
      if (o.examName) state.examName = o.examName;
    } catch (e) { /* ignore */ }
  }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function classNo(label) {
    var m = String(label).match(/\d+/);
    return m ? parseInt(m[0], 10) : 9999;
  }
  function sortClasses(list) {
    return list.slice().sort(function (a, b) { return classNo(a) - classNo(b) || String(a).localeCompare(String(b)); });
  }
  function dataClasses() {
    if (!state.cur) return [];
    return Core.classesOf(state.cur.students).map(function (c) { return c.label; });
  }
  function allClassLabels() {
    var set = {};
    dataClasses().forEach(function (c) { set[c] = 1; });
    state.tiers.forEach(function (t) { t.classes.forEach(function (c) { set[c] = 1; }); });
    return sortClasses(Object.keys(set));
  }
  function tierOf(label) {
    var t = Core.tierOfClass(state.tiers, label);
    return t ? t.name : '未归类';
  }

  /* ---------------- import ---------------- */
  function readWorkbook(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function (e) {
        try {
          var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: false });
          var sheets = wb.SheetNames.map(function (name) {
            var ws = wb.Sheets[name];
            return {
              name: name,
              rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false, defval: null })
            };
          });
          resolve(sheets);
        } catch (err) { reject(err); }
      };
      fr.onerror = function () { reject(new Error('文件读取失败')); };
      fr.readAsArrayBuffer(file);
    });
  }

  function ingest(kind, file) {
    toast('正在解析 ' + file.name + ' …');
    readWorkbook(file).then(function (sheets) {
      var ds = Core.buildDataset(sheets);
      if (!ds.ok) { toast(ds.error || '解析失败', true); return; }
      if (!ds.students.length) { toast('没有识别到有效学生：请检查表头是否含「姓名/学号」、得分列是否识别', true); return; }
      ds.fileName = file.name;
      Core.ensureTotals(ds.students);
      Core.computeRanks(ds.students);
      var infoRecords = Object.keys(state.info).map(function (k) { return state.info[k]; });
      if (infoRecords.length) Core.applyInfo(ds.students, infoRecords);
      if (kind === 'cur') {
        state.cur = ds;
        state.ui.gradePage = 1;
        state.ui.stuPage = 1;
        if (Object.keys(state.texts).length) { state.texts = {}; saveCfg(); }
        initSpecialMap();
      } else { state.prev = ds; }
      rebuild();
      renderAll();
      toast(kind === 'cur' ? ('已导入 ' + ds.students.length + ' 名学生') : ('上次考试已导入 ' + ds.students.length + ' 人'));
      if (kind === 'cur') updateDropLabels();
    }).catch(function (err) {
      toast('解析出错：' + (err && err.message ? err.message : err), true);
    });
  }

  function updateDropLabels() {
    $('fileCurName').textContent = state.cur ? state.cur.fileName + ' · ' + state.cur.students.length + ' 人' : '未选择文件';
    $('filePrevName').textContent = state.prev ? state.prev.fileName + ' · ' + state.prev.students.length + ' 人' : '未选择文件';
    $('dropCur').className = 'card drop' + (state.cur ? ' has-file' : '');
    $('dropPrev').className = 'card drop' + (state.prev ? ' has-file' : '');
  }

  function rebuild() {
    if (state.cur) {
      Core.ensureTotals(state.cur.students);
      Core.computeRanks(state.cur.students);
    }
    if (state.prev) {
      Core.ensureTotals(state.prev.students);
      Core.computeRanks(state.prev.students);
      state.prevMap = {};
      state.prev.students.forEach(function (s) { state.prevMap[s.sid] = s; });
    } else {
      state.prevMap = null;
    }
    updateStatus();
  }

  function updateStatus() {
    var pill = $('dataStatus');
    if (state.cur) {
      pill.textContent = '已导入 ' + state.cur.students.length + ' 人' + (state.prev ? ' · 含上次考试' : '');
      pill.className = 'status-pill on';
      $('btnExportAll').disabled = false;
      $('btnGotoStudents').disabled = false;
    } else {
      pill.textContent = '未导入数据';
      pill.className = 'status-pill';
      $('btnExportAll').disabled = true;
      $('btnGotoStudents').disabled = true;
    }
  }

  /* ---------------- navigation ---------------- */
  function go(page) {
    if (page !== 'import' && !state.cur) { toast('请先导入本次考试', true); return; }
    state.page = page;
    var pages = document.querySelectorAll('.page');
    for (var i = 0; i < pages.length; i++) pages[i].hidden = pages[i].id !== 'page-' + page;
    var navs = document.querySelectorAll('.nav-item');
    for (var j = 0; j < navs.length; j++) navs[j].classList.toggle('active', navs[j].getAttribute('data-page') === page);
    if (page === 'import') renderImport();
    if (page === 'students') renderStudents();
    if (page === 'tiers') renderTiers();
    if (page === 'stats') renderStats();
    if (page === 'analysis') renderAnalysis();
    if (page === 'awards') renderAwards();
    if (page === 'export') renderExportPage();
    window.scrollTo(0, 0);
  }

  function renderAll() {
    renderImport();
    if (state.cur) {
      renderStudents();
      renderTiers();
      renderStats();
      renderAnalysis();
      renderAwards();
      renderExportPage();
    }
  }

  /* ---------------- page: import ---------------- */
  function renderImport() {
    updateDropLabels();
    var card = $('importSummaryCard');
    var issues = $('issuesCard');
    if (!state.cur) { card.hidden = true; issues.hidden = true; return; }
    card.hidden = false;
    $('baseSheetNote').textContent = '基准表：' + (state.cur.baseSheet || '—');

    var clsCount = dataClasses().length;
    var absent = state.cur.absent.length;
    var sheets = state.cur.sheetInfo.length;
    $('importStats').innerHTML = [
      statHtml(state.cur.students.length, '有效学生'),
      statHtml(clsCount, '班级数'),
      statHtml(absent, '缺考 / 删除'),
      statHtml(sheets, '工作表')
    ].join('');

    $('sheetList').innerHTML = state.cur.sheetInfo.map(function (s) {
      var cls = 'sheet-chip' + (s.isGrade ? ' base' : '') + (s.ok ? '' : ' bad');
      return '<span class="' + cls + '">' + esc(s.name) + ' · ' + (s.ok ? s.students + ' 人' : '解析失败') + '</span>';
    }).join('');

    var list = state.cur.warnings.slice();
    if (state.prev) {
      list.push('已导入上次考试：' + state.prev.fileName + '（' + state.prev.students.length + ' 人），将计算位次进退');
      state.prev.warnings.forEach(function (w) { list.push('上次考试：' + w); });
    } else {
      list.push('未导入上次考试，进退列将显示「—」');
    }
    if (!list.length) list.push('未发现异常');
    $('issueList').innerHTML = list.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('');
    issues.hidden = false;
  }

  function statHtml(value, label) {
    return '<div class="stat"><b>' + esc(value) + '</b><span>' + esc(label) + '</span></div>';
  }

  /* ---------------- page: students ---------------- */
  function renderStudents() {
    if (!state.cur) return;
    var students = state.cur.students;
    $('studentStats').innerHTML = [
      statHtml(students.length, '在校学生'),
      statHtml(dataClasses().length, '班级'),
      statHtml(state.cur.absent.length, '缺考已删除'),
      statHtml(avgTotal(students), '总分平均')
    ].join('');

    var sel = $('studentClassFilter');
    var curVal = state.ui.stuClass;
    sel.innerHTML = '<option value="">全部班级</option>' + dataClasses().map(function (c) {
      return '<option value="' + esc(c) + '"' + (c === curVal ? ' selected' : '') + '>' + esc(c) + '</option>';
    }).join('');

    var q = state.ui.stuSearch.trim();
    var list = students.filter(function (s) {
      if (state.ui.stuClass && s.clsLabel !== state.ui.stuClass) return false;
      if (q && s.sid.indexOf(q) < 0 && s.name.indexOf(q) < 0) return false;
      return true;
    });
    $('studentCount').textContent = '共 ' + list.length + ' 人';
    var pageRows = list;

    var head = ['学号', '姓名', '班级', '总分', '语数外'].concat(SUBJECTS).concat(['操作']);
    var html = '<table class="grid"><thead><tr>' + head.map(function (h, i) {
      return '<th class="' + (i < 3 ? 'left' : '') + '">' + esc(h) + '</th>';
    }).join('') + '</tr></thead><tbody>';
    pageRows.forEach(function (s) {
      html += '<tr><td class="left">' + esc(s.sid) + '</td><td class="left">' + esc(s.name) + '</td><td class="left">' + esc(s.clsLabel) + '</td>';
      html += '<td class="num">' + fmtScore(s.scores['总分']) + '</td><td class="num">' + fmtScore(s.scores['语数外总分']) + '</td>';
      SUBJECTS.forEach(function (k) { html += '<td class="num">' + fmtScore(s.scores[k]) + '</td>'; });
      html += '<td><button class="btn small ghost" data-del="' + esc(s.sid) + '">删除</button></td></tr>';
    });
    html += '</tbody></table>';
    $('studentTable').innerHTML = html;
    $('studentTable').querySelectorAll('[data-del]').forEach(function (b) {
      b.addEventListener('click', function () { confirmDelete(b.getAttribute('data-del')); });
    });


    $('absentCount').textContent = state.cur.absent.length ? state.cur.absent.length + ' 人' : '无';
    if (!state.cur.absent.length) {
      $('absentTable').innerHTML = '<div class="muted" style="padding:10px">没有缺考或已删除的学生。</div>';
    } else {
      var ah = '<table class="grid"><thead><tr><th class="left">学号</th><th class="left">姓名</th><th class="left">班级</th></tr></thead><tbody>';
      state.cur.absent.forEach(function (s) {
        ah += '<tr><td class="left">' + esc(s.sid) + '</td><td class="left">' + esc(s.name) + '</td><td class="left">' + esc(s.clsLabel) + '</td></tr>';
      });
      ah += '</tbody></table>';
      $('absentTable').innerHTML = ah;
    }
  }

  function avgTotal(students) {
    var sum = 0, n = 0;
    students.forEach(function (s) { if (s.scores['总分'] !== null) { sum += s.scores['总分']; n++; } });
    return n ? Math.round(sum / n * 10) / 10 : '—';
  }

  function confirmDelete(sid) {
    var s = state.cur.students.filter(function (x) { return x.sid === sid; })[0];
    if (!s) return;
    openDialog('删除学生', '<p>确定删除 <b>' + esc(s.name) + '</b>（' + esc(s.sid) + '，' + esc(s.clsLabel) + '）吗？删除后所有统计立即重算。</p>',
      '<button class="btn" data-close="1">取消</button><button class="btn primary" id="confirmDel">删除</button>');
    $('confirmDel').addEventListener('click', function () {
      state.cur.students = state.cur.students.filter(function (x) { return x.sid !== sid; });
      closeDialog();
      rebuild();
      renderAll();
      toast('已删除 ' + s.name);
    });
  }

  function addStudentDialog() {
    var clsOptions = dataClasses().map(function (c) { return '<option value="' + esc(c) + '">' + esc(c) + '</option>'; }).join('');
    var scoreRows = SUBJECTS.map(function (k) {
      return '<div class="form-row"><label>' + k + '</label><input class="input" data-score="' + k + '" type="number" step="0.5" min="0"></div>';
    }).join('');
    openDialog('新增学生',
      '<div class="form-row"><label>学号</label><input class="input" id="addSid"></div>' +
      '<div class="form-row"><label>姓名</label><input class="input" id="addName"></div>' +
      '<div class="form-row"><label>班级</label><select class="input" id="addCls">' + clsOptions + '</select></div>' +
      '<div class="form-row"><label>总分</label><input class="input" id="addTotal" type="number" step="0.5" min="0" placeholder="留空则自动求和"></div>' +
      scoreRows +
      '<div class="form-note">学号不能与现有学生重复；各科可留空。</div>',
      '<button class="btn" data-close="1">取消</button><button class="btn primary" id="confirmAdd">添加</button>');
    $('confirmAdd').addEventListener('click', function () {
      var sid = Core.normSid($('addSid').value);
      var name = $('addName').value.trim();
      if (!sid || !name) { toast('学号和姓名必填', true); return; }
      if (state.cur.students.some(function (s) { return s.sid === sid; })) { toast('学号已存在', true); return; }
      var clsLabel = $('addCls').value;
      var scores = {};
      SUBJECTS.forEach(function (k) {
        var el = document.querySelector('[data-score="' + k + '"]');
        scores[k] = el && el.value !== '' ? Number(el.value) : null;
      });
      var totalEl = $('addTotal');
      scores['总分'] = totalEl.value !== '' ? Number(totalEl.value) : null;
      scores['语数外总分'] = null;
      state.cur.students.push({
        sid: sid,
        examNo: '',
        name: name,
        cls: classNo(clsLabel),
        clsRaw: clsLabel,
        clsLabel: clsLabel,
        scores: scores,
        srcRanks: {},
        status: 'normal'
      });
      closeDialog();
      rebuild();
      renderAll();
      toast('已添加 ' + name);
    });
  }

  /* ---------------- page: tiers ---------------- */
  function renderTiers() {
    var board = $('tierBoard');
    var assigned = {};
    state.tiers.forEach(function (t) { t.classes.forEach(function (c) { assigned[c] = 1; }); });
    board.innerHTML = state.tiers.map(function (t, ti) {
      var chips = t.classes.map(function (c) {
        return '<span class="cls-chip">' + esc(c) + '<button data-rm="' + ti + '" data-cls="' + esc(c) + '" title="移出">×</button></span>';
      }).join('');
      var available = allClassLabels().filter(function (c) { return t.classes.indexOf(c) < 0; });
      var opts = available.map(function (c) {
        return '<option value="' + esc(c) + '">' + esc(c) + '</option>';
      }).join('');
      return '<div class="tier-card">' +
        '<div class="tier-head"><span class="tier-name">' + esc(t.name) + '</span>' +
        '<span class="tier-count">' + t.classes.length + ' 个班</span>' +
        '<button class="icon-btn" data-deltier="' + ti + '" title="删除层次">×</button></div>' +
        '<div class="chip-row">' + (chips || '<span class="muted">暂无班级</span>') + '</div>' +
        '<div class="tier-add"><select class="input" data-addsel="' + ti + '"><option value="">选择班级…</option>' + opts + '</select>' +
        '<button class="btn small" data-addcls="' + ti + '">加入</button></div>' +
        '</div>';
    }).join('');

    var un = allClassLabels().filter(function (c) { return !assigned[c]; });
    $('unassignedClasses').innerHTML = un.length ? un.map(function (c) {
      return '<span class="cls-chip">' + esc(c) + '</span>';
    }).join('') : '<span class="muted">全部班级都已归类。</span>';

    board.querySelectorAll('[data-rm]').forEach(function (b) {
      b.addEventListener('click', function () {
        var ti = Number(b.getAttribute('data-rm'));
        var c = b.getAttribute('data-cls');
        state.tiers[ti].classes = state.tiers[ti].classes.filter(function (x) { return x !== c; });
        saveCfg(); renderTiers(); renderStats(); renderAnalysis(); renderAwards();
      });
    });
    board.querySelectorAll('[data-addcls]').forEach(function (b) {
      b.addEventListener('click', function () {
        var ti = Number(b.getAttribute('data-addcls'));
        var sel = board.querySelector('[data-addsel="' + ti + '"]');
        var c = sel.value;
        if (!c) { toast('请选择班级', true); return; }
        state.tiers.forEach(function (t) { t.classes = t.classes.filter(function (x) { return x !== c; }); });
        state.tiers[ti].classes.push(c);
        saveCfg(); renderTiers(); renderStats(); renderAnalysis(); renderAwards();
      });
    });
    board.querySelectorAll('[data-deltier]').forEach(function (b) {
      b.addEventListener('click', function () {
        var ti = Number(b.getAttribute('data-deltier'));
        if (state.tiers.length <= 1) { toast('至少保留一个层次', true); return; }
        state.tiers.splice(ti, 1);
        saveCfg(); renderTiers(); renderStats(); renderAnalysis(); renderAwards();
      });
    });
  }

  /* ---------------- page: stats ---------------- */
  function renderStats() {
    if (!state.cur) return;
    var tabs = document.querySelectorAll('#statTabs .subtab');
    for (var i = 0; i < tabs.length; i++) tabs[i].classList.toggle('active', tabs[i].getAttribute('data-tab') === state.ui.statTab);
    var panels = document.querySelectorAll('.tab-panel');
    for (var j = 0; j < panels.length; j++) panels[j].hidden = panels[j].getAttribute('data-panel') !== state.ui.statTab;

    if (state.ui.statTab === 'grade') renderGradeTable();
    if (state.ui.statTab === 'class') renderClassTable();
    if (state.ui.statTab === 'means') renderMeansTable();
    if (state.ui.statTab === 'diff') renderDiffTable();
  }

  function mainRows() {
    return Core.buildMainRows(state.cur.students, state.prevMap);
  }

  function renderGradeTable() {
    var pack = mainRows();
    var headers = pack.headers;
    var rows = pack.rows;
    var q = state.ui.gradeSearch.trim();
    var cf = state.ui.gradeClass;
    var filtered = rows.filter(function (r) {
      if (cf && r[3] !== cf) return false;
      if (q && String(r[0]).indexOf(q) < 0 && String(r[2]).indexOf(q) < 0) return false;
      return true;
    });
    var sort = state.ui.gradeSort;
    if (sort.col !== null && sort.col !== undefined) {
      filtered.sort(function (a, b) {
        var va = a[sort.col], vb = b[sort.col];
        var na = typeof va === 'number', nb = typeof vb === 'number';
        if (na && nb) return (va - vb) * sort.dir;
        if (na) return -1;
        if (nb) return 1;
        return String(va).localeCompare(String(vb)) * sort.dir;
      });
    }
    var size = state.ui.gradeSize;
    var totalPages = Math.max(1, Math.ceil(filtered.length / size));
    if (state.ui.gradePage > totalPages) state.ui.gradePage = totalPages;
    var pageRows = filtered.slice((state.ui.gradePage - 1) * size, state.ui.gradePage * size);

    var html = '<table class="grid" id="gradeGrid"><thead><tr>';
    headers.forEach(function (h, i) {
      var cls = (i < 4 ? 'left' : '') + (sort.col === i ? (sort.dir > 0 ? ' sort-asc' : ' sort-desc') : '');
      html += '<th class="' + cls + '" data-col="' + i + '">' + esc(h) + '</th>';
    });
    html += '</tr></thead><tbody>';
    pageRows.forEach(function (r) {
      html += '<tr>';
      r.forEach(function (v, i) {
        var isLeft = i < 4;
        var isProg = headers[i].indexOf('进退') >= 0;
        var isScore = i >= 4 && (i - 4) % 5 === 0;
        var cell = isProg ? (v === null || v === undefined ? '' : progCell(v)) : (isScore ? fmtScore(v) : num(v));
        html += '<td class="' + (isLeft ? 'left' : 'num') + (isProg ? ' prog' : '') + '">' + cell + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    $('gradeTable').innerHTML = html;

    $('gradeTable').querySelectorAll('th[data-col]').forEach(function (th) {
      th.addEventListener('click', function () {
        var col = Number(th.getAttribute('data-col'));
        if (state.ui.gradeSort.col === col) state.ui.gradeSort.dir = -state.ui.gradeSort.dir;
        else { state.ui.gradeSort.col = col; state.ui.gradeSort.dir = -1; }
        renderGradeTable();
      });
    });

    $('gradeCount').textContent = '共 ' + filtered.length + ' 人';
    renderGradePager(filtered.length);
    var sel = $('gradeClassFilter');
    var curVal = state.ui.gradeClass;
    sel.innerHTML = '<option value="">全部班级</option>' + dataClasses().map(function (c) {
      return '<option value="' + esc(c) + '"' + (c === curVal ? ' selected' : '') + '>' + esc(c) + '</option>';
    }).join('');
  }

  function renderGradePager(total) {
    var size = state.ui.gradeSize;
    var pages = Math.max(1, Math.ceil(total / size));
    var page = state.ui.gradePage;
    var html = '<span class="pager-info">第 ' + page + ' / ' + pages + ' 页</span>';
    html += '<button data-pg="' + Math.max(1, page - 1) + '"' + (page <= 1 ? ' disabled' : '') + '>上一页</button>';
    var start = Math.max(1, page - 3), end = Math.min(pages, start + 6);
    for (var p = start; p <= end; p++) html += '<button data-pg="' + p + '" class="' + (p === page ? 'active' : '') + '">' + p + '</button>';
    html += '<button data-pg="' + Math.min(pages, page + 1) + '"' + (page >= pages ? ' disabled' : '') + '>下一页</button>';
    var el = $('gradePager');
    el.innerHTML = html;
    el.querySelectorAll('[data-pg]').forEach(function (b) {
      b.addEventListener('click', function () { state.ui.gradePage = Number(b.getAttribute('data-pg')); renderGradeTable(); });
    });
  }

  function renderClassTable() {
    var classes = dataClasses();
    var sel = $('classSelect');
    if (!state.ui.classSel || classes.indexOf(state.ui.classSel) < 0) state.ui.classSel = classes[0] || '';
    sel.innerHTML = classes.map(function (c) {
      return '<option value="' + esc(c) + '"' + (c === state.ui.classSel ? ' selected' : '') + '>' + esc(c) + '</option>';
    }).join('');
    var pack = mainRows();
    var idx = state.cur.students.map(function (s, i) { return s.clsLabel === state.ui.classSel ? i : -1; }).filter(function (i) { return i >= 0; });
    var rows = idx.map(function (i) { return pack.rows[i]; });
    var html = '<table class="grid"><thead><tr>' + pack.headers.map(function (h, i) {
      return '<th class="' + (i < 4 ? 'left' : '') + '">' + esc(h) + '</th>';
    }).join('') + '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr>';
      r.forEach(function (v, i) {
        var isProg = pack.headers[i].indexOf('进退') >= 0;
        var isScore = i >= 4 && (i - 4) % 5 === 0;
        html += '<td class="' + (i < 4 ? 'left' : 'num') + '">' + (isProg ? (v === null || v === undefined ? '' : progCell(v)) : (isScore ? fmtScore(v) : num(v))) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    $('classTable').innerHTML = html;
    $('classSheetCount').textContent = state.ui.classSel + ' · ' + rows.length + ' 人';
  }

  function renderMeansTable() {
    var cm = Core.classMeans(state.cur.students, ITEMS);
    var rk = Core.classRankTable(cm.means, cm.classes, ITEMS);
    var tiers = state.tiers;
    var html = '<table class="grid twohead"><thead><tr><th class="left" rowspan="2">班级</th><th class="left" rowspan="2">层次</th><th rowspan="2">人数</th>';
    ITEMS.forEach(function (item) { html += '<th colspan="3">' + esc(item) + '</th>'; });
    html += '</tr><tr>';
    ITEMS.forEach(function () { html += '<th>均分</th><th>年级排名</th><th>层次排名</th>'; });
    html += '</tr></thead><tbody>';
    cm.classes.forEach(function (c) {
      var tier = Core.tierOfClass(tiers, c.label);
      html += '<tr><td class="left">' + esc(c.label) + '</td><td class="left">' + esc(tier ? tier.name : '未归类') + '</td><td class="num">' + c.count + '</td>';
      ITEMS.forEach(function (item) {
        var tr = null;
        if (tier) {
          var vals = tier.classes.map(function (x) { return cm.means[x] ? cm.means[x][item] : null; });
          var rr = Core.rankValues(vals);
          var k = tier.classes.indexOf(c.label);
          tr = k >= 0 ? rr[k] : null;
        }
        html += '<td class="num">' + fmtScore(cm.means[c.label][item]) + '</td><td class="num">' + fmt(rk[c.label][item]) + '</td><td class="num">' + fmt(tr) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    $('meansTable').innerHTML = html;
  }

  function renderDiffTable() {
    var groups = Core.diffTable(state.cur.students, state.prev ? state.prev.students : null, state.tiers, ITEMS, state.teachers);
    var html = '';
    groups.forEach(function (g) {
      html += '<div style="margin-bottom:22px">';
      html += '<div class="card-head" style="margin-bottom:8px"><h2>' + esc(g.name) + '组</h2><span class="muted">' + g.rows.length + ' 个班 · 组末平均行</span></div>';
      html += '<div class="table-wrap"><table class="grid diff"><thead><tr><th class="left" rowspan="2">班级</th><th class="left" rowspan="2">班主任</th>';
      ITEMS.forEach(function (item) { html += '<th colspan="6">' + esc(item) + '</th>'; });
      html += '</tr><tr>';
      ITEMS.forEach(function () {
        html += '<th>均分</th><th>组内排名</th><th>与组均差值</th><th>上次名次</th><th>进退</th><th>代课教师</th>';
      });
      html += '</tr></thead><tbody>';
      g.rows.forEach(function (row) {
        html += '<tr><td class="left">' + esc(row.cls) + '</td>';
        html += '<td><input class="cell-input" data-hcls="' + esc(row.cls) + '" value="' + esc(getHeadTeacher(row.cls)) + '"></td>';
        ITEMS.forEach(function (item) {
          var d = row.items[item];
          html += '<td class="num">' + fmtScore(d.mean) + '</td><td class="num">' + fmt(d.rank) + '</td><td class="num">' + deltaCell(d.delta) + '</td><td class="num">' + fmt(d.prevRank) + '</td><td class="num">' + progCell(d.progress === null ? null : d.progress) + '</td>';
          html += '<td><input class="cell-input" data-tcls="' + esc(row.cls) + '" data-titem="' + esc(item) + '" value="' + esc(getTeacher(row.cls, item)) + '"></td>';
        });
        html += '</tr>';
      });
      var avg = g.avg;
      html += '<tr class="total-row"><td class="left">' + esc(avg.cls) + '</td><td></td>';
      ITEMS.forEach(function (item) {
        var d = avg.items[item];
        html += '<td class="num">' + fmtScore(d.mean) + '</td><td class="num">—</td><td class="num">—</td><td class="num">—</td><td class="num">—</td><td></td>';
      });
      html += '</tr></tbody></table></div></div>';
    });
    $('diffTable').innerHTML = html;
    $('diffTable').querySelectorAll('input[data-tcls]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        setTeacher(inp.getAttribute('data-tcls'), inp.getAttribute('data-titem'), inp.value.trim());
      });
    });
    $('diffTable').querySelectorAll('input[data-hcls]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        state.headTeachers[inp.getAttribute('data-hcls')] = inp.value.trim();
        saveCfg();
      });
    });
    updateTeacherStatus();
  }

  function getHeadTeacher(cls) { return state.headTeachers[cls] || ''; }
  function getTeacher(cls, item) { return (state.teachers[cls] || {})[item] || ''; }
  function setTeacher(cls, item, val) {
    if (!state.teachers[cls]) state.teachers[cls] = {};
    state.teachers[cls][item] = val;
    saveCfg();
  }

  function updateTeacherStatus() {
    if (!state.cur || !$('teacherStatus')) return;
    var classes = dataClasses();
    var hc = 0, tc = 0;
    classes.forEach(function (c) {
      if (state.headTeachers[c]) hc++;
      tc += Object.keys(state.teachers[c] || {}).length;
    });
    $('teacherStatus').textContent = (hc || tc) ? ('班主任 ' + hc + ' 个班 · 代课教师 ' + tc + ' 条') : '';
  }

  function importTeachers(file) {
    readWorkbook(file).then(function (sheets) {
      var head = {}, teachers = {}, warnings = [];
      sheets.forEach(function (sh) {
        var p = Core.parseTeacherRows(sh.rows || []);
        if (!p.ok) { warnings.push(sh.name + '：' + p.error); return; }
        Object.keys(p.head).forEach(function (k) { head[k] = p.head[k]; });
        Object.keys(p.teachers).forEach(function (cls) {
          teachers[cls] = teachers[cls] || {};
          Object.keys(p.teachers[cls]).forEach(function (it) { teachers[cls][it] = p.teachers[cls][it]; });
        });
        (p.warnings || []).forEach(function (w) { warnings.push(sh.name + '：' + w); });
      });
      var hc = Object.keys(head).length, tc = 0;
      if (!hc && !Object.keys(teachers).length) { toast('没有识别出班主任或教师姓名', true); return; }
      Object.keys(head).forEach(function (k) { state.headTeachers[k] = head[k]; });
      Object.keys(teachers).forEach(function (cls) {
        if (!state.teachers[cls]) state.teachers[cls] = {};
        Object.keys(teachers[cls]).forEach(function (it) { state.teachers[cls][it] = teachers[cls][it]; tc++; });
      });
      saveCfg();
      renderDiffTable();
      updateTeacherStatus();
      toast('已识别 班主任 ' + hc + ' 个班，代课教师 ' + tc + ' 条' + (warnings.length ? ('；' + warnings.length + ' 条未识别') : ''));
    }).catch(function (e) { toast('解析出错：' + (e && e.message ? e.message : e), true); });
  }

  /* ---------------- page: analysis ---------------- */
  function renderAnalysis() {
    if (!state.cur) return;
    var classes = dataClasses();
    var sel = $('analysisClass');
    if (!state.ui.analysisCls || classes.indexOf(state.ui.analysisCls) < 0) state.ui.analysisCls = classes[0] || '';
    sel.innerHTML = classes.map(function (c) {
      return '<option value="' + esc(c) + '"' + (c === state.ui.analysisCls ? ' selected' : '') + '>' + esc(c) + '</option>';
    }).join('');

    var cls = state.ui.analysisCls;
    if (!cls) { $('analysisBody').innerHTML = ''; return; }
    var bands = Core.bandDefs(state.bands);
    var a = Core.analyzeClass(state.cur.students, cls, state.tiers, bands, ITEMS);
    var enrolled = (state.cur.enrolled && state.cur.enrolled[cls]) || a.count;
    var text = state.texts[cls];
    if (!text) { text = Core.autoText(a); state.texts[cls] = text; saveCfg(); }

    var maxBand = Math.max.apply(null, a.bands.map(function (b) { return b.count; })) || 1;
    var bandHtml = a.bands.map(function (b) {
      var w = Math.round(b.count / maxBand * 100);
      return '<div class="band-row"><span class="band-label">' + esc(b.label) + '</span>' +
        '<span class="band-track"><span class="band-fill' + (b.label === a.bands[0].label ? ' top' : '') + '" style="width:' + w + '%"></span></span>' +
        '<span class="band-num">' + b.count + ' 人</span><span class="band-num">' + b.pct + '%</span></div>';
    }).join('');

    var subjHtml = '<table class="grid"><thead><tr><th class="left">学科</th><th>班级均分</th><th>层次均分</th><th>层次内名次</th></tr></thead><tbody>';
    ITEMS.forEach(function (item) {
      subjHtml += '<tr><td class="left">' + esc(item) + '</td><td class="num">' + fmtScore(a.means[item]) + '</td><td class="num">' + fmtScore(a.tierAvg[item]) + '</td><td class="num">' + fmt(a.tierRanks[item]) + '</td></tr>';
    });
    subjHtml += '</tbody></table>';

    $('analysisBody').innerHTML =
      '<div class="card">' +
        '<div class="report-head"><span class="report-title">' + esc(cls) + ' 成绩分析</span>' +
        '<span class="report-meta">层次：' + esc(a.tier) + ' · 应考 ' + enrolled + ' 人 · 实考 ' + a.count + ' 人</span>' +
        '<span class="report-meta">总分平均 <b>' + fmtScore(a.avg) + '</b> · 最高 ' + fmtScore(a.max) + ' · 最低 ' + fmtScore(a.min) + '</span></div>' +
        '<div class="analysis-grid">' +
          '<div><h3 style="font-size:13px;margin-bottom:10px">总分分数段</h3><div class="band-bars">' + bandHtml + '</div></div>' +
          '<div><h3 style="font-size:13px;margin-bottom:10px">各学科均分与层次名次</h3><div class="table-wrap short">' + subjHtml + '</div></div>' +
        '</div>' +
      '</div>' +
      '<div class="card">' +
        '<div class="card-head"><h2>分析与建议</h2><button class="btn small ghost" id="btnRegenText">按数据重新生成</button></div>' +
        '<div class="text-block"><label>存在问题</label><textarea data-text="problem">' + esc(text.problem) + '</textarea></div>' +
        '<div class="text-block"><label>改进措施</label><textarea data-text="measure">' + esc(text.measure) + '</textarea></div>' +
        '<div class="text-block"><label>教学建议</label><textarea data-text="advice">' + esc(text.advice) + '</textarea></div>' +
      '</div>';

    $('btnRegenText').addEventListener('click', function () {
      state.texts[cls] = Core.autoText(a);
      saveCfg();
      renderAnalysis();
      toast('已按当前数据重新生成');
    });
    $('analysisBody').querySelectorAll('[data-text]').forEach(function (ta) {
      ta.addEventListener('input', function () {
        if (!state.texts[cls]) state.texts[cls] = {};
        state.texts[cls][ta.getAttribute('data-text')] = ta.value;
        clearTimeout(ta._t);
        ta._t = setTimeout(saveCfg, 400);
      });
    });
    renderSpecial();
  }


  /* ---------------- 全年级专项统计 ---------------- */
  function withZhongkao() { return state.cur ? state.cur.students.filter(function (s) { return s.zhongkao !== null && s.zhongkao !== undefined; }).length : 0; }
  function withCategory() { return state.cur ? state.cur.students.filter(function (s) { return !!s.category; }).length : 0; }

  function isXiufan(s) { return /休返|休学|返校/.test(s.category || ''); }
  function isZidi(s) { return /子弟|子女|教师子/.test(s.category || ''); }
  function hasZk(s) { return s.zhongkao !== null && s.zhongkao !== undefined; }
  function isSpecial(s) { return isXiufan(s) || isZidi(s); }
  function defaultSpecialExclude(s) {
    if (isXiufan(s)) return true;
    if (isZidi(s) && !hasZk(s)) return true;
    return false;
  }
  function specialStudents() {
    return state.cur ? state.cur.students.filter(isSpecial) : [];
  }
  function initSpecialMap() {
    var map = {};
    specialStudents().forEach(function (s) { map[s.sid] = defaultSpecialExclude(s); });
    state.specialMap = map;
  }
  function excludedSet() {
    var set = {};
    if (!state.cur) return set;
    state.cur.students.forEach(function (s) {
      if (state.specialMap[s.sid] === true || state.excludes[s.sid]) set[s.sid] = true;
    });
    return set;
  }
  function analysisBase() {
    var set = excludedSet();
    return state.cur.students.filter(function (s) { return !set[s.sid]; });
  }

  function renderSpecial() {
    if (!state.cur) return;
    var nz = withZhongkao(), nc = withCategory();
    $('infoStatus').textContent = (nz || nc) ? ('中考成绩 ' + nz + ' 人 · 类别 ' + nc + ' 人') : '尚未导入中考成绩 / 类别';
    if (document.activeElement !== $('zkCutA')) $('zkCutA').value = state.zkCutA;
    if (document.activeElement !== $('zkCutB')) $('zkCutB').value = state.zkCutB;
    $('zkCard').hidden = !nz;
    var selCount = Object.keys(state.excludes).length;
    $('excludeCard').hidden = !(nc || selCount);
    if (nz) renderZkTable();
    if (nc || selCount) renderExclude();
  }

  function renderZkTable() {
    var base = analysisBase();
    var splits = Core.zhongkaoSplits(base, [state.zkCutA, state.zkCutB], ITEMS);
    var total = Core.statBlock(base, ITEMS);
    var noZk = Core.statBlock(base.filter(function (s) { return !hasZk(s); }), ITEMS);
    var n = base.length;
    var ZK_ITEMS = ['语数外总分'].concat(SUBJECTS);
    var head = ['分线', '组别', '人数', '占比', '中考均分', '本次总分均分', '最高', '最低'].concat(ZK_ITEMS.map(function (i) { return i + '均分'; }));
    var html = '<table class="grid"><thead><tr>' + head.map(function (h, i) { return '<th class="' + (i < 2 ? 'left' : '') + '">' + esc(h) + '</th>'; }).join('') + '</tr></thead><tbody>';
    function row(cut, label, b, muted) {
      var cells = [cut, label, b.n, (n ? Math.round(b.n / n * 1000) / 10 : 0) + '%', fmtScore(b.zkAvg), fmtScore(b.avg), fmtScore(b.max), fmtScore(b.min)];
      ZK_ITEMS.forEach(function (it) { cells.push(fmtScore(b.means[it])); });
      return '<tr' + (muted ? ' class="total-row"' : '') + '>' + cells.map(function (v, i) { return '<td class="' + (i < 2 ? 'left' : 'num') + '">' + (i < 2 ? esc(v) : v) + '</td>'; }).join('') + '</tr>';
    }
    splits.forEach(function (g) { html += row(g.cut + ' 分线', g.label, g); });
    html += row('—', '未录入中考', noZk);
    html += row('—', '全年级（剔除后）', total, true);
    html += '</tbody></table>';
    $('zkTable').innerHTML = html;
    var removed = state.cur.students.length - n;
    $('zkNote').textContent = '本校线 ' + state.zkCutA + ' · 对比校线 ' + state.zkCutB + ' · 统计 ' + n + ' 人' + (removed ? '（已剔除 ' + removed + ' 人）' : '');
  }

  function renderExclude() {
    var students = state.cur.students;
    var sp = specialStudents();
    $('exBulk').innerHTML =
      '<button type="button" class="btn small" id="bulkXfOn">休返全选</button>' +
      '<button type="button" class="btn small" id="bulkXfOff">休返全不选</button>' +
      '<button type="button" class="btn small" id="bulkZdOn">子弟全选</button>' +
      '<button type="button" class="btn small" id="bulkZdOff">子弟全不选</button>' +
      '<button type="button" class="btn small ghost" id="bulkReset">恢复默认</button>';

    var specialHtml = '<table class="grid"><thead><tr><th class="left">剔除</th><th class="left">学号</th><th class="left">姓名</th><th class="left">班级</th><th class="left">类别</th><th>中考分数</th><th class="left">默认</th></tr></thead><tbody>';
    if (!sp.length) {
      specialHtml += '<tr><td class="left" colspan="7"><span class="muted">没有识别到子弟 / 休返学生（可先导入学生档案）</span></td></tr>';
    }
    sp.forEach(function (s) {
      var on = state.specialMap[s.sid] === true;
      specialHtml += '<tr><td class="left"><input type="checkbox" data-sp="' + esc(s.sid) + '"' + (on ? ' checked' : '') + '></td>' +
        '<td class="left">' + esc(s.sid) + '</td><td class="left">' + esc(s.name) + '</td><td class="left">' + esc(s.clsLabel) + '</td>' +
        '<td class="left">' + esc(s.category || '') + '</td><td class="num">' + fmt(s.zhongkao) + '</td>' +
        '<td class="left">' + (defaultSpecialExclude(s) ? '剔除' : '保留') + '</td></tr>';
    });
    specialHtml += '</tbody></table>';
    $('exSpecial').innerHTML = specialHtml;

    var manual = students.filter(function (s) { return state.excludes[s.sid] && !isSpecial(s); });
    $('exChips').innerHTML = manual.length ? manual.map(function (s) {
      return '<span class="cls-chip">' + esc(s.name) + '（' + esc(s.clsLabel) + '）<button type="button" data-unex="' + esc(s.sid) + '" title="移出">×</button></span>';
    }).join('') : '<span class="muted">未手动指定其他学生</span>';

    var set = excludedSet();
    var res = Core.exclusionStats(students, set, ITEMS);
    $('excludeNote').textContent = '当前剔除 ' + res.excludedCount + ' 人（可在下表逐人勾选）';
    var head = ['指标', '全体', '剔除后', '被剔除', '差值（剔除后−全体）'];
    var rows = [
      ['人数', res.all.n, res.left.n, res.out.n, res.left.n - res.all.n],
      ['总分均分', res.all.avg, res.left.avg, res.out.avg, diff(res.all.avg, res.left.avg)],
      ['最高分', res.all.max, res.left.max, res.out.max, diff(res.all.max, res.left.max)],
      ['最低分', res.all.min, res.left.min, res.out.min, diff(res.all.min, res.left.min)]
    ];
    ['语数外总分'].concat(SUBJECTS).forEach(function (it) {
      rows.push([it + '均分', res.all.means[it], res.left.means[it], res.out.means[it], diff(res.all.means[it], res.left.means[it])]);
    });
    var html = '<table class="grid"><thead><tr>' + head.map(function (h) { return '<th class="' + (h === '指标' ? 'left' : '') + '">' + esc(h) + '</th>'; }).join('') + '</tr></thead><tbody>';
    rows.forEach(function (r) {
      var d = r[4];
      var isCount = r[0] === '人数';
      var f = isCount ? fmt : fmtScore;
      html += '<tr><td class="left">' + esc(r[0]) + '</td><td class="num">' + f(r[1]) + '</td><td class="num">' + f(r[2]) + '</td><td class="num">' + f(r[3]) + '</td><td class="num">' + (d === null || d === undefined ? '—' : (d > 0 ? '+' : '') + f(d)) + '</td></tr>';
    });
    html += '</tbody></table>';
    $('exTable').innerHTML = html;

    $('exSpecial').querySelectorAll('[data-sp]').forEach(function (cb) {
      cb.addEventListener('change', function () {
        state.specialMap[cb.getAttribute('data-sp')] = cb.checked;
        renderExclude();
        renderZkTable();
      });
    });
    $('exChips').querySelectorAll('[data-unex]').forEach(function (b) {
      b.addEventListener('click', function () {
        delete state.excludes[b.getAttribute('data-unex')];
        renderExclude();
      });
    });
    function setAll(kind, val) {
      specialStudents().forEach(function (s) {
        if (kind === 'xf' && isXiufan(s)) state.specialMap[s.sid] = val;
        if (kind === 'zd' && isZidi(s)) state.specialMap[s.sid] = val;
      });
      renderExclude();
      renderZkTable();
    }
    $('bulkXfOn').onclick = function () { setAll('xf', true); };
    $('bulkXfOff').onclick = function () { setAll('xf', false); };
    $('bulkZdOn').onclick = function () { setAll('zd', true); };
    $('bulkZdOff').onclick = function () { setAll('zd', false); };
    $('bulkReset').onclick = function () { initSpecialMap(); renderExclude(); renderZkTable(); };
  }

  function diff(a, b) {
    if (a === null || a === undefined || b === null || b === undefined) return null;
    return Math.round((b - a) * 10) / 10;
  }

  function renderExResults() {
    var q = $('exSearch').value.trim();
    if (!q) { $('exResults').innerHTML = ''; return; }
    var list = state.cur.students.filter(function (s) {
      return !state.excludes[s.sid] && (s.sid.indexOf(q) >= 0 || s.name.indexOf(q) >= 0);
    }).slice(0, 8);
    $('exResults').innerHTML = list.map(function (s) {
      return '<button type="button" class="btn small" data-add="' + esc(s.sid) + '">' + esc(s.name) + '（' + esc(s.clsLabel) + '）加入剔除</button>';
    }).join('') || '<span class="muted">没有匹配学生</span>';
    $('exResults').querySelectorAll('[data-add]').forEach(function (b) {
      b.addEventListener('click', function () {
        state.excludes[b.getAttribute('data-add')] = true;
        $('exSearch').value = '';
        $('exResults').innerHTML = '';
        renderExclude();
      });
    });
  }

  function importInfo(file) {
    readWorkbook(file).then(function (sheets) {
      var records = [];
      sheets.forEach(function (sh) {
        var p = Core.parseInfoRows(sh.rows || []);
        if (p.ok) records = records.concat(p.records);
      });
      if (!records.length) { toast('没有识别到「学号 + 中考总分/类别」', true); return; }
      var merged = {};
      records.forEach(function (r) {
        var cur = merged[r.sid] || { sid: r.sid, zhongkao: null, category: '' };
        if (r.zhongkao !== null && r.zhongkao !== undefined) cur.zhongkao = r.zhongkao;
        if (r.category) cur.category = r.category;
        merged[r.sid] = cur;
      });
      state.info = merged;
      var list = Object.keys(merged).map(function (k) { return merged[k]; });
      var res = Core.applyInfo(state.cur.students, list);
      if (state.prev) Core.applyInfo(state.prev.students, list);
      initSpecialMap();
      renderSpecial();
      toast('已匹配 ' + res.matched + ' 人（中考 ' + res.withZk + '，类别 ' + res.withCat + '）' +
        (res.unmatched.length ? '，另有 ' + res.unmatched.length + ' 人不在本次名单' : ''));
    }).catch(function (e) { toast('解析出错：' + (e && e.message ? e.message : e), true); });
  }

  function zkAoa() {
    var base = analysisBase();
    var splits = Core.zhongkaoSplits(base, [state.zkCutA, state.zkCutB], ITEMS);
    var total = Core.statBlock(base, ITEMS);
    var noZk = Core.statBlock(base.filter(function (s) { return !hasZk(s); }), ITEMS);
    var n = base.length;
    var ZK_ITEMS = ['语数外总分'].concat(SUBJECTS);
    var head = ['分线', '组别', '人数', '占比%', '中考均分', '本次总分均分', '最高分', '最低分'].concat(ZK_ITEMS.map(function (i) { return i + '均分'; }));
    var rows = [head];
    function pct(b) { return n ? Math.round(b.n / n * 1000) / 10 : 0; }
    function add(cut, b) {
      rows.push([cut, b.label || '', b.n, pct(b), b.zkAvg, b.avg, b.max, b.min].concat(ZK_ITEMS.map(function (i) { return b.means[i]; })));
    }
    splits.forEach(function (g) { add(g.cut + ' 分线', g); });
    noZk.label = '未录入中考';
    add('—', noZk);
    total.label = '全年级（剔除后）';
    var rowTotal = ['—', total.label, total.n, 100, total.zkAvg, total.avg, total.max, total.min].concat(ZK_ITEMS.map(function (i) { return total.means[i]; }));
    rows.push(rowTotal);
    return rows;
  }

  function excludeAoa() {
    var set = excludedSet();
    var res = Core.exclusionStats(state.cur.students, set, ITEMS);
    var rows = [['指标', '全体', '剔除后', '被剔除', '差值'],
      ['人数', res.all.n, res.left.n, res.out.n, res.left.n - res.all.n],
      ['总分均分', res.all.avg, res.left.avg, res.out.avg, diff(res.all.avg, res.left.avg)],
      ['最高分', res.all.max, res.left.max, res.out.max, diff(res.all.max, res.left.max)],
      ['最低分', res.all.min, res.left.min, res.out.min, diff(res.all.min, res.left.min)]];
    ['语数外总分'].concat(SUBJECTS).forEach(function (it) { rows.push([it + '均分', res.all.means[it], res.left.means[it], res.out.means[it], diff(res.all.means[it], res.left.means[it])]); });
    return rows;
  }

  function zkMissingAoa() {
    var rows = [['学号', '姓名', '班级', '类别', '说明']];
    var set = excludedSet();
    state.cur.students.forEach(function (s) {
      if (hasZk(s)) return;
      rows.push([s.sid, s.name, s.clsLabel, s.category || '', set[s.sid] ? '无中考分数（已剔除）' : '无中考分数（计入未录入）']);
    });
    if (rows.length === 1) rows.push(['—', '—', '—', '—', '无']);
    return rows;
  }
  function excludeListAoa() {
    var rows = [['学号', '姓名', '班级', '类别', '来源']];
    var set = excludedSet();
    state.cur.students.forEach(function (s) {
      if (!set[s.sid]) return;
      var reason;
      if (state.excludes[s.sid]) reason = '手动';
      else if (isXiufan(s)) reason = '休返';
      else if (isZidi(s)) reason = hasZk(s) ? '子弟（勾选）' : '子弟无中考分';
      else reason = '其他';
      rows.push([s.sid, s.name, s.clsLabel, s.category || '', reason]);
    });
    if (rows.length === 1) rows.push(['—', '—', '—', '—', '无']);
    return rows;
  }

  function exportSpecial() {
    var wb = XLSX.utils.book_new();
    if (withZhongkao()) {
      addSheet(wb, '中考分组统计', zkAoa());
      addSheet(wb, '中考未录入', zkMissingAoa());
    }
    if (withCategory() || Object.keys(state.excludes).length) {
      addSheet(wb, '剔除后统计', excludeAoa());
      addSheet(wb, '剔除名单', excludeListAoa());
    }
    if (!wb.SheetNames.length) { toast('还没有中考成绩或类别数据', true); return; }
    downloadWb(wb, safeName() + '_专项统计.xlsx');
  }

  /* ---------------- page: awards ---------------- */
  function renderAwards() {
    if (!state.cur) return;
    if (document.activeElement !== $('awardPct')) $('awardPct').value = state.awardPct;
    var res = Core.awards(state.cur.students, state.tiers, state.awardPct);
    var total = state.cur.students.length;
    $('awardPreview').textContent = '全年级 ' + total + ' 人 × ' + state.awardPct + '% ≈ ' + Math.round(total * state.awardPct / 100) + ' 人（各层次同比例）';
    $('awardTotalCount').textContent = '共 ' + res.excellent.length + ' 人';

    var eh = '<table class="grid"><thead><tr><th class="left">层次</th><th>年级位次</th><th class="left">学号</th><th class="left">姓名</th><th class="left">班级</th><th>总分</th></tr></thead><tbody>';
    res.excellent.forEach(function (s) {
      eh += '<tr><td class="left">' + esc(s.tier) + '</td><td class="num">' + fmt(s.rank) + '</td><td class="left">' + esc(s.sid) + '</td><td class="left">' + esc(s.name) + '</td><td class="left">' + esc(s.clsLabel) + '</td><td class="num">' + fmtScore(s.total) + '</td></tr>';
    });
    eh += '</tbody></table>';
    $('awardExcellent').innerHTML = eh;

    var th = '<table class="grid"><thead><tr><th class="left">学科</th><th>最高分</th><th class="left">姓名（班级）</th></tr></thead><tbody>';
    res.top.forEach(function (t) {
      var names = t.students.map(function (s) { return esc(s.name) + '（' + esc(s.clsLabel) + '）'; }).join('、');
      th += '<tr><td class="left">' + esc(t.item) + '</td><td class="num">' + fmt(t.score) + '</td><td class="left">' + (names || '—') + '</td></tr>';
    });
    th += '</tbody></table>';
    $('awardTop').innerHTML = th;
  }

  /* ---------------- page: export ---------------- */
  /* ---------------- export builders ---------------- */
  function setCols(ws, aoa) {
    var widths = [];
    var n = aoa[0] ? aoa[0].length : 0;
    for (var c = 0; c < n; c++) {
      var w = 8;
      for (var r = 0; r < Math.min(aoa.length, 80); r++) {
        var v = aoa[r][c];
        var s = v === null || v === undefined ? '' : String(v);
        var len = 0;
        for (var k = 0; k < s.length; k++) len += s.charCodeAt(k) > 255 ? 2 : 1;
        if (len > w) w = len;
      }
      widths.push({ wch: Math.max(6, Math.min(30, w + 1)) });
    }
    ws['!cols'] = widths;
    return ws;
  }

  var SCORE_HEADER_RE = /均分|平均|得分|分数|最高|最低|总分|差值/;
  var RANK_HEADER_RE = /排名|位次|班次|级次|名次|人数|占比|进退|序号/;
  function isScoreHeader(h) {
    var s = String(h === null || h === undefined ? '' : h).trim();
    if (!s) return false;
    if (RANK_HEADER_RE.test(s)) return false;
    if (SCORE_HEADER_RE.test(s)) return true;
    return ITEMS.indexOf(s) >= 0;
  }
  function applyScoreFormats(ws, aoa, name) {
    var heads = aoa[0] || [];
    var cols = [];
    heads.forEach(function (h, c) { if (isScoreHeader(h)) cols.push(c); });
    if (name === '剔除后统计') {
      for (var r = 1; r < aoa.length; r++) {
        if (aoa[r][0] === '人数') continue;
        for (var c = 1; c <= 4; c++) cols.push(c);
      }
      cols = cols.filter(function (v, i, a) { return a.indexOf(v) === i; });
    }
    if (!cols.length) return;
    for (var row = 1; row < aoa.length; row++) {
      for (var k = 0; k < cols.length; k++) {
        var cell = ws[XLSX.utils.encode_cell({ r: row, c: cols[k] })];
        if (cell && cell.t === 'n' && !Number.isInteger(cell.v)) cell.z = '0.00';
      }
    }
  }
  function addSheet(wb, name, aoa) {
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    applyScoreFormats(ws, aoa, name);
    setCols(ws, aoa);
    var base = String(name).slice(0, 31);
    var finalName = base, i = 2;
    while (wb.SheetNames.indexOf(finalName) >= 0) { finalName = String(base).slice(0, 28) + '(' + i + ')'; i++; }
    XLSX.utils.book_append_sheet(wb, ws, finalName);
    return ws;
  }
  function downloadWb(wb, filename) {
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    var blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
  }
  function safeName() {
    var n = (state.examName || '成绩汇总').replace(/[\\/:*?"<>|]/g, '_').trim();
    return n || '成绩汇总';
  }
  function mainAoa() {
    var pack = mainRows();
    return [pack.headers].concat(pack.rows.map(function (r) {
      return r.map(function (v) { return v === 'new' ? '新进' : (v === undefined ? null : v); });
    }));
  }
  function classAoa(cls) {
    var pack = mainRows();
    var rows = [];
    state.cur.students.forEach(function (s, i) {
      if (s.clsLabel === cls) rows.push(pack.rows[i].map(function (v) { return v === 'new' ? '新进' : (v === undefined ? null : v); }));
    });
    return [pack.headers].concat(rows);
  }
  function progressAoa() {
    var head = ['学号', '姓名', '班级'];
    ITEMS.forEach(function (item) {
      head.push(item + '上次班次', item + '本次班次', item + '班次进退');
      head.push(item + '上次级次', item + '本次级次', item + '级次进退');
    });
    var rows = [head];
    state.cur.students.forEach(function (s) {
      var r = [s.sid, s.name, s.clsLabel];
      ITEMS.forEach(function (item) {
        var pv = Core.progressValues(s, state.prevMap, item);
        var pr = state.prevMap && state.prevMap[s.sid] && state.prevMap[s.sid].ranks[item] ? state.prevMap[s.sid].ranks[item] : null;
        var cr = s.ranks[item] || {};
        r.push(pr ? pr.cls : null, cr.cls === undefined ? null : cr.cls, pv.cls);
        r.push(pr ? pr.grade : null, cr.grade === undefined ? null : cr.grade, pv.grade);
      });
      rows.push(r);
    });
    return rows;
  }
  function meansAoa() {
    var cm = Core.classMeans(state.cur.students, ITEMS);
    var rk = Core.classRankTable(cm.means, cm.classes, ITEMS);
    var head = ['班级', '层次', '人数'];
    ITEMS.forEach(function (item) { head.push(item, '年级排名', '层次排名'); });
    var rows = [head];
    cm.classes.forEach(function (c) {
      var tier = Core.tierOfClass(state.tiers, c.label);
      var row = [c.label, tier ? tier.name : '未归类', c.count];
      ITEMS.forEach(function (item) {
        var tr = null;
        if (tier) {
          var vals = tier.classes.map(function (x) { return cm.means[x] ? cm.means[x][item] : null; });
          var rr = Core.rankValues(vals);
          var k = tier.classes.indexOf(c.label);
          tr = k >= 0 ? rr[k] : null;
        }
        row.push(cm.means[c.label][item], rk[c.label][item], tr);
      });
      rows.push(row);
    });
    return rows;
  }
  function diffAoa() {
    var groups = Core.diffTable(state.cur.students, state.prev ? state.prev.students : null, state.tiers, ITEMS, state.teachers);
    var head = ['层次', '班级', '班主任'];
    ITEMS.forEach(function (item) { head.push(item, '组内排名', '与组均差值', '上次名次', '进退', '代课教师'); });
    var rows = [head];
    groups.forEach(function (g) {
      g.rows.forEach(function (row) {
        var r = [g.name, row.cls, getHeadTeacher(row.cls)];
        ITEMS.forEach(function (item) {
          var d = row.items[item];
          r.push(d.mean, d.rank, d.delta, d.prevRank, d.progress === null ? null : d.progress, getTeacher(row.cls, item));
        });
        rows.push(r);
      });
      var avg = [g.name, g.avg.cls, ''];
      ITEMS.forEach(function (item) { avg.push(g.avg.items[item].mean, null, 0, null, null, ''); });
      rows.push(avg);
    });
    return rows;
  }
  function analysisAoa() {
    var bands = Core.bandDefs(state.bands);
    var ANALY_ITEMS = ['语数外总分'].concat(SUBJECTS);
    var head = ['班级', '层次', '应考人数', '实考人数', '总分平均', '最高分', '最低分', '总分层次名次'];
    ANALY_ITEMS.forEach(function (item) { head.push(item + '均分', item + '层次名次'); });
    head.push('存在问题', '改进措施', '教学建议');
    var rows = [head];
    dataClasses().forEach(function (cls) {
      var a = Core.analyzeClass(state.cur.students, cls, state.tiers, bands, ITEMS);
      var t = state.texts[cls] || Core.autoText(a);
      var enrolled = (state.cur.enrolled && state.cur.enrolled[cls]) || a.count;
      var r = [cls, a.tier, enrolled, a.count, a.avg, a.max, a.min, a.tierRanks['总分']];
      ANALY_ITEMS.forEach(function (item) { r.push(a.means[item], a.tierRanks[item]); });
      r.push(t.problem, t.measure, t.advice);
      rows.push(r);
    });
    return rows;
  }
  function bandsAoa() {
    var bands = Core.bandDefs(state.bands);
    var head = ['班级', '层次', '实考人数'].concat(bands.map(function (b) { return b.label; }));
    var rows = [head];
    dataClasses().forEach(function (cls) {
      var a = Core.analyzeClass(state.cur.students, cls, state.tiers, bands, ITEMS);
      var r = [cls, a.tier, a.count].concat(a.bands.map(function (b) { return b.count; }));
      rows.push(r);
    });
    return rows;
  }
  function excellentAoa() {
    var res = Core.awards(state.cur.students, state.tiers, state.awardPct);
    return [['层次', '年级位次', '学号', '姓名', '班级', '总分']].concat(res.excellent.map(function (s) {
      return [s.tier, s.rank, s.sid, s.name, s.clsLabel, s.total];
    }));
  }
  function topAoa() {
    var res = Core.awards(state.cur.students, state.tiers, state.awardPct);
    return [['学科', '最高分', '姓名', '班级']].concat(res.top.reduce(function (arr, t) {
      if (!t.students.length) arr.push([t.item, t.score, '', '']);
      t.students.forEach(function (s) { arr.push([t.item, t.score, s.name, s.clsLabel]); });
      return arr;
    }, []));
  }
  function awardSheets(wb) {
    addSheet(wb, '总分优秀', excellentAoa());
    addSheet(wb, '单科第一', topAoa());
  }
  function absentAoa() {
    var rows = [['学号', '姓名', '班级', '说明']];
    state.cur.absent.forEach(function (s) { rows.push([s.sid, s.name, s.clsLabel, '缺考 / 已删除']); });
    (state.cur.orphans || []).forEach(function (s) { rows.push([s.sid, s.name, s.clsLabel, '仅在班级分表出现，按总表剔除']); });
    if (rows.length === 1) rows.push(['—', '—', '—', '无']);
    return rows;
  }
  function issuesAoa() {
    var rows = [['类型', '说明']];
    state.cur.warnings.forEach(function (w) { rows.push(['提示', w]); });
    if (state.cur.orphans) state.cur.orphans.forEach(function (o) { rows.push(['分表多余学生', o.sheet + '：' + o.sid + ' ' + o.name + ' ' + o.clsLabel]); });
    return rows;
  }


  function diffStyledSheet(wb) {
    var groups = Core.diffTable(state.cur.students, state.prev ? state.prev.students : null, state.tiers, ITEMS, state.teachers);
    var h1 = new Array(66).fill(null);
    var h2 = new Array(66).fill('');
    h1[0] = '班级';
    h2[0] = '';
    var layout = [];
    var pos = 1;
    var SUBJ = ['语文', '数学', '英语', '物理', '化学', '生物', '历史', '地理', '政治'];
    ITEMS.forEach(function (item) {
      var cols, labels;
      if (item === '总分') { cols = 6; labels = ['班主任', '总分', '排名', '差值', '上次名次', '进退']; }
      else if (item === '语数外总分') { cols = 5; labels = ['均分', '排名', '差值', '上次名次', '进退']; }
      else { cols = 6; labels = ['均分', '排名', '差值', '上次名次', '进退', '教师']; }
      h1[pos] = item === '语数外总分' ? '语数外' : item;
      labels.forEach(function (l, i) { h2[pos + i] = l; });
      layout.push({ item: item, start: pos, cols: cols });
      pos += cols;
    });
    var aoa = [h1, h2];
    var meanCols = [];
    groups.forEach(function (g) {
      g.rows.forEach(function (row) {
        var r = new Array(66).fill('');
        r[0] = row.cls;
        layout.forEach(function (L) {
          var d = row.items[L.item];
          if (L.item === '总分') {
            r[L.start] = getHeadTeacher(row.cls);
            r[L.start + 1] = d.mean;
            r[L.start + 2] = d.rank;
            r[L.start + 3] = d.delta;
            r[L.start + 4] = d.prevRank;
            r[L.start + 5] = d.progress;
          } else if (L.item === '语数外总分') {
            r[L.start] = d.mean;
            r[L.start + 1] = d.rank;
            r[L.start + 2] = d.delta;
            r[L.start + 3] = d.prevRank;
            r[L.start + 4] = d.progress;
          } else {
            r[L.start] = d.mean;
            r[L.start + 1] = d.rank;
            r[L.start + 2] = d.delta;
            r[L.start + 3] = d.prevRank;
            r[L.start + 4] = d.progress;
            r[L.start + 5] = getTeacher(row.cls, L.item);
          }
        });
        aoa.push(r);
      });
      var a = new Array(66).fill('');
      a[0] = '平均';
      layout.forEach(function (L) {
        var m = g.avg.items[L.item].mean;
        if (L.item === '总分') { a[L.start + 1] = m; a[L.start + 3] = 0; }
        else if (L.item === '语数外总分') { a[L.start] = m; a[L.start + 2] = 0; }
        else { a[L.start] = m; a[L.start + 2] = 0; }
      });
      aoa.push(a);
    });

    var ws = addSheet(wb, '班级差值表', aoa);
    var merges = [{ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }];
    layout.forEach(function (L) { merges.push({ s: { r: 0, c: L.start }, e: { r: 0, c: L.start + L.cols - 1 } }); });
    ws['!merges'] = merges;
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: 65 } });
    var widths = [{ wch: 10 }];
    layout.forEach(function (L) {
      if (L.item === '总分') widths.push({ wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 8 });
      else if (L.item === '语数外总分') widths.push({ wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 8 });
      else widths.push({ wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 8 }, { wch: 10 });
    });
    ws['!cols'] = widths;
    ws['!rows'] = [{ hpt: 26 }, { hpt: 26 }];

    layout.forEach(function (L) {
      if (L.item === '总分') meanCols.push(L.start + 1, L.start + 3);
      else meanCols.push(L.start, L.start + 2);
    });
    for (var r2 = 2; r2 < aoa.length; r2++) {
      meanCols.forEach(function (c) {
        var cell = ws[XLSX.utils.encode_cell({ r: r2, c: c })];
        if (cell && cell.t === 'n' && !Number.isInteger(cell.v)) cell.z = '0.00';
      });
    }
    return ws;
  }

  function buildAllWorkbook() {
    var wb = XLSX.utils.book_new();
    addSheet(wb, '年级总表', mainAoa());
    addSheet(wb, '进退分析', progressAoa());
    dataClasses().forEach(function (cls) { addSheet(wb, cls, classAoa(cls)); });
    addSheet(wb, '班级均分', meansAoa());
    diffStyledSheet(wb);
    awardSheets(wb);
    addSheet(wb, '成绩分析', analysisAoa());
    addSheet(wb, '分数段统计', bandsAoa());
    if (withZhongkao()) {
      addSheet(wb, '中考分组统计', zkAoa());
      addSheet(wb, '中考未录入', zkMissingAoa());
    }
    if (withCategory() || Object.keys(state.excludes).length) {
      addSheet(wb, '剔除后统计', excludeAoa());
      addSheet(wb, '剔除名单', excludeListAoa());
    }
    addSheet(wb, '缺考名单', absentAoa());
    addSheet(wb, '数据质检', issuesAoa());
    return wb;
  }
  function buildClassFile(cls, withAnalysis, withBands) {
    var wb = XLSX.utils.book_new();
    addSheet(wb, cls, classAoa(cls));
    var bands = Core.bandDefs(state.bands);
    var a = Core.analyzeClass(state.cur.students, cls, state.tiers, bands, ITEMS);
    if (withAnalysis) {
      var t = state.texts[cls] || Core.autoText(a);
      var enrolled = (state.cur.enrolled && state.cur.enrolled[cls]) || a.count;
      addSheet(wb, '成绩分析', [
        ['班级', '层次', '应考人数', '实考人数', '总分平均', '最高分', '最低分', '存在问题', '改进措施', '教学建议'],
        [cls, a.tier, enrolled, a.count, a.avg, a.max, a.min, t.problem, t.measure, t.advice]
      ]);
    }
    if (withBands) {
      addSheet(wb, '分数段', [['分数段', '人数', '占比%']].concat(a.bands.map(function (b) { return [b.label, b.count, b.pct]; })));
    }
    return wb;
  }
  function buildClassWorkbook(cls) {
    return buildClassFile(cls, true, true);
  }


  function wbFromAoa(sheet, aoa) {
    var wb = XLSX.utils.book_new();
    addSheet(wb, sheet, aoa);
    return wb;
  }
  function exportFile(sheet, aoa, suffix) {
    return { name: safeName() + '_' + suffix + '.xlsx', wb: wbFromAoa(sheet, aoa) };
  }

  var EXPORT_MODULES = [
    { id: 'grade', label: '年级总表', desc: '全部学生成绩、位次、进退', files: function () { return [exportFile('年级总表', mainAoa(), '年级总表')]; } },
    { id: 'progress', label: '进退分析', desc: '上次/本次/进退（班级与年级）', files: function () { return [exportFile('进退分析', progressAoa(), '进退分析')]; } },
    { id: 'classes', label: '班级分表（独立文件）', desc: '勾选班级，每班一个独立 xlsx（单个 sheet）', defaultOff: true, files: function () {
      return dataClasses().filter(function (cls) { return state.exportClasses[cls] !== false; }).map(function (cls) {
        return { name: cls + '.xlsx', wb: buildClassFile(cls, state.exportClassAnalysis, state.exportClassBands) };
      });
    } },
    { id: 'means', label: '班级均分统计', desc: '各科均分、年级排名、层次排名', files: function () { return [exportFile('班级均分', meansAoa(), '班级均分')]; } },
    { id: 'diff', label: '班级差值表', desc: '按层次分组，两行合并表头（与模板一致）', files: function () {
      var wb = XLSX.utils.book_new();
      diffStyledSheet(wb);
      return [{ name: safeName() + '_班级差值表.xlsx', wb: wb }];
    } },
    { id: 'excellent', label: '总分优秀', desc: '分层等比例优秀名单', files: function () { return [exportFile('总分优秀', excellentAoa(), '总分优秀')]; } },
    { id: 'top', label: '单科第一', desc: '各科年级第一（同分并列）', files: function () { return [exportFile('单科第一', topAoa(), '单科第一')]; } },
    { id: 'analysis', label: '成绩分析', desc: '各班分析表与文字板块', files: function () { return [exportFile('成绩分析', analysisAoa(), '成绩分析')]; } },
    { id: 'bands', label: '分数段统计', desc: '各班总分分数段人数', files: function () { return [exportFile('分数段统计', bandsAoa(), '分数段统计')]; } },
    { id: 'zk', label: '中考分组统计', desc: '688 / 699 两条线以上以下', avail: function () { return withZhongkao() > 0; }, files: function () { return [exportFile('中考分组统计', zkAoa(), '中考分组统计')]; } },
    { id: 'zkMiss', label: '中考未录入', desc: '无中考分数的学生清单', avail: function () { return withZhongkao() > 0; }, files: function () { return [exportFile('中考未录入', zkMissingAoa(), '中考未录入')]; } },
    { id: 'exclude', label: '剔除后统计', desc: '全体 / 剔除后 / 被剔除 / 差值', avail: function () { return withCategory() > 0 || Object.keys(state.excludes).length > 0; }, files: function () { return [exportFile('剔除后统计', excludeAoa(), '剔除后统计')]; } },
    { id: 'excludeList', label: '剔除名单', desc: '休返、无中考分子弟、手动剔除', avail: function () { return withCategory() > 0 || Object.keys(state.excludes).length > 0; }, files: function () { return [exportFile('剔除名单', excludeListAoa(), '剔除名单')]; } },
    { id: 'absent', label: '缺考名单', desc: '缺考 / 分表多余学生', files: function () { return [exportFile('缺考名单', absentAoa(), '缺考名单')]; } },
    { id: 'issues', label: '数据质检', desc: '导入提示与异常清单', files: function () { return [exportFile('数据质检', issuesAoa(), '数据质检')]; } }
  ];

  function moduleFiles(id) {
    for (var i = 0; i < EXPORT_MODULES.length; i++) {
      if (EXPORT_MODULES[i].id === id) return EXPORT_MODULES[i].files();
    }
    return [];
  }

  function downloadFiles(files) {
    if (!files.length) { toast('没有可下载的文件', true); return; }
    if (files.length === 1) { downloadWb(files[0].wb, files[0].name); toast('已下载 ' + files[0].name); return; }
    var i = 0;
    (function next() {
      if (i >= files.length) { toast('已下载 ' + files.length + ' 个文件'); return; }
      var f = files[i++];
      downloadWb(f.wb, f.name);
      toast('正在下载 ' + i + ' / ' + files.length + '：' + f.name);
      setTimeout(next, 600);
    })();
  }

  function isModuleSelected(m) {
    var v = state.exportSel[m.id];
    if (v === undefined) return !m.defaultOff;
    return !!v;
  }

  function renderExportPage() {
    if (!state.cur || !$('exportList')) return;
    var mods = EXPORT_MODULES.filter(function (m) { return !m.avail || m.avail(); });
    var html = mods.map(function (m) {
      var on = isModuleSelected(m);
      return '<div class="export-row">' +
        '<input type="checkbox" data-ex="' + m.id + '"' + (on ? ' checked' : '') + '>' +
        '<div><b>' + esc(m.label) + '</b><small>' + esc(m.desc || '') + '</small></div>' +
        '<button type="button" class="btn small" data-dl="' + m.id + '">单独下载</button>' +
        '</div>';
    }).join('');
    var classes = dataClasses();
    if (classes.length) {
      html += '<div class="export-class-panel">' +
        '<div class="toolbar"><span class="muted">班级分表：勾选要导出的班级（每班一个独立文件）</span><span class="spacer"></span>' +
        '<label class="class-check"><input type="checkbox" id="clsWithAnalysis"' + (state.exportClassAnalysis ? ' checked' : '') + '> 含成绩分析</label>' +
        '<label class="class-check"><input type="checkbox" id="clsWithBands"' + (state.exportClassBands ? ' checked' : '') + '> 含分数段统计</label>' +
        '<button type="button" class="btn small" id="clsAll">全选</button>' +
        '<button type="button" class="btn small" id="clsNone">全不选</button></div>' +
        '<div class="chip-row">' +
        classes.map(function (c) {
          var on = state.exportClasses[c] !== false;
          return '<label class="class-check"><input type="checkbox" data-cls="' + esc(c) + '"' + (on ? ' checked' : '') + '> ' + esc(c) + '</label>';
        }).join('') +
        '</div></div>';
    }
    $('exportList').innerHTML = html;
    $('exportList').querySelectorAll('[data-ex]').forEach(function (cb) {
      cb.addEventListener('change', function () {
        state.exportSel[cb.getAttribute('data-ex')] = cb.checked;
      });
    });
    $('exportList').querySelectorAll('[data-dl]').forEach(function (b) {
      b.addEventListener('click', function () {
        downloadFiles(moduleFiles(b.getAttribute('data-dl')));
      });
    });
    $('exportList').querySelectorAll('[data-cls]').forEach(function (cb) {
      cb.addEventListener('change', function () {
        state.exportClasses[cb.getAttribute('data-cls')] = cb.checked;
      });
    });
    if ($('clsWithAnalysis')) $('clsWithAnalysis').onchange = function () { state.exportClassAnalysis = this.checked; };
    if ($('clsWithBands')) $('clsWithBands').onchange = function () { state.exportClassBands = this.checked; };
    if ($('clsAll')) $('clsAll').onclick = function () { dataClasses().forEach(function (c) { state.exportClasses[c] = true; }); renderExportPage(); };
    if ($('clsNone')) $('clsNone').onclick = function () { dataClasses().forEach(function (c) { state.exportClasses[c] = false; }); renderExportPage(); };
  }

  function downloadSelected() {
    var mods = EXPORT_MODULES.filter(function (m) { return (!m.avail || m.avail()) && isModuleSelected(m); });
    var files = [];
    mods.forEach(function (m) { files = files.concat(moduleFiles(m.id)); });
    downloadFiles(files);
  }

  function exportAll() { downloadWb(buildAllWorkbook(), safeName() + '_全部结果.xlsx'); toast('已生成全部结果'); }
  function exportGrade() { var wb = XLSX.utils.book_new(); addSheet(wb, '年级总表', mainAoa()); downloadWb(wb, safeName() + '_年级总表.xlsx'); }
  function exportMeans() { var wb = XLSX.utils.book_new(); addSheet(wb, '班级均分', meansAoa()); downloadWb(wb, safeName() + '_班级均分.xlsx'); }
  function exportDiff() { var wb = XLSX.utils.book_new(); diffStyledSheet(wb); downloadWb(wb, safeName() + '_班级差值表.xlsx'); }
  function exportAwards() { var wb = XLSX.utils.book_new(); awardSheets(wb); downloadWb(wb, safeName() + '_奖励名单.xlsx'); }
  function exportAnalysis() { var wb = XLSX.utils.book_new(); addSheet(wb, '成绩分析', analysisAoa()); addSheet(wb, '分数段统计', bandsAoa()); downloadWb(wb, safeName() + '_成绩分析.xlsx'); }
  function exportOneClass(cls) { downloadWb(buildClassWorkbook(cls), cls + '.xlsx'); }
  function exportClasses() {
    var classes = dataClasses();
    if (!classes.length) return;
    var i = 0;
    (function next() {
      if (i >= classes.length) { toast('已导出 ' + classes.length + ' 个班级文件'); return; }
      var cls = classes[i++];
      exportOneClass(cls);
      toast('正在导出 ' + cls + '（' + i + '/' + classes.length + '）');
      setTimeout(next, 550);
    })();
  }

  /* ---------------- dialog ---------------- */
  function openDialog(title, body, foot) {
    $('dialogTitle').textContent = title;
    $('dialogBody').innerHTML = body;
    $('dialogFoot').innerHTML = foot || '<button class="btn" data-close="1">关闭</button>';
    $('overlay').hidden = false;
    $('dialogFoot').querySelectorAll('[data-close]').forEach(function (b) { b.addEventListener('click', closeDialog); });
  }
  function closeDialog() { $('overlay').hidden = true; }

  /* ---------------- bind ---------------- */
  function bind() {
    document.querySelectorAll('.nav-item').forEach(function (b) {
      b.addEventListener('click', function () { go(b.getAttribute('data-page')); });
    });

    $('examName').value = state.examName;
    $('examName').addEventListener('input', function () { state.examName = this.value; clearTimeout(this._t); this._t = setTimeout(saveCfg, 400); });

    $('btnReset').addEventListener('click', function () {
      openDialog('重置', '<p>将清空已导入的数据与本地配置（班级层次、教师、分析文字）。确定吗？</p>',
        '<button class="btn" data-close="1">取消</button><button class="btn primary" id="confirmReset">确定重置</button>');
      $('confirmReset').addEventListener('click', function () {
        try { localStorage.removeItem(LS_KEY); } catch (e) {}
        location.reload();
      });
    });

    ['cur', 'prev'].forEach(function (kind) {
      var drop = $('drop' + (kind === 'cur' ? 'Cur' : 'Prev'));
      var input = $('file' + (kind === 'cur' ? 'Cur' : 'Prev'));
      drop.addEventListener('click', function () { input.click(); });
      input.addEventListener('change', function () { if (input.files[0]) ingest(kind, input.files[0]); input.value = ''; });
      drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
      drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
      drop.addEventListener('drop', function (e) {
        e.preventDefault();
        drop.classList.remove('over');
        var f = e.dataTransfer.files && e.dataTransfer.files[0];
        if (f) ingest(kind, f);
      });
    });

    $('btnGotoStudents').addEventListener('click', function () { go('students'); });
    $('btnExportAll').addEventListener('click', exportAll);
    $('btnAddStudent').addEventListener('click', addStudentDialog);

    $('studentSearch').addEventListener('input', function () { state.ui.stuSearch = this.value; state.ui.stuPage = 1; renderStudents(); });
    $('studentClassFilter').addEventListener('change', function () { state.ui.stuClass = this.value; state.ui.stuPage = 1; renderStudents(); });

    $('btnAddTier').addEventListener('click', function () {
      var name = $('newTierName').value.trim();
      if (!name) { toast('请输入层次名称', true); return; }
      if (state.tiers.some(function (t) { return t.name === name; })) { toast('层次已存在', true); return; }
      state.tiers.push({ name: name, classes: [] });
      $('newTierName').value = '';
      saveCfg(); renderTiers(); renderStats(); renderAnalysis(); renderAwards();
    });
    $('btnResetTiers').addEventListener('click', function () {
      state.tiers = clone(DEFAULT_TIERS);
      saveCfg(); renderTiers(); renderStats(); renderAnalysis(); renderAwards();
      toast('已恢复默认层次');
    });

    document.querySelectorAll('#statTabs .subtab').forEach(function (b) {
      b.addEventListener('click', function () { state.ui.statTab = b.getAttribute('data-tab'); renderStats(); });
    });
    $('gradeSearch').addEventListener('input', function () { state.ui.gradeSearch = this.value; state.ui.gradePage = 1; renderGradeTable(); });
    $('gradeClassFilter').addEventListener('change', function () { state.ui.gradeClass = this.value; state.ui.gradePage = 1; renderGradeTable(); });
    $('gradePageSize').addEventListener('change', function () { state.ui.gradeSize = Number(this.value); state.ui.gradePage = 1; renderGradeTable(); });
    $('classSelect').addEventListener('change', function () { state.ui.classSel = this.value; renderClassTable(); });
    $('analysisClass').addEventListener('change', function () { state.ui.analysisCls = this.value; renderAnalysis(); });

    $('btnImportInfo').addEventListener('click', function () { $('fileInfo').click(); });
    $('fileInfo').addEventListener('change', function () {
      if (this.files && this.files[0]) importInfo(this.files[0]);
      this.value = '';
    });
    function bindCut(input, key) {
      $(input).addEventListener('change', function () {
        var v = Number(this.value);
        if (isFinite(v)) { state[key] = v; saveCfg(); if (withZhongkao()) renderZkTable(); }
      });
    }
    bindCut('zkCutA', 'zkCutA');
    bindCut('zkCutB', 'zkCutB');
    $('exSearch').addEventListener('input', renderExResults);
    $('btnExportSpecial').addEventListener('click', exportSpecial);
    $('awardPct').addEventListener('input', function () {
      var v = Number(this.value);
      if (isFinite(v) && v > 0 && v <= 100) { state.awardPct = v; saveCfg(); renderAwards(); }
    });

    $('btnExportGrade').addEventListener('click', exportGrade);
    $('btnExportClass').addEventListener('click', function () { exportOneClass(state.ui.classSel); });
    $('btnExportMeans').addEventListener('click', exportMeans);
    $('btnExportDiff').addEventListener('click', exportDiff);
    $('btnImportTeachers').addEventListener('click', function () { $('fileTeachers').click(); });
    $('fileTeachers').addEventListener('change', function () {
      if (this.files && this.files[0]) importTeachers(this.files[0]);
      this.value = '';
    });
    $('btnExportAwards').addEventListener('click', exportAwards);
    $('btnExportAnalysis').addEventListener('click', function () { if (state.ui.analysisCls) exportOneClass(state.ui.analysisCls); });
    $('btnPrintAnalysis').addEventListener('click', function () {
      document.body.classList.add('printing');
      window.print();
      setTimeout(function () { document.body.classList.remove('printing'); }, 500);
    });

    $('exSelectAll').addEventListener('click', function () {
      EXPORT_MODULES.forEach(function (m) { if (!m.avail || m.avail()) state.exportSel[m.id] = true; });
      renderExportPage();
    });
    $('exSelectNone').addEventListener('click', function () {
      EXPORT_MODULES.forEach(function (m) { state.exportSel[m.id] = false; });
      renderExportPage();
    });
    $('exDownloadSel').addEventListener('click', downloadSelected);
    $('exMerge').addEventListener('click', exportAll);

    $('dialogClose').addEventListener('click', closeDialog);
    $('overlay').addEventListener('click', function (e) { if (e.target === $('overlay')) closeDialog(); });
  }

  /* ---------------- test/debug hook ---------------- */
  window.GradeApp = {
    state: state,
    rebuild: rebuild,
    renderAll: renderAll,
    go: go,
    buildAllWorkbook: buildAllWorkbook,
    buildClassWorkbook: buildClassWorkbook,
    mainAoa: mainAoa,
    meansAoa: meansAoa,
    diffAoa: diffAoa,
    analysisAoa: analysisAoa,
    progressAoa: progressAoa,
    bandsAoa: bandsAoa,
    moduleFiles: moduleFiles,
    EXPORT_MODULES: EXPORT_MODULES,
    awards: function (pct) { return Core.awards(state.cur.students, state.tiers, pct); }
  };

  /* ---------------- boot ---------------- */
  loadCfg();
  bind();
  updateStatus();
  go('import');
})();
