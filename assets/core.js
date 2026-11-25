/* 成绩汇总工作台 · 核心计算（纯函数，可在浏览器与 Node 中运行） */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) { module.exports = factory(); }
  else { root.GradeCore = factory(); }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var SUBJECTS = ['语文', '数学', '英语', '物理', '化学', '生物', '历史', '地理', '政治'];
  var ITEMS = ['总分', '语数外总分'].concat(SUBJECTS);
  var THREE = ['语文', '数学', '英语'];

  var ALIAS = {
    sid: ['学号', '学生学号', '学籍号', '学生编号', '学籍编号', '学生id', '学生ID', 'ID'],
    examNo: ['考号', '准考证号', '考试号'],
    name: ['姓名', '学生姓名'],
    cls: ['班级', '分班', '班号', '行政班', '所在班级', '教学班', '班别', '班级名称'],
    status: ['缺考', '缺考标记', '考试状态', '状态', '备注']
  };

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
  }
  function compact(v) { return str(v).replace(/\s+/g, ''); }

  function toNum(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = compact(v);
    if (!s) return null;
    if (/^[-—–]+$/.test(s)) return null;
    if (/缺考|未考|未上传|免考|弃考|无成绩|无/.test(s)) return null;
    s = s.replace(/分$/, '');
    var n = Number(s);
    return isFinite(n) ? n : null;
  }

  function normSid(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') {
      return Number.isInteger(v) ? String(v) : String(v).replace(/\.0+$/, '');
    }
    return compact(v);
  }

  function parseClassNo(v) {
    var s = compact(v);
    if (!s) return null;
    var m = s.match(/\d+/);
    if (!m) return null;
    var n = parseInt(m[0], 10);
    return isFinite(n) ? n : null;
  }

  function classLabel(cls) {
    if (cls === null || cls === undefined || cls === '') return '未分班';
    if (typeof cls === 'number') return '高一' + cls;
    return String(cls);
  }

  function detectItem(h) {
    if (!h) return null;
    if (h.indexOf('语数外') >= 0 || h.indexOf('语数英') >= 0) return '语数外总分';
    if (h.indexOf('总分') >= 0 || h === '总成绩' || h === '合计' || h === '总计') return '总分';
    for (var i = 0; i < SUBJECTS.length; i++) {
      if (h.indexOf(SUBJECTS[i]) >= 0) return SUBJECTS[i];
    }
    return null;
  }

  function detectField(h, item) {
    if (!h) return null;
    if (h.indexOf('等级') >= 0) return 'level';
    if (h.indexOf('年级位次') >= 0 || h.indexOf('年级排名') >= 0 || h.indexOf('年级名次') >= 0 ||
        h.indexOf('年次') >= 0 || h.indexOf('级名次') >= 0 || h.indexOf('校次') >= 0) return 'gradeRank';
    if (h.indexOf('班级位次') >= 0 || h.indexOf('班级排名') >= 0 || h.indexOf('班级名次') >= 0 ||
        h.indexOf('班次') >= 0 || h.indexOf('班名次') >= 0) return 'classRank';
    if (h === item || h.indexOf('得分') >= 0 || h.indexOf('分数') >= 0 || h.indexOf('成绩') >= 0) return 'score';
    return null;
  }

  function matchAlias(h, list) {
    for (var i = 0; i < list.length; i++) { if (h === list[i]) return true; }
    return false;
  }

  function detectHeaderRow(rows) {
    var limit = Math.min(rows.length, 12);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [];
      var cells = row.map(compact);
      var hasName = cells.indexOf('姓名') >= 0 || cells.indexOf('学生姓名') >= 0;
      var hasId = cells.indexOf('学号') >= 0 || cells.indexOf('考号') >= 0 || cells.indexOf('准考证号') >= 0;
      if (hasName && hasId) return r;
    }
    return -1;
  }

  function parseHeader(cells) {
    var cols = { sid: -1, examNo: -1, name: -1, cls: -1, status: -1 };
    var extras = { zhongkao: -1, category: -1 };
    var items = {};
    var current = null;
    for (var i = 0; i < cells.length; i++) {
      var h = compact(cells[i]);
      if (!h) continue;

      if (h.indexOf('中考') >= 0 && h.indexOf('位次') < 0 && h.indexOf('等级') < 0) {
        if (extras.zhongkao < 0) extras.zhongkao = i;
        current = null;
        continue;
      }
      if (matchAlias(h, ['类别', '学生类别', '身份', '学生类型'])) {
        if (extras.category < 0) extras.category = i;
        continue;
      }

      var baseHit = null;
      for (var k in ALIAS) {
        if (matchAlias(h, ALIAS[k])) { baseHit = k; break; }
      }
      if (baseHit) {
        if (cols[baseHit] < 0) cols[baseHit] = i;
        if (baseHit === 'name') current = null;
        continue;
      }

      var item = detectItem(h);
      if (item) {
        current = item;
        if (!items[item]) items[item] = { score: -1, gradeRank: -1, classRank: -1, level: -1 };
        var f = detectField(h, item);
        if (f && items[item][f] < 0) items[item][f] = i;
        continue;
      }

      if (current) {
        var f2 = detectField(h, current);
        if (f2 && items[current][f2] < 0) { items[current][f2] = i; continue; }
      }
    }
    return { cols: cols, items: items, extras: extras };
  }

  function isEmptyRow(row) {
    if (!row) return true;
    for (var i = 0; i < row.length; i++) {
      if (str(row[i]) !== '') return false;
    }
    return true;
  }

  function isSubHeaderRow(row, topRow) {
    if (!row) return false;
    var cells = row.map(compact);
    var fieldHits = 0;
    cells.forEach(function (c) {
      if (/^(等级|得分|年次|班次|年级位次|班级位次|位次|分数|成绩|名次)$/.test(c)) fieldHits++;
    });
    if (fieldHits < 2) return false;
    var top = (topRow || []).map(compact);
    var nameIdx = top.indexOf('姓名');
    if (nameIdx >= 0 && cells[nameIdx]) return false;
    var subjHits = 0;
    top.forEach(function (c) { if (detectItem(c)) subjHits++; });
    return subjHits >= 2;
  }

  function mergeHeaderRows(top, sub) {
    var n = Math.max(top ? top.length : 0, sub ? sub.length : 0);
    var out = [];
    for (var i = 0; i < n; i++) {
      var a = top && top[i] !== undefined ? compact(top[i]) : '';
      var b = sub && sub[i] !== undefined ? compact(sub[i]) : '';
      out.push(a && b ? a + b : (a || b));
    }
    return out;
  }

  function parseRows(rows) {
    var headerIdx = detectHeaderRow(rows);
    if (headerIdx < 0) {
      return { ok: false, error: '未找到表头行（需同时包含「姓名」和「学号/考号」）', students: [], absent: [], headerIdx: -1 };
    }
    var headerCells = rows[headerIdx] || [];
    var dataStart = headerIdx + 1;
    if (isSubHeaderRow(rows[headerIdx + 1], headerCells)) {
      headerCells = mergeHeaderRows(headerCells, rows[headerIdx + 1] || []);
      dataStart = headerIdx + 2;
    }
    var hdr = parseHeader(headerCells);
    if (hdr.cols.name < 0) {
      return { ok: false, error: '表头缺少「姓名」列', students: [], absent: [], headerIdx: headerIdx };
    }
    if (hdr.cols.sid < 0 && hdr.cols.examNo >= 0) hdr.cols.sid = hdr.cols.examNo;
    if (hdr.cols.sid < 0) {
      return { ok: false, error: '表头缺少「学号」列', students: [], absent: [], headerIdx: headerIdx };
    }

    var students = [];
    var absent = [];
    var seen = {};
    var warnings = [];
    for (var r = dataStart; r < rows.length; r++) {
      var row = rows[r] || [];
      if (isEmptyRow(row)) continue;
      var sid = normSid(row[hdr.cols.sid]);
      var name = str(row[hdr.cols.name]);
      if (!sid && !name) continue;
      if (!sid) { warnings.push('第 ' + (r + 1) + ' 行缺少学号，已跳过'); continue; }

      var scores = {};
      var srcRanks = {};
      var hasAny = false;
      for (var it in hdr.items) {
        var c = hdr.items[it];
        var v = c.score >= 0 ? toNum(row[c.score]) : null;
        scores[it] = v;
        if (v !== null) hasAny = true;
        if (c.gradeRank >= 0 || c.classRank >= 0) {
          srcRanks[it] = {
            grade: c.gradeRank >= 0 ? toNum(row[c.gradeRank]) : null,
            cls: c.classRank >= 0 ? toNum(row[c.classRank]) : null
          };
        }
      }

      var statusText = hdr.cols.status >= 0 ? str(row[hdr.cols.status]) : '';
      var absentFlag = /缺考|未考|未上传|免考|弃考/.test(statusText);
      if (sid && !hasAny && !absentFlag) absentFlag = true;

      var rawCls = hdr.cols.cls >= 0 ? row[hdr.cols.cls] : '';
      var clsNo = parseClassNo(rawCls);

      var stu = {
        sid: sid,
        examNo: (function () {
          var v = hdr.cols.examNo >= 0 ? str(row[hdr.cols.examNo]) : '';
          return /^[-—–]+$/.test(v) ? '' : v;
        })(),
        name: name,
        cls: clsNo,
        clsRaw: str(rawCls),
        clsLabel: clsNo !== null ? classLabel(clsNo) : (str(rawCls) || '未分班'),
        scores: scores,
        srcRanks: srcRanks,
        zhongkao: hdr.extras && hdr.extras.zhongkao >= 0 ? toNum(row[hdr.extras.zhongkao]) : null,
        category: (function () {
          var c = hdr.extras && hdr.extras.category >= 0 ? str(row[hdr.extras.category]) : '';
          return /^[-—–]+$/.test(c) ? '' : c;
        })(),
        status: absentFlag ? 'absent' : 'normal'
      };

      if (seen[sid]) { warnings.push('学号重复：' + sid + '（' + name + '），仅保留首条'); continue; }
      seen[sid] = true;

      if (absentFlag) absent.push(stu); else students.push(stu);
    }
    return { ok: true, students: students, absent: absent, warnings: warnings, headerIdx: headerIdx, header: hdr };
  }

  function isGradeSheet(name) {
    var n = compact(name);
    return n.indexOf('年级') >= 0 || n.indexOf('总表') >= 0 || n === '全部' || n === '总成绩';
  }

  function buildDataset(sheets, opts) {
    opts = opts || {};
    if (!sheets || !sheets.length) return { ok: false, error: '没有可解析的工作表' };

    var parsed = [];
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var p = parseRows(sh.rows || []);
      p.name = sh.name;
      p.isGrade = isGradeSheet(sh.name);
      parsed.push(p);
    }

    var valid = parsed.filter(function (p) { return p.ok; });
    if (!valid.length) return { ok: false, error: parsed[0] && parsed[0].error ? parsed[0].error : '所有工作表都无法解析' };

    var grade = valid.filter(function (p) { return p.isGrade; });
    var base = grade.length ? grade.sort(function (a, b) { return b.students.length - a.students.length; })[0]
      : valid.slice().sort(function (a, b) { return b.students.length - a.students.length; })[0];

    var students = base.students.slice();
    var absent = base.absent.slice();
    var warnings = [];
    var baseIds = {};
    students.forEach(function (s) { baseIds[s.sid] = true; });
    absent.forEach(function (s) { baseIds[s.sid] = true; });

    var classSheets = valid.filter(function (p) { return p !== base; });
    var orphans = [];
    var enrolled = {};
    students.forEach(function (s) { enrolled[s.clsLabel] = (enrolled[s.clsLabel] || 0) + 1; });
    absent.forEach(function (s) { enrolled[s.clsLabel] = (enrolled[s.clsLabel] || 0) + 1; });

    classSheets.forEach(function (p) {
      p.students.concat(p.absent).forEach(function (s) {
        if (!baseIds[s.sid]) {
          orphans.push({ sid: s.sid, name: s.name, clsLabel: s.clsLabel, sheet: p.name });
        }
      });
      if (!p.ok) warnings.push('工作表「' + p.name + '」解析失败：' + p.error);
    });

    if (orphans.length) {
      warnings.push('有 ' + orphans.length + ' 名学生只出现在班级分表、不在年级总表，已按总表为准剔除');
    }
    var diffCls = [];
    students.forEach(function (s) { if (!s.cls) diffCls.push(s.name); });
    if (diffCls.length) warnings.push('有 ' + diffCls.length + ' 名学生没有可识别的班级');

    var sheetInfo = parsed.map(function (p) {
      return { name: p.name, ok: p.ok, error: p.error || '', students: p.ok ? p.students.length : 0, absent: p.ok ? p.absent.length : 0, isGrade: p.isGrade };
    });

    return {
      ok: true,
      students: students,
      absent: absent,
      orphans: orphans,
      warnings: warnings.concat(base.warnings || []),
      enrolled: enrolled,
      sheetInfo: sheetInfo,
      baseSheet: base.name
    };
  }

  function ensureTotals(students) {
    students.forEach(function (s) {
      var sc = s.scores || (s.scores = {});
      if (sc['语数外总分'] === null || sc['语数外总分'] === undefined) {
        var sum3 = 0, any3 = false;
        THREE.forEach(function (k) { if (sc[k] !== null && sc[k] !== undefined) { sum3 += sc[k]; any3 = true; } });
        sc['语数外总分'] = any3 ? round1(sum3) : null;
      }
      if (sc['总分'] === null || sc['总分'] === undefined) {
        var sum = 0, any = false;
        SUBJECTS.forEach(function (k) { if (sc[k] !== null && sc[k] !== undefined) { sum += sc[k]; any = true; } });
        sc['总分'] = any ? round1(sum) : null;
      }
    });
    return students;
  }

  function round1(n) { return Math.round(n * 10) / 10; }

  function rankValues(values) {
    var idx = [];
    for (var i = 0; i < values.length; i++) {
      if (values[i] !== null && values[i] !== undefined) idx.push([values[i], i]);
    }
    idx.sort(function (a, b) { return b[0] - a[0]; });
    var out = new Array(values.length);
    for (var k = 0; k < out.length; k++) out[k] = null;
    var p = 0;
    while (p < idx.length) {
      var q = p;
      while (q < idx.length && idx[q][0] === idx[p][0]) q++;
      var rank = p + 1;
      for (var m = p; m < q; m++) out[idx[m][1]] = rank;
      p = q;
    }
    return out;
  }

  function computeRanks(students) {
    ITEMS.forEach(function (item) {
      var vals = students.map(function (s) { return s.scores[item]; });
      var gr = rankValues(vals);
      var byCls = {};
      students.forEach(function (s, i) {
        if (!byCls[s.clsLabel]) byCls[s.clsLabel] = [];
        byCls[s.clsLabel].push(i);
      });
      var clsRanks = new Array(students.length);
      for (var k = 0; k < clsRanks.length; k++) clsRanks[k] = null;
      for (var key in byCls) {
        var idxs = byCls[key];
        var sub = idxs.map(function (i) { return vals[i]; });
        var rr = rankValues(sub);
        idxs.forEach(function (i, j) { clsRanks[i] = rr[j]; });
      }
      students.forEach(function (s, i) {
        if (!s.ranks) s.ranks = {};
        s.ranks[item] = { grade: gr[i], cls: clsRanks[i] };
      });
    });
    return students;
  }

  function classesOf(students) {
    var map = {};
    students.forEach(function (s) {
      var key = s.clsLabel;
      if (!map[key]) map[key] = { label: key, cls: s.cls, count: 0, sid: s.cls };
      map[key].count++;
    });
    return Object.keys(map).map(function (k) { return map[k]; }).sort(function (a, b) {
      var an = a.cls === null || a.cls === undefined ? 999 : a.cls;
      var bn = b.cls === null || b.cls === undefined ? 999 : b.cls;
      return an - bn;
    });
  }

  function classMeans(students, items) {
    items = items || ITEMS;
    var classes = classesOf(students);
    var out = {};
    classes.forEach(function (c) { out[c.label] = {}; });
    items.forEach(function (item) {
      var acc = {};
      students.forEach(function (s) {
        if (s.scores[item] === null || s.scores[item] === undefined) return;
        if (!acc[s.clsLabel]) acc[s.clsLabel] = { sum: 0, n: 0 };
        acc[s.clsLabel].sum += s.scores[item];
        acc[s.clsLabel].n++;
      });
      classes.forEach(function (c) {
        var a = acc[c.label];
        out[c.label][item] = a && a.n ? round1(a.sum / a.n) : null;
      });
    });
    return { classes: classes, means: out };
  }

  function classRankTable(means, classes, items) {
    items = items || ITEMS;
    var table = {};
    items.forEach(function (item) {
      var vals = classes.map(function (c) { return means[c.label][item]; });
      table[item] = rankValues(vals);
    });
    var out = {};
    classes.forEach(function (c, i) {
      out[c.label] = {};
      items.forEach(function (item) { out[c.label][item] = table[item][i]; });
    });
    return out;
  }

  function tierOfClass(tiers, clsLabel) {
    for (var i = 0; i < tiers.length; i++) {
      if (tiers[i].classes.indexOf(clsLabel) >= 0) return tiers[i];
    }
    return null;
  }

  function tierMeans(students, tiers, items) {
    items = items || ITEMS;
    var out = {};
    tiers.forEach(function (t) {
      out[t.name] = {};
      items.forEach(function (item) {
        var sum = 0, n = 0;
        students.forEach(function (s) {
          if (t.classes.indexOf(s.clsLabel) < 0) return;
          var v = s.scores[item];
          if (v === null || v === undefined) return;
          sum += v; n++;
        });
        out[t.name][item] = n ? round1(sum / n) : null;
      });
    });
    return out;
  }

  function progressValues(s, prevMap, item) {
    var out = { cls: null, grade: null };
    if (!s.ranks || !s.ranks[item]) return out;
    if (prevMap === null || prevMap === undefined) return out;
    var p = prevMap[s.sid];
    if (!p || !p.ranks || !p.ranks[item]) return out;
    if (s.ranks[item].cls !== null && p.ranks[item].cls !== null) out.cls = p.ranks[item].cls - s.ranks[item].cls;
    if (s.ranks[item].grade !== null && p.ranks[item].grade !== null) out.grade = p.ranks[item].grade - s.ranks[item].grade;
    return out;
  }

  function progressOf(s, prevMap, item) {
    return progressValues(s, prevMap, item).grade;
  }

  function buildMainRows(students, prevMap) {
    var headers = ['学号', '考号', '姓名', '班级'];
    ITEMS.forEach(function (item) {
      headers.push(item, '班次', '班次进退', '级次', '级次进退');
    });
    var rows = students.map(function (s) {
      var row = [s.sid, s.examNo || '', s.name, s.clsLabel];
      ITEMS.forEach(function (item) {
        var p = progressValues(s, prevMap, item);
        row.push(s.scores[item] === undefined ? null : s.scores[item]);
        row.push(s.ranks && s.ranks[item] ? s.ranks[item].cls : null);
        row.push(p.cls);
        row.push(s.ranks && s.ranks[item] ? s.ranks[item].grade : null);
        row.push(p.grade);
      });
      return row;
    });
    return { headers: headers, rows: rows };
  }

  function diffTable(students, prevStudents, tiers, reqItems, teachers) {
    var items = reqItems || ITEMS;
    teachers = teachers || {};
    var cur = classMeans(students, items);
    var prev = prevStudents ? classMeans(prevStudents, items) : null;
    var tm = tierMeans(students, tiers, items);

    var groups = tiers.map(function (t) {
      var classes = cur.classes.filter(function (c) { return t.classes.indexOf(c.label) >= 0; });
      var rankCur = {}, rankPrev = {};
      items.forEach(function (item) {
        var vals = classes.map(function (c) { return cur.means[c.label][item]; });
        var rr = rankValues(vals);
        classes.forEach(function (c, i) {
          if (!rankCur[c.label]) rankCur[c.label] = {};
          rankCur[c.label][item] = rr[i];
        });
        if (prev) {
          var pvals = classes.map(function (c) { return prev.means[c.label] ? prev.means[c.label][item] : null; });
          var prr = rankValues(pvals);
          classes.forEach(function (c, i) {
            if (!rankPrev[c.label]) rankPrev[c.label] = {};
            rankPrev[c.label][item] = prr[i];
          });
        }
      });
      var rows = classes.map(function (c) {
        var row = { cls: c.label, count: c.count, teacher: teachers[c.label] || '', items: {} };
        items.forEach(function (item) {
          var m = cur.means[c.label][item];
          var tmean = tm[t.name][item];
          var curRankVal = rankCur[c.label][item];
          var prevRankVal = prev && rankPrev[c.label] ? rankPrev[c.label][item] : null;
          row.items[item] = {
            mean: m,
            rank: curRankVal,
            delta: (m !== null && tmean !== null) ? round1(m - tmean) : null,
            prevRank: prevRankVal,
            progress: (prevRankVal !== null && curRankVal !== null) ? prevRankVal - curRankVal : null
          };
        });
        return row;
      });
      var avg = { cls: t.name + '平均', isAvg: true, count: classes.reduce(function (a, c) { return a + c.count; }, 0), teacher: '', items: {} };
      items.forEach(function (item) {
        avg.items[item] = { mean: tm[t.name][item], rank: null, delta: 0, prevRank: null, progress: null };
      });
      return { name: t.name, rows: rows, avg: avg };
    });
    return groups;
  }

  function bandDefs(thresholds) {
    var cuts = thresholds && thresholds.length ? thresholds.slice().sort(function (a, b) { return b - a; }) : [850, 800, 750, 700, 650, 600, 550, 500];
    var bands = [];
    for (var i = 0; i < cuts.length; i++) {
      var hi = cuts[i];
      var lo = i + 1 < cuts.length ? cuts[i + 1] : null;
      if (i === 0) bands.push({ label: hi + '以上', min: hi, max: Infinity });
      bands.push({ label: lo === null ? (hi + '以下') : (lo + '–' + hi), min: lo === null ? -Infinity : lo, max: hi });
    }
    return bands;
  }

  function bandStats(scores, bands) {
    var counts = bands.map(function () { return 0; });
    var total = 0;
    scores.forEach(function (v) {
      if (v === null || v === undefined) return;
      total++;
      for (var i = 0; i < bands.length; i++) {
        if (v >= bands[i].min && v < bands[i].max) { counts[i]++; break; }
      }
    });
    return bands.map(function (b, i) {
      return { label: b.label, count: counts[i], pct: total ? round1(counts[i] / total * 100) : 0 };
    });
  }

  function analyzeClass(students, clsLabel, tiers, bands, items) {
    items = items || ITEMS;
    var clsStudents = students.filter(function (s) { return s.clsLabel === clsLabel; });
    var tier = tierOfClass(tiers, clsLabel);
    var totals = clsStudents.map(function (s) { return s.scores['总分']; }).filter(function (v) { return v !== null && v !== undefined; });
    var sum = totals.reduce(function (a, b) { return a + b; }, 0);
    var avg = totals.length ? round1(sum / totals.length) : null;
    var max = totals.length ? Math.max.apply(null, totals) : null;
    var min = totals.length ? Math.min.apply(null, totals) : null;

    var tierAvg = {};
    items.forEach(function (item) {
      var s = 0, n = 0;
      students.forEach(function (st) {
        if (!tier || tier.classes.indexOf(st.clsLabel) < 0) return;
        var v = st.scores[item];
        if (v === null || v === undefined) return;
        s += v; n++;
      });
      tierAvg[item] = n ? round1(s / n) : null;
    });

    var means = {};
    var tierRanks = {};
    var cm = classMeans(students, items);
    var ranks = classRankTable(cm.means, cm.classes, items);
    items.forEach(function (item) {
      means[item] = cm.means[clsLabel] ? cm.means[clsLabel][item] : null;
    });
    if (tier) {
      items.forEach(function (item) {
        var vals = tier.classes.map(function (c) { return cm.means[c] ? cm.means[c][item] : null; });
        var rr = rankValues(vals);
        var idx = tier.classes.indexOf(clsLabel);
        tierRanks[item] = idx >= 0 ? rr[idx] : null;
      });
    }

    return {
      cls: clsLabel,
      tier: tier ? tier.name : '—',
      count: clsStudents.length,
      avg: avg, max: max, min: min,
      bands: bandStats(totals, bands),
      means: means, tierRanks: tierRanks, tierAvg: tierAvg,
      students: clsStudents
    };
  }

  function awards(students, tiers, pct) {
    pct = pct || 16.7;
    var excellent = [];
    tiers.forEach(function (t) {
      var list = students.filter(function (s) { return t.classes.indexOf(s.clsLabel) >= 0 && s.scores['总分'] !== null; });
      list.sort(function (a, b) { return b.scores['总分'] - a.scores['总分']; });
      var k = Math.max(1, Math.round(list.length * pct / 100));
      list.slice(0, k).forEach(function (s) {
        excellent.push({ tier: t.name, sid: s.sid, name: s.name, clsLabel: s.clsLabel, total: s.scores['总分'], rank: s.ranks['总分'].grade });
      });
    });

    var top = ITEMS.map(function (item) {
      var best = null;
      students.forEach(function (s) {
        var v = s.scores[item];
        if (v === null || v === undefined) return;
        if (best === null || v > best) best = v;
      });
      if (best === null) return { item: item, score: null, students: [] };
      var list = students.filter(function (s) { return s.scores[item] === best; });
      return { item: item, score: best, students: list.map(function (s) { return { sid: s.sid, name: s.name, clsLabel: s.clsLabel }; }) };
    });
    return { excellent: excellent, top: top };
  }

  function autoText(analysis) {
    if (!analysis || !analysis.count) return { problem: '', measure: '', advice: '' };
    var avg = analysis.avg === null ? 0 : analysis.avg;
    var diff = analysis.tierAvg['总分'] === null ? 0 : round1(avg - analysis.tierAvg['总分']);
    var weak = null, weakGap = 0;
    SUBJECTS.forEach(function (k) {
      var g = (analysis.tierAvg[k] !== null && analysis.means[k] !== null) ? round1(analysis.means[k] - analysis.tierAvg[k]) : null;
      if (g !== null && (weak === null || g < weakGap)) { weak = k; weakGap = g; }
    });
    var problem = '本班实考 ' + analysis.count + ' 人，总分平均 ' + avg + ' 分，' +
      (diff >= 0 ? '高于' : '低于') + '本层次平均 ' + Math.abs(diff) + ' 分。';
    if (weak) problem += weak + '均分较层次平均' + (weakGap >= 0 ? '高 ' : '低 ') + Math.abs(weakGap) + ' 分，是本班相对薄弱学科。';
    var bands = analysis.bands || [];
    var topBand = bands[0];
    if (topBand && topBand.count > 0) problem += '高分段（' + topBand.label + '）' + topBand.count + ' 人。';

    var measure = '针对' + (weak || '薄弱学科') + '，建议梳理高频失分点，安排专项训练与面批；' +
      '对总分' + (diff >= 0 ? '优势' : '滞后') + '的学生分层布置任务，落实错题回炉。';
    var advice = '用好本次数据：按学科位次定位临界生，' +
      '每周跟踪一次小测，月末对比本次基线，动态调整辅导名单。';
    return { problem: problem, measure: measure, advice: advice };
  }


  /* ---------------- 学生信息（中考 / 类别） ---------------- */
  var INFO_ALIAS = {
    sid: ['学号', '学生学号', '学籍号', '学生编号', '学籍编号', '学生id', '学生ID', 'ID', '考号', '准考证号'],
    zhongkao: ['中考总分', '中考成绩', '中考分数', '中考得分', '中考'],
    category: ['类别', '学生类别', '身份', '学生类型', '备注', '标记']
  };

  function findAliasIndex(cells, list) {
    for (var i = 0; i < cells.length; i++) {
      if (matchAlias(cells[i], list)) return i;
    }
    return -1;
  }

  function parseInfoRows(rows) {
    if (!rows || !rows.length) return { ok: false, error: '表为空' };
    var headerIdx = -1, sidCol = -1, zkCol = -1, catCol = -1;
    for (var r = 0; r < Math.min(rows.length, 10); r++) {
      var cells = (rows[r] || []).map(compact);
      var s = findAliasIndex(cells, INFO_ALIAS.sid);
      var z = -1, c = -1;
      for (var i = 0; i < cells.length; i++) {
        if (z < 0 && cells[i].indexOf('中考') >= 0) z = i;
        if (c < 0 && matchAlias(cells[i], INFO_ALIAS.category)) c = i;
      }
      if (s >= 0 && (z >= 0 || c >= 0)) { headerIdx = r; sidCol = s; zkCol = z; catCol = c; break; }
    }
    if (headerIdx < 0) return { ok: false, error: '未找到「学号」与「中考总分」或「类别」列' };
    var records = [];
    for (var k = headerIdx + 1; k < rows.length; k++) {
      var row = rows[k] || [];
      var sid = normSid(row[sidCol]);
      if (!sid) continue;
      var zk = zkCol >= 0 ? toNum(row[zkCol]) : null;
      var cat = catCol >= 0 ? str(row[catCol]) : '';
      records.push({ sid: sid, zhongkao: zk, category: cat });
    }
    return { ok: true, records: records };
  }

  function applyInfo(students, records) {
    var map = {};
    records.forEach(function (r) { map[r.sid] = r; });
    var matched = 0, withZk = 0, withCat = 0;
    students.forEach(function (s) {
      var r = map[s.sid];
      if (!r) return;
      matched++;
      if (r.zhongkao !== null && r.zhongkao !== undefined) { s.zhongkao = r.zhongkao; withZk++; }
      if (r.category) { s.category = r.category; withCat++; }
    });
    var ids = {};
    students.forEach(function (s) { ids[s.sid] = true; });
    var unmatched = records.filter(function (r) { return !ids[r.sid]; });
    return { matched: matched, withZk: withZk, withCat: withCat, unmatched: unmatched };
  }


  /* ---------------- 班主任 / 任课教师表 ---------------- */
  function normSubjectName(h) {
    var s = compact(h).replace(/教师|老师|任课|科任|姓名/g, '');
    if (!s) return null;
    for (var i = 0; i < ITEMS.length; i++) { if (s === ITEMS[i]) return ITEMS[i]; }
    if (s.indexOf('语数外') >= 0) return '语数外总分';
    return null;
  }

  function normClassLabel(v) {
    var n = parseClassNo(v);
    return n !== null ? classLabel(n) : str(v);
  }

  function splitClassList(v) {
    var s = str(v);
    if (!s) return [];
    var parts = s.split(/[,，、;；\/\s]+/);
    var out = [];
    parts.forEach(function (p) {
      if (!p) return;
      var lab = normClassLabel(p);
      if (lab && out.indexOf(lab) < 0) out.push(lab);
    });
    return out;
  }

  function parseTeacherRows(rows) {
    if (!rows || !rows.length) return { ok: false, error: '空表' };
    var headerIdx = -1, header = null;
    for (var r = 0; r < Math.min(rows.length, 12); r++) {
      var cells = (rows[r] || []).map(compact);
      var hasCls = cells.some(function (c) { return /^班级$|^班$|班级名称|教学班|行政班/.test(c); });
      if (!hasCls) continue;
      var hasAny = cells.some(function (c) { return /班主任|学科|科目|教师|老师|语文|数学|英语/.test(c); });
      if (hasAny) { headerIdx = r; header = cells; break; }
    }
    if (headerIdx < 0) return { ok: false, error: '未找到含「班级」的表头' };

    var clsCol = -1, headCol = -1, subjCol = -1, teacherCol = -1, roleCol = -1;
    header.forEach(function (c, i) {
      if (clsCol < 0 && /^班级$|^班$|班级名称|教学班|行政班/.test(c)) clsCol = i;
      if (roleCol < 0 && /^职务$|^角色$/.test(c)) roleCol = i;
      if (subjCol < 0 && /^学科$|^科目$|学科名称$|科目名称$/.test(c)) subjCol = i;
    });
    header.forEach(function (c, i) {
      if (i === clsCol || i === roleCol) return;
      if (headCol < 0 && c.indexOf('班主任') >= 0) headCol = i;
    });
    header.forEach(function (c, i) {
      if (i === clsCol || i === headCol || i === roleCol) return;
      if (teacherCol < 0 && /教师|老师|姓名/.test(c)) teacherCol = i;
    });

    var subjCols = [];
    header.forEach(function (c, i) {
      if (i === clsCol || i === headCol || i === teacherCol || i === roleCol) return;
      var item = normSubjectName(c);
      if (item) subjCols.push({ i: i, item: item });
    });

    var head = {}, teachers = {}, warnings = [];
    var mode = (subjCol >= 0 && teacherCol >= 0) ? 'long' : (subjCols.length ? 'wide' : 'head');

    for (var k = headerIdx + 1; k < rows.length; k++) {
      var row = rows[k] || [];
      if (isEmptyRow(row)) continue;
      var clsList = splitClassList(row[clsCol]);
      if (!clsList.length) continue;

      if (mode === 'long') {
        var subjText = compact(row[subjCol]);
        var teacher = str(row[teacherCol]);
        var role = roleCol >= 0 ? compact(row[roleCol]) : '';
        if (!teacher) continue;
        if (role.indexOf('班主任') >= 0 || subjText.indexOf('班主任') >= 0) {
          clsList.forEach(function (c) { head[c] = teacher; });
          continue;
        }
        var item = normSubjectName(subjText);
        if (!item) { warnings.push(clsList[0] + '：无法识别学科「' + subjText + '」'); continue; }
        clsList.forEach(function (c) {
          if (!teachers[c]) teachers[c] = {};
          teachers[c][item] = teacher;
        });
      } else if (mode === 'wide') {
        if (headCol >= 0) {
          var h = str(row[headCol]);
          if (h) clsList.forEach(function (c) { head[c] = h; });
        }
        subjCols.forEach(function (sc) {
          var t = str(row[sc.i]);
          if (!t) return;
          clsList.forEach(function (c) {
            if (!teachers[c]) teachers[c] = {};
            teachers[c][sc.item] = t;
          });
        });
      } else {
        var nameCol = headCol >= 0 ? headCol : teacherCol;
        if (nameCol >= 0) {
          var nm = str(row[nameCol]);
          if (nm) clsList.forEach(function (c) { head[c] = nm; });
        }
      }
    }
    if (!Object.keys(head).length && !Object.keys(teachers).length) {
      return { ok: false, error: '没有识别到教师姓名' };
    }
    return { ok: true, mode: mode, head: head, teachers: teachers, warnings: warnings };
  }

  function statBlock(students, items) {
    items = items || ITEMS;
    var vals = [];
    students.forEach(function (s) { if (s.scores['总分'] !== null && s.scores['总分'] !== undefined) vals.push(s.scores['总分']); });
    var sum = vals.reduce(function (a, b) { return a + b; }, 0);
    var zk = [];
    students.forEach(function (s) { if (s.zhongkao !== null && s.zhongkao !== undefined) zk.push(s.zhongkao); });
    var zkSum = zk.reduce(function (a, b) { return a + b; }, 0);
    var means = {};
    items.forEach(function (item) {
      var s = 0, n = 0;
      students.forEach(function (st) {
        var v = st.scores[item];
        if (v === null || v === undefined) return;
        s += v; n++;
      });
      means[item] = n ? round1(s / n) : null;
    });
    return {
      n: students.length,
      avg: vals.length ? round1(sum / vals.length) : null,
      max: vals.length ? Math.max.apply(null, vals) : null,
      min: vals.length ? Math.min.apply(null, vals) : null,
      zkAvg: zk.length ? round1(zkSum / zk.length) : null,
      means: means
    };
  }

  function zhongkaoSplits(students, cuts, items) {
    cuts = cuts && cuts.length ? cuts : [688, 699];
    var withZk = students.filter(function (s) { return s.zhongkao !== null && s.zhongkao !== undefined; });
    var out = [];
    cuts.forEach(function (cut) {
      var up = [], down = [];
      withZk.forEach(function (s) { (s.zhongkao >= cut ? up : down).push(s); });
      var a = statBlock(up, items);
      a.cut = cut; a.side = 'up'; a.label = cut + ' 分及以上';
      var b = statBlock(down, items);
      b.cut = cut; b.side = 'down'; b.label = cut + ' 分以下';
      out.push(a, b);
    });
    return out;
  }

  function exclusionStats(students, excludedSet, items) {
    var left = [], out = [];
    students.forEach(function (s) { (excludedSet[s.sid] ? out : left).push(s); });
    return {
      all: statBlock(students, items),
      left: statBlock(left, items),
      out: statBlock(out, items),
      excludedCount: out.length
    };
  }

  return {
    SUBJECTS: SUBJECTS,
    ITEMS: ITEMS,
    str: str,
    toNum: toNum,
    normSid: normSid,
    parseClassNo: parseClassNo,
    classLabel: classLabel,
    detectHeaderRow: detectHeaderRow,
    parseHeader: parseHeader,
    parseRows: parseRows,
    buildDataset: buildDataset,
    ensureTotals: ensureTotals,
    rankValues: rankValues,
    computeRanks: computeRanks,
    classesOf: classesOf,
    classMeans: classMeans,
    classRankTable: classRankTable,
    tierOfClass: tierOfClass,
    tierMeans: tierMeans,
    buildMainRows: buildMainRows,
    diffTable: diffTable,
    bandDefs: bandDefs,
    bandStats: bandStats,
    analyzeClass: analyzeClass,
    awards: awards,
    autoText: autoText,
    parseInfoRows: parseInfoRows,
    parseTeacherRows: parseTeacherRows,
    applyInfo: applyInfo,
    statBlock: statBlock,
    zhongkaoSplits: zhongkaoSplits,
    exclusionStats: exclusionStats,
    progressOf: progressOf,
    progressValues: progressValues,
    round1: round1
  };
});
