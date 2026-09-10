/* CJ나눔재단 KPI 대시보드 - 클라이언트 사이드 렌더링 엔진
 *
 * 1) 페이지 최초 로드시: index.html에 이미 서버(Python)가 그려둔 차트를 그대로 보여준다 (변경 없음).
 * 2) 사용자가 xlsx 파일을 업로드하면: 이 파일 안의 parseWorkbook()이 tools/build_data.py와 동일한
 *    "라벨 텍스트 스캔 + forward-fill" 방식으로 원본 구글시트(xlsx)를 파싱하고, renderAll()이 모든
 *    카드/차트/진단 문구를 새 데이터로 다시 그린다. 이 갱신은 업로드한 사람의 브라우저에만 반영되며
 *    (새로고침하면 사라짐), 실제로 공개 링크를 갱신하려면 로컬에서 publish_to_github.bat을 실행해야 한다.
 * 3) "원본으로 되돌리기": index.html에 함께 심어둔 window.__BASELINE_TABLES__(원본 데이터 스냅샷)로
 *    다시 렌더링하고, 처음 페이지 로드시 저장해둔 진단/처방 문구(손으로 쓴 원본 분석)를 그대로 복원한다.
 *    (업로드로 새로 생성하는 진단/처방 문구는 100% 자동 생성된 일반 규칙 기반 문구다 - 손으로 쓴 원본
 *    분석만큼 구체적이지 않을 수 있다는 점을 업로드 후 화면에 안내한다.)
 */
(function () {
  "use strict";

  const C_IG = "#D6336C";
  const C_YT = "#E8590C";
  const C_TARGET = "#868E96";
  const C_GOOD = "#2B8A3E";
  const C_WARN = "#E67700";
  const C_BAD = "#C92A2A";
  const GRID = "#E9ECEF";
  const FONT_COLOR = "#212529";
  const CHANNEL_KO = { IG: "인스타그램", YT: "유튜브" };

  // ---------------------------------------------------------------------
  // 공용 헬퍼
  // ---------------------------------------------------------------------
  function num(v) {
    if (v === null || v === undefined || v === "" || v === "-") return null;
    if (typeof v === "string" && v.startsWith("#")) return null;
    if (typeof v === "number" && Number.isNaN(v)) return null;
    return v;
  }
  function cleanLabel(v) {
    if (typeof v !== "string") return v;
    return v.replace(/\n/g, " ").replace(/\s+/g, " ").trim();
  }
  function cleanMetric(v) {
    v = cleanLabel(v);
    if (typeof v !== "string") return v;
    return v.replace(/\s*\([^)]*\)$/, "").trim();
  }
  function fmt(n) {
    if (n === null || n === undefined || Number.isNaN(n)) return "-";
    return Number(n).toLocaleString("en-US", { maximumFractionDigits: 0 });
  }
  function fmtWon(n) {
    return n === null || n === undefined ? "-" : "₩" + fmt(n);
  }
  function pct(rate, digits) {
    if (rate === null || rate === undefined || Number.isNaN(rate)) return "-";
    return (rate * 100).toFixed(digits === undefined ? 0 : digits) + "%";
  }
  function sortByKey(arr, keyFn) {
    return [...arr].sort((a, b) => {
      const ka = keyFn(a), kb = keyFn(b);
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    });
  }

  // ---------------------------------------------------------------------
  // xlsx(SheetJS workbook) → 정제 테이블 (tools/build_data.py 를 그대로 JS로 옮긴 것)
  // ---------------------------------------------------------------------
  function sheetToRows(ws) {
    // 원본 시트가 A열을 전혀 안 쓰면(B열부터 데이터 시작) SheetJS는 실제 사용 범위만 반환해 인덱스가
    // 한 칸씩 밀린다. openpyxl(Python)은 항상 A열(=index0)부터 세므로, 여기서도 A열부터 강제로 포함시켜
    // build_data.py와 동일한 컬럼 인덱스를 쓸 수 있게 맞춘다.
    const ref = ws["!ref"] || "A1:A1";
    const range = XLSX.utils.decode_range(ref);
    range.s.c = 0;
    range.s.r = 0;
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, range: XLSX.utils.encode_range(range) });
  }
  // Python의 row_values(ws, r)와 동일: r은 1부터 시작하는 시트 행 번호, 반환 배열의 index0 = A열
  function rowValues(rows, r) {
    return rows[r - 1] || [];
  }

  function parseAnnualSummary(rowsKPI) {
    const out = [];
    for (let r = 1; r <= 8; r++) {
      const vals = rowValues(rowsKPI, r);
      const channel = vals[1];
      if (channel !== "인스타그램" && channel !== "유튜브") continue;
      const stockMetric = channel === "인스타그램" ? "팔로워" : "구독자";
      out.push({
        channel, metric: stockMetric, actual_cumulative: num(vals[2]), annual_target: num(vals[4]),
        annual_progress_rate: num(vals[5]), monthly_target: num(vals[6]), monthly_progress_rate: num(vals[7]),
      });
      if (num(vals[8]) !== null) {
        out.push({
          channel, metric: "인게이지먼트", actual_cumulative: num(vals[8]), annual_target: num(vals[10]),
          annual_progress_rate: num(vals[11]), monthly_target: num(vals[12]), monthly_progress_rate: num(vals[13]),
        });
      }
      if (num(vals[14]) !== null) {
        out.push({
          channel, metric: "조회수", actual_cumulative: num(vals[14]), annual_target: num(vals[16]),
          annual_progress_rate: num(vals[17]), monthly_target: num(vals[18]), monthly_progress_rate: num(vals[19]),
        });
      }
    }
    return out;
  }

  function parseLadderAndBudgetPlan(rowsKPI) {
    let planTitleRow = null;
    for (let r = 1; r <= 29; r++) {
      if (rowValues(rowsKPI, r)[1] === "채널별 KPI 계획") { planTitleRow = r; break; }
    }
    if (!planTitleRow) throw new Error('"채널별 KPI 계획" 표를 찾지 못했습니다.');

    let monthHeaderRow = null;
    for (let r = planTitleRow; r < planTitleRow + 5; r++) {
      const vals = rowValues(rowsKPI, r);
      if (vals.some((v) => typeof v === "string" && /^\d{2}월$/.test(v))) { monthHeaderRow = r; break; }
    }
    if (!monthHeaderRow) throw new Error("월별 목표 계획표의 월 헤더 행을 찾지 못했습니다.");

    const monthCols = {};
    rowValues(rowsKPI, monthHeaderRow).forEach((v, idx) => {
      if (typeof v === "string" && /^\d{2}월$/.test(v)) monthCols[v] = idx + 1;
    });

    const ladderRaw = [];
    const budgetPlan = [];
    let curChannel = null, curMetric = null;
    let r = monthHeaderRow + 1;
    let blankStreak = 0;
    while (blankStreak < 3 && r < monthHeaderRow + 40) {
      const vals = rowValues(rowsKPI, r);
      const b = vals[1], cc = vals[2], label = cleanLabel(vals[3]);
      if (b) curChannel = b;
      if (cc && cc !== "광고") curMetric = cleanMetric(cc);
      if (![vals[1], vals[2], vals[3], vals[4]].some(Boolean)) { blankStreak++; r++; continue; }
      blankStreak = 0;
      if (label === "KPI (누적)") {
        Object.entries(monthCols).forEach(([mon, col]) => {
          const t = num(vals[col - 1]);
          if (t !== null) ladderRaw.push({ channel: curChannel, metric: curMetric, month: `2026-${mon.slice(0, 2)}`, val: t, isRate: false });
        });
      } else if (label === "KPI 달성률") {
        Object.entries(monthCols).forEach(([mon, col]) => {
          const rate = num(vals[col - 1]);
          if (rate !== null) ladderRaw.push({ channel: curChannel, metric: curMetric, month: `2026-${mon.slice(0, 2)}`, val: rate, isRate: true });
        });
      } else if (label && String(label).includes("예산")) {
        Object.entries(monthCols).forEach(([mon, col]) => {
          const v = num(vals[col - 1]);
          if (v !== null) budgetPlan.push({ channel: curChannel, month: `2026-${mon.slice(0, 2)}`, planned_ad_budget: v });
        });
      }
      r++;
    }
    const targets = {}, rates = {};
    ladderRaw.forEach((row) => {
      const key = row.channel + "|" + row.metric + "|" + row.month;
      if (row.isRate) rates[key] = row.val; else targets[key] = row.val;
    });
    const ladder = Object.keys(targets).map((key) => {
      const [channel, metric, month] = key.split("|");
      return { channel, metric, month, target_cumulative: targets[key], target_achievement_rate: rates[key] === undefined ? null : rates[key] };
    });
    return { ladder: sortByKey(ladder, (d) => d.channel + d.metric + d.month), budgetPlan };
  }

  function parseMonthlyActual(rowsMonthly) {
    let mHeaderRow = null;
    for (let r = 1; r <= 14; r++) {
      if (rowValues(rowsMonthly, r)[0] === "채널") { mHeaderRow = r; break; }
    }
    if (!mHeaderRow) throw new Error('"KPI(monthly)" 시트의 헤더 행을 찾지 못했습니다.');

    const hdrVals = rowValues(rowsMonthly, mHeaderRow);
    let baselineCol = null;
    const monthActualCols = {};
    hdrVals.forEach((v, idx) => {
      if (typeof v !== "string") return;
      const col = idx + 1;
      if (v.includes("연초 기준")) baselineCol = col;
      const m = v.trim().match(/^(\d{1,2})월\s*실적$/);
      if (m) monthActualCols[`2026-${String(m[1]).padStart(2, "0")}`] = col;
    });

    const kindMap = { 팔로워: "stock", 구독자: "stock", 인게이지먼트: "flow", 조회수: "flow" };
    let curChannel = null;
    const out = [];
    for (let r = mHeaderRow + 1; r < mHeaderRow + 5; r++) {
      const vals = rowValues(rowsMonthly, r);
      const channel = vals[0], metric = vals[1];
      if (channel) curChannel = CHANNEL_KO[channel] || channel;
      if (!(metric in kindMap)) continue;
      const kind = kindMap[metric];
      let cum = 0;
      if (baselineCol) {
        const base = num(vals[baselineCol - 1]);
        if (base !== null) out.push({ channel: curChannel, metric, metric_type: kind, month: "2025_baseline", actual_value: base, actual_cumulative: base });
      }
      Object.keys(monthActualCols).sort().forEach((month) => {
        const v = num(vals[monthActualCols[month] - 1]);
        if (v === null) return;
        if (kind === "flow") {
          cum += v;
          out.push({ channel: curChannel, metric, metric_type: kind, month, actual_value: v, actual_cumulative: cum });
        } else {
          out.push({ channel: curChannel, metric, metric_type: kind, month, actual_value: v, actual_cumulative: v });
        }
      });
    }
    return out;
  }

  function extractMediaTable(rows, sectionTitle, fields) {
    let titleRow = null;
    for (let r = 1; r <= 59; r++) {
      if (rowValues(rows, r)[1] === sectionTitle) { titleRow = r; break; }
    }
    if (!titleRow) throw new Error(`"${sectionTitle}" 표를 찾지 못했습니다.`);
    const out = [];
    for (let r = titleRow + 1; r < titleRow + 15; r++) {
      const vals = rowValues(rows, r);
      const b = vals[1];
      let month;
      if (typeof b === "string" && b.trim() === "TOTAL") month = "TOTAL";
      else if (b instanceof Date) {
        // SheetJS의 Excel 일련번호→Date 변환에는 부동소수점 오차로 자정 대비 최대 수십 초 정도의 오차가
        // 생길 수 있다(예: "2026-03-01 00:00:00"이 "2026-02-28 23:59:08"로 변환됨). 정오(+12시간)만큼
        // 밀어서 판단하면 이런 오차로 날짜가 하루/한 달 밀려 읽히는 것을 안전하게 방지할 수 있다.
        const nudged = new Date(b.getTime() + 12 * 60 * 60 * 1000);
        month = `${nudged.getUTCFullYear()}-${String(nudged.getUTCMonth() + 1).padStart(2, "0")}`;
      }
      else if (typeof b === "string" && /^\d{4}-\d{2}/.test(b)) month = b.slice(0, 7);
      else continue;
      const row = { month };
      fields.forEach((f, i) => { row[f] = num(vals[2 + i]); });
      out.push(row);
    }
    return out;
  }

  function parseBudgetSummaryAndH2Plan(rowsH2) {
    const budgetSummary = [];
    for (let r = 1; r <= 9; r++) {
      const vals = rowValues(rowsH2, r);
      const label = vals[1];
      if (["26년 총예산", "5~7월 소진금액", "잔여예산"].includes(label)) budgetSummary.push({ item: label, value: num(vals[2]) });
    }
    let planTitleRow2 = null;
    for (let r = 1; r <= 19; r++) {
      if (rowValues(rowsH2, r)[1] === "매체") { planTitleRow2 = r; break; }
    }
    const h2MediaPlan = [];
    if (planTitleRow2) {
      for (let r = planTitleRow2 + 1; r < planTitleRow2 + 6; r++) {
        const vals = rowValues(rowsH2, r);
        const label = vals[1];
        if (!label || label === "합계") continue;
        h2MediaPlan.push({
          media_item: label, spend_may: num(vals[2]), spend_jun: num(vals[3]), spend_jul: num(vals[4]),
          budget_aug_nov_total: num(vals[5]), budget_aug_nov_monthly_even: num(vals[6]),
          expected_result: num(vals[7]), unit_cost_note: vals[8] === undefined ? null : vals[8],
        });
      }
    }
    return { budgetSummary, h2MediaPlan };
  }

  function parseH2Cascade(rowsH2) {
    let cascadeTitleRow = null;
    for (let r = 1; r <= 24; r++) {
      const vals = rowValues(rowsH2, r);
      if (vals[1] === "채널" && vals[2] === "KPI") { cascadeTitleRow = r; break; }
    }
    if (!cascadeTitleRow) throw new Error("하반기 KPI 캐스케이드 표를 찾지 못했습니다.");

    let monthHdrRow = null;
    for (let r = cascadeTitleRow; r < cascadeTitleRow + 4; r++) {
      if (rowValues(rowsH2, r).some((v) => v === "8월")) { monthHdrRow = r; break; }
    }
    if (!monthHdrRow) throw new Error("하반기 월 헤더 행을 찾지 못했습니다.");

    const h2MonthCols = {};
    rowValues(rowsH2, monthHdrRow).forEach((v, idx) => {
      if (typeof v === "string" && /^(8|9|10|11)월$/.test(v)) h2MonthCols[v] = idx + 1;
    });

    const rows = [];
    let curChannel = null, curMetric = null;
    let r = monthHdrRow + 1;
    let blankStreak = 0;
    while (blankStreak < 3 && r < monthHdrRow + 20) {
      const vals = rowValues(rowsH2, r);
      const channel = vals[1], metric = vals[2], label = vals[3];
      if (channel) curChannel = channel;
      if (metric && metric !== "광고") curMetric = metric === "구독" ? "구독자" : metric;
      if (![vals[1], vals[2], vals[3], vals[4]].some(Boolean)) { blankStreak++; r++; continue; }
      blankStreak = 0;
      if (label === "KPI (누적)" || label === "예상 실적 (증가분)" || label === "KPI 달성률" || (label && String(label).includes("예산"))) {
        Object.entries(h2MonthCols).forEach(([mon, col]) => {
          const v = num(vals[col - 1]);
          const monthKey = `2026-${mon.slice(0, -1).padStart(2, "0")}`;
          rows.push({ channel: curChannel, metric: curMetric, month: monthKey, label, v });
        });
      }
      r++;
      const nextVals = rowValues(rowsH2, r);
      if (nextVals.some((v) => typeof v === "string" && v.includes("공동작업자 콘텐츠 성과 포함"))) break;
    }

    const labelMap = { "KPI (누적)": "kpi_cumulative_target", "예상 실적 (증가분)": "expected_actual_increment", "KPI 달성률": "kpi_achievement_rate" };
    const pivot = {};
    rows.forEach((row) => {
      const key = row.channel + "|" + row.metric + "|" + row.month;
      pivot[key] = pivot[key] || { channel: row.channel, metric: row.metric, month: row.month };
      if (labelMap[row.label]) pivot[key][labelMap[row.label]] = row.v;
      else if (row.label && String(row.label).includes("예산")) pivot[key].ad_budget = row.v;
    });
    return sortByKey(Object.values(pivot), (d) => d.channel + d.metric + d.month);
  }

  const IG_PERF_FIELDS = ["spend", "engagement_total", "like", "comment", "share", "scrap", "cpe", "impressions",
    "clicks", "reach", "post_engagement", "profile_visit", "ctr", "reach_rate", "cpm", "cpc", "cpa_engagement", "cpa_profile_visit"];
  const YT_PERF_FIELDS = ["spend", "engagement_total", "like", "comment", "cumulative_monthly_views", "published_views",
    "impressions", "clicks", "views", "ctr", "vtr", "cpm", "cpc", "cpv"];

  function parseWorkbook(wb) {
    if (!wb.SheetNames || wb.SheetNames.length < 3) {
      throw new Error("시트가 3개(KPI / KPI(monthly) / 하반기 계획) 있는 원본 구글시트 xlsx 파일이 맞는지 확인해주세요.");
    }
    const rowsKPI = sheetToRows(wb.Sheets[wb.SheetNames[0]]);
    const rowsMonthly = sheetToRows(wb.Sheets[wb.SheetNames[1]]);
    const rowsH2 = sheetToRows(wb.Sheets[wb.SheetNames[2]]);

    const annual = parseAnnualSummary(rowsKPI);
    const { ladder } = parseLadderAndBudgetPlan(rowsKPI);
    const actual = parseMonthlyActual(rowsMonthly);
    const igPerf = extractMediaTable(rowsKPI, "인스타그램 운영 성과", IG_PERF_FIELDS);
    const ytPerf = extractMediaTable(rowsKPI, "유튜브 운영 성과", YT_PERF_FIELDS);
    const { budgetSummary, h2MediaPlan } = parseBudgetSummaryAndH2Plan(rowsH2);
    const h2Cascade = parseH2Cascade(rowsH2);

    if (!annual.length) throw new Error("채널별 KPI 요약 데이터를 찾지 못했습니다 (KPI 시트 구조를 확인해주세요).");
    return { annual, ladder, actual, igPerf, ytPerf, budgetSummary, h2MediaPlan, h2Cascade };
  }

  // ---------------------------------------------------------------------
  // 차트 렌더링 (build_static_site.py의 Plotly 차트 로직을 그대로 옮김)
  // ---------------------------------------------------------------------
  function styleFig(height, title, monthXaxis) {
    const layout = {
      height: height,
      template: "plotly_white",
      font: { color: FONT_COLOR, size: 13 },
      margin: { l: 10, r: 10, t: title ? 72 : 20, b: 10 },
      legend: { orientation: "h", yanchor: "top", y: 0.90, xanchor: "left", x: 0 },
      hovermode: "x unified",
      xaxis: monthXaxis
        ? { showgrid: false, zeroline: false, type: "category", categoryorder: "category ascending" }
        : { showgrid: false, zeroline: false },
      yaxis: { showgrid: true, gridcolor: GRID, zeroline: false },
    };
    if (title) layout.title = { text: title, x: 0.01, xanchor: "left", y: 0.98, yanchor: "top" };
    return layout;
  }
  // responsive:true - 컨테이너 크기가 바뀌면(탭 전환, 창 크기 조절 등) ResizeObserver로 항상 다시 맞춤.
  // 이게 없으면 최초 렌더링 시점의 컨테이너 폭으로 고정되어 버려서, 레이아웃이 나중에 자리잡을 때
  // 실제보다 넓게 그려진 채 잘려 보이는 문제가 생길 수 있다.
  const PCFG = { displaylogo: false, responsive: true };

  function filterSorted(rows, pred) {
    return rows.filter(pred).sort((a, b) => (a.month < b.month ? -1 : a.month > b.month ? 1 : 0));
  }

  function renderKpiCards(annual) {
    let html = "";
    annual.forEach((row) => {
      const rate = row.annual_progress_rate;
      const color = rate >= 0.95 ? C_GOOD : rate >= 0.8 ? C_WARN : C_BAD;
      const icon = rate >= 0.95 ? "🟢" : rate >= 0.8 ? "🟡" : "🔴";
      html += `
      <div class="kpi-card">
        <div class="kpi-label">${row.channel} · ${row.metric}</div>
        <div class="kpi-value">${fmt(row.actual_cumulative)}</div>
        <div class="kpi-target">연간 목표 ${fmt(row.annual_target)}</div>
        <div class="kpi-rate" style="color:${color}">${icon} ${pct(rate, 1)}</div>
        <div class="kpi-sub">월간 목표 대비 ${pct(row.monthly_progress_rate, 0)}</div>
        <div class="kpi-bar-track"><div class="kpi-bar-fill" style="width:${Math.min(rate || 0, 1) * 100}%;background:${color}"></div></div>
      </div>`;
    });
    document.getElementById("kpiGrid").innerHTML = html;
  }

  function renderCascadeCharts(tables) {
    const order = [
      { channel: "인스타그램", metric: "팔로워", color: C_IG, id: "cascade0" },
      { channel: "인스타그램", metric: "인게이지먼트", color: C_IG, id: "cascade1" },
      { channel: "유튜브", metric: "구독자", color: C_YT, id: "cascade2" },
      { channel: "유튜브", metric: "조회수", color: C_YT, id: "cascade3" },
    ];
    order.forEach(({ channel, metric, color, id }) => {
      const lad = filterSorted(tables.ladder, (d) => d.channel === channel && d.metric === metric);
      const act = filterSorted(tables.actual, (d) => d.channel === channel && d.metric === metric && d.month !== "2025_baseline");
      const traceBar = { x: lad.map((d) => d.month), y: lad.map((d) => d.target_cumulative), name: "누적 목표", type: "bar", marker: { color: GRID, line: { width: 0 } } };
      const traceLine = { x: act.map((d) => d.month), y: act.map((d) => d.actual_cumulative), name: "누적 실적", mode: "lines+markers", line: { color: color, width: 3 }, marker: { size: 8 } };
      Plotly.react(id, [traceBar, traceLine], styleFig(320, `${channel} · ${metric}`, true), PCFG);
    });
  }

  function insightIcon(rate) {
    return rate >= 100 ? "🟢" : rate >= 90 ? "🟡" : "🔴";
  }

  function monthlyInsightRows(tables, channel, metric) {
    const lad = {}, act = {};
    tables.ladder.filter((d) => d.channel === channel && d.metric === metric).forEach((d) => { lad[d.month] = d.target_cumulative; });
    tables.actual.filter((d) => d.channel === channel && d.metric === metric && d.month !== "2025_baseline").forEach((d) => { act[d.month] = d.actual_cumulative; });
    const months = Object.keys(lad).filter((m) => m in act).sort();
    const rows = [];
    months.forEach((month) => {
      const target = lad[month];
      if (!target) return;
      const value = act[month];
      rows.push([month, target, value, (value / target) * 100]);
    });
    return rows;
  }

  function renderMonthlyInsights(tables) {
    const order = [
      { channel: "인스타그램", metric: "팔로워" }, { channel: "인스타그램", metric: "인게이지먼트" },
      { channel: "유튜브", metric: "구독자" }, { channel: "유튜브", metric: "조회수" },
    ];
    let html = "";
    order.forEach(({ channel, metric }) => {
      const rows = monthlyInsightRows(tables, channel, metric);
      if (!rows.length) {
        html += `<div class="chart-half"><b>${channel} · ${metric}</b><div class="caption">월별 목표 데이터가 없습니다.</div></div>`;
        return;
      }
      const worst = rows.reduce((m, r) => (r[3] < m[3] ? r : m), rows[0]);
      const latest = rows[rows.length - 1];
      let trend = "";
      if (rows.length >= 2) {
        const prev = rows[rows.length - 2][3];
        trend = " · 전월 대비 " + (latest[3] > prev ? "개선 중 ⬆" : latest[3] < prev ? "악화 중 ⬇" : "동일");
      }
      const items = rows.map((r) => `<div class="insight-item">${insightIcon(r[3])} <b>${r[0]}</b>: 목표 ${fmt(r[1])} 대비 실적 ${fmt(r[2])} (${r[3].toFixed(0)}%)</div>`).join("");
      html += `
      <div class="chart-half insight-panel">
        <b>${channel} · ${metric}</b>
        <div class="insight-summary">${insightIcon(worst[3])} 최저 달성월 ${worst[0]} (${worst[3].toFixed(0)}%) ·
          ${insightIcon(latest[3])} 최신 ${latest[0]} (${latest[3].toFixed(0)}%)${trend}</div>
        <details><summary>월별 상세 보기</summary><div class="insight-list">${items}</div></details>
      </div>`;
    });
    document.getElementById("insightPanelRow").innerHTML = html;
  }

  function declineBoxHtml(rows, unitLabel) {
    const vals = rows.map((d) => d.actual_value);
    const declineMonths = [];
    for (let i = 1; i < vals.length; i++) if (vals[i] - vals[i - 1] < 0) declineMonths.push(rows[i].month);
    if (!declineMonths.length) return `✅ ${unitLabel} 순감소 없이 꾸준히 증가/유지되었습니다.`;
    return `⚠️ ${vals.length}개월 중 <b>${declineMonths.length}개월</b> ${unitLabel} 순감소 발생 (${declineMonths.join(", ")})`;
  }

  function renderTrendCharts(tables) {
    const igF = filterSorted(tables.actual, (d) => d.channel === "인스타그램" && d.metric === "팔로워" && d.month !== "2025_baseline");
    const igFVals = igF.map((d) => d.actual_value);
    const igFColors = igFVals.map((v, i) => (i === 0 ? 0 : v - igFVals[i - 1]) < 0 ? C_BAD : C_GOOD);
    Plotly.react("igf", [{ x: igF.map((d) => d.month), y: igFVals, mode: "lines+markers", line: { color: C_IG, width: 3 }, marker: { size: 9, color: igFColors }, name: "IG 팔로워(월말)" }],
      styleFig(320, "인스타그램 팔로워 수 추이 (월말 절대치)", true), PCFG);
    document.getElementById("igFDeclineBox").innerHTML = declineBoxHtml(igF, "팔로워");

    const ytS = filterSorted(tables.actual, (d) => d.channel === "유튜브" && d.metric === "구독자" && d.month !== "2025_baseline");
    const ytSVals = ytS.map((d) => d.actual_value);
    const ytSColors = ytSVals.map((v, i) => (i === 0 ? 0 : v - ytSVals[i - 1]) < 0 ? C_BAD : C_GOOD);
    Plotly.react("yts", [{ x: ytS.map((d) => d.month), y: ytSVals, mode: "lines+markers", line: { color: C_YT, width: 3 }, marker: { size: 9, color: ytSColors }, name: "YT 구독자(월말)" }],
      styleFig(320, "유튜브 구독자 수 추이 (월말 절대치)", true), PCFG);
    document.getElementById("ytSDeclineBox").innerHTML = declineBoxHtml(ytS, "구독자");

    const igE = filterSorted(tables.actual, (d) => d.channel === "인스타그램" && d.metric === "인게이지먼트" && d.month !== "2025_baseline");
    Plotly.react("ige", [{ x: igE.map((d) => d.month), y: igE.map((d) => d.actual_value), type: "bar", name: "월별 인게이지먼트", marker: { color: C_IG } }],
      styleFig(320, "인스타그램 월별 인게이지먼트 발생량", true), PCFG);

    const ytV = filterSorted(tables.actual, (d) => d.channel === "유튜브" && d.metric === "조회수" && d.month !== "2025_baseline");
    Plotly.react("ytv", [{ x: ytV.map((d) => d.month), y: ytV.map((d) => d.actual_value), type: "bar", name: "월별 조회수", marker: { color: C_YT } }],
      styleFig(320, "유튜브 월별 조회수 발생량", true), PCFG);
  }

  function pearson(xs, ys) {
    const n = xs.length;
    if (n < 2) return null;
    const mx = xs.reduce((a, b) => a + b, 0) / n;
    const my = ys.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0, syy = 0;
    for (let i = 0; i < n; i++) {
      const dx = xs[i] - mx, dy = ys[i] - my;
      sxy += dx * dy; sxx += dx * dx; syy += dy * dy;
    }
    const denom = Math.sqrt(sxx * syy);
    return denom === 0 ? null : sxy / denom;
  }
  function corrWithSpend(rows, cols) {
    const out = [];
    cols.forEach((c) => {
      const pairs = rows.filter((r) => r.spend != null && r[c] != null);
      if (pairs.length < 2) return;
      const v = pearson(pairs.map((r) => r.spend), pairs.map((r) => r[c]));
      if (v !== null) out.push({ col: c, val: v });
    });
    out.sort((a, b) => a.val - b.val);
    return out;
  }
  function renderCorrChart(id, rows, cols) {
    const corr = corrWithSpend(rows, cols);
    const trace = { x: corr.map((c) => c.val), y: corr.map((c) => c.col), type: "bar", orientation: "h", marker: { color: corr.map((c) => (c.val < 0 ? C_BAD : C_GOOD)) } };
    const layout = styleFig(320, "소진금액과의 상관계수 (Pearson)", false);
    layout.xaxis = Object.assign({}, layout.xaxis, { range: [-1, 1] });
    Plotly.react(id, [trace], layout, PCFG);
  }

  function renderCrossAnalysis(tables) {
    const igPerf = tables.igPerf.filter((r) => r.month !== "TOTAL");
    const ytPerf = tables.ytPerf.filter((r) => r.month !== "TOTAL");

    let d = filterSorted(igPerf.filter((r) => r.spend != null), () => true);
    let layout = styleFig(360, "인스타그램: 소진금액 vs 인게이지먼트", true);
    layout.yaxis = Object.assign({}, layout.yaxis, { title: { text: "소진 금액(원)" } });
    layout.yaxis2 = { showgrid: false, zeroline: false, overlaying: "y", side: "right", title: { text: "인게이지먼트" } };
    Plotly.react("igdual", [
      { x: d.map((r) => r.month), y: d.map((r) => r.spend), name: "소진 금액(원)", type: "bar", marker: { color: "#FFC9DE" } },
      { x: d.map((r) => r.month), y: d.map((r) => r.engagement_total), name: "인게이지먼트", mode: "lines+markers", line: { color: C_IG, width: 3 }, yaxis: "y2" },
    ], layout, PCFG);

    d = filterSorted(ytPerf.filter((r) => r.spend != null), () => true);
    layout = styleFig(360, "유튜브: 소진금액 vs 조회수(페이드)", true);
    layout.yaxis = Object.assign({}, layout.yaxis, { title: { text: "소진 금액(원)" } });
    layout.yaxis2 = { showgrid: false, zeroline: false, overlaying: "y", side: "right", title: { text: "조회수" } };
    Plotly.react("ytdual", [
      { x: d.map((r) => r.month), y: d.map((r) => r.spend), name: "소진 금액(원)", type: "bar", marker: { color: "#FFD8A8" } },
      { x: d.map((r) => r.month), y: d.map((r) => r.views), name: "조회수(페이드)", mode: "lines+markers", line: { color: C_YT, width: 3 }, yaxis: "y2" },
    ], layout, PCFG);

    const igNum = igPerf.filter((r) => r.spend > 0);
    const ytNum = ytPerf.filter((r) => r.spend > 0);
    renderCorrChart("igcorr", igNum, ["engagement_total", "cpe", "impressions", "clicks", "reach", "profile_visit", "ctr", "cpc"]);
    renderCorrChart("ytcorr", ytNum, ["engagement_total", "impressions", "clicks", "views", "ctr", "vtr", "cpv"]);
    document.getElementById("igCorrNLabel").textContent = `${igNum.length}개월, 광고 집행월만`;
    document.getElementById("ytCorrNLabel").textContent = `${ytNum.length}개월, 광고 집행월만`;
    document.getElementById("igCorrCaption").textContent = `소진금액↑일수록 CPC·CPE는 낮아지고(음의 상관) 도달·참여는 늘어나는 경향이면 예산 확대가 유효하다는 신호입니다. 다만 표본이 ${igNum.length}개월뿐이라 참고용 지표입니다.`;
    document.getElementById("ytCorrCaption").textContent = `유튜브는 표본이 ${ytNum.length}개월뿐이라 상관계수의 통계적 의미는 제한적입니다 — 방향성 참고용입니다.`;

    d = filterSorted(igPerf.filter((r) => r.cpe != null), () => true);
    Plotly.react("igeff", [
      { x: d.map((r) => r.month), y: d.map((r) => r.cpe), name: "CPE(참여당 단가)", mode: "lines+markers", line: { color: C_IG } },
      { x: d.map((r) => r.month), y: d.map((r) => r.cpc), name: "CPC(클릭당 단가)", mode: "lines+markers", line: { color: "#862E9C" } },
    ], styleFig(320, "인스타그램 CPE / CPC 추이 (원)", true), PCFG);

    d = filterSorted(ytPerf.filter((r) => r.cpv != null), () => true);
    Plotly.react("yteff", [
      { x: d.map((r) => r.month), y: d.map((r) => r.cpv), name: "CPV(조회당 단가)", mode: "lines+markers", line: { color: C_YT } },
      { x: d.map((r) => r.month), y: d.map((r) => (r.vtr == null ? null : r.vtr * 1000)), name: "VTR×1000 (스케일 보정)", mode: "lines+markers", line: { color: "#1971C2" } },
    ], styleFig(320, "유튜브 CPV / VTR 추이", true), PCFG);
  }

  function renderRankChart(annual) {
    const ranked = sortByKey(annual, (r) => r.annual_progress_rate);
    const colors = ranked.map((r) => (r.annual_progress_rate < 0.8 ? C_BAD : r.annual_progress_rate < 0.95 ? C_WARN : C_GOOD));
    const trace = {
      x: ranked.map((r) => r.annual_progress_rate * 100), y: ranked.map((r) => `${r.channel} · ${r.metric}`),
      type: "bar", orientation: "h", marker: { color: colors },
      text: ranked.map((r) => pct(r.annual_progress_rate, 0)), textposition: "outside",
    };
    const layout = styleFig(Math.max(220, 50 + ranked.length * 45), "연간 목표 달성률 순위 (낮은 순)", false);
    layout.shapes = [{ type: "line", x0: 100, x1: 100, y0: 0, y1: 1, yref: "paper", line: { dash: "dash", color: C_TARGET } }];
    const maxVal = Math.max(120, Math.max(...ranked.map((r) => r.annual_progress_rate * 100)) * 1.1);
    layout.xaxis = Object.assign({}, layout.xaxis, { range: [0, maxVal] });
    Plotly.react("rankchart", [trace], layout, PCFG);
  }

  // ---------------------------------------------------------------------
  // 자동 진단 / 처방 (업로드된 데이터에 대해서만 사용 - 규칙 기반 범용 생성기)
  // ---------------------------------------------------------------------
  function contentDependencyNotes(actualRows) {
    const flow = actualRows.filter((d) => d.metric_type === "flow" && d.month !== "2025_baseline");
    const byKey = {};
    flow.forEach((d) => {
      const key = d.channel + "|" + d.metric;
      (byKey[key] = byKey[key] || []).push(d);
    });
    const notes = [];
    Object.entries(byKey).forEach(([key, rows]) => {
      const total = rows.reduce((s, r) => s + (r.actual_value || 0), 0);
      if (total <= 0) return;
      const maxRow = rows.reduce((m, r) => (r.actual_value > m.actual_value ? r : m), rows[0]);
      const share = maxRow.actual_value / total;
      if (share > 0.5) {
        const [channel, metric] = key.split("|");
        notes.push(`<li><b>${channel} · ${metric}</b>: ${maxRow.month} 한 달이 전체 발생량의 ${pct(share, 0)}를 차지합니다 — 특정 월 의존도가 높아 지속가능성 점검이 필요합니다.</li>`);
      }
    });
    return notes;
  }
  function momAnomalyNotes(perf, label) {
    const rows = filterSorted(perf.filter((r) => r.month !== "TOTAL"), () => true);
    const notes = [];
    ["ctr", "cpc", "cpe", "cpv"].forEach((f) => {
      if (!rows.some((r) => r[f] != null)) return;
      for (let i = 1; i < rows.length; i++) {
        const prev = rows[i - 1][f], cur = rows[i][f];
        if (prev == null || cur == null || prev === 0) continue;
        const change = (cur - prev) / Math.abs(prev);
        if (Math.abs(change) > 0.5) {
          notes.push(`<li><b>${label} ${rows[i].month}</b>: ${f.toUpperCase()}가 전월 대비 ${pct(Math.abs(change), 0)} ${change > 0 ? "상승" : "하락"}했습니다 — 소재/타겟/캠페인 목표 설정을 점검해보세요.</li>`);
        }
      }
    });
    return notes;
  }

  function makeItemHtml(icon, num, numSuffix, title, bodyHtml) {
    return `<div class="item">
    <div class="item-head">
      <span class="item-icon">${icon}</span>
      <span class="item-num">${num}${numSuffix}</span>
      <span class="item-title">${title}</span>
      <button class="item-del" type="button" hidden title="이 항목 삭제">✕</button>
    </div>
    <p class="item-body">${bodyHtml}</p>
  </div>`;
  }

  function generateDiagnosis(annual) {
    const ranked = sortByKey(annual, (r) => r.annual_progress_rate);
    let html = "";
    ranked.forEach((row, i) => {
      const rate = row.annual_progress_rate;
      const icon = rate >= 0.95 ? "🟢" : rate >= 0.8 ? "🟡" : "🔴";
      const gap = row.annual_target - row.actual_cumulative;
      const title = `${row.channel} · ${row.metric} — 연간 진척률 ${pct(rate, 0)}`;
      let body = `누적 실적 ${fmt(row.actual_cumulative)} / 연간 목표 ${fmt(row.annual_target)} (${pct(rate, 0)}). `;
      body += gap > 0 ? `잔여 ${fmt(gap)}이 남아 있습니다. ` : `목표를 ${fmt(-gap)} 초과 달성했습니다. `;
      body += `월간 목표(캐스케이드) 대비로는 ${pct(row.monthly_progress_rate, 0)} 수준입니다. `;
      if (rate < 0.8) body += `<strong>목표 대비 뒤처져 있어 우선 점검이 필요합니다.</strong>`;
      else if (rate > 1.1) body += `목표를 크게 초과했다면, 특정 월에 성과가 쏠려있지 않은지 함께 확인해보는 것이 좋습니다.`;
      else body += `대체로 정상 궤도 위에 있습니다.`;
      html += makeItemHtml(icon, i + 1, ".", title, body);
    });
    return html;
  }

  function generatePrescription(tables) {
    const ranked = sortByKey(tables.annual, (r) => r.annual_progress_rate);
    const worst = ranked[0], best = ranked[ranked.length - 1];
    const bs = {};
    tables.budgetSummary.forEach((r) => { bs[r.item] = r.value; });

    const items = [];

    let body1;
    if (bs["26년 총예산"] != null && worst && best) {
      body1 = `가장 뒤처진 지표는 <b>${worst.channel} · ${worst.metric}</b>(${pct(worst.annual_progress_rate, 0)}), 가장 앞서 있는 지표는
      <b>${best.channel} · ${best.metric}</b>(${pct(best.annual_progress_rate, 0)})입니다. 연간 총예산 ${fmtWon(bs["26년 총예산"])} 중
      아직 집행되지 않은 잔여 예산이 있다면, 초과 달성 중인 지표에 배정된 몫 일부를 뒤처진 지표 쪽으로 전환하는 것을 검토해볼 만합니다.`;
    } else {
      body1 = `예산 데이터가 부족해 구체적인 배분 제안은 생략합니다.`;
    }
    items.push(["잔여 예산 재배분 검토", body1]);

    const contentNotes = contentDependencyNotes(tables.actual);
    items.push(["콘텐츠/캠페인 의존도 리스크",
      contentNotes.length ? `<ul>${contentNotes.join("")}</ul>` : `특정 월에 쏠린 성과는 발견되지 않았습니다.`]);

    const anomalyNotes = [...momAnomalyNotes(tables.igPerf, "인스타그램"), ...momAnomalyNotes(tables.ytPerf, "유튜브")];
    items.push(["월별 효율지표 급변 감지 (전월 대비 ±50% 이상)",
      anomalyNotes.length ? `<ul>${anomalyNotes.join("")}</ul>` : `전월 대비 50% 이상 급변한 효율지표는 발견되지 않았습니다.`]);

    items.push(["하반기 단가 가정 재검증",
      `하반기 집행 계획(참고 테이블의 unit_cost_note)에 적힌 단가 가정과 실제 집행 단가(CPE/CPV 등)를 비교해, 계획대로 예산을
      소진했을 때 목표한 물량을 실제로 달성할 수 있는지 재확인하는 것을 권장합니다.`]);

    let html = "";
    items.forEach((it, i) => { html += makeItemHtml("", i + 1, ")", it[0], it[1]); });
    return html;
  }

  function arrayToTable(rows, columns) {
    if (!rows || !rows.length) return '<table class="ref-table"><tbody><tr><td>데이터 없음</td></tr></tbody></table>';
    let html = '<table class="ref-table"><thead><tr>';
    columns.forEach((c) => { html += `<th>${c}</th>`; });
    html += "</tr></thead><tbody>";
    rows.forEach((row) => {
      html += "<tr>";
      columns.forEach((c) => {
        let v = row[c];
        if (v === null || v === undefined) v = "";
        else if (typeof v === "number") v = v.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
        html += `<td>${v}</td>`;
      });
      html += "</tr>";
    });
    html += "</tbody></table>";
    return html;
  }

  const H2_CASCADE_COLS = ["channel", "metric", "month", "kpi_cumulative_target", "expected_actual_increment", "kpi_achievement_rate", "ad_budget"];
  const H2_MEDIA_PLAN_COLS = ["media_item", "spend_may", "spend_jun", "spend_jul", "budget_aug_nov_total", "budget_aug_nov_monthly_even", "expected_result", "unit_cost_note"];

  // ---------------------------------------------------------------------
  // 전체 렌더 오케스트레이션
  // ---------------------------------------------------------------------
  // ---------------------------------------------------------------------
  // 개요 탭 기간(월별) 필터 - 캐스케이드 차트 + 월별 인사이트에만 적용
  // ---------------------------------------------------------------------
  let currentTables = null;

  function getAvailableMonths(tables) {
    const months = new Set();
    tables.ladder.forEach((d) => months.add(d.month));
    tables.actual.forEach((d) => { if (d.month !== "2025_baseline") months.add(d.month); });
    return Array.from(months).sort();
  }

  function fillMonthSelect(sel, months, preferredValue) {
    sel.innerHTML = months.map((m) => `<option value="${m}">${m}</option>`).join("");
    sel.value = months.includes(preferredValue) ? preferredValue : months[months.length - 1];
  }

  function populatePeriodSelectors(tables) {
    const startSel = document.getElementById("periodStart");
    const endSel = document.getElementById("periodEnd");
    if (!startSel || !endSel) return;
    const months = getAvailableMonths(tables);
    if (!months.length) return;
    startSel.innerHTML = months.map((m) => `<option value="${m}">${m}</option>`).join("");
    startSel.value = months[0];
    endSel.innerHTML = months.map((m) => `<option value="${m}">${m}</option>`).join("");
    endSel.value = months[months.length - 1];
  }

  function filterTablesByPeriod(tables, start, end) {
    return Object.assign({}, tables, {
      ladder: tables.ladder.filter((d) => d.month >= start && d.month <= end),
      actual: tables.actual.filter((d) => d.month >= start && d.month <= end),
    });
  }

  function applyPeriodFilter() {
    const startSel = document.getElementById("periodStart");
    const endSel = document.getElementById("periodEnd");
    if (!currentTables || !startSel || !endSel || !startSel.value || !endSel.value) return;
    const start = startSel.value, end = endSel.value;
    const filtered = filterTablesByPeriod(currentTables, start, end);
    renderCascadeCharts(filtered);
    renderMonthlyInsights(filtered);
    const caption = document.getElementById("periodCaption");
    if (caption) caption.textContent = `📅 ${start} ~ ${end} 기간만 표시 중`;
  }

  function initPeriodFilter() {
    const startSel = document.getElementById("periodStart");
    const endSel = document.getElementById("periodEnd");
    if (!startSel || !endSel) return;
    startSel.addEventListener("change", function () {
      // 종료월 선택지를 시작월 이후로 제한 (시작월 > 종료월이 되는 것을 방지)
      const months = getAvailableMonths(currentTables).filter((m) => m >= startSel.value);
      const prevEnd = endSel.value;
      fillMonthSelect(endSel, months, prevEnd);
      applyPeriodFilter();
    });
    endSel.addEventListener("change", applyPeriodFilter);
  }

  function renderAll(tables, opts) {
    opts = opts || {};
    currentTables = tables;
    renderKpiCards(tables.annual);
    populatePeriodSelectors(tables);
    applyPeriodFilter();
    renderTrendCharts(tables);
    renderCrossAnalysis(tables);
    renderRankChart(tables.annual);
    document.getElementById("h2CascadeTable").innerHTML = arrayToTable(tables.h2Cascade, H2_CASCADE_COLS);
    document.getElementById("h2MediaPlanTable").innerHTML = arrayToTable(tables.h2MediaPlan, H2_MEDIA_PLAN_COLS);
    if (opts.autoDiagnosis) {
      document.getElementById("diagnosisContent").innerHTML = generateDiagnosis(tables.annual);
      document.getElementById("prescriptionContent").innerHTML = generatePrescription(tables);
    }
  }

  // ---------------------------------------------------------------------
  // 업로드 UI 연결
  // ---------------------------------------------------------------------
  let originalDiagnosisHTML = null;
  let originalPrescriptionHTML = null;

  function showStatus(msg, kind) {
    const el = document.getElementById("uploadStatus");
    el.hidden = false;
    el.innerHTML = msg;
    el.className = "upload-status " + kind;
  }

  // ---------------------------------------------------------------------
  // 진단/처방 코멘트를 페이지에서 바로 편집 ("여기서 바로 수정하기" 버튼)
  // ---------------------------------------------------------------------
  function renumberItems() {
    document.querySelectorAll("#diagnosisContent .item").forEach((el, i) => {
      const n = el.querySelector(".item-num");
      if (n) n.textContent = (i + 1) + ".";
    });
    document.querySelectorAll("#prescriptionContent .item").forEach((el, i) => {
      const n = el.querySelector(".item-num");
      if (n) n.textContent = (i + 1) + ")";
    });
  }

  function makeEditableItemEl(kind) {
    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `<div class="item-head">
      <span class="item-icon" contenteditable="true">${kind === "diag" ? "⚪" : ""}</span>
      <span class="item-num"></span>
      <span class="item-title" contenteditable="true">새 항목 제목을 입력하세요</span>
      <button class="item-del" type="button" title="이 항목 삭제">✕</button>
    </div>
    <p class="item-body" contenteditable="true">내용을 입력하세요.</p>`;
    return div;
  }

  function setCommentsEditable(state) {
    document.querySelectorAll(
      "#diagnosisContent .item-icon, #diagnosisContent .item-title, #diagnosisContent .item-body, " +
      "#prescriptionContent .item-icon, #prescriptionContent .item-title, #prescriptionContent .item-body"
    ).forEach((el) => { el.contentEditable = state ? "true" : "false"; });
    document.querySelectorAll("#diagnosisContent .item-del, #prescriptionContent .item-del").forEach((btn) => { btn.hidden = !state; });
    document.getElementById("t4").classList.toggle("editing", state);
    document.getElementById("editToolbar").hidden = !state;
    document.getElementById("startEditBtn").hidden = state;
  }

  function buildSheetPasteText() {
    const lines = ["구분\t아이콘\t제목\t내용"];
    document.querySelectorAll("#diagnosisContent .item").forEach((el) => {
      lines.push(["진단", el.querySelector(".item-icon").textContent.trim(),
        el.querySelector(".item-title").textContent.trim(),
        el.querySelector(".item-body").textContent.trim()].join("\t"));
    });
    document.querySelectorAll("#prescriptionContent .item").forEach((el) => {
      lines.push(["처방", el.querySelector(".item-icon").textContent.trim(),
        el.querySelector(".item-title").textContent.trim(),
        el.querySelector(".item-body").textContent.trim()].join("\t"));
    });
    return lines.join("\n");
  }

  function copyPlainText(text, onDone) {
    function legacyCopy() {
      const ta = document.createElement("textarea");
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand("copy"); } catch (e) { /* 무시 */ }
      document.body.removeChild(ta);
      onDone();
    }
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(onDone).catch(legacyCopy);
    } else {
      legacyCopy();
    }
  }

  let diagEditSnapshot = null;
  let presEditSnapshot = null;

  function initCommentEditing() {
    const startBtn = document.getElementById("startEditBtn");
    if (!startBtn) return;

    renumberItems();

    startBtn.addEventListener("click", function () {
      diagEditSnapshot = document.getElementById("diagnosisContent").innerHTML;
      presEditSnapshot = document.getElementById("prescriptionContent").innerHTML;
      setCommentsEditable(true);
    });

    document.getElementById("addDiagBtn").addEventListener("click", function () {
      document.getElementById("diagnosisContent").appendChild(makeEditableItemEl("diag"));
      renumberItems();
    });
    document.getElementById("addPresBtn").addEventListener("click", function () {
      document.getElementById("prescriptionContent").appendChild(makeEditableItemEl("pres"));
      renumberItems();
    });

    document.addEventListener("click", function (e) {
      if (e.target.classList.contains("item-del")) {
        e.target.closest(".item").remove();
        renumberItems();
      }
    });

    document.getElementById("doneEditBtn").addEventListener("click", function () {
      setCommentsEditable(false);
    });

    document.getElementById("cancelEditBtn").addEventListener("click", function () {
      if (diagEditSnapshot !== null) document.getElementById("diagnosisContent").innerHTML = diagEditSnapshot;
      if (presEditSnapshot !== null) document.getElementById("prescriptionContent").innerHTML = presEditSnapshot;
      renumberItems();
      setCommentsEditable(false);
    });

    document.getElementById("copyForSheetBtn").addEventListener("click", function () {
      const btn = this;
      const original = btn.textContent;
      copyPlainText(buildSheetPasteText(), function () {
        btn.textContent = "✓ 복사됨 — 구글시트에 붙여넣으세요";
        setTimeout(function () { btn.textContent = original; }, 2500);
      });
    });
  }

  function init() {
    originalDiagnosisHTML = document.getElementById("diagnosisContent").innerHTML;
    originalPrescriptionHTML = document.getElementById("prescriptionContent").innerHTML;
    initCommentEditing();
    initPeriodFilter();

    document.getElementById("xlsxFileInput").addEventListener("change", function (e) {
      const file = e.target.files[0];
      if (!file) return;
      showStatus("⏳ 파일을 읽는 중...", "busy");
      const reader = new FileReader();
      reader.onload = function (ev) {
        try {
          const data = new Uint8Array(ev.target.result);
          const wb = XLSX.read(data, { type: "array", cellDates: true });
          const tables = parseWorkbook(wb);
          renderAll(tables, { autoDiagnosis: true });
          setCommentsEditable(false);
          renumberItems();
          document.getElementById("revertBtn").hidden = false;
          document.getElementById("kpiSnapshotTitle").textContent = "채널별 핵심 KPI 스냅샷 (업로드한 데이터 기준)";
          showStatus(`✅ "${file.name}" 데이터로 이 브라우저에서만 갱신했습니다. 진단·처방 문구는 자동 생성된 일반 분석입니다
            (새로고침하면 원본으로 돌아가고, 다른 방문자에게는 공유되지 않습니다). 실제 공개 대시보드를 갱신하려면
            로컬에서 <code>publish_to_github.bat</code>을 실행하세요.`, "ok");
        } catch (err) {
          console.error(err);
          showStatus("❌ 파일을 읽지 못했습니다: " + err.message, "err");
        } finally {
          e.target.value = "";
        }
      };
      reader.onerror = function () { showStatus("❌ 파일을 읽는 중 오류가 발생했습니다.", "err"); };
      reader.readAsArrayBuffer(file);
    });

    document.getElementById("revertBtn").addEventListener("click", function () {
      renderAll(window.__BASELINE_TABLES__, { autoDiagnosis: false });
      document.getElementById("diagnosisContent").innerHTML = originalDiagnosisHTML;
      document.getElementById("prescriptionContent").innerHTML = originalPrescriptionHTML;
      setCommentsEditable(false);
      renumberItems();
      document.getElementById("kpiSnapshotTitle").textContent = "채널별 핵심 KPI 스냅샷 (2026년 08월 기준)";
      document.getElementById("revertBtn").hidden = true;
      showStatus("↩ 원본 데이터로 되돌렸습니다.", "ok");
    });

    // 최초 페이지 로드시 캐스케이드/인사이트는 Python이 이미 정적으로 그려뒀지만, 기간 선택 드롭다운
    // 자체는 JS가 채워야 하므로 원본 데이터로 한 번 초기화한다 (내용은 동일하게 다시 그려짐).
    if (window.__BASELINE_TABLES__) {
      currentTables = window.__BASELINE_TABLES__;
      populatePeriodSelectors(currentTables);
      applyPeriodFilter();
    }
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
