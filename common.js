// common.js
(function () {
  // ====== External URLs ======
  const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS7WIgpgiy-yAbJNaLbF3Vm4p9qPxi_WpHtW8yWi4MeCUigElvsY7y3T-E6OG5kIxH711SSwfBPtFG7/pub?gid=0&single=true&output=csv";
  const WPL_JSON_URL = "wpllist.json";

  // ===== 상태 =====
  let allRows = [], rows = [];
  let selectedStatus = null;
  let selectedTitle  = null;
  let detailStatusFilter = "";

  // ===== DOM =====
  const updateDate      = document.getElementById("updateDate");
  const regionBox       = document.getElementById("regionBox");
  const projBox         = document.getElementById("projBox");
  const ownerGrpBox     = document.getElementById("ownerGrpBox");
  const monthBox        = document.getElementById("monthBox");
  const tbody           = document.getElementById("tbody");
  const statusBadge     = document.getElementById("statusBadge");
  const selStatusTxt    = document.getElementById("selStatusTxt");
  const clearStatus     = document.getElementById("clearStatus");
  const titleBadge      = document.getElementById("titleBadge");
  const selTitleTxt     = document.getElementById("selTitleTxt");
  const clearTitle      = document.getElementById("clearTitle");
  const resetBtn        = document.getElementById("resetBtn");
  const detailStatusSel = document.getElementById("detailStatusFilter");

  // ===== 요약 패널 & Page# Total =====
  const summaryPanel      = document.getElementById("summaryPanel");
  const summaryTblBody    = document.querySelector("#summaryTbl tbody");
  const summaryPageTotal  = document.getElementById("summaryPageTotal");
  const pageTotalVal      = document.getElementById("pageTotalVal");

  // ===== 팝오버 =====
  const commentPop      = document.getElementById('commentPop');
  const commentPopTitle = document.getElementById('commentPopTitle');
  const commentPopText  = document.getElementById('commentPopText');
  if (commentPop) {
    commentPop.style.position = 'fixed';
    commentPop.style.display  = 'none';
    commentPop.style.zIndex   = '9999';
  }

  // ===== WPL JSON =====
  let wplInfoByLocale = {};
  async function loadWPLJson() {
    try {
      const res = await fetch(WPL_JSON_URL, { cache: "no-store" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      wplInfoByLocale = await res.json();
      console.log("WPL JSON loaded:", Object.keys(wplInfoByLocale).length);
    } catch (err) {
      console.error("Failed to load WPL JSON:", err);
      wplInfoByLocale = {};
    }
  }
  function getWPLByLocale(locale) {
    const k = locale || "";
    return (
      wplInfoByLocale[k] ||
      wplInfoByLocale[k.toLowerCase?.()] ||
      wplInfoByLocale[k.toUpperCase?.()] ||
      null
    );
  }

  // ===== 초기 로드 =====
  window.addEventListener("load", async () => {
    await loadWPLJson();
    if (location.protocol === 'file:') {
      console.log('file:// 모드 - CSV fetch 생략.');
      return;
    }
    try {
      await loadFromUrl(DATA_URL);
    } catch (err) {
      console.error(err);
      alert("데이터를 불러오지 못했습니다.");
    }
  });

  // ===== CSV URL 로딩 =====
  async function loadFromUrl(url) {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error("HTTP Error " + res.status);
    const text = await res.text();
    const wb = XLSX.read(text, { type: "string" });
    await loadWorkbook(wb);
  }

  // ===== 업로드 파일 로딩 =====
  async function loadFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    await loadWorkbook(wb);
  }

  // ===== 유틸 =====
  const norm = (s) => (s ?? "").toString().trim();
  const parseDate = (s) => {
    if (!s) return null;
    if (s instanceof Date) return s;
    if (!isNaN(s) && typeof s === "number") {
      const d = XLSX.SSF.parse_date_code(s);
      if (d) return new Date(Date.UTC(d.y, d.m - 1, d.d));
    }
    const d = new Date(s);
    return isNaN(+d) ? null : d;
  };
  const escapeHtml = (s) =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
  function parsePageNum(v) {
    if (v == null) return 1;
    if (typeof v === "number" && !isNaN(v)) return Math.max(1, Math.round(v));
    const s = String(v);
    const nums = [...s.matchAll(/\d+/g)].map(m => parseInt(m[0], 10)).filter(n => !isNaN(n));
    if (nums.length === 0) return 1;
    const sum = nums.reduce((a,b)=>a+b,0);
    return Math.max(1, sum);
  }

  // ===== Locale =====
  const normalizeLocale = (loc) => {
    const v = (loc || "").toString().trim();
    if (/^I$/i.test(v)) return "IR-ar";
    const m = v.match(/^L-([A-Za-z]{2})$/i);
    if (m) return `Levant-${m[1].toLowerCase()}`;
    return v;
  };
  const extractLocaleFromProject = (p) => {
    if (!p) return "";
    const raw = String(p);
    const s = raw.replace(/\s*\(.*?\)\s*/g, "").trim();
    const parts = s.split(" - ");
    return normalizeLocale(parts.length > 1 ? parts[parts.length - 1] : s);
  };

  // ===== 담당자 그룹 =====
  const emailGroupMap = {
    "sejun.kim@lge.com": "김세준책임",
    "jinho1031.kim@lge.com": "김진호선임",
    "joanna.ryu@lge.com": "류예원책임",
  };
  const extractEmail = (s) => {
    if (!s) return "";
    const m = String(s).match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    return m ? m[0].toLowerCase() : String(s).toLowerCase();
  };
  const groupRequester = (v) => emailGroupMap[extractEmail(v)] || "그외";

  // ===== 상태 그룹 =====
  function mapStatusGrouped(raw) {
    const s = (raw || "").toString().trim();
    const sl = s.toLowerCase();
    if (sl === "change req. need close") return null;
    if (/(^|\s)(closed|publishing)(\s|$)/i.test(s)) return "완료";
    if (/(^|\s)(client\s*review|wpl\s*review)(\s|$)/i.test(s)) return "법인리뷰";
    if (/(^|\s)(new\s*request)(\s|$)/i.test(s)) return "사전검토";
    if (/(^|\s)(in\s*progress|request\s*clarification)(\s|$)/i.test(s)) return "진행중";
    if (/(^|\s)(cancelled)(\s|$)/i.test(s)) return "취소";
    return null;
  }

  // ===== 워크북 처리 =====
  async function loadWorkbook(wb) {
    const sheetName = wb.SheetNames.includes("raw") ? "raw" : wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];

    let raw = [];
    const ref = sheet["!ref"];
    if (ref) {
      const range = XLSX.utils.decode_range(ref);
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      const rowsMeta = sheet["!rows"] || [];
      const visible = aoa.filter((row, i) => !(rowsMeta[range.s.r + i] && rowsMeta[range.s.r + i].hidden));
      if (visible.length) {
        const headers = visible[0].map((h) => String(h ?? "").trim());
        for (let i = 1; i < visible.length; i++) {
          const obj = {};
          for (let c = 0; c < headers.length; c++) obj[headers[c]] = visible[i][c];
          raw.push(obj);
        }
      }
    } else {
      raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
    }

    allRows = normalizeRows(raw);
    rows = allRows.slice();
    renderFilters(rows);
    updateRangeBadge(allRows);
    renderAll();
  }

  // ===== 정규화 =====
  function normalizeRows(raw) {
    return raw
      .map((r) => {
        const region   = norm(r["Region"] || "");
        const proj     = norm(r["Project Name"] || r["Project"] || "");
        const locale   = extractLocaleFromProject(proj) || normalizeLocale(norm(r["Locale"] || ""));
        const subcat   = norm(r["Sub Category"] || "");
        const ownerRaw = r["BU Requestor Name"] || r["Requestor"] || r["Requester"] || "";
        const ownerGrp = groupRequester(ownerRaw);
        const date     = parseDate(r["Target staging date"]);
        const month    = date ? `${date.getUTCMonth() + 1}월` : null;
        const id       = norm(r["PTT Task ID"] || r["GR ID"] || r["ID"] || "");
        const pgRaw    = r["Pg#"] ?? r["Page#"] ?? r["PgNo"] ?? r["Page No"] ?? r["Page"];
        const pageNo   = pgRaw == null ? "" : (typeof pgRaw === "number" ? String(pgRaw) : norm(pgRaw));
        const pageNum  = parsePageNum(pgRaw);
        const cancelReason = norm(r["Cancelled Reason"] || r["Cancel Reason"] || "");
        const lastComment  = norm(r["Last comment"] || r["Last Comment"] || r["Last update"] || "");
        let rawTitle  = norm(r["Global Request Title"] || r["Title"] || r["Description"] || "");
        rawTitle      = rawTitle.replace(/^GR\d+-W\d+-\s*/i, "");
        rawTitle      = rawTitle.replace(/^GR[0-9A-Za-z\-]+-\s*/i, "");
        const title   = rawTitle;
        const statusRaw = r["Task Status in PTT"] || r["Status"] || "";
        const statusGrp = mapStatusGrouped(statusRaw);
        return { region, proj, locale, subcat, ownerGrp, date, month, id, statusGrp, title, pageNo, pageNum, cancelReason, lastComment };
      })
      .filter((r) => (r.region || r.id || r.title) && r.statusGrp);
  }

  // ===== 필터 렌더 =====
  function uniq(arr) { return [...new Set(arr.filter((v) => v !== null && v !== ""))].sort((a, b) => (a > b ? 1 : a < b ? -1 : 0)); }
  function sortMonthDesc(arr) { return [...arr].filter(v => v).sort((a, b) => parseInt(b) - parseInt(a)); }
  function fillBox(box, values) {
    if (!box) return;
    box.innerHTML = values.map((v) => `
      <label class="filter-item">
        <input type="checkbox" value="${escapeHtml(String(v))}"/>
        <span class="checkmark"></span>
        <span>${escapeHtml(String(v))}</span>
      </label>`).join("");
    box.querySelectorAll('input[type="checkbox"]').forEach((cb) => cb.addEventListener("change", applyFilters));
  }
  function getChecked(box) {
    if (!box) return new Set();
    return new Set([...box.querySelectorAll('input[type="checkbox"]:checked')].map((cb) => cb.value));
  }
  function updateLocaleBoxByRegion() {
    if (!projBox) return;
    const rset = getChecked(regionBox);
    let base = allRows;
    if (rset.size) base = allRows.filter(r => rset.has(r.region));
    const oldChecked = getChecked(projBox);
    const candidates = uniq(base.map(r => r.locale));
    fillBox(projBox, candidates);
    projBox.querySelectorAll('input[type="checkbox"]').forEach(cb => { if (oldChecked.has(cb.value)) cb.checked = true; });
  }
  function renderFilters(data) {
    fillBox(regionBox, uniq(data.map((r) => r.region)));
    fillBox(projBox,   uniq(data.map((r) => r.locale)));
    const order = ["김세준책임", "김진호선임", "이유진사원", "류예원책임", "그외"];
    const present = new Set(data.map((r) => r.ownerGrp));
    if (ownerGrpBox) {
      ownerGrpBox.innerHTML = order.filter((x) => present.has(x)).map((v) => `
        <label class="filter-item">
          <input type="checkbox" value="${v}"/><span class="checkmark"></span><span>${v}</span>
        </label>`).join("");
      ownerGrpBox.querySelectorAll('input[type="checkbox"]').forEach((cb) => cb.addEventListener("change", applyFilters));
    }
    fillBox(monthBox, sortMonthDesc(uniq(data.map((r) => r.month))));
  }

  // ===== KPI & 범위 =====
  function computeKPIs(data) {
    const total = data.length;
    return {
      total,
      newCnt: data.filter((r) => r.statusGrp === "사전검토").length,
      inProg: data.filter((r) => r.statusGrp === "진행중").length,
    };
  }
  function updateRangeBadge(base) {
    const dates = base.map((r) => r.date).filter(Boolean).sort((a, b) => a - b);
    if (dates.length) {
      const s = dates[0].toISOString().slice(0, 10);
      const e = dates[dates.length - 1].toISOString().slice(0, 10);
      if (updateDate) updateDate.textContent = `Row-data Date: ${s} ~ ${e} (${base.length.toLocaleString()}건)`;
    } else {
      if (updateDate) updateDate.textContent = `Row-data Date: -`;
    }
  }

  // ===== 도넛 =====
  function computeStatusShare(data) {
    const total = data.length || 0;
    const counts = { 완료: 0, 법인리뷰: 0, 사전검토: 0, 진행중: 0, 취소: 0 };
    for (const r of data) counts[r.statusGrp] = (counts[r.statusGrp] || 0) + 1;
    return { total, counts };
  }
  function renderStatusDonut() {
    const { total, counts } = computeStatusShare(rows);
    const palette = { 완료: "#16a34a", 법인리뷰: "#a78bfa", 사전검토: "#60a5fa", 진행중: "#f59e0b", 취소: "#ef4444" };
    const data = Object.keys(counts).map((name) => {
      const y = total ? (counts[name] / total) * 100 : 0;
      return { name, y, raw: counts[name], color: palette[name], sliced: selectedStatus === name, selected: selectedStatus === name };
    });
    Highcharts.chart("chartStatusDonut", {
      chart: { backgroundColor: "transparent", type: "pie" },
      title: { text: "" },
      credits: { enabled: false },
      exporting: { enabled: false },
      tooltip: { pointFormatter: function () { return `<span style="color:${this.color}">\u25CF</span> <b>${this.name}</b><br/>비율: <b>${Highcharts.numberFormat(this.y, 1)}%</b><br/>건수: <b>${this.raw.toLocaleString()}</b>`; } },
      accessibility: { point: { valueSuffix: "%" } },
      plotOptions: {
        pie: {
          innerSize: "60%",
          dataLabels: { enabled: true, formatter: function () { if (this.y < 3) return null; return `${this.point.name}: ${Highcharts.numberFormat(this.y, 1)}%`; }, style: { color: "#e6e8ec", textOutline: "none" } },
          showInLegend: true,
          point: { events: { click: function () { selectedStatus = selectedStatus === this.name ? null : this.name; if (selectedStatus) { selStatusTxt.textContent = selectedStatus; statusBadge.style.display = "inline-flex"; } else { statusBadge.style.display = "none"; } renderStatusDonut(); renderTitleStatusChart(); renderSummaryForSelectedTitle(); renderTable(); } } }
        }
      },
      legend: { itemStyle: { color: "#e5e7eb" }, itemHoverStyle: { color: "#d1d5db" } },
      series: [{ name: "비중", data }],
    });
  }

  // ===== 제목별 100% 스택 =====
  function computeTitleStatusByPages(data) {
    const bucket = new Map();
    for (const r of data) {
      const t = r.title || "—";
      if (!bucket.has(t)) bucket.set(t, { totalPages: 0, countsPages: { 완료:0, 법인리뷰:0, 사전검토:0, 진행중:0, 취소:0 } });
      const b = bucket.get(t);
      const p = r.pageNum || 1;
      b.totalPages += p;
      b.countsPages[r.statusGrp] = (b.countsPages[r.statusGrp] || 0) + p;
    }
    const arr = [...bucket.entries()].map(([title, obj]) => ({ title, totalPages: obj.totalPages, countsPages: obj.countsPages }));
    arr.sort((a, b) => b.totalPages - a.totalPages);
    return arr;
  }
  function renderTitleStatusChart() {
    const scoped = selectedStatus ? rows.filter(r => r.statusGrp === selectedStatus) : rows;
    const top = computeTitleStatusByPages(scoped);
    const barPointWidth = 30, barGap = 20, basePadding = 80;
    const h = Math.max(240, top.length * (barPointWidth + barGap) + basePadding);
    const container = document.getElementById("chartTitleStatus");
    if (container) container.style.height = `${h}px`;
    const categories = top.map(x => x.title.length > 50 ? x.title.slice(0,49) + "…" : x.title);
    const palette = { 완료:"#16a34a", 법인리뷰:"#a78bfa", 사전검토:"#60a5fa", 진행중:"#f59e0b", 취소:"#ef4444" };
    const statuses = ["완료","법인리뷰","사전검토","진행중","취소"];
    const seriesByGroup = { 완료:[], 법인리뷰:[], 사전검토:[], 진행중:[], 취소:[] };
    for (const item of top) for (const s of statuses) seriesByGroup[s].push(item.countsPages[s] || 0);
    const makePoint = (g, idx, y) => {
      if (!selectedTitle) return { y };
      const titleAtIdx = top[idx]?.title || "";
      if (titleAtIdx === selectedTitle) return { y };
      const dim = Highcharts.color(palette[g]).setOpacity(0.35).get();
      return { y, color: dim };
    };
    Highcharts.chart('chartTitleStatus', {
      chart: { backgroundColor:'transparent', type:'bar', height: h },
      title: { text:''},
      xAxis: { categories, labels: { style: { color:'#cbd5e1' } }, lineColor: 'rgba(255,255,255,.15)' },
      yAxis: { min: 0, title: { text: '비율(%)', align:'high' }, labels: { formatter(){ return this.value + '%'; }, style:{ color:'#cbd5e1' } }, gridLineColor:'rgba(255,255,255,.08)' },
      tooltip: { shared: true, formatter: function(){ const idx = this.points?.[0]?.point?.index ?? this.point.index; const item = top[idx]; const title = item?.title ?? ''; const totalPages = item?.totalPages ?? 0; let html = `<b>${escapeHtml(title)}</b><br/>총 Page#: ${totalPages.toLocaleString()}<br/>`; this.points.forEach(p => { const pct = totalPages ? (p.y/totalPages*100) : 0; html += `<span style="color:${p.color}">\u25CF</span> ${p.series.name}: <b>${p.y.toLocaleString()} Page</b> (${Highcharts.numberFormat(pct,1)}%)<br/>`; }); return html; } },
      plotOptions: { series: { stacking: 'percent', pointPadding: 0, groupPadding: 0.08, dataLabels: { enabled: true, formatter(){ return this.percentage > 15 ? Math.round(this.percentage) + '%' : null; } } }, bar: { borderWidth: 0, pointWidth: barPointWidth, point: { events: { click: function(){ const idx = this.index; const fullTitle = top[idx]?.title ?? this.category; selectedTitle = (selectedTitle === fullTitle) ? null : fullTitle; if (selectedTitle){ selTitleTxt.textContent = selectedTitle; titleBadge.style.display = 'inline-flex'; } else { titleBadge.style.display = 'none'; } renderTitleStatusChart(); renderSummaryForSelectedTitle(); detailStatusFilter = ""; if (detailStatusSel) detailStatusSel.value = ""; renderTable(); const listPanel = document.querySelector('.panel.list:last-of-type'); if (listPanel) listPanel.scrollIntoView({ behavior: 'smooth', block: 'start' }); } } } } },
      legend: { itemStyle:{ color:'#e5e7eb' }, itemHoverStyle:{ color:'#d1d5db' } },
      credits: { enabled:false },
      exporting: { enabled:false },
      series: statuses.map(g => ({ name: g, color: palette[g], data: seriesByGroup[g].map((y, idx) => makePoint(g, idx, y)) }))
    });
  }

  // ===== 요약 =====
  function renderSummaryForSelectedTitle() {
    if (!summaryPanel || !summaryTblBody || !summaryPageTotal) return;
    const base = selectedStatus ? rows.filter(r => r.statusGrp === selectedStatus) : rows;
    if (!selectedTitle) { summaryTblBody.innerHTML = ""; summaryPageTotal.textContent = "0"; return; }
    const list = base.filter(r => r.title === selectedTitle);
    const totalPages = list.reduce((a,b)=>a+(b.pageNum||1),0);
    const statuses = ["완료","법인리뷰","사전검토","진행중","취소"];
    const pageByStatus = Object.fromEntries(statuses.map(s=>[s,0]));
    list.forEach(r => { pageByStatus[r.statusGrp] += (r.pageNum || 1); });
    summaryTblBody.innerHTML = statuses.map(s => {
      const pages = pageByStatus[s] || 0;
      const pct = totalPages ? (pages/totalPages*100) : 0;
      return `<tr class="row"><td>${s}</td><td>${pages.toLocaleString()}</td><td>${Highcharts.numberFormat(pct,1)}%</td></tr>`;
    }).join("");
    summaryPageTotal.textContent = totalPages.toLocaleString();
  }

  // ===== Page# Total =====
  function updateDetailPageTotalBadge(currentList) {
    if (!pageTotalVal) return;
    const pageSum = currentList.reduce((a,b)=>a+(b.pageNum||1),0);
    pageTotalVal.textContent = `${pageSum.toLocaleString()}`;
  }

  // ===== 필터 적용 =====
  function applyFilters() {
    updateLocaleBoxByRegion();
    const rset = getChecked(regionBox);
    const pset = getChecked(projBox);
    const oset = getChecked(ownerGrpBox);
    const mset = getChecked(monthBox);
    rows = allRows.filter((r) => {
      if (rset.size && !rset.has(r.region)) return false;
      if (pset.size && !pset.has(r.locale)) return false;
      if (oset.size && !oset.has(r.ownerGrp)) return false;
      if (mset.size && !(r.month && mset.has(r.month))) return false;
      return true;
    });
    renderAll();
  }

  // ===== 렌더 전체 =====
  function renderAll() {
    const k = computeKPIs(rows);
    const elTotal = document.querySelector("#kpi-total .val");
    if (elTotal) elTotal.textContent = k.total.toLocaleString();
    const elNew = document.querySelector("#kpi-new .val");
    if (elNew) elNew.textContent = k.newCnt.toLocaleString();
    const elInp = document.querySelector("#kpi-inp .val");
    if (elInp) elInp.textContent = k.inProg.toLocaleString();
    renderStatusDonut();
    renderTitleStatusChart();
    renderSummaryForSelectedTitle();
    renderTable();
  }

  // ===== 상세 테이블 =====
  function renderTable() {
    const base = selectedStatus ? rows.filter(r => r.statusGrp === selectedStatus) : rows;
    const list0 = selectedTitle ? base.filter(r => r.title === selectedTitle) : base;
    const list = list0.filter(r => !detailStatusFilter || r.statusGrp === detailStatusFilter);
    updateDetailPageTotalBadge(list);
    tbody.innerHTML = "";
    list.forEach(r => {
      const isCancelled = r.statusGrp === '취소';
      const lastCmt = (r.lastComment || '').trim();
      const cancelReason = (r.cancelReason || '').trim();
      const showBtn = isCancelled || lastCmt.length > 0;
      let btnHtml = '';
      if (showBtn) {
        const tipText = isCancelled ? cancelReason : lastCmt;
        const btnLabel = isCancelled ? '!' : 'i';
        btnHtml = `<button class="tip-btn ${isCancelled ? 'danger' : ''}" data-tip="${escapeHtml(tipText || '메모 없음')}">${btnLabel}</button>`;
      }
      const wpl = getWPLByLocale(r.locale);
      const wplName = wpl?.name || "-";
      const wplEmail = wpl?.email || "";
      const wplNameEsc = escapeHtml(wplName);
      const wplEmailEsc = escapeHtml(wplEmail);
      const emailBtnHtml = wplEmail ? `<button class="tip-btn" data-type="email" data-name="${wplNameEsc}" data-tip="${wplEmailEsc}">@</button>` : "";
      const tr = document.createElement("tr");
      tr.className = "row";
      tr.innerHTML = `
        <td>${escapeHtml(r.region)}</td>
        <td>${escapeHtml(r.locale)}</td>
        <td>${escapeHtml(r.ownerGrp)}</td>
        <td>${r.date ? r.date.toISOString().slice(0, 10) : ""}</td>
        <td><span class="chip">${escapeHtml(r.statusGrp)}${btnHtml}</span></td>
        <td>${escapeHtml(r.pageNo || "-")}</td>
        <td>
          <span class="chip">
            <span class="wpl-name" title="${wplNameEsc}">${wplNameEsc}</span>
            ${emailBtnHtml}
          </span>
        </td>
        <td title="${escapeHtml(r.title)}">${escapeHtml(r.title.length > 60 ? r.title.slice(0, 59) + "…" : r.title)}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // ===== 이벤트 =====
  const excelInput = document.getElementById("excelInput");
  if (excelInput) {
    excelInput.addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      if (!f) return;
      await loadFile(f);
    });
  }
  document.addEventListener("click", (e) => {
    const btn = e.target.closest(".btn-mini");
    if (!btn) return;
    const type = btn.dataset.all;
    const boxMap = { region: regionBox, proj: projBox, ownergrp: ownerGrpBox, month: monthBox };
    const box = boxMap[type];
    if (!box) return;
    box.querySelectorAll('input[type="checkbox"]').forEach((cb) => (cb.checked = false));
    applyFilters();
  });
  clearStatus?.addEventListener("click", () => {
    selectedStatus = null;
    statusBadge.style.display = "none";
    renderStatusDonut();
    renderTitleStatusChart();
    renderSummaryForSelectedTitle();
    renderTable();
  });
  clearTitle?.addEventListener("click", () => {
    selectedTitle = null;
    titleBadge.style.display = "none";
    renderTitleStatusChart();
    renderSummaryForSelectedTitle();
    renderTable();
  });
  detailStatusSel?.addEventListener("change", (e) => {
    detailStatusFilter = e.target.value || "";
    renderSummaryForSelectedTitle();
    renderTable();
  });
  function resetFilters() {
    [regionBox, projBox, ownerGrpBox, monthBox].forEach(box => {
      if (!box) return;
      box.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
    });
    selectedStatus = null;
    selectedTitle  = null;
    detailStatusFilter = "";
    if (detailStatusSel) detailStatusSel.value = "";
    if (statusBadge) statusBadge.style.display = "none";
    if (titleBadge)  titleBadge.style.display  = "none";
    rows = allRows.slice();
    renderAll();
  }
  resetBtn?.addEventListener('click', resetFilters);
  document.addEventListener('click', (e) => {
    const btn = e.target.closest('.tip-btn');
    if (!btn || !commentPop || !commentPopTitle || !commentPopText) return;
    const msg = btn.dataset.tip || '내용 없음';
    const isCancelled = btn.classList.contains('danger');
    const isEmail = btn.dataset.type === 'email';
    const fullName = btn.dataset.name || 'Email';
    commentPopTitle.textContent = isEmail ? fullName : (isCancelled ? 'Cancelled Reason' : 'Last Comment');
    commentPopText.textContent  = msg;
    const rect = btn.getBoundingClientRect();
    const margin = 8, pad = 12;
    const vpW = document.documentElement.clientWidth;
    const vpH = document.documentElement.clientHeight;
    let top  = rect.top - 4;
    let left = rect.left + rect.width + margin;
    commentPop.style.visibility = 'hidden';
    commentPop.style.display = 'block';
    const popW = commentPop.offsetWidth;
    const popH = commentPop.offsetHeight;
    if (left + popW + pad > vpW) left = Math.max(margin, rect.left - popW - margin);
    if (top + popH + pad > vpH)  top  = Math.max(margin, vpH - popH - pad);
    commentPop.style.top  = `${top}px`;
    commentPop.style.left = `${left}px`;
    commentPop.style.visibility = 'visible';
  });
  document.addEventListener('click', (e) => {
    if (!commentPop) return;
    const isTipBtn  = e.target.closest('.tip-btn');
    const insidePop = e.target.closest('#commentPop');
    if (!isTipBtn && !insidePop) commentPop.style.display = 'none';
  });
  window.addEventListener('scroll', () => {
    if (commentPop && commentPop.style.display === 'block') commentPop.style.display = 'none';
  });

  // ===== 초기 =====
  renderFilters([]);
})();
