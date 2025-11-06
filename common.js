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
    const emailBtnHtml = wplEmail
      ? `<button class="tip-btn" data-type="email" data-name="${wplNameEsc}" data-tip="${wplEmailEsc}">@</button>`
      : "";

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
      <td title="${escapeHtml(r.title)}">
        ${escapeHtml(r.title.length > 60 ? r.title.slice(0, 59) + "…" : r.title)}
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// ===== Tip 버튼 (이름을 팝업 제목에 표시) =====
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
