// 公聽會倒數計時 sticky bar
// 五場公聽會時程：5/20 南、5/22 北、5/23 兒少、5/28 中、6/4 東
// 過期場次標灰、最近未來場次標亮並倒數
// 全部結束後切換為「公聽會已結束、進入立法院審議」

(function () {
  const HEARINGS = [
    { id: 'south', date: '2026-05-20', label: '5/20 南', name: '南區（高雄）' },
    { id: 'north', date: '2026-05-22', label: '5/22 北', name: '北區（台北）' },
    { id: 'kids',  date: '2026-05-23', label: '5/23 兒少', name: '兒少場（視訊）' },
    { id: 'central', date: '2026-05-28', label: '5/28 中', name: '中區（台中）' },
    { id: 'east',  date: '2026-06-04', label: '6/4 東', name: '東區（花蓮）' },
  ];

  function renderBar() {
    const bar = document.querySelector('.active-campaign-bar');
    if (!bar) return;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let nextHearing = null;
    const sessionsHTML = HEARINGS.map(h => {
      const hDate = new Date(h.date + 'T00:00:00');
      const past = hDate < today;
      const isNext = !past && !nextHearing;
      if (isNext) nextHearing = { ...h, date: hDate };
      const cls = past ? 'past' : (isNext ? 'next' : '');
      return `<span class="session ${cls}" title="${h.name}">${h.label}</span>`;
    }).join('');

    let countdownHTML = '';
    if (nextHearing) {
      const diffMs = nextHearing.date - today;
      const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
      if (diffDays === 0) {
        countdownHTML = `<span class="campaign-countdown">今日 ${nextHearing.name}</span>`;
      } else {
        countdownHTML = `<span class="campaign-countdown">距離 ${nextHearing.name} ${diffDays} 天</span>`;
      }
    } else {
      // 全部結束
      bar.innerHTML = `
        <span class="campaign-icon">📋</span>
        <span class="campaign-label">公聽會已結束</span>
        <span class="campaign-hearings">下一步：立法院社會福利及衛生環境委員會審議</span>
        <a href="index.html" class="campaign-countdown" style="text-decoration:none;color:white;">查看本盟主張</a>
      `;
      return;
    }

    bar.innerHTML = `
      <span class="campaign-icon">🔴</span>
      <span class="campaign-label">進行中倡議：兒少權法 165 條修法</span>
      <span class="campaign-hearings">${sessionsHTML}</span>
      ${countdownHTML}
    `;
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', renderBar);
  } else {
    renderBar();
  }
})();
