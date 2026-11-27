const worker = new Worker("./worker.js");
let reqCounter = 0;
const pending = new Map();
let workbookData = null;

function callWorker(type, payload = {}) {
  return new Promise((resolve, reject) => {
    const id = ++reqCounter;
    pending.set(id, { resolve, reject });
    worker.postMessage({ id, type, payload });
  });
}

worker.onmessage = (event) => {
  const { id, ok, data, error } = event.data;
  const entry = pending.get(id);
  if (!entry) return;
  pending.delete(id);
  if (ok) entry.resolve(data);
  else entry.reject(new Error(error));
};

const statusText = document.getElementById("statusText");
const unitText = document.getElementById("unitText");
const sectorCountText = document.getElementById("sectorCountText");
const classificationSelect = document.getElementById("classificationSelect");
const tradeScopeSelect = document.getElementById("tradeScopeSelect");
const sheetSelect = document.getElementById("sheetSelect");
const sectorSearch = document.getElementById("sectorSearch");
const sectorList = document.getElementById("sectorList");
const targetSectorSelect = document.getElementById("targetSectorSelect");
const demandShockInput = document.getElementById("demandShockInput");
const shockLabelInput = document.getElementById("shockLabelInput");
const runAnalysisBtn = document.getElementById("runAnalysisBtn");
const kpiGrid = document.getElementById("kpiGrid");
const insightBox = document.getElementById("insightBox");
const resultTableWrap = document.getElementById("resultTableWrap");
const compareTableWrap = document.getElementById("compareTableWrap");

function getFilteredSheets() {
  if (!workbookData) return [];
  const clsFilter = classificationSelect.value;
  const scopeFilter = tradeScopeSelect.value;
  let list = workbookData.sheets.filter((s) => s.sectorCount > 0);
  const byScope = list.filter((s) => s.tradeScope === scopeFilter);
  if (byScope.length) list = byScope;
  if (clsFilter !== "ALL") {
    const byCls = list.filter((s) => s.classification === clsFilter);
    if (byCls.length) list = byCls;
  }
  return list;
}

function renderSectors(sectors, keyword = "") {
  const q = keyword.trim().toLowerCase();
  const filtered = q ? sectors.filter((s) => s.toLowerCase().includes(q)) : sectors;
  sectorCountText.textContent = `${filtered.length} / ${sectors.length}개`;
  sectorList.innerHTML = filtered
    .slice(0, 150)
    .map((s) => `<span class="sector-chip" data-sector="${s}" title="클릭 시 충격 대상 부문으로 선택">${s}</span>`)
    .join("");

  sectorList.querySelectorAll(".sector-chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      targetSectorSelect.value = chip.dataset.sector;
    });
  });
}

function populateSheets() {
  const sheets = getFilteredSheets();
  sheetSelect.disabled = sheets.length === 0;
  sheetSelect.innerHTML = sheets
    .map((s) => `<option value="${s.sheetName}">${s.sheetName}</option>`)
    .join("");
  onSheetChange();
}

function onSheetChange() {
  const sheets = getFilteredSheets();
  const chosen = sheets.find((s) => s.sheetName === sheetSelect.value) || sheets[0];
  if (!chosen) return;

  const sectors = chosen.sectors || [];
  renderSectors(sectors, sectorSearch.value);

  targetSectorSelect.disabled = sectors.length === 0;
  targetSectorSelect.innerHTML = sectors.map((s) => `<option value="${s}">${s}</option>`).join("");

  const preferred = sectors.find((s) => s.includes("집적회로") || s.includes("반도체") || s.includes("제조업"));
  if (preferred) targetSectorSelect.value = preferred;

  runAnalysisBtn.disabled = sectors.length === 0;
}

async function loadExcelBuffer(buffer) {
  statusText.textContent = "실측 2023년 투입산출표(8개 시트 · 최대 380×380 행렬) 연산 중...";
  const data = await callWorker("LOAD_EXCEL", { buffer });
  workbookData = data;
  unitText.textContent = data.unit;
  statusText.textContent = `로딩 완료 (총 ${data.sheets.length}개 시트 · 실측 2023년 데이터)`;
  populateSheets();
}

document.getElementById("defaultBtn").addEventListener("click", async () => {
  try {
    statusText.textContent = "data/io-2023.xlsx 불러오는 중...";
    const resp = await fetch("./data/io-2023.xlsx");
    if (!resp.ok) throw new Error("data/io-2023.xlsx 파일을 찾을 수 없습니다.");
    await loadExcelBuffer(await resp.arrayBuffer());
  } catch (err) {
    statusText.textContent = `오류: ${err.message}`;
  }
});

document.getElementById("fileInput").addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    await loadExcelBuffer(await file.arrayBuffer());
  } catch (err) {
    statusText.textContent = `오류: ${err.message}`;
  }
});

classificationSelect.addEventListener("change", () => workbookData && populateSheets());
tradeScopeSelect.addEventListener("change", () => workbookData && populateSheets());
sheetSelect.addEventListener("change", onSheetChange);

sectorSearch.addEventListener("input", () => {
  const chosen = getFilteredSheets().find((s) => s.sheetName === sheetSelect.value);
  if (chosen) renderSectors(chosen.sectors || [], sectorSearch.value);
});
document.getElementById("clearSearchBtn").addEventListener("click", () => {
  sectorSearch.value = "";
  onSheetChange();
});

document.querySelectorAll(".preset-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    if (!workbookData) return;
    const preset = btn.dataset.preset;
    const chosen = getFilteredSheets().find((s) => s.sheetName === sheetSelect.value);
    const secs = chosen?.sectors || [];

    if (preset === "ic_export") {
      demandShockInput.value = 10000000;
      shockLabelInput.value = "집적회로·반도체 수출 10조원 증가";
      const m = secs.find((s) => s.includes("집적회로") || s.includes("반도체") || s.includes("제조업"));
      if (m) targetSectorSelect.value = m;
    } else if (preset === "semi_equip") {
      demandShockInput.value = 3000000;
      shockLabelInput.value = "반도체 제조용 기계 설비투자 3조원 확대";
      const m = secs.find((s) => s.includes("반도체 제조용") || s.includes("반도체·디스플레이") || s.includes("일반기계") || s.includes("제조업"));
      if (m) targetSectorSelect.value = m;
    } else if (preset === "power_grid") {
      demandShockInput.value = 2000000;
      shockLabelInput.value = "전력공급·변압기·배전 인프라 2조원 확충";
      const m = secs.find((s) => s.includes("전력공급") || s.includes("화력") || s.includes("전력·가스"));
      if (m) targetSectorSelect.value = m;
    } else if (preset === "sw_ai") {
      demandShockInput.value = 2000000;
      shockLabelInput.value = "소프트웨어·IT서비스 디지털 전환 2조원 투자";
      const m = secs.find((s) => s.includes("소프트웨어") || s.includes("정보통신") || s.includes("서비스업"));
      if (m) targetSectorSelect.value = m;
    }
    runAnalysisBtn.click();
  });
});

runAnalysisBtn.addEventListener("click", async () => {
  try {
    statusText.textContent = "산업연관 파급효과 계산 중...";
    const res = await callWorker("RUN_ANALYSIS", {
      sheetName: sheetSelect.value,
      targetSector: targetSectorSelect.value,
      demandShock: demandShockInput.value,
      shockLabel: shockLabelInput.value,
    });
    statusText.textContent = "분석 완료";

    const shockTrillion = (res.demandShock / 1000000).toFixed(2);
    const outTrillion = (res.totalOutputEffect / 1000000).toFixed(2);
    const vaTrillion = (res.totalVaEffect / 1000000).toFixed(2);

    kpiGrid.innerHTML = `
      <div class="kpi-card">
        <div class="kpi-title">총 생산유발액 (${res.tradeScope})</div>
        <div class="kpi-value">${res.totalOutputEffect.toLocaleString()} 백만원</div>
        <div class="kpi-sub">약 ${outTrillion}조 원 · 생산유발계수 ${res.targetOutputMultiplier}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">총 부가가치유발액 (${res.tradeScope})</div>
        <div class="kpi-value">${res.totalVaEffect.toLocaleString()} 백만원</div>
        <div class="kpi-sub">약 ${vaTrillion}조 원 · 부가가치유발계수 ${res.targetVaMultiplier}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">영향력계수(후방) / 감응도계수(전방)</div>
        <div class="kpi-value">${res.targetPowerIndex} / ${res.targetSensitivityIndex}</div>
        <div class="kpi-sub">전 산업 평균 1.0000 대비 상대 지수</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">산업조직론적 연쇄 유형</div>
        <div class="kpi-value" style="font-size:0.98rem;">${res.targetLinkageType}</div>
        <div class="kpi-sub">${res.classification} 기준</div>
      </div>
    `;

    insightBox.style.display = "block";
    insightBox.innerHTML = `
      <strong>2023년 실측 분석 요약 (${res.shockLabel || "사용자 시나리오"}):</strong>
      [${res.classification} · ${res.tradeScope}] 기준 <strong>${res.targetSector}</strong> 부문에
      <strong>${res.demandShock.toLocaleString()} 백만원(약 ${shockTrillion}조 원)</strong>의 최종수요 충격이 발생할 경우,
      전 산업에 총 <strong>${res.totalOutputEffect.toLocaleString()} 백만원(약 ${outTrillion}조 원)</strong>의 생산유발과
      <strong>${res.totalVaEffect.toLocaleString()} 백만원(약 ${vaTrillion}조 원)</strong>의 부가가치유발이 나타납니다.
      국산거래표와 총거래표를 전환해 비교하면 중간재 수입으로 인한 해외 누출 효과까지 확인할 수 있습니다.
    `;

    resultTableWrap.innerHTML = `
      <table>
        <thead>
          <tr>
            <th>순위</th>
            <th>부문코드</th>
            <th>산업 부문명</th>
            <th>생산유발액 (백만원)</th>
            <th>부가가치유발액 (백만원)</th>
            <th>생산유발계수</th>
            <th>부가가치유발계수</th>
            <th>영향력계수(후방)</th>
            <th>감응도계수(전방)</th>
          </tr>
        </thead>
        <tbody>
          ${res.topRows.map((r, i) => `
            <tr>
              <td>${i + 1}</td>
              <td>${r.code}</td>
              <td><strong>${r.name}</strong></td>
              <td>${r.deltaOutput.toLocaleString()}</td>
              <td>${r.deltaVa.toLocaleString()}</td>
              <td>${r.outputMultiplier}</td>
              <td>${r.vaMultiplier}</td>
              <td>${r.powerIndex}</td>
              <td>${r.sensitivityIndex}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;

    compareTableWrap.innerHTML = `
      <table>
        <thead>
          <tr>
            <th>분류 체계</th>
            <th>시트명 (부문 수)</th>
            <th>대응 충격 부문</th>
            <th>총 생산유발액 (백만원)</th>
            <th>총 부가가치유발액 (백만원)</th>
            <th>상위 3개 부문 집중도 CR3 (%)</th>
          </tr>
        </thead>
        <tbody>
          ${res.crossLevel.map((c) => `
            <tr>
              <td><strong>${c.classification}</strong></td>
              <td>${c.sheetName}</td>
              <td>${c.matchedSector}</td>
              <td>${c.totalOutputEffect.toLocaleString()}</td>
              <td>${c.totalVaEffect.toLocaleString()}</td>
              <td>${c.cr3Pct}%</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  } catch (err) {
    statusText.textContent = `분석 오류: ${err.message}`;
  }
});

(async () => {
  try {
    await callWorker("PING");
    statusText.textContent = "Pyodide 준비 완료 — '기본 Excel 로딩 (data/io-2023.xlsx)' 버튼을 클릭하세요.";
  } catch (err) {
    statusText.textContent = `초기화 실패: ${err.message}`;
  }
})();
