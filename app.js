// 반배정 웹앱 (브라우저 전용) v1
// - 엑셀 파일은 서버로 전송하지 않고, FileReader로 브라우저 메모리에서만 읽습니다.
// - 1단계: 미리보기/기본 통계/그래프

const fileInput = document.getElementById("fileInput");
const filePill = document.getElementById("filePill");
const rowsPill = document.getElementById("rowsPill");
const errorsDiv = document.getElementById("errors");
const statsDiv = document.getElementById("stats");
const table = document.getElementById("previewTable");

// Charts
let genderChart, needsChart, acadChart, peerChart, parentChart;

function safeString(x){ return (x === null || x === undefined) ? "" : String(x).trim(); }
function ynTo01(x){
  const v = safeString(x).toUpperCase();
  if (v === "Y" || v === "1" || v === "TRUE") return 1;
  return 0;
}

function countBy(arr){
  const m = new Map();
  for (const v of arr){
    const k = safeString(v) || "(비어있음)";
    m.set(k, (m.get(k) || 0) + 1);
  }
  return m;
}

function renderTable(rows, maxRows=20){
  table.innerHTML = "";
  if (!rows || rows.length === 0) return;

  const cols = Object.keys(rows[0]);
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  for (const c of cols){
    const th = document.createElement("th");
    th.textContent = c;
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const r of rows.slice(0, maxRows)){
    const tr = document.createElement("tr");
    for (const c of cols){
      const td = document.createElement("td");
      td.textContent = safeString(r[c]);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
}

function setErrors(msg){
  errorsDiv.textContent = msg || "";
}

function renderStats(rows){
  const n = rows.length;
  const names = rows.map(r => safeString(r["학생명"] || r["이름"] || r["성명"]));
  const gender = rows.map(r => safeString(r["성별"] || r["남녀"]));
  const acad = rows.map(r => safeString(r["학업성취"] || r["학업성취(3단계)"]));
  const peer = rows.map(r => safeString(r["교우관계"] || r["교우관계(3단계)"]));
  const parent = rows.map(r => safeString(r["학부모민원"] || r["학부모민원(3단계)"]));
  const special = rows.map(r => ynTo01(r["특수여부"]));
  const adhd = rows.map(r => ynTo01(r["ADHD여부"]));

  const male = gender.filter(g => g === "남").length;
  const female = gender.filter(g => g === "여").length;

  const specN = special.reduce((a,b)=>a+b,0);
  const adhdN = adhd.reduce((a,b)=>a+b,0);

  const uniqueNames = new Set(names.filter(Boolean)).size;

  statsDiv.innerHTML = `
    <div class="row" style="gap:8px;">
      <span class="pill ok">총 ${n}명</span>
      <span class="pill">이름(유니크) ${uniqueNames}</span>
      <span class="pill">남 ${male} · 여 ${female}</span>
      <span class="pill">특수 ${specN}</span>
      <span class="pill">ADHD ${adhdN}</span>
    </div>
    <div style="height:10px"></div>
    <div class="small">※ 다음 단계에서 “반배정 후 반별 분포/편차” 그래프를 추가합니다.</div>
  `;

  // Pills
  rowsPill.textContent = `${n}명`;

  // Charts
  const genderMap = countBy(gender);
  drawPie("genderChart", ["남","여","기타"], [genderMap.get("남")||0, genderMap.get("여")||0, n - (genderMap.get("남")||0) - (genderMap.get("여")||0)], (c)=>{ genderChart = c; }, ()=>genderChart);

  drawBar("needsChart", ["특수","ADHD"], [specN, adhdN], (c)=>{ needsChart = c; }, ()=>needsChart);

  drawBar("acadChart", ["좋음","보통","나쁨","기타"], dist3(acad), (c)=>{ acadChart = c; }, ()=>acadChart);
  drawBar("peerChart", ["좋음","보통","나쁨","기타"], dist3(peer), (c)=>{ peerChart = c; }, ()=>peerChart);
  drawBar("parentChart", ["좋음","보통","나쁨","기타"], dist3(parent), (c)=>{ parentChart = c; }, ()=>parentChart);
}

function dist3(arr){
  let good=0, mid=0, bad=0, etc=0;
  for (const v of arr){
    if (v === "좋음") good++;
    else if (v === "보통") mid++;
    else if (v === "나쁨") bad++;
    else if (v) etc++;
  }
  return [good, mid, bad, etc];
}

function drawPie(canvasId, labels, data, setRef, getRef){
  const ctx = document.getElementById(canvasId);
  const existing = getRef();
  if (existing) existing.destroy();
  const chart = new Chart(ctx, {
    type: "doughnut",
    data: { labels, datasets: [{ data }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } }
  });
  setRef(chart);
}

function drawBar(canvasId, labels, data, setRef, getRef){
  const ctx = document.getElementById(canvasId);
  const existing = getRef();
  if (existing) existing.destroy();
  const chart = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets: [{ data }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });
  setRef(chart);
}

fileInput.addEventListener("change", async (e) => {
  setErrors("");
  const file = e.target.files?.[0];
  if (!file){
    filePill.textContent = "엑셀 미선택";
    return;
  }
  filePill.textContent = `선택됨: ${file.name}`;

  try{
    const data = await file.arrayBuffer(); // stays local
    const wb = XLSX.read(data, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

    if (!json || json.length === 0){
      setErrors("엑셀에서 데이터를 읽지 못했어요. 첫 시트에 데이터가 있는지 확인해주세요.");
      return;
    }

    // Render
    renderTable(json, 20);
    renderStats(json);

  } catch(err){
    console.error(err);
    setErrors("엑셀을 읽는 중 오류가 발생했습니다: " + (err?.message || String(err)));
  }
});
