// result page for v3
function safeParse(s){ try{ return JSON.parse(s); }catch(e){ return null; } }

const metaPill = document.getElementById("metaPill");
const scorePill = document.getElementById("scorePill");
const sepPill = document.getElementById("sepPill");
const carePill = document.getElementById("carePill");
const classSummary = document.getElementById("classSummary");
const violationsDiv = document.getElementById("violations");
const table = document.getElementById("resultTable");
const classFilter = document.getElementById("classFilter");
const tableMeta = document.getElementById("tableMeta");
const downloadBtn = document.getElementById("downloadBtn");
const backBtn = document.getElementById("backBtn");

backBtn.addEventListener("click", ()=> window.close());

let payload = safeParse(sessionStorage.getItem("classAssignResultV3"));
if (!payload) payload = safeParse(localStorage.getItem("classAssignResultV3"));

if (!payload){
  document.body.innerHTML = "<div style='color:#fff;padding:24px;font-family:system-ui;'>결과를 찾을 수 없어요. 먼저 설정 화면에서 반배정을 실행해주세요.</div>";
  throw new Error("No payload");
}

metaPill.textContent = `${payload.meta.total}명 · ${payload.meta.classCount}반 · ${payload.meta.iterations.toLocaleString()}회 · seed ${payload.meta.seed} · ${payload.meta.elapsedMs.toLocaleString()}ms`;
scorePill.textContent = `Score: ${Math.round(payload.best.score).toLocaleString()}`;
sepPill.textContent = `분리 위반: ${payload.best.sepViol.toLocaleString()}쌍`;
carePill.textContent = `배려 미충족: ${payload.best.careMiss.toLocaleString()}쌍`;

function renderClassSummary(){
  const C = payload.meta.classCount;
  const cnt = payload.arrays.cnt;
  const male = payload.arrays.male;
  const female = payload.arrays.female;
  const spec = payload.arrays.spec;
  const adhd = payload.arrays.adhd;

  let html = "<div style='overflow:auto'><table><thead><tr><th>반</th><th>인원</th><th>남</th><th>여</th><th>특수</th><th>ADHD</th></tr></thead><tbody>";
  for (let c=0;c<C;c++){
    html += `<tr><td>${c+1}</td><td>${cnt[c]}</td><td>${male[c]}</td><td>${female[c]}</td><td>${spec[c]}</td><td>${adhd[c]}</td></tr>`;
  }
  html += "</tbody></table></div>";
  classSummary.innerHTML = html;
}

function buildViolationReport(){
  // Recompute per-code counts from resultRows (lightweight)
  const rows = payload.resultRows;
  const sepMap = new Map(); // code -> Map(class -> count)
  const careMap = new Map();

  function add(map, code, cls){
    if (!map.has(code)) map.set(code, new Map());
    const m = map.get(code);
    m.set(cls, (m.get(cls)||0)+1);
  }

  for (const r of rows){
    const cls = r["반"];
    const sepCodes = (r["분리요청코드"]||"").split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
    const careCodes = (r["배려요청코드"]||"").split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
    for (const c of sepCodes) add(sepMap, c, cls);
    for (const c of careCodes) add(careMap, c, cls);
  }

  const worstSep = [];
  for (const [code, m] of sepMap.entries()){
    for (const [cls, k] of m.entries()){
      if (k >= 2){
        worstSep.push({code, cls, k});
      }
    }
  }
  worstSep.sort((a,b)=>b.k-a.k);

  const worstCare = [];
  for (const [code, m] of careMap.entries()){
    let total = 0;
    for (const k of m.values()) total += k;
    if (total < 2) continue;
    const classes = [...m.keys()].sort((a,b)=>a-b);
    if (classes.length >= 2){
      worstCare.push({code, classes: classes.join(","), total});
    }
  }

  let html = "";
  html += `<div class="small"><b>분리 위반(상위 10)</b></div>`;
  if (worstSep.length===0) html += `<div class="small">- 위반 없음</div>`;
  else {
    html += "<div style='overflow:auto;max-height:160px;'><table><thead><tr><th>코드</th><th>반</th><th>동반 인원</th></tr></thead><tbody>";
    for (const x of worstSep.slice(0,10)){
      html += `<tr><td>${x.code}</td><td>${x.cls}</td><td>${x.k}</td></tr>`;
    }
    html += "</tbody></table></div>";
  }

  html += `<div style="height:10px"></div><div class="small"><b>배려 분산(상위 10)</b></div>`;
  if (worstCare.length===0) html += `<div class="small">- 분산 없음(또는 코드 없음)</div>`;
  else {
    html += "<div style='overflow:auto;max-height:160px;'><table><thead><tr><th>코드</th><th>분산된 반</th><th>총 인원</th></tr></thead><tbody>";
    for (const x of worstCare.slice(0,10)){
      html += `<tr><td>${x.code}</td><td>${x.classes}</td><td>${x.total}</td></tr>`;
    }
    html += "</tbody></table></div>";
  }

  violationsDiv.innerHTML = html;
}

function renderTable(filterClass){
  const rows = payload.resultRows.slice().sort((a,b)=>a["반"]-b["반"] || a["학생명"].localeCompare(b["학생명"]));
  const filtered = (filterClass==="all") ? rows : rows.filter(r=>String(r["반"])===String(filterClass));

  table.innerHTML = "";
  if (filtered.length===0) return;

  const cols = Object.keys(filtered[0]);
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
  for (const r of filtered){
    const tr = document.createElement("tr");
    for (const c of cols){
      const td = document.createElement("td");
      td.textContent = (r[c]===null||r[c]===undefined) ? "" : String(r[c]);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  tableMeta.textContent = `표시 ${filtered.length}명`;
}

function setupFilter(){
  classFilter.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "all";
  optAll.textContent = "전체";
  classFilter.appendChild(optAll);

  for (let c=1;c<=payload.meta.classCount;c++){
    const o = document.createElement("option");
    o.value = String(c);
    o.textContent = `${c}반`;
    classFilter.appendChild(o);
  }
  classFilter.addEventListener("change", ()=> renderTable(classFilter.value));
}

function drawCharts(){
  // non-responsive to avoid loops
  Chart.defaults.responsive = false;
  Chart.defaults.animation = false;

  const labels = Array.from({length: payload.meta.classCount}, (_,i)=>`${i+1}반`);
  const cnt = payload.arrays.cnt;
  const male = payload.arrays.male;
  const female = payload.arrays.female;

  const cntCtx = document.getElementById("cntChart");
  const gCtx = document.getElementById("genderChart");

  new Chart(cntCtx, {
    type:"bar",
    data:{ labels, datasets:[{ label:"인원", data: cnt }]},
    options:{ responsive:false, animation:false, plugins:{ legend:{ display:false }}, scales:{ y:{ beginAtZero:true } } }
  });

  new Chart(gCtx, {
    type:"bar",
    data:{ labels, datasets:[
      { label:"남", data: male, stack:"g" },
      { label:"여", data: female, stack:"g" },
    ]},
    options:{ responsive:false, animation:false, plugins:{ legend:{ position:"bottom" }}, scales:{ x:{ stacked:true }, y:{ stacked:true, beginAtZero:true } } }
  });
}

downloadBtn.addEventListener("click", ()=>{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(payload.resultRows);
  XLSX.utils.book_append_sheet(wb, ws, "반배정결과");

  const metaSheet = XLSX.utils.aoa_to_sheet([
    ["항목","값"],
    ["총원", payload.meta.total],
    ["반 수", payload.meta.classCount],
    ["시뮬레이션", payload.meta.iterations],
    ["시드", payload.meta.seed],
    ["경과(ms)", payload.meta.elapsedMs],
    ["학업 가중치", payload.meta.weights.wAcad],
    ["교우 가중치", payload.meta.weights.wPeer],
    ["민원 가중치", payload.meta.weights.wParent],
    ["분리강도", payload.meta.sepStrength],
    ["배려강도", payload.meta.careStrength],
    ["분리 위반(쌍)", payload.best.sepViol],
    ["배려 미충족(쌍)", payload.best.careMiss],
    ["Score", payload.best.score]
  ]);
  XLSX.utils.book_append_sheet(wb, metaSheet, "설정요약");

  XLSX.writeFile(wb, `반배정_결과_${payload.meta.total}명_${payload.meta.classCount}반_seed${payload.meta.seed}.xlsx`);
});

renderClassSummary();
buildViolationReport();
setupFilter();
renderTable("all");
drawCharts();
