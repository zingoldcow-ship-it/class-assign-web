console.log('class-assign webapp v3.2.1 loaded');

const REQUIRED_COLUMNS = ['학생명','성별','학업성취','교우관계','학부모민원','특수여부','ADHD여부','분리요청코드','배려요청코드','비고'];

// ===== v3.2: Header normalization (공백/유사명 자동 인식) =====
function normHeader(h){
  return String(h ?? "")
    .replace(/\u00A0/g, " ")         // NBSP
    .replace(/[ \t\r\n]+/g, "")      // 모든 공백 제거
    .trim();
}

// 입력 헤더(정규화) -> 표준 헤더로 매핑
const HEADER_ALIASES = {
  "학생명": ["학생명","이름","성명","학생이름","학생명칭"],
  "성별": ["성별","성","남녀","성구분","성별(남/여)"],
  "학업성취": ["학업성취","학업성취도","성취","학업","성취도","학업수준","학업성취(3단계)"],
  "교우관계": ["교우관계","교우","대인관계","친구관계","교우관계(3단계)"],
  "학부모민원": ["학부모민원","민원","학부모","학부모요청","학부모민원(3단계)"],
  "특수여부": ["특수여부","특수","특수학급","특수대상","특수교육대상","특수유무"],
  "ADHD여부": ["ADHD여부","ADHD","adhd여부","주의력결핍","주의력","ADHD유무"],
  "분리요청코드": ["분리요청코드","분리코드","분리요청","분리","분리요청 코드"],
  "배려요청코드": ["배려요청코드","배려코드","배려요청","배려","배려요청 코드"],
  "비고": ["비고","특이사항","메모","참고","기타","비고(특이사항)"]
};

// 정규화된 입력 헤더 -> 표준 헤더 반환(없으면 null)
function mapToStandardHeader(inputHeader){
  const nh = normHeader(inputHeader);
  if (!nh) return null;
  for (const [std, aliases] of Object.entries(HEADER_ALIASES)){
    for (const a of aliases){
      if (nh === normHeader(a)) return std;
    }
  }
  // 이미 표준과 동일한 경우
  for (const std of Object.keys(HEADER_ALIASES)){
    if (nh === normHeader(std)) return std;
  }
  return null;
}

// 입력 헤더 리스트를 받아서: {stdHeader: originalHeader} 형태로 돌려줌
function buildHeaderMap(headers){
  const map = {};
  headers.forEach(h=>{
    const std = mapToStandardHeader(h);
    if (std && !map[std]) map[std] = h; // 첫 매칭 우선
  });
  return map;
}

// 반배정 웹앱 v3 (브라우저 전용)
// - 엑셀 로컬 파싱
// - 가중치/분리/배려 강도 반영 시뮬레이션
// - 결과를 새 탭(result.html)에서 보기 (sessionStorage 사용)

const fileInput = document.getElementById("fileInput");
const filePill = document.getElementById("filePill");
const rowsPill = document.getElementById("rowsPill");
const errorsDiv = document.getElementById("errors");
const statsDiv = document.getElementById("stats");
const table = document.getElementById("previewTable");

const classCountEl = document.getElementById("classCount");
const iterationsEl = document.getElementById("iterations");
const seedEl = document.getElementById("seed");
const wAcad = document.getElementById("wAcad");
const wPeer = document.getElementById("wPeer");
const wParent = document.getElementById("wParent");
const wAcadV = document.getElementById("wAcadV");
const wPeerV = document.getElementById("wPeerV");
const wParentV = document.getElementById("wParentV");
const sepStrengthEl = document.getElementById("sepStrength");
const careStrengthEl = document.getElementById("careStrength");
const runBtn = document.getElementById("runBtn");
const overlay = document.getElementById("overlay");

// ===== Tab UI =====
const tabSetupBtn = document.getElementById("tabSetup");
const tabResultBtn = document.getElementById("tabResult");
const setupTab = document.getElementById("setupTab");
const resultTab = document.getElementById("resultTab");
const statusPill = document.getElementById("statusPill");

function showTab(which){
  const isSetup = (which==="setup");
  setupTab.style.display = isSetup ? "" : "none";
  resultTab.style.display = isSetup ? "none" : "";
  tabSetupBtn.classList.toggle("active", isSetup);
  tabResultBtn.classList.toggle("active", !isSetup);
}
tabSetupBtn.addEventListener("click", ()=> showTab("setup"));
tabResultBtn.addEventListener("click", ()=> showTab("result"));

const progressTxt = document.getElementById("progressTxt");

let rawRows = null;
let studentRows = null;

function syncWeights(){
  wAcadV.textContent = wAcad.value;
  wPeerV.textContent = wPeer.value;
  wParentV.textContent = wParent.value;
}
[wAcad,wPeer,wParent].forEach(el=>el.addEventListener("input",syncWeights));
syncWeights();

function showOverlay(on, msg){
  overlay.style.display = on ? "flex" : "none";
  if (msg) progressTxt.textContent = msg;
}

function safeString(x){ return (x===null||x===undefined) ? "" : String(x).trim(); }
function ynTo01(x){
  const v = safeString(x).toUpperCase();
  if (v==="Y"||v==="1"||v==="TRUE") return 1;
  return 0;
}
function level3ToScore(x){
  const v = safeString(x);
  if (v==="좋음") return 1;
  if (v==="보통") return 0;
  if (v==="나쁨") return -1;
  return 0;
}
function splitCodes(x){
  const s = safeString(x);
  if (!s) return [];
  return s.split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
}

function renderTable(rows, maxRows=20){
  table.innerHTML = "";
  if (!rows || rows.length===0) return;
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
  for (const r of rows.slice(0,maxRows)){
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

function setErrors(msg){ errorsDiv.textContent = msg || ""; }

function summarize(rows){
  const n = rows.length;
  const gender = rows.map(r=>r.gender);
  const male = gender.filter(g=>g==="남").length;
  const female = gender.filter(g=>g==="여").length;
  const specN = rows.reduce((a,r)=>a+r.special,0);
  const adhdN = rows.reduce((a,r)=>a+r.adhd,0);
  statsDiv.innerHTML = `
    <div style="display:flex; gap:8px; flex-wrap:wrap;">
      <span class="pill ok">총 ${n}명</span>
      <span class="pill">남 ${male} · 여 ${female}</span>
      <span class="pill">특수 ${specN}</span>
      <span class="pill">ADHD ${adhdN}</span>
      <span class="pill">분리코드 ${rows.reduce((a,r)=>a+(r.sepCodes.length>0),0)}명</span>
      <span class="pill">배려코드 ${rows.reduce((a,r)=>a+(r.careCodes.length>0),0)}명</span>
    </div>
    <div style="height:10px"></div>
    <div class="small">* 실행을 누르면 결과는 새 탭에서 표시됩니다.</div>
  `;
}

function normalizeRow(r){
  // Flexible headers (Korean variants)
  const name = safeString(r["학생명"]||r["이름"]||r["성명"]);
  const gender = safeString(r["성별"]||r["남녀"]);
  const acad = safeString(r["학업성취"]||r["학업성취(3단계)"]);
  const peer = safeString(r["교우관계"]||r["교우관계(3단계)"]);
  const parent = safeString(r["학부모민원"]||r["학부모민원(3단계)"]);
  const special = ynTo01(r["특수여부"]||r["특수"]||r["특수여부(Y/N)"]);
  const adhd = ynTo01(r["ADHD여부"]||r["adhd여부"]||r["ADHD"]||r["ADHD여부(Y/N)"]);
  const note = safeString(r["비고"]||r["특이사항"]||r["메모"]);
  const sepCodes = splitCodes(r["분리요청코드"]||r["분리코드"]||r["분리"]);
  const careCodes = splitCodes(r["배려요청코드"]||r["배려코드"]||r["배려"]);
  return {
    name, gender,
    acad, peer, parent,
    acadS: level3ToScore(acad),
    peerS: level3ToScore(peer),
    parentS: level3ToScore(parent),
    special, adhd,
    note,
    sepCodes, careCodes
  };
}

// Simple seeded RNG
function mulberry32(seed){
  let t = seed >>> 0;
  return function(){
    t += 0x6D2B79F5;
    let x = Math.imul(t ^ (t >>> 15), 1 | t);
    x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
    return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
  }
}
function randInt(rng, n){ return Math.floor(rng()*n); }

function comb2(k){ return k>1 ? (k*(k-1))/2 : 0; }

function buildCodeGroups(rows){
  const sep = new Map(); // code -> indices
  const care = new Map();
  rows.forEach((r, idx)=>{
    for (const c of r.sepCodes){
      if (!sep.has(c)) sep.set(c, []);
      sep.get(c).push(idx);
    }
    for (const c of r.careCodes){
      if (!care.has(c)) care.set(c, []);
      care.get(c).push(idx);
    }
  });
  // Filter singletons (no effect)
  for (const [k,v] of [...sep.entries()]) if (v.length<2) sep.delete(k);
  for (const [k,v] of [...care.entries()]) if (v.length<2) care.delete(k);
  return {sep, care};
}

function scoreAssignment(rows, assign, classCount, weights, groups){
  // assign[i] in [0..classCount-1]
  const C = classCount;
  const cnt = new Array(C).fill(0);
  const male = new Array(C).fill(0);
  const female = new Array(C).fill(0);
  const spec = new Array(C).fill(0);
  const adhd = new Array(C).fill(0);
  const acadSum = new Array(C).fill(0);
  const peerSum = new Array(C).fill(0);
  const parentSum = new Array(C).fill(0);

  for (let i=0;i<rows.length;i++){
    const c = assign[i];
    cnt[c] += 1;
    if (rows[i].gender==="남") male[c] += 1;
    else if (rows[i].gender==="여") female[c] += 1;
    spec[c] += rows[i].special;
    adhd[c] += rows[i].adhd;
    acadSum[c] += rows[i].acadS;
    peerSum[c] += rows[i].peerS;
    parentSum[c] += rows[i].parentS;
  }

  // helper variance
  function variance(arr){
    const m = arr.reduce((a,b)=>a+b,0)/arr.length;
    let v=0;
    for (const x of arr){ const d=x-m; v += d*d; }
    return v/arr.length;
  }

  // base balance terms
  const vCnt = variance(cnt);
  const vMale = variance(male);
  const vFemale = variance(female);
  const vSpec = variance(spec);
  const vAdhd = variance(adhd);

  // 3-level sums: variance of sums
  const vAcad = variance(acadSum);
  const vPeer = variance(peerSum);
  const vParent = variance(parentSum);

  let score =
    80*vCnt +
    30*(vMale+vFemale) +
    120*vSpec +
    90*vAdhd +
    weights.wAcad*vAcad +
    weights.wPeer*vPeer +
    weights.wParent*vParent;

  // Separation penalties (violations)
  let sepViol = 0;
  for (const [code, idxs] of groups.sep.entries()){
    const perClass = new Map();
    for (const i of idxs){
      const c = assign[i];
      perClass.set(c, (perClass.get(c)||0)+1);
    }
    for (const k of perClass.values()){
      sepViol += comb2(k);
    }
  }
  score += weights.sepPenalty * sepViol;

  // Care penalties (how much split)
  let careMiss = 0;
  for (const [code, idxs] of groups.care.entries()){
    const totalPairs = comb2(idxs.length);
    if (totalPairs === 0) continue;
    const perClass = new Map();
    for (const i of idxs){
      const c = assign[i];
      perClass.set(c, (perClass.get(c)||0)+1);
    }
    let within = 0;
    for (const k of perClass.values()) within += comb2(k);
    careMiss += (totalPairs - within); // pairs not together
  }
  score += weights.carePenalty * careMiss;

  return {score, sepViol, careMiss, cnt, male, female, spec, adhd, acadSum, peerSum, parentSum};
}

function initialAssignment(rows, classCount, rng){
  // simple round-robin with gender mixing
  const idxs = rows.map((_,i)=>i);
  // shuffle
  for (let i=idxs.length-1;i>0;i--){
    const j = randInt(rng, i+1);
    [idxs[i], idxs[j]] = [idxs[j], idxs[i]];
  }
  const assign = new Array(rows.length).fill(0);
  let c = 0;
  for (const i of idxs){
    assign[i] = c;
    c = (c+1) % classCount;
  }
  return assign;
}

async function optimize(rows, classCount, iterations, seed, weights, groups){
  const rng = mulberry32(seed);
  let assign = initialAssignment(rows, classCount, rng);
  let best = scoreAssignment(rows, assign, classCount, weights, groups);
  let bestAssign = assign.slice();

  // Local search: random swaps
  const n = rows.length;
  const reportEvery = Math.max(200, Math.floor(iterations/30));

  for (let t=1; t<=iterations; t++){
    const i = randInt(rng, n);
    let j = randInt(rng, n);
    if (j===i) j = (j+1)%n;

    const ci = assign[i], cj = assign[j];
    if (ci===cj) continue;

    // swap
    assign[i]=cj; assign[j]=ci;

    const s = scoreAssignment(rows, assign, classCount, weights, groups);

    // accept if better (greedy). light escape: accept slightly worse rarely
    let accept = false;
    if (s.score <= best.score){
      accept = true;
    } else {
      // small probability with temperature
      const temp = Math.max(0.02, 1 - t/iterations);
      const prob = Math.exp(-(s.score - best.score) / (5000*temp));
      if (rng() < prob) accept = true;
    }

    if (accept){
      if (s.score < best.score){
        best = s;
        bestAssign = assign.slice();
      }
    } else {
      // revert
      assign[i]=ci; assign[j]=cj;
    }

    if (t % reportEvery === 0){
      progressTxt.textContent = `시뮬레이션 ${t.toLocaleString()} / ${iterations.toLocaleString()} (현재 best score: ${Math.round(best.score).toLocaleString()})`;
      // yield to UI
      await new Promise(r=>setTimeout(r, 0));
    }
  }

  return {best, bestAssign};
}

function strengthToPenalty(strength, kind){
  // kind: 'sep' or 'care'
  if (kind==='sep'){
    if (strength==='strict') return 500000; // near-hard
    if (strength==='strong') return 120000;
    if (strength==='medium') return 50000;
    return 15000;
  } else {
    if (strength==='strong') return 5000;
    if (strength==='medium') return 2000;
    return 700;
  }
}

function validateRows(rows){
  const missing = [];
  // minimal required
  const nameOk = rows.some(r=>r.name);
  const genderOk = rows.some(r=>r.gender);
  if (!nameOk) missing.push("학생명");
  if (!genderOk) missing.push("성별");
  if (missing.length) return `엑셀에 필요한 열(또는 값)이 부족합니다: ${missing.join(", ")}`;
  return null;
}

fileInput.addEventListener("change", async (e)=>{
  setErrors("");
  const file = e.target.files?.[0];
  if (!file){
    filePill.textContent = "엑셀 미선택";
    runBtn.disabled = true;
    return;
  }
  filePill.textContent = `선택됨: ${file.name}`;

  showOverlay(true, "엑셀을 읽는 중…");
  await new Promise(r=>setTimeout(r, 10));

  try{
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    rawRows = XLSX.utils.sheet_to_json(ws, {defval:""});
    if (!rawRows || rawRows.length===0){
      setErrors("엑셀에서 데이터를 읽지 못했어요. 첫 시트에 데이터가 있는지 확인해주세요.");
      runBtn.disabled = true;
      return;
    }
    studentRows = rawRows.map(normalizeRow);
    const err = validateRows(studentRows);
    if (err){
      setErrors(err);
      runBtn.disabled = true;
      return;
    }
    rowsPill.textContent = `${studentRows.length}명`;
    renderTable(rawRows, 20);
    summarize(studentRows);
    statusPill.textContent = "엑셀 로드됨";
    runBtn.disabled = !(studentRows && studentRows.length>0);
  } catch(err){
    console.error(err);
    setErrors("엑셀을 읽는 중 오류: " + (err?.message || String(err)));
    runBtn.disabled = true;
  } finally{
    showOverlay(false);
  }
});

runBtn.addEventListener("click", async ()=>{
  if (!studentRows || studentRows.length===0) return;

  const classCount = Math.max(2, Math.min(30, parseInt(classCountEl.value||"10",10)));
  const iterations = Math.max(200, Math.min(60000, parseInt(iterationsEl.value||"8000",10)));
  const seed = Math.max(0, Math.min(999999, parseInt(seedEl.value||"42",10)));

  const weights = {
    wAcad: parseInt(wAcad.value,10),
    wPeer: parseInt(wPeer.value,10),
    wParent: parseInt(wParent.value,10),
    sepPenalty: strengthToPenalty(sepStrengthEl.value, 'sep'),
    carePenalty: strengthToPenalty(careStrengthEl.value, 'care')
  };

  showOverlay(true, "코드 그룹(분리/배려)을 구성하는 중…");
  await new Promise(r=>setTimeout(r, 10));
  const groups = buildCodeGroups(studentRows);

  showOverlay(true, "시뮬레이션을 시작합니다…");
  const start = performance.now();
  const {best, bestAssign} = await optimize(studentRows, classCount, iterations, seed, weights, groups);
  const elapsedMs = Math.round(performance.now() - start);

  // Build result rows with class number (1..)
  const resultRows = studentRows.map((r,i)=>({
    반: bestAssign[i] + 1,
    학생명: r.name,
    성별: r.gender,
    학업성취: r.acad,
    교우관계: r.peer,
    학부모민원: r.parent,
    특수여부: r.special ? "Y" : "N",
    ADHD여부: r.adhd ? "Y" : "N",
    분리요청코드: r.sepCodes.join(","),
    배려요청코드: r.careCodes.join(","),
    비고: r.note
  }));

  const payload = {
    meta: { total: studentRows.length, classCount, iterations, seed, elapsedMs, weights, sepStrength: sepStrengthEl.value, careStrength: careStrengthEl.value },
    best: { score: best.score, sepViol: best.sepViol, careMiss: best.careMiss },
    arrays: { cnt: best.cnt, male: best.male, female: best.female, spec: best.spec, adhd: best.adhd },
    resultRows
  };

  showOverlay(false);
  // enable result tab and render
  tabResultBtn.disabled = false;
  statusPill.textContent = "완료";
  try{ renderResult(payload); }catch(e){ console.error(e); }
  showTab("result");
});


// ===== In-app Result Tab Rendering (v3.1) =====

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
function renderResult(payload){
  // (payload is provided by app.js after computation)
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
  if (typeof Chart === "undefined"){
    console.warn("Chart.js not loaded");
    return;
  }
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
}
