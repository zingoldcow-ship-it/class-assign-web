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
    runBtn.disabled = false;
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

  try{
    sessionStorage.setItem("classAssignResultV3", JSON.stringify(payload));
  } catch(e){
    // fallback: localStorage if sessionStorage fails
    localStorage.setItem("classAssignResultV3", JSON.stringify(payload));
  }

  showOverlay(false);
  window.open("./result.html", "_blank");
});
