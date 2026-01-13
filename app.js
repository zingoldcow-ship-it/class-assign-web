// 반배정 웹앱 v3.2.5 (브라우저 전용)
// - 전역 오염 최소화(IIFE)
// - 결과 탭 테이블/요약 렌더링 버그 수정
// - 함수/변수 중복 선언(Identifier already declared) 오류 수정

console.log('class-assign webapp v3.3.2 loaded');

// 앱 전체를 IIFE로 감싸 전역변수/함수 충돌을 줄입니다.
(()=>{
  // ----- Global error hooks (UI) -----
  window.addEventListener('error', (e)=>{
    try{
      const errorsDiv = document.getElementById("errors");
      const statusPill = document.getElementById("statusPill");
      if (statusPill) statusPill.textContent = "오류";
      if (errorsDiv){
        errorsDiv.textContent = "스크립트 오류: " + (e?.message || e) + (e?.filename ? ("\n" + e.filename + ":" + e.lineno) : "");
      }
    }catch(_){/* noop */}
  });
  window.addEventListener('unhandledrejection', (e)=>{
    try{
      const errorsDiv = document.getElementById("errors");
      const statusPill = document.getElementById("statusPill");
      if (statusPill) statusPill.textContent = "오류";
      if (errorsDiv){
        errorsDiv.textContent = "비동기 오류: " + (e?.reason?.message || e?.reason || e);
      }
    }catch(_){/* noop */}
  });

  // ----- Column policy -----
  const REQUIRED_COLUMNS = ['학생명','성별','생년월일','학업성취','교우관계','학부모민원','특수여부','ADHD여부','분리요청학생','배려요청학생','비고'];

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
    "생년월일": ["생년월일","생년","생일","출생일","출생","생년월일(yyyy-mm-dd)","생년월일(YYYY-MM-DD)"],
    "분리요청학생": ["분리요청학생","분리요청코드","분리코드","분리요청","분리","분리요청 코드"],
    "배려요청학생": ["배려요청학생","배려요청코드","배려코드","배려요청","배려","배려요청 코드"],
    "비고": ["비고","특이사항","메모","참고","기타","비고(특이사항)"]
  };

  function mapToStandardHeader(inputHeader){
    const nh = normHeader(inputHeader);
    if (!nh) return null;
    for (const [std, aliases] of Object.entries(HEADER_ALIASES)){
      for (const a of aliases){
        if (nh === normHeader(a)) return std;
      }
    }
    for (const std of Object.keys(HEADER_ALIASES)){
      if (nh === normHeader(std)) return std;
    }
    return null;
  }

  function buildHeaderMap(headers){
    const map = {};
    headers.forEach(h=>{
      const std = mapToStandardHeader(h);
      if (std && !map[std]) map[std] = h; // 첫 매칭 우선
    });
    return map;
  }

  function parseWorksheetToObjects(ws){
    // 헤더가 1행이 아닐 수도 있어(안내문/제목/빈 줄). 헤더 행을 자동으로 찾는다.
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw:false });
    if (!matrix || matrix.length===0) return {rows:[], headerRow:-1, score:0};

    let bestHeaderRow = -1;
    let bestScore = -1;
    let bestHeaders = [];
    const scanLimit = Math.min(30, matrix.length);
    for (let r=0; r<scanLimit; r++){
      const row = matrix[r] || [];
      const headers = row.map(x=>safeString(x)).filter(Boolean);
      if (!headers.length) continue;
      const map = buildHeaderMap(headers);
      const score = Object.keys(map).length + (map["학생명"]?5:0) + (map["성별"]?3:0);
      if (score > bestScore){
        bestScore = score;
        bestHeaderRow = r;
        bestHeaders = row.map(x=>safeString(x));
      }
    }

    if (bestHeaderRow < 0) return {rows:[], headerRow:-1, score:0};

    // 표준 헤더로 변환(가능한 경우)하여 키를 안정화
    const stdHeaders = bestHeaders.map(h=> mapToStandardHeader(h) || safeString(h));
    const out = [];
    for (let r = bestHeaderRow + 1; r < matrix.length; r++){
      const row = matrix[r] || [];
      // 완전 빈 줄은 스킵
      const hasAny = row.some(v=>safeString(v)!=="");
      if (!hasAny) continue;
      const obj = {};
      for (let c=0; c<stdHeaders.length; c++){
        const key = stdHeaders[c];
        if (!key) continue;
        obj[key] = row[c] ?? "";
      }
      // 학생명 없으면 데이터 행으로 보기 어려우므로 스킵
      if (!safeString(obj["학생명"]||obj["이름"]||obj["성명"])) continue;
      out.push(obj);
    }
    return {rows: out, headerRow: bestHeaderRow, score: bestScore};
  }

  function pickBestSheetName(workbook){
    let best = workbook.SheetNames[0];
    let bestRows = -1;
    let bestScore = -1;
    for (const name of workbook.SheetNames){
      const ws = workbook.Sheets[name];
      const parsed = parseWorksheetToObjects(ws);
      const n = parsed.rows.length;
      // 우선순위: (1) 읽힌 행 수, (2) 헤더 매칭 점수
      if (n > bestRows || (n===bestRows && parsed.score > bestScore)){
        bestRows = n;
        bestScore = parsed.score;
        best = name;
      }
    }
    return best;
  }

  // ----- DOM refs -----
  const fileInput = document.getElementById("fileInput");
  const filePill = document.getElementById("filePill");
  const rowsPill = document.getElementById("rowsPill");
  const errorsDiv = document.getElementById("errors");
  const statsDiv = document.getElementById("stats");
  const previewTableEl = document.getElementById("previewTable");

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

  // ----- Tab UI -----
  const tabSetupBtn = document.getElementById("tabSetup");
  const tabResultBtn = document.getElementById("tabResult");
  const setupTab = document.getElementById("setupTab");
  const resultTab = document.getElementById("resultTab");
  const statusPill = document.getElementById("statusPill");
  const progressTxt = document.getElementById("progressTxt");

  function showTab(which){
    const isSetup = (which === "setup");
    setupTab.style.display = isSetup ? "" : "none";
    resultTab.style.display = isSetup ? "none" : "";
    tabSetupBtn.classList.toggle("active", isSetup);
    tabResultBtn.classList.toggle("active", !isSetup);
  }
  tabSetupBtn?.addEventListener("click", ()=> showTab("setup"));
  tabResultBtn?.addEventListener("click", ()=> showTab("result"));

  let rawRows = null;
  let studentRows = null;
  let lastHeaders = null; // 업로드 엑셀의 열 순서 기억(결과 다운로드에 반영)

  function syncWeights(){
    wAcadV.textContent = wAcad.value;
    wPeerV.textContent = wPeer.value;
    wParentV.textContent = wParent.value;
  }
  [wAcad,wPeer,wParent].forEach(el=>el.addEventListener("input", syncWeights));
  syncWeights();

  function showOverlay(on, msg){
    overlay.style.display = on ? "flex" : "none";
    if (msg) progressTxt.textContent = msg;
  }

  // Run button state helper (DOM 속성/프로퍼티 불일치로 버튼이 계속 비활성처럼 보이는 문제 방지)
  function setRunEnabled(enabled){
    if (!runBtn) return;
    runBtn.disabled = !enabled;
    // disabled 속성은 HTML에 남아있으면 브라우저에 따라 UI가 계속 비활성처럼 보일 수 있어 명시적으로 정리
    if (enabled) runBtn.removeAttribute("disabled");
    else runBtn.setAttribute("disabled", "");
  }


  // 버튼 기본 타입 보정(어떤 컨테이너/테마에서는 submit으로 오작동할 수 있어 명시)
  try{ if (runBtn) runBtn.type = "button"; }catch(e){}

  // 실행 버튼 워치독: 엑셀을 정상 로드했는데도 disabled 속성이 남아 클릭이 안 되는 경우를 방지
  // (캐시/브라우저별 disabled attribute 잔존 이슈 대응)
  setInterval(()=>{
    if (!runBtn) return;
    const shouldEnable = !!(studentRows && studentRows.length>0);
    if (shouldEnable && (runBtn.disabled || runBtn.hasAttribute("disabled"))){
      setRunEnabled(true);
    }
  }, 500);


  function safeString(x){ return (x===null||x===undefined) ? "" : String(x).trim(); }
  function ynTo01(x){
    const v = safeString(x).toUpperCase();
    if (v === "Y" || v === "1" || v === "TRUE") return 1;
    return 0;
  }
  function level3ToScore(x){
    const v = safeString(x);
    if (v === "좋음") return 1;
    if (v === "보통") return 0;
    if (v === "나쁨") return -1;
    return 0;
  }
  function splitCodes(x){
    const s = safeString(x);
    if (!s) return [];
    return s.split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
  }

  // ----- Setup preview table -----
  function renderSetupPreviewTable(rows, maxRows=20){
    previewTableEl.innerHTML = "";
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
    previewTableEl.appendChild(thead);

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
    previewTableEl.appendChild(tbody);
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
        <span class="pill">분리학생 ${rows.reduce((a,r)=>a+(r.sepCodes.length>0),0)}명</span>
        <span class="pill">배려학생 ${rows.reduce((a,r)=>a+(r.careCodes.length>0),0)}명</span>
      </div>
      <div style="height:10px"></div>
      <div class="small">* 실행을 누르면 결과는 같은 화면의 “결과” 탭에서 표시됩니다.</div>
    `;
  }

  function normalizeRow(r){
    const name = safeString(r["학생명"]||r["이름"]||r["성명"]);
    const gender = safeString(r["성별"]||r["남녀"]);
    const birth = safeString(r["생년월일"]||r["생년"]||r["생일"]||r["출생일"]);
    const acad = safeString(r["학업성취"]||r["학업성취(3단계)"]);
    const peer = safeString(r["교우관계"]||r["교우관계(3단계)"]);
    const parent = safeString(r["학부모민원"]||r["학부모민원(3단계)"]);
    const special = ynTo01(r["특수여부"]||r["특수"]||r["특수여부(Y/N)"]);
    const adhd = ynTo01(r["ADHD여부"]||r["adhd여부"]||r["ADHD"]||r["ADHD여부(Y/N)"]);
    const note = safeString(r["비고"]||r["특이사항"]||r["메모"]);
        const sepCodes = splitCodes(r["분리요청학생"]||r["분리요청학생"]||r["분리코드"]||r["분리"]);
        const careCodes = splitCodes(r["배려요청학생"]||r["배려요청학생"]||r["배려코드"]||r["배려"]);
        return {
      _raw: {...r},
      name, gender, birth,
      acad, peer, parent,
      acadS: level3ToScore(acad),
      peerS: level3ToScore(peer),
      parentS: level3ToScore(parent),
      special, adhd,
      note,
      sepCodes, careCodes
    };
  }

  // ----- RNG / scoring -----
  function mulberry32(seed){
    let t = seed >>> 0;
    return function(){
      t += 0x6D2B79F5;
      let x = Math.imul(t ^ (t >>> 15), 1 | t);
      x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
      return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
    };
  }
  function randInt(rng, n){ return Math.floor(rng()*n); }
  function comb2(k){ return k>1 ? (k*(k-1))/2 : 0; }

  function buildCodeGroups(rows){
    const sep = new Map();
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
    for (const [k,v] of [...sep.entries()]) if (v.length < 2) sep.delete(k);
    for (const [k,v] of [...care.entries()]) if (v.length < 2) care.delete(k);
    return {sep, care};
  }

  function scoreAssignment(rows, assign, classCount, weights, groups){
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
      if (rows[i].gender === "남") male[c] += 1;
      else if (rows[i].gender === "여") female[c] += 1;
      spec[c] += rows[i].special;
      adhd[c] += rows[i].adhd;
      acadSum[c] += rows[i].acadS;
      peerSum[c] += rows[i].peerS;
      parentSum[c] += rows[i].parentS;
    }

    function variance(arr){
      const m = arr.reduce((a,b)=>a+b,0)/arr.length;
      let v = 0;
      for (const x of arr){ const d=x-m; v += d*d; }
      return v/arr.length;
    }

    const vCnt = variance(cnt);
    const vMale = variance(male);
    const vFemale = variance(female);
    const vSpec = variance(spec);
    const vAdhd = variance(adhd);

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

    // Separation penalties
    let sepViol = 0;
    for (const [, idxs] of groups.sep.entries()){
      const perClass = new Map();
      for (const i of idxs){
        const c = assign[i];
        perClass.set(c, (perClass.get(c)||0)+1);
      }
      for (const k of perClass.values()) sepViol += comb2(k);
    }
    score += weights.sepPenalty * sepViol;

    // Care penalties
    let careMiss = 0;
    for (const [, idxs] of groups.care.entries()){
      const totalPairs = comb2(idxs.length);
      if (totalPairs === 0) continue;
      const perClass = new Map();
      for (const i of idxs){
        const c = assign[i];
        perClass.set(c, (perClass.get(c)||0)+1);
      }
      let within = 0;
      for (const k of perClass.values()) within += comb2(k);
      careMiss += (totalPairs - within);
    }
    score += weights.carePenalty * careMiss;

    return {score, sepViol, careMiss, cnt, male, female, spec, adhd, acadSum, peerSum, parentSum};

  }

  // 분리/배려 '쌍' 수치가 너무 크게 느껴질 수 있어, 실제로 영향을 받은 '학생 수'도 계산합니다.
  function computeViolationStudentCounts(assign, groups){
    const sepSet = new Set();
    const careSet = new Set();

    // 분리: 같은 반에 2명 이상 함께 배정된 그룹 구성원을 '위반 학생'으로 집계
    for (const [, idxs] of groups.sep.entries()){
      const perClass = new Map();
      for (const i of idxs){
        const c = assign[i];
        if (!perClass.has(c)) perClass.set(c, []);
        perClass.get(c).push(i);
      }
      for (const arr of perClass.values()){
        if (arr.length >= 2){
          for (const i of arr) sepSet.add(i);
        }
      }
    }

    // 배려: 같은 그룹이 여러 반으로 나뉜 경우, 해당 그룹 구성원을 '미충족 학생'으로 집계
    for (const [, idxs] of groups.care.entries()){
      const classes = new Set();
      for (const i of idxs) classes.add(assign[i]);
      if (classes.size >= 2){
        for (const i of idxs) careSet.add(i);
      }
    }

    return { sepStudents: sepSet.size, careStudents: careSet.size };
  }



  function initialAssignment(rows, classCount, rng){
    const idxs = rows.map((_,i)=>i);
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

    const n = rows.length;
    const reportEvery = Math.max(200, Math.floor(iterations/30));

    for (let t=1; t<=iterations; t++){
      const i = randInt(rng, n);
      let j = randInt(rng, n);
      if (j === i) j = (j+1) % n;

      const ci = assign[i], cj = assign[j];
      if (ci === cj) continue;

      assign[i] = cj; assign[j] = ci;
      const s = scoreAssignment(rows, assign, classCount, weights, groups);

      let accept = false;
      if (s.score <= best.score){
        accept = true;
      } else {
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
        assign[i] = ci; assign[j] = cj;
      }

      if (t % reportEvery === 0){
        progressTxt.textContent = `시뮬레이션 ${t.toLocaleString()} / ${iterations.toLocaleString()} (현재 best score: ${Math.round(best.score).toLocaleString()})`;
        await new Promise(r=>setTimeout(r, 0));
      }
    }

    return {best, bestAssign};
  }

  function strengthToPenalty(strength, kind){
    if (kind === 'sep'){
      if (strength === 'strict') return 500000;
      if (strength === 'strong') return 120000;
      if (strength === 'medium') return 50000;
      return 15000;
    }
    // care (같은 반 선호): 기존보다 강도를 더 크게 반영
    if (strength === 'strict') return 50000;   // 가능한 한 같은 반
    if (strength === 'strong') return 20000;
    if (strength === 'medium') return 8000;
    return 2000;
  }

  function validateRows(rows){
    const missing = [];
    const nameOk = rows.some(r=>r.name);
    const genderOk = rows.some(r=>r.gender);
    if (!nameOk) missing.push("학생명");
    if (!genderOk) missing.push("성별");
    if (missing.length) return `엑셀에 필요한 열(또는 값)이 부족합니다: ${missing.join(", ")}`;
    return null;
  }

  // ----- File load -----
  fileInput?.addEventListener("change", async (e)=>{
    setErrors("");
    const file = e.target.files?.[0];
    if (!file){
      filePill.textContent = "엑셀 미선택";
      setRunEnabled(false);
      return;
    }
    filePill.textContent = `선택됨: ${file.name}`;

    showOverlay(true, "엑셀을 읽는 중…");
    await new Promise(r=>setTimeout(r, 10));

    try{
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, {type:"array"});
      const bestName = pickBestSheetName(wb);
      const ws = wb.Sheets[bestName];
      const parsed = parseWorksheetToObjects(ws);
      rawRows = parsed.rows;
      if (!rawRows || rawRows.length === 0){
        setErrors("엑셀에서 데이터를 읽지 못했어요. '학생명/성별' 헤더가 있는 시트와 표(헤더 행)를 확인해주세요.");
        setRunEnabled(false);
        return;
      }

      // (선택) 헤더 매핑 점검 - 현재는 normalizeRow가 유연하게 읽음
      const headers = Object.keys(rawRows[0] || {});
      lastHeaders = headers.slice();
      const headerMap = buildHeaderMap(headers);
      const missingStd = REQUIRED_COLUMNS.filter(c=> !headerMap[c] && !headers.some(h=>normHeader(h)===normHeader(c)));
      // 학생명/성별은 validateRows로 다시 확인

      studentRows = rawRows.map(normalizeRow);
      const err = validateRows(studentRows);
      if (err){
        setErrors(err);
        setRunEnabled(false);
        return;
      }

      rowsPill.textContent = `${studentRows.length}명`;
      renderSetupPreviewTable(rawRows, 20);
      summarize(studentRows);
      statusPill.textContent = missingStd.length ? `엑셀 로드됨(권장열 누락: ${missingStd.join(', ')})` : "엑셀 로드됨";
      setRunEnabled(!!(studentRows && studentRows.length>0));
    } catch(err){
      console.error(err);
      setErrors("엑셀을 읽는 중 오류: " + (err?.message || String(err)));
      setRunEnabled(false);
    } finally{
      showOverlay(false);
    }
  });

  // ----- Run optimization -----
  runBtn?.addEventListener("click", async (ev)=>{
    try{ ev?.preventDefault?.(); ev?.stopPropagation?.(); }catch(e){}

    if (!studentRows || studentRows.length === 0) return;

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

    let payload = null;
    try{
showOverlay(true, "코드 그룹(분리/배려)을 구성하는 중…");
    await new Promise(r=>setTimeout(r, 10));
    const groups = buildCodeGroups(studentRows);

    showOverlay(true, "시뮬레이션을 시작합니다…");
    const start = performance.now();
    const {best, bestAssign} = await optimize(studentRows, classCount, iterations, seed, weights, groups);
    const elapsedMs = Math.round(performance.now() - start);

    const violCount = computeViolationStudentCounts(bestAssign, groups);

    // 결과 다운로드는 '업로드 양식의 열'을 최대한 유지하고, 표준 열은 값이 확실히 채워지도록 보정합니다.
    const resultRows = studentRows.map((r,i)=>{
      const base = Object.assign({}, rawRows[i] || {});
      base["반"] = bestAssign[i] + 1;

      // 표준 열 보정(업로드 파일의 표기와 무관하게 결과가 안정적으로 나오도록)
      base["학생명"] = r.name;
      base["성별"] = r.gender;
      if (r.birth) base["생년월일"] = r.birth;
      base["학업성취"] = r.acad;
      base["교우관계"] = r.peer;
      base["학부모민원"] = r.parent;
      base["특수여부"] = r.special ? "Y" : "N";
      base["ADHD여부"] = r.adhd ? "Y" : "N";
      base["비고"] = r.note;

      // 분리/배려는 '학생' 표기로 통일하여 하나의 열만 남깁니다.
      base["분리요청학생"] = r.sepCodes.join(",");
      base["배려요청학생"] = r.careCodes.join(",");
      delete base["분리요청코드"];
      delete base["배려요청코드"];
      return base;
    });

    const payload = {
      meta: { total: studentRows.length, classCount, iterations, seed, elapsedMs, weights, sepStrength: sepStrengthEl.value, careStrength: careStrengthEl.value },
      best: { score: best.score, sepPairs: best.sepViol, carePairs: best.careMiss, sepStudents: violCount.sepStudents, careStudents: violCount.careStudents },
      arrays: { cnt: best.cnt, male: best.male, female: best.female, spec: best.spec, adhd: best.adhd },
      resultRows
    };

    showOverlay(false);
    tabResultBtn.disabled = false;
    statusPill.textContent = "완료";
    try{ renderResult(payload); }catch(e){ console.error(e); setErrors("결과 렌더링 오류: " + (e?.message || e)); }
    showTab("result");
    }catch(e){ console.error(e); setErrors("실행 중 오류: " + (e?.message || e)); }
    finally{ showOverlay(false); }

  });

  // ===== Result Tab Rendering =====
  const metaPill = document.getElementById("metaPill");
  const scorePill = document.getElementById("scorePill");
  const sepPill = document.getElementById("sepPill");
  const carePill = document.getElementById("carePill");
  const classSummary = document.getElementById("classSummary");
  const violationsDiv = document.getElementById("violations");
  const resultTableEl = document.getElementById("resultTable");
  const classFilter = document.getElementById("classFilter");
  const tableMeta = document.getElementById("tableMeta");
  const downloadBtn = document.getElementById("downloadBtn");

  function renderResult(payload){
    metaPill.textContent = `${payload.meta.total}명 · ${payload.meta.classCount}반 · ${payload.meta.iterations.toLocaleString()}회 · seed ${payload.meta.seed} · ${payload.meta.elapsedMs.toLocaleString()}ms`;
    scorePill.textContent = `Score: ${Math.round(payload.best.score).toLocaleString()}`;
    sepPill.textContent = `분리 미충족: ${payload.best.sepStudents.toLocaleString()}명 (위반 ${payload.best.sepPairs.toLocaleString()}쌍)`;
    carePill.textContent = `배려 미충족: ${payload.best.careStudents.toLocaleString()}명 (미충족 ${payload.best.carePairs.toLocaleString()}쌍)`;

    function renderClassSummary(){
      const C = payload.meta.classCount;
      const {cnt, male, female, spec, adhd} = payload.arrays;

      let html = "<div style='overflow:auto'><table><thead><tr><th>반</th><th>인원</th><th>남</th><th>여</th><th>특수</th><th>ADHD</th></tr></thead><tbody>";
      for (let c=0;c<C;c++){
        html += `<tr><td>${c+1}</td><td>${cnt[c]}</td><td>${male[c]}</td><td>${female[c]}</td><td>${spec[c]}</td><td>${adhd[c]}</td></tr>`;
      }
      html += "</tbody></table></div>";
      classSummary.innerHTML = html;
    }

    function buildViolationReport(){
      const rows = payload.resultRows;
      const sepMap = new Map();
      const careMap = new Map();

      function add(map, code, cls){
        if (!map.has(code)) map.set(code, new Map());
        const m = map.get(code);
        m.set(cls, (m.get(cls)||0)+1);
      }

      for (const r of rows){
        const cls = r["반"];
        const sepCodes = (r["분리요청학생"]||"").split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
        const careCodes = (r["배려요청학생"]||"").split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
        for (const c of sepCodes) add(sepMap, c, cls);
        for (const c of careCodes) add(careMap, c, cls);
      }

      const worstSep = [];
      for (const [code, m] of sepMap.entries()){
        for (const [cls, k] of m.entries()){
          if (k >= 2) worstSep.push({code, cls, k});
        }
      }
      worstSep.sort((a,b)=>b.k-a.k);

      const worstCare = [];
      for (const [code, m] of careMap.entries()){
        let total = 0;
        for (const k of m.values()) total += k;
        if (total < 2) continue;
        const classes = [...m.keys()].sort((a,b)=>a-b);
        if (classes.length >= 2) worstCare.push({code, classes: classes.join(","), total});
      }

      let html = "";
      html += `<div class="small"><b>분리 위반(상위 10)</b></div>`;
      if (worstSep.length === 0) html += `<div class="small">- 위반 없음</div>`;
      else {
        html += "<div style='overflow:auto;max-height:160px;'><table><thead><tr><th>코드</th><th>반</th><th>동반 인원</th></tr></thead><tbody>";
        for (const x of worstSep.slice(0,10)) html += `<tr><td>${x.code}</td><td>${x.cls}</td><td>${x.k}</td></tr>`;
        html += "</tbody></table></div>";
      }

      html += `<div style="height:10px"></div><div class="small"><b>배려 분산(상위 10)</b></div>`;
      if (worstCare.length === 0) html += `<div class="small">- 분산 없음(또는 코드 없음)</div>`;
      else {
        html += "<div style='overflow:auto;max-height:160px;'><table><thead><tr><th>코드</th><th>분산된 반</th><th>총 인원</th></tr></thead><tbody>";
        for (const x of worstCare.slice(0,10)) html += `<tr><td>${x.code}</td><td>${x.classes}</td><td>${x.total}</td></tr>`;
        html += "</tbody></table></div>";
      }

      violationsDiv.innerHTML = html;
    }

    function renderResultTable(filterClass){
      const all = payload.resultRows.slice().sort((a,b)=>a["반"]-b["반"] || String(a["학생명"]).localeCompare(String(b["학생명"])));
      const filtered = (filterClass === "all") ? all : all.filter(r=>String(r["반"]) === String(filterClass));

      resultTableEl.innerHTML = "";
      if (filtered.length === 0){
        tableMeta.textContent = "표시 0명";
        return;
      }

      const cols = Object.keys(filtered[0]);
      const thead = document.createElement("thead");
      const trh = document.createElement("tr");
      for (const c of cols){
        const th = document.createElement("th");
        th.textContent = c;
        trh.appendChild(th);
      }
      thead.appendChild(trh);
      resultTableEl.appendChild(thead);

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
      resultTableEl.appendChild(tbody);
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

      // 중복 리스너 방지
      classFilter.onchange = () => renderResultTable(classFilter.value);
    }

    function drawCharts(){
      if (typeof Chart === "undefined"){
        console.warn("Chart.js not loaded");
        return;
      }
      Chart.defaults.responsive = false;
      Chart.defaults.animation = false;

      const labels = Array.from({length: payload.meta.classCount}, (_,i)=>`${i+1}반`);
      const {cnt, male, female} = payload.arrays;

      const cntCtx = document.getElementById("cntChart");
      const gCtx = document.getElementById("genderChart");

      // 이전 차트가 있으면 제거(재실행 시 겹침 방지)
      if (cntCtx?._chartInstance) cntCtx._chartInstance.destroy();
      if (gCtx?._chartInstance) gCtx._chartInstance.destroy();

      cntCtx._chartInstance = new Chart(cntCtx, {
        type:"bar",
        data:{ labels, datasets:[{ label:"인원", data: cnt }]},
        options:{ responsive:false, animation:false, plugins:{ legend:{ display:false }}, scales:{ y:{ beginAtZero:true } } }
      });

      gCtx._chartInstance = new Chart(gCtx, {
        type:"bar",
        data:{ labels, datasets:[
          { label:"남", data: male, stack:"g" },
          { label:"여", data: female, stack:"g" },
        ]},
        options:{ responsive:false, animation:false, plugins:{ legend:{ position:"bottom" }}, scales:{ x:{ stacked:true }, y:{ stacked:true, beginAtZero:true } } }
      });
    }

    downloadBtn.onclick = ()=>{
      // 다운로드 파일은 "새로운 반"(1반→2반→...) 기준으로 정렬
      const sorted = payload.resultRows.slice().sort((a,b)=>
        (a["반"]-b["반"]) || String(a["학생명"]||"").localeCompare(String(b["학생명"]||""))
      );

      const wb = XLSX.utils.book_new();
      // 업로드한 양식의 열 순서를 최대한 유지하고, 분리/배려 열은 '학생' 표기로 통일
      const baseHeaders = (lastHeaders && lastHeaders.length) ? lastHeaders : Object.keys(sorted[0] || {});
      const mappedHeaders = baseHeaders.map(h=>{
        const nh = normHeader(h);
        if (nh === normHeader("분리요청코드") || nh === normHeader("분리요청학생")) return "분리요청학생";
        if (nh === normHeader("배려요청코드") || nh === normHeader("배려요청학생")) return "배려요청학생";
        return h;
      });
      // 중복 제거(순서 유지)
      const seen = new Set();
      const ordered = [];
      for (const h of mappedHeaders){
        if (!h) continue;
        if (!seen.has(h)){ seen.add(h); ordered.push(h); }
      }
      // '반'은 항상 맨 앞에
      const headerOrder = ["반", ...ordered.filter(h=>h!=="반")];

      const ws = XLSX.utils.json_to_sheet(sorted, { header: headerOrder });
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
        ["분리 미충족(명)", payload.best.sepStudents],
        ["배려 미충족(명)", payload.best.careStudents],
        ["분리 위반(쌍)", payload.best.sepPairs],
        ["배려 미충족(쌍)", payload.best.carePairs],
        ["Score", payload.best.score]
      ]);
      XLSX.utils.book_append_sheet(wb, metaSheet, "설정요약");

      XLSX.writeFile(wb, `반배정_결과_${payload.meta.total}명_${payload.meta.classCount}반_seed${payload.meta.seed}.xlsx`);
    };

    renderClassSummary();
    buildViolationReport();
    setupFilter();
    renderResultTable("all");
    drawCharts();
  }

})();
