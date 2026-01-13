// 반배정 웹앱 v3.2.5 (브라우저 전용)
// - 전역 오염 최소화(IIFE)
// - 결과 탭 테이블/요약 렌더링 버그 수정
// - 함수/변수 중복 선언(Identifier already declared) 오류 수정

console.log('class-assign webapp v3.4.2 loaded');

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

  function escapeHtml(s){
    return String(s ?? "").replace(/[&<>"']/g, (ch)=>({
      "&":"&amp;",
      "<":"&lt;",
      ">":"&gt;",
      '"':"&quot;",
      "'":"&#39;"
    }[ch]));
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
  const adhdCapEl = document.getElementById("adhdCap");
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
    // 흔한 '없음' 표기들은 코드로 취급하지 않습니다.
    const NO = new Set(["-","–","—","없음","없","x","X","0","N","n","미입력","무"]);
    return s
      .split(/[,\s;]+/)
      .map(t=>t.trim())
      .filter(t=>t && !NO.has(t));
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

  function setErrors(msg){
    const m = msg || "";
    if (errorsDiv) errorsDiv.textContent = m;
    // 결과 탭에 있을 때도 에러가 보이도록 리포트 영역에 함께 표시
    try{
      if (!m) return;
      const v = document.getElementById('violations');
      if (v) v.innerHTML = `<div class="danger">${escapeHtml(m)}</div>`;
    }catch(e){}
  }

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

    // '코드' 뿐 아니라 '학생명(또는 학번 형태)'를 직접 넣는 경우도 반영합니다.
    // 예: 배려요청학생에 '학생026'을 넣으면, 해당 학생과 같은 반이 되도록(가능하면) 처리
    const nameToIdx = new Map();
    rows.forEach((r, i)=>{ if (r?.name) nameToIdx.set(String(r.name).trim(), i); });
    rows.forEach((r, idx)=>{
      for (const c of r.sepCodes){
        if (!sep.has(c)) sep.set(c, []);
        sep.get(c).push(idx);
      }
      for (const c of r.careCodes){
        if (!care.has(c)) care.set(c, []);
        care.get(c).push(idx);
      }

      // 1:1 요청(상대 학생명을 직접 적은 경우)도 그룹으로 추가
      for (const c of r.sepCodes){
        const j = nameToIdx.get(String(c).trim());
        if (j !== undefined && j !== idx){
          const a = Math.min(idx, j), b = Math.max(idx, j);
          const key = `@sep_pair:${a}-${b}`;
          if (!sep.has(key)) sep.set(key, []);
          sep.get(key).push(a, b);
        }
      }
      for (const c of r.careCodes){
        const j = nameToIdx.get(String(c).trim());
        if (j !== undefined && j !== idx){
          const a = Math.min(idx, j), b = Math.max(idx, j);
          const key = `@care_pair:${a}-${b}`;
          if (!care.has(key)) care.set(key, []);
          care.get(key).push(a, b);
        }
      }
    });

    // 중복 제거
    for (const [k,v] of [...sep.entries()]) sep.set(k, Array.from(new Set(v)));
    for (const [k,v] of [...care.entries()]) care.set(k, Array.from(new Set(v)));
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
    const vAcad = variance(acadSum);
    const vPeer = variance(peerSum);
    const vParent = variance(parentSum);

    // 성비는 현장에서는 1순위로 중요해요. '분산(분산도)'만 보지 말고,
    // 각 반의 인원수에 비례한 '기대 남학생 수'에서 얼마나 벗어났는지(제곱오차)를 크게 벌점 줍니다.
    const totalMale = male.reduce((a,b)=>a+b,0);
    const total = rows.length || 1;
    const maleRatio = totalMale / total;
    let genderSqErr = 0;
    for (let c=0;c<C;c++){
      const expectedMale = cnt[c] * maleRatio;
      const d = male[c] - expectedMale;
      genderSqErr += d*d;
    }

    // 2순위: 특수(가능하면 0~1명/반). 총 특수 인원과 반 수를 바탕으로 이상적인 최소/최대치를 만들고,
    // 최대치를 초과하는 반은 매우 큰 벌점을 부여합니다.
    const totalSpec = spec.reduce((a,b)=>a+b,0);
    const specExpected = totalSpec / (C || 1);
    const specIdealMin = Math.floor(specExpected);
    const specIdealMax = Math.ceil(specExpected);

    let specSqErr = 0;
    let specOverflow = 0;
    for (let c=0;c<C;c++){
      const d = spec[c] - specExpected;
      specSqErr += d*d;
      const over = spec[c] - specIdealMax;
      if (over > 0) specOverflow += over*over;
    }

    // 3순위: ADHD(가능하면 0~1명/반, 여건상 불가하면 균등 분산).
    const totalAdhd = adhd.reduce((a,b)=>a+b,0);
    const adhdExpected = totalAdhd / (C || 1);
    const adhdIdealMin = Math.floor(adhdExpected);
    const adhdIdealMax = Math.ceil(adhdExpected);

    let adhdSqErr = 0;
    let adhdOverflow = 0;
    // 사용자가 'ADHD 반당 최대 인원'을 지정한 경우, 초과 반은 매우 큰 벌점(하드캡)을 부여합니다.
    let adhdHardCapOverflow = 0;
    for (let c=0;c<C;c++){
      const d = adhd[c] - adhdExpected;
      adhdSqErr += d*d;
      const over = adhd[c] - adhdIdealMax;
      if (over > 0) adhdOverflow += over*over;

      if (weights && typeof weights.adhdCap === 'number' && isFinite(weights.adhdCap)){
        const overCap = adhd[c] - weights.adhdCap;
        if (overCap > 0) adhdHardCapOverflow += overCap*overCap;
      }
    }

    let score =
      80*vCnt +
      260*genderSqErr +
      800*specSqErr +
      6000*specOverflow +
      350*adhdSqErr +
      2500*adhdOverflow +
      // 하드캡은 다른 항목보다 우선해서 지키도록 큰 벌점
      2000000*adhdHardCapOverflow +
      weights.wParent*vParent +
      weights.wAcad*vAcad +
      weights.wPeer*vPeer;

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

  // 분리/배려 '미충족'을 쌍(pair)뿐 아니라 '학생 목록'으로도 보여주기 위한 상세 계산
  // - 분리 미충족: 같은 분리요청학생(코드/상대학생)이 같은 반에 함께 배정된 학생
  // - 배려 미충족: 같은 배려요청학생(코드/상대학생) 친구가 자기 반에 1명도 없는 학생
  function computeUnsatisfiedDetails(rows, assign, groups){
    const sepSet = new Set();
    const careSet = new Set();
    const sepItems = [];
    const careItems = [];

    function labelForKey(key, kind){
      const s = String(key);
      if (s.startsWith('@sep_pair:') || s.startsWith('@care_pair:')){
        const parts = s.split(':')[1] || '';
        const ab = parts.split('-');
        const a = parseInt(ab[0],10); const b = parseInt(ab[1],10);
        const an = rows[a]?.name || `#${a}`;
        const bn = rows[b]?.name || `#${b}`;
        return `${an} ↔ ${bn}`;
      }
      return s;
    }

    // 분리
    for (const [code, idxs] of groups.sep.entries()){
      const perClass = new Map();
      for (const i of idxs){
        const c = assign[i];
        if (!perClass.has(c)) perClass.set(c, []);
        perClass.get(c).push(i);
      }
      for (const [cls, arr] of perClass.entries()){
        if (arr.length < 2) continue;
        const names = arr.map(i=>rows[i]?.name || '').filter(Boolean);
        for (const i of arr){
          sepSet.add(i);
          sepItems.push({
            kind: '분리',
            code: labelForKey(code,'sep'),
            student: rows[i]?.name || '',
            classNo: (assign[i] + 1),
            withStudents: names.join(', ')
          });
        }
      }
    }

    // 배려
    for (const [code, idxs] of groups.care.entries()){
      if (idxs.length < 2) continue;
      const classCounts = new Map();
      for (const i of idxs){
        const c = assign[i];
        classCounts.set(c, (classCounts.get(c) || 0) + 1);
      }
      const groupNames = idxs.map(i=>rows[i]?.name || '').filter(Boolean);
      for (const i of idxs){
        const c = assign[i];
        if ((classCounts.get(c) || 0) < 2){
          careSet.add(i);
          careItems.push({
            kind: '배려',
            code: labelForKey(code,'care'),
            student: rows[i]?.name || '',
            classNo: (assign[i] + 1),
            group: groupNames.join(', ')
          });
        }
      }
    }

    // 보기 좋게 정렬
    sepItems.sort((a,b)=>a.classNo-b.classNo || String(a.code).localeCompare(String(b.code)) || String(a.student).localeCompare(String(b.student)));
    careItems.sort((a,b)=>a.classNo-b.classNo || String(a.code).localeCompare(String(b.code)) || String(a.student).localeCompare(String(b.student)));

    return {
      sepStudents: sepSet.size,
      careStudents: careSet.size,
      sepItems,
      careItems
    };
  }

  function initialAssignment(rows, classCount, rng){
    // 초기 배치는 '성비'를 최대한 맞춘 상태에서 시작하면 결과가 훨씬 안정적입니다.
    // 1) 남/여를 각각 셔플
    // 2) 각 집합을 round-robin으로 반에 배치
    const males = [];
    const females = [];
    const others = [];
    rows.forEach((r,i)=>{
      if (r.gender === '남') males.push(i);
      else if (r.gender === '여') females.push(i);
      else others.push(i);
    });
    function shuffle(a){
      for (let i=a.length-1;i>0;i--){
        const j = randInt(rng, i+1);
        [a[i], a[j]] = [a[j], a[i]];
      }
    }
    shuffle(males); shuffle(females); shuffle(others);

    const assign = new Array(rows.length).fill(0);
    let c = 0;
    for (const i of males){ assign[i] = c; c = (c+1)%classCount; }
    for (const i of females){ assign[i] = c; c = (c+1)%classCount; }
    for (const i of others){ assign[i] = c; c = (c+1)%classCount; }
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
    // care (같은 반 선호): 반 구성에서는 '후순위'인 경우가 많아 가중치를 낮게 둡니다.
    // (성비/인원균등/특수/ADHD/학업/교우/민원 + 분리요청을 우선 반영)
    if (strength === 'strict') return 20000;   // 가능한 한 같은 반
    if (strength === 'strong') return 8000;
    if (strength === 'medium') return 3000;
    return 1000;
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

    // 먼저 결과 탭을 열고 진행 상황을 표시(클릭했는데 반응 없는 것처럼 보이는 문제 방지)
    try{
      if (tabResultBtn){
        tabResultBtn.disabled = false;
    try{ tabResultBtn.removeAttribute('disabled'); }catch(e){}
        tabResultBtn.removeAttribute("disabled");
      }
      showTab("result");
    }catch(e){}

    const classCount = Math.max(2, Math.min(30, parseInt(classCountEl.value||"10",10)));
    const iterations = Math.max(200, Math.min(60000, parseInt(iterationsEl.value||"8000",10)));
    const seed = Math.max(0, Math.min(999999, parseInt(seedEl.value||"42",10)));

    const weights = {
      wAcad: parseInt(wAcad.value,10),
      wPeer: parseInt(wPeer.value,10),
      wParent: parseInt(wParent.value,10),
      sepPenalty: strengthToPenalty(sepStrengthEl.value, 'sep'),
      carePenalty: strengthToPenalty(careStrengthEl.value, 'care'),
      // ADHD 반당 최대 인원(선택). auto면 제한 없음.
      adhdCap: (adhdCapEl && String(adhdCapEl.value||'auto') !== 'auto') ? Math.max(1, Math.min(5, parseInt(adhdCapEl.value,10))) : null
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

    const unsat = computeUnsatisfiedDetails(studentRows, bestAssign, groups);

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
      best: { score: best.score, sepPairs: best.sepViol, carePairs: best.careMiss, sepStudents: unsat.sepStudents, careStudents: unsat.careStudents },
      arrays: { cnt: best.cnt, male: best.male, female: best.female, spec: best.spec, adhd: best.adhd },
      resultRows,
      unsatisfied: { sepItems: unsat.sepItems, careItems: unsat.careItems }
    };

    showOverlay(false);
    tabResultBtn.disabled = false;
    try{ tabResultBtn.removeAttribute('disabled'); }catch(e){}
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
  const unsatSepMeta = document.getElementById("unsatSepMeta");
  const unsatCareMeta = document.getElementById("unsatCareMeta");
  const unsatSepTable = document.getElementById("unsatSepTable");
  const unsatCareTable = document.getElementById("unsatCareTable");

  function renderResult(payload){
    metaPill.textContent = `${payload.meta.total}명 · ${payload.meta.classCount}반 · ${payload.meta.iterations.toLocaleString()}회 · seed ${payload.meta.seed} · ${payload.meta.elapsedMs.toLocaleString()}ms`;
    scorePill.textContent = `Score: ${Math.round(payload.best.score).toLocaleString()}`;
    sepPill.textContent = `분리 미충족: ${payload.best.sepStudents.toLocaleString()}명 (위반 ${payload.best.sepPairs.toLocaleString()}쌍)`;
    carePill.textContent = `배려 미충족: ${payload.best.careStudents.toLocaleString()}명 (미충족 ${payload.best.carePairs.toLocaleString()}쌍)`;

    function renderUnsatisfiedTables(){
      const sep = payload?.unsatisfied?.sepItems || [];
      const care = payload?.unsatisfied?.careItems || [];

      if (unsatSepMeta) unsatSepMeta.textContent = `(${sep.length.toLocaleString()}건)`;
      if (unsatCareMeta) unsatCareMeta.textContent = `(${care.length.toLocaleString()}건)`;

      function renderTable(el, cols, rows){
        if (!el) return;
        el.innerHTML = '';
        const thead = document.createElement('thead');
        const trh = document.createElement('tr');
        for (const c of cols){
          const th = document.createElement('th');
          th.textContent = c;
          trh.appendChild(th);
        }
        thead.appendChild(trh);
        el.appendChild(thead);

        const tbody = document.createElement('tbody');
        for (const r of rows){
          const tr = document.createElement('tr');
          for (const c of cols){
            const td = document.createElement('td');
            td.textContent = (r[c]===null||r[c]===undefined) ? '' : String(r[c]);
            tr.appendChild(td);
          }
          tbody.appendChild(tr);
        }
        el.appendChild(tbody);
      }

      renderTable(
        unsatSepTable,
        ["학생명","반","분리요청","같은반 학생"],
        sep.map(x=>({"학생명":x.student,"반":x.classNo,"분리요청":x.code,"같은반 학생":x.withStudents}))
      );

      renderTable(
        unsatCareTable,
        ["학생명","반","배려요청","같은 코드 그룹"],
        care.map(x=>({"학생명":x.student,"반":x.classNo,"배려요청":x.code,"같은 코드 그룹":x.group}))
      );
    }

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

      // ----- helper: 3단계(좋음/보통/나쁨) 반별 요약 -----
      function buildLevelReport(title, field){
        const C = payload.meta.classCount;
        const buckets = Array.from({length:C}, ()=>({good:0, normal:0, bad:0, total:0}));
        for (const r of rows){
          const cls = Math.max(1, parseInt(r["반"],10)) - 1;
          if (cls < 0 || cls >= C) continue;
          const v = safeString(r[field]);
          if (v === "좋음") buckets[cls].good++;
          else if (v === "나쁨") buckets[cls].bad++;
          else buckets[cls].normal++;
          buckets[cls].total++;
        }

        // 평균 점수(좋음=+1, 보통=0, 나쁨=-1)로 상/하위 반을 뽑아 보여줌
        const scored = buckets.map((b, i)=>{
          const score = (b.good - b.bad) / Math.max(1, b.total);
          return {cls:i+1, score, ...b};
        });
        scored.sort((a,b)=>b.score-a.score);

        const top = scored.slice(0,5);
        const bottom = scored.slice(-5).reverse();
        const max = scored[0];
        const min = scored[scored.length-1];
        const range = (max.score - min.score);

        let out = "";
        out += `<div class="small"><b>${escapeHtml(title)}(반별 요약)</b></div>`;
        out += `<div class="small">- 평균점수 범위: ${range.toFixed(2)} (최고: ${max.cls}반 ${max.score.toFixed(2)}, 최저: ${min.cls}반 ${min.score.toFixed(2)})</div>`;
        out += "<div style='overflow:auto;max-height:200px;'><table><thead><tr><th>구분</th><th>반</th><th>평균점수</th><th>좋음</th><th>보통</th><th>나쁨</th></tr></thead><tbody>";
        for (const x of top){ out += `<tr><td>상위</td><td>${x.cls}</td><td>${x.score.toFixed(2)}</td><td>${x.good}</td><td>${x.normal}</td><td>${x.bad}</td></tr>`; }
        for (const x of bottom){ out += `<tr><td>하위</td><td>${x.cls}</td><td>${x.score.toFixed(2)}</td><td>${x.good}</td><td>${x.normal}</td><td>${x.bad}</td></tr>`; }
        out += "</tbody></table></div>";
        return out;
      }

      let html = "";

      // ----- 요약(핵심 지표) -----
      const clsCnt = payload?.meta?.classCount || 0;
      const {cnt, male, female, spec, adhd} = payload.arrays || {};
      function maxDev(arr){
        if(!arr || arr.length===0) return 0;
        const avg = arr.reduce((a,b)=>a+b,0)/arr.length;
        return Math.max(...arr.map(v=>Math.abs(v-avg)));
      }
      const genderDev = (male&&female) ? Math.max(maxDev(male), maxDev(female)) : 0;
      const specDev = maxDev(spec||[]);
      const adhdDev = maxDev(adhd||[]);
      html += `<div class="small"><b>요약</b></div>`;
      html += `<div class="small">- 성비(남/여) 반별 편차(평균 대비): 최대 ${genderDev.toFixed(1)}명</div>`;
      html += `<div class="small">- 특수 반별 편차(평균 대비): 최대 ${specDev.toFixed(1)}명</div>`;
      html += `<div class="small">- ADHD 반별 편차(평균 대비): 최대 ${adhdDev.toFixed(1)}명</div>`;
      if(payload?.meta?.adhdCap && payload.meta.adhdCap !== "auto"){
        // 상한 초과 반 수
        const cap = Number(payload.meta.adhdCap);
        if(!Number.isNaN(cap) && (adhd||[]).length){
          const over = (adhd||[]).filter(v=>v>cap).length;
          html += `<div class="small">- ADHD 반당 최대 ${cap}명 제한: 초과 반 ${over}개</div>`;
        }
      }
      html += `<div style="height:10px"></div>`;
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

      // ----- 추가 리포트: 학부모민원/학업성취/교우관계 (토글 탭) -----
      html += `<div style="height:14px"></div>`;
      html += `<div class="small"><b>3단계 요약(토글)</b></div>`;
      html += `
        <div class="tabs" style="margin-top:10px" id="levelTabs">
          <button type="button" class="tabBtn active" data-target="parent">학부모민원</button>
          <button type="button" class="tabBtn" data-target="acad">학업성취</button>
          <button type="button" class="tabBtn" data-target="peer">교우관계</button>
        </div>
        <div style="height:10px"></div>
        <div id="levelSection-parent" class="levelSection">${buildLevelReport("학부모민원", "학부모민원")}</div>
        <div id="levelSection-acad" class="levelSection" style="display:none">${buildLevelReport("학업성취", "학업성취")}</div>
        <div id="levelSection-peer" class="levelSection" style="display:none">${buildLevelReport("교우관계", "교우관계")}</div>
      `;

      violationsDiv.innerHTML = html;

      // 탭 클릭 이벤트(동적 생성된 DOM에 바인딩)
      try{
        const tabWrap = violationsDiv.querySelector('#levelTabs');
        if(tabWrap){
          const btns = Array.from(tabWrap.querySelectorAll('button.tabBtn'));
          const sections = {
            parent: violationsDiv.querySelector('#levelSection-parent'),
            acad: violationsDiv.querySelector('#levelSection-acad'),
            peer: violationsDiv.querySelector('#levelSection-peer')
          };
          btns.forEach(btn=>{
            btn.addEventListener('click', ()=>{
              const key = btn.getAttribute('data-target');
              btns.forEach(b=>b.classList.toggle('active', b===btn));
              Object.keys(sections).forEach(k=>{
                if(sections[k]) sections[k].style.display = (k===key) ? '' : 'none';
              });
            });
          });
        }
      }catch(e){ console.warn('level tabs init failed', e); }
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
      if(!classFilter) return;
      classFilter.innerHTML = "";
      // classCount가 비정상(0/undefined)일 경우 결과행에서 최대 반 번호로 보정
      let classCount = payload?.meta?.classCount;
      if(!classCount || classCount<1){
        classCount = Math.max(0, ...(payload.resultRows||[]).map(r=>Number(r["반"]||r["새반"]||0))) || 0;
      }
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
      const cntWrap = document.getElementById("cntChart")?.parentElement;
      const gWrap = document.getElementById("genderChart")?.parentElement;
      function showChartMsg(wrap,msg){
        if(!wrap) return;
        // canvas는 유지하되 안내문 추가
        let note = wrap.querySelector(".chartNote");
        if(!note){ note=document.createElement("div"); note.className="chartNote small"; note.style.marginTop="8px"; wrap.appendChild(note); }
        note.textContent = msg;
      }
      if (typeof Chart === "undefined"){
        console.warn("Chart.js not loaded");
        showChartMsg(cntWrap, "그래프 라이브러리(Chart.js)가 로드되지 않아 그래프를 표시할 수 없습니다.");
        showChartMsg(gWrap, "그래프 라이브러리(Chart.js)가 로드되지 않아 그래프를 표시할 수 없습니다.");
        return;
      }
      try{
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

      }catch(err){
        console.error(err);
        const msg = (err && err.message) ? err.message : String(err);
        const cntWrap = document.getElementById('cntChart')?.parentElement;
        const gWrap = document.getElementById('genderChart')?.parentElement;
        const show=(wrap,m)=>{ if(!wrap) return; let note=wrap.querySelector('.chartNote'); if(!note){ note=document.createElement('div'); note.className='chartNote small'; note.style.marginTop='8px'; wrap.appendChild(note);} note.textContent=m; };
        show(cntWrap, '그래프 생성 중 오류: '+msg);
        show(gWrap, '그래프 생성 중 오류: '+msg);
      }
    }

    // 렌더 실행 순서
    renderUnsatisfiedTables();

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
    try{ buildViolationReport(); }catch(err){
      console.error(err);
      const e=document.getElementById("errors");
      if(e) e.textContent = "리포트 생성 중 오류: " + ((err&&err.message)?err.message:String(err));
      if(violationsDiv) violationsDiv.textContent = "리포트를 생성하지 못했습니다.";
    }
    setupFilter();
    renderResultTable("all");
    drawCharts();
  }

})();