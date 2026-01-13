// 반배정 웹앱 v3.2.5 (브라우저 전용)
// - 전역 오염 최소화(IIFE)
// - 결과 탭 테이블/요약 렌더링 버그 수정
// - 함수/변수 중복 선언(Identifier already declared) 오류 수정

console.log('class-assign webapp v3.2.5 loaded');

// 앱 전체를 IIFE로 감싸 전역변수/함수 충돌을 줄입니다.
(()=>
  // ----- Sample template (embedded, for GitHub Pages 안정 다운로드) -----
  const SAMPLE_TEMPLATE_FILENAME = "반배정_양식_샘플(토글적용).xlsx";
  const SAMPLE_XLSX_BASE64 = "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";
  function downloadBase64Xlsx(base64Str, filename){
    const binary = atob(base64Str);
    const bytes = new Uint8Array(binary.length);
    for (let i=0;i<binary.length;i++) bytes[i]=binary.charCodeAt(i);
    const blob = new Blob([bytes], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename || "template.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(()=>URL.revokeObjectURL(url), 1000);
  }

{
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

  function pickBestSheetName(workbook){
    let best = workbook.SheetNames[0];
    let bestRows = -1;
    for (const name of workbook.SheetNames){
      const ws = workbook.Sheets[name];
      const arr = XLSX.utils.sheet_to_json(ws, {defval:"", raw:false});
      if (arr.length > bestRows){
        bestRows = arr.length;
        best = name;
      }
    }
    return best;
  }

  // ----- Robust sheet/header detection -----
  function pickStudentSheetAndHeaderRow(workbook){
    // Returns {sheetName, headerRowIdx} where headerRowIdx is 0-based row index for header.
    const requiredCore = ["학생명","성별"];
    let best = null;

    for (const name of workbook.SheetNames){
      const ws = workbook.Sheets[name];
      const grid = XLSX.utils.sheet_to_json(ws, {header:1, defval:"", raw:false});
      if (!grid || grid.length === 0) continue;

      const maxScan = Math.min(30, grid.length);
      for (let r=0;r<maxScan;r++){
        const row = grid[r] || [];
        const stdHits = new Set();
        for (const cell of row){
          const std = mapToStandardHeader(String(cell||""));
          if (std) stdHits.add(std);
        }
        for (const cell of row){
          const s = String(cell||"");
          for (const std of Object.keys(HEADER_ALIASES)){
            if (normHeader(s)===normHeader(std)) stdHits.add(std);
          }
        }

        const hitCount = stdHits.size;
        const hasCore = requiredCore.every(c=> stdHits.has(c) || row.some(cell=> normHeader(cell||"")===normHeader(c)));
        if (!hasCore) continue;

        const score = hitCount * 10 + (grid.length - r); // prefer more headers + earlier in sheet
        if (!best || score > best.score){
          best = {sheetName:name, headerRowIdx:r, score};
        }
      }
    }

    if (best) return best;

    // fallback: 기존 방식(가장 행 많은 시트, 0행을 헤더로)
    return {sheetName: pickBestSheetName(workbook), headerRowIdx: 0, score:0};
  }


  // ----- DOM refs -----
  const fileInput = document.getElementById("fileInput");
  const templateDownloadBtn = document.getElementById("templateDownload");
  templateDownloadBtn?.addEventListener("click", (ev)=>{
    ev.preventDefault();
    downloadBase64Xlsx(SAMPLE_XLSX_BASE64, SAMPLE_TEMPLATE_FILENAME);
  });

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
    // care
    if (strength === 'strong') return 5000;
    if (strength === 'medium') return 2000;
    return 700;
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
      runBtn.disabled = true;
      return;
    }
    filePill.textContent = `선택됨: ${file.name}`;

    showOverlay(true, "엑셀을 읽는 중…");
    await new Promise(r=>setTimeout(r, 10));

    try{
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, {type:"array"});
      const pick = pickStudentSheetAndHeaderRow(wb);
      const bestName = pick.sheetName;
      const ws = wb.Sheets[bestName];
      // headerRowIdx가 0이 아니면 해당 행을 헤더로 사용합니다.
      rawRows = XLSX.utils.sheet_to_json(ws, {defval:"", range: pick.headerRowIdx});
      if (!rawRows || rawRows.length === 0){
        setErrors("엑셀에서 데이터를 읽지 못했어요. 데이터가 있는 시트/행(헤더)이 맞는지 확인해주세요.");
        runBtn.disabled = true;
        return;
      }

      // (선택) 헤더 매핑 점검 - 현재는 normalizeRow가 유연하게 읽음
      const headers = Object.keys(rawRows[0] || {});
      const headerMap = buildHeaderMap(headers);
      const missingStd = REQUIRED_COLUMNS.filter(c=> !headerMap[c] && !headers.some(h=>normHeader(h)===normHeader(c)));
      // 학생명/성별은 validateRows로 다시 확인

      studentRows = rawRows.map(normalizeRow);
      const err = validateRows(studentRows);
      if (err){
        setErrors(err);
        runBtn.disabled = true;
        return;
      }

      rowsPill.textContent = `${studentRows.length}명`;
      renderSetupPreviewTable(rawRows, 20);
      summarize(studentRows);
      statusPill.textContent = missingStd.length ? `엑셀 로드됨(권장열 누락: ${missingStd.join(', ')})` : "엑셀 로드됨";
      runBtn.disabled = !(studentRows && studentRows.length>0);
    } catch(err){
      console.error(err);
      setErrors("엑셀을 읽는 중 오류: " + (err?.message || String(err)));
      runBtn.disabled = true;
    } finally{
      showOverlay(false);
    }
  });

  // ----- Run optimization -----
  runBtn?.addEventListener("click", async ()=>{
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

    showOverlay(true, "코드 그룹(분리/배려)을 구성하는 중…");
    await new Promise(r=>setTimeout(r, 10));
    const groups = buildCodeGroups(studentRows);

    showOverlay(true, "시뮬레이션을 시작합니다…");
    const start = performance.now();
    const {best, bestAssign} = await optimize(studentRows, classCount, iterations, seed, weights, groups);
    const elapsedMs = Math.round(performance.now() - start);

    const resultRows = studentRows.map((r,i)=>( {
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

    // 다운로드/공유용 결과는 "새로운 반" 기준으로 1반 → 2반 ... 순서로 정렬
    resultRows.sort((a,b)=> (a.반 - b.반) || String(a.학생명).localeCompare(String(b.학생명), 'ko'));

    // 다운로드/표시용: 새 반(1반→) 기준으로 정렬
    resultRows.sort((a,b)=> (a["반"]-b["반"]) || String(a["학생명"]).localeCompare(String(b["학생명"])) );

    const payload = {
      meta: { total: studentRows.length, classCount, iterations, seed, elapsedMs, weights, sepStrength: sepStrengthEl.value, careStrength: careStrengthEl.value },
      best: { score: best.score, sepViol: best.sepViol, careMiss: best.careMiss },
      arrays: { cnt: best.cnt, male: best.male, female: best.female, spec: best.spec, adhd: best.adhd },
      resultRows
    };

    showOverlay(false);
    tabResultBtn.disabled = false;
    statusPill.textContent = "완료";
    try{ renderResult(payload); }catch(e){ console.error(e); setErrors("결과 렌더링 오류: " + (e?.message || e)); }
    showTab("result");
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
    sepPill.textContent = `분리 위반: ${payload.best.sepViol.toLocaleString()}쌍`;
    carePill.textContent = `배려 미충족: ${payload.best.careMiss.toLocaleString()}쌍`;

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
        const sepCodes = (r["분리요청코드"]||"").split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
        const careCodes = (r["배려요청코드"]||"").split(/[,\s;]+/).map(t=>t.trim()).filter(Boolean);
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
    };

    renderClassSummary();
    buildViolationReport();
    setupFilter();
    renderResultTable("all");
    drawCharts();
  }

})();
