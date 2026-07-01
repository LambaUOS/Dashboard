/* ============================================================================
 *  탄소흡수 산정 에이전트 (Carbon Sequestration Agent) — 브라우저 내장형
 *  ----------------------------------------------------------------------------
 *  조경 유지관리 내역서(xlsx) → 수목 데이터 추출 → 수종분류 →
 *  DBH 상대생장식(W=a·D^b) 기반 연간 CO₂ 흡수량 산정 → 순공사비 산출 →
 *  검증(이상치/누락) → 표·차트·PDF 보고서 → 자연어 챗봇.
 *
 *  설계 원칙(명세 준수):
 *   - 계산식은 투명·정확. 화학상수 44/12, IPCC 기본 CF=0.5는 확정값.
 *   - 수종별 계수(a,b,R,생장률 등)는 "검증 전 기본값"으로 명시하고
 *     채팅에서 직접 수정·확정(HITL). 임의추정 금지 → 자료 없으면 "확인 필요".
 *   - 가정·예외·최종 확정은 사용자 승인(Human-in-the-Loop).
 *
 *  의존성: ../vendor/xlsx.full.min.js (SheetJS), ./assets/inventory-data.js (내장 xlsx)
 *  외부 서버 불필요 — 100% 브라우저/오프라인 동작. (LLM은 Ollama 있으면 사용, 없으면 규칙기반)
 * ========================================================================== */
(function () {
  'use strict';

  // ===========================================================================
  // 0. 유틸리티
  // ===========================================================================
  const $id = (s) => document.getElementById(s);
  const el = (tag, cls, html) => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (html != null) e.innerHTML = html;
    return e;
  };
  const norm = (v) => (v == null ? '' : String(v)).replace(/\s+/g, '').trim();
  const round2 = (v) => (v == null || isNaN(v) ? null : Math.round(v * 100) / 100);
  const numOrNull = (v) => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/,/g, '').trim();
    if (s === '' || /#REF!|#N\/A|#VALUE!/i.test(s)) return null;
    const n = Number(s);
    return isFinite(n) ? n : null;
  };
  const fmt = (v, d = 2) =>
    v == null || isNaN(v)
      ? '—'
      : Number(v).toLocaleString('ko-KR', { minimumFractionDigits: d, maximumFractionDigits: d });
  const fmtInt = (v) => (v == null || isNaN(v) ? '—' : Number(v).toLocaleString('ko-KR'));
  const esc = (s) =>
    String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  const CO2_C = 44 / 12; // 탄소 → CO₂ 분자량비 (확정 상수)

  // 익명화(필터링) 마스크 — 순공사비/단가성 금액·수량·녹지면적 등 민감 수치는
  // 결과 표·PDF·챗봇 응답에서 모두 '***'로 가린다. 단가 컬럼은 아예 노출하지 않는다.
  const MASK = '***';

  // ===========================================================================
  // 1. 도메인 지식 — 수종 사전 & 탄소계수 DB (모두 수정 가능, 출처/신뢰도 표기)
  // ===========================================================================
  // 수종 사전: 성상 추론 보조 (성상 컬럼이 비어있을 때 사용)
  const SPECIES_DICT = {
    소나무:   { form: '교목', leaf: '침엽', ever: '상록', sci: 'Pinus densiflora' },
    곰솔:     { form: '교목', leaf: '침엽', ever: '상록', sci: 'Pinus thunbergii' },
    잣나무:   { form: '교목', leaf: '침엽', ever: '상록', sci: 'Pinus koraiensis' },
    홍가시나무:{ form: '교목', leaf: '활엽', ever: '상록', sci: 'Photinia glabra' },
    은목서:   { form: '교목', leaf: '활엽', ever: '상록', sci: 'Osmanthus asiaticus' },
    먼나무:   { form: '교목', leaf: '활엽', ever: '상록', sci: 'Ilex rotunda' },
    느티나무: { form: '교목', leaf: '활엽', ever: '낙엽', sci: 'Zelkova serrata' },
    이팝나무: { form: '교목', leaf: '활엽', ever: '낙엽', sci: 'Chionanthus retusus' },
    벚나무:   { form: '교목', leaf: '활엽', ever: '낙엽', sci: 'Prunus serrulata' },
    단풍나무: { form: '교목', leaf: '활엽', ever: '낙엽', sci: 'Acer palmatum' },
    철쭉:     { form: '관목', leaf: '활엽', ever: '낙엽', sci: 'Rhododendron schlippenbachii' },
    명자:     { form: '관목', leaf: '활엽', ever: '낙엽', sci: 'Chaenomeles speciosa' },
    영산홍:   { form: '관목', leaf: '활엽', ever: '상록', sci: 'Rhododendron indicum' },
    회양목:   { form: '관목', leaf: '활엽', ever: '상록', sci: 'Buxus' },
  };

  // 탄소계수 DB (기본 가정값 — 검증 전). 사용자가 채팅 계수편집기에서 수정/확정.
  //  교목: 지상부 바이오매스 W(kg) = a · D^b  (D = 흉고직경 DBH, cm)
  //        총바이오매스 = W·(1+R),  탄소 = ×CF,  CO₂ = ×44/12
  //        연간흡수량 = [W(D+g) − W(D)]·(1+R)·CF·(44/12),  g=연간직경생장(cm/yr)
  //  관목: 면적기반 — 연간CO₂(kg/yr) = 면적(㎡) · areaCoeff(kgCO₂/㎡/yr)
  const DEFAULT_COEFF = {
    trees: {
      상록침엽: { a: 0.0613, b: 2.55, R: 0.25, CF: 0.5, g: 0.5,
        src: '기본 가정값(검증 전) — 국립산림과학원 침엽수 상대생장식으로 교체 권장', conf: 'low' },
      상록활엽: { a: 0.1115, b: 2.53, R: 0.27, CF: 0.5, g: 0.5,
        src: '기본 가정값(검증 전) — 상록활엽 유사종 식 적용, 검증 필요', conf: 'low' },
      낙엽침엽: { a: 0.0590, b: 2.56, R: 0.24, CF: 0.5, g: 0.6,
        src: '기본 가정값(검증 전)', conf: 'low' },
      낙엽활엽: { a: 0.1120, b: 2.54, R: 0.26, CF: 0.5, g: 0.6,
        src: '기본 가정값(검증 전) — 활엽수 일반식, 검증 필요', conf: 'low' },
    },
    shrubs: {
      관목: { areaCoeff: null, CF: 0.5,
        src: '자료 없음 — 면적기반 관목 탄소계수 미공표. 사용자 확인 필요', conf: 'none' },
    },
  };

  // ===========================================================================
  // 2. 전역 상태
  // ===========================================================================
  const STATE = {
    workbook: null,
    sheets: [],          // [{name, rows}]
    parsed: null,        // {treeRecords, areaRecords, unitPrices, maintRows, logs}
    classified: null,    // [{...record, cls}]
    results: null,       // {bySpecies, totals, lines}
    cost: null,          // {lines, total, unmatched}
    validation: null,    // {outliers, missing, warnings}
    report: null,        // 명세 출력 스키마 JSON
    coeff: structuredCloneSafe(DEFAULT_COEFF),
    coeffApproved: false,
    baseYear: 2026,
    ai: { apiKey: '', model: 'gpt-4o-mini', baseUrl: 'https://api.openai.com/v1', busy: false },
  };
  function structuredCloneSafe(o) { return JSON.parse(JSON.stringify(o)); }

  // ===========================================================================
  // 3. 도구(Tools) — 결정론적 엔진
  // ===========================================================================

  // ---- 3.1 parseExcelTool : 워크북 → 구조화 데이터 -------------------------
  function parseExcel(workbook) {
    const logs = [];
    const sheets = workbook.SheetNames.map((name) => ({
      name,
      rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: '' }),
    }));
    STATE.sheets = sheets;

    // 시트 식별: 이름/헤더에 키워드 우선순위
    const findSheet = (kw) =>
      sheets.find((s) => kw.some((k) => norm(s.name).includes(k)));

    const treeSheet =
      findSheet(['수목']) ||
      sheets.find((s) => s.rows.some((r) => r.some((c) => norm(c) === '수목명')));
    const greenSheet = findSheet(['녹지']) || findSheet(['식재']);
    const priceSheet = findSheet(['일위대가']) || findSheet(['단가']);

    const treeRecords = treeSheet ? parseTreeSheet(treeSheet, logs) : [];
    const unitPrices = priceSheet ? parsePriceSheet(priceSheet, logs) : {};
    const maintRows = []
      .concat(treeSheet ? parseMaintSchedule(treeSheet, logs) : [])
      .concat(greenSheet ? parseMaintSchedule(greenSheet, logs) : []);
    const areaRecords = greenSheet ? parseAreaSheet(greenSheet, logs) : [];

    if (!treeSheet) logs.push('⚠ 수목 집계표 시트를 찾지 못했습니다.');
    else logs.push(`시트 식별: 수목="${treeSheet.name}"` +
      (priceSheet ? `, 단가="${priceSheet.name}"` : '') +
      (greenSheet ? `, 녹지="${greenSheet.name}"` : ''));

    STATE.parsed = { treeRecords, areaRecords, unitPrices, maintRows, logs,
      sheetNames: sheets.map((s) => s.name) };
    return STATE.parsed;
  }

  // 수목 집계표 파싱 — 성상 섹션 상속 + 소계/합계 분리
  function parseTreeSheet(sheet, logs) {
    const rows = sheet.rows;
    // 헤더행: '수목명'+'규격' 동시 포함
    let hi = -1;
    for (let i = 0; i < rows.length; i++) {
      const set = rows[i].map(norm);
      if (set.includes('수목명') && set.includes('규격')) { hi = i; break; }
    }
    if (hi < 0) { logs.push('⚠ 수목 시트 헤더(수목명/규격)를 찾지 못함'); return []; }
    const H = rows[hi].map(norm);
    const col = (label) => H.indexOf(label);
    const cDiv = col('구분'), cForm = col('성상'), cName = col('수목명'),
      cSpec = col('규격'), cUnit = col('단위'),
      cQtyT = col('수목수량') >= 0 ? col('수목수량') : col('수량'),
      cQtyM = col('유지관리수량'), cStd = col('유지관리기준'), cCalc = col('산출');

    const SUBTOTAL = new Set(['소계', '합계', '소 계', '합 계']);
    const FORM_MAP = (s) => {
      const n = norm(s);
      if (!n) return null;
      if (n.includes('관목')) return { form: '관목', ever: n.includes('상록') ? '상록' : (n.includes('낙엽') ? '낙엽' : null) };
      if (n.includes('교목') || n.includes('수')) {
        return { form: '교목',
          ever: n.includes('상록') ? '상록' : (n.includes('낙엽') ? '낙엽' : null),
          leaf: n.includes('침엽') ? '침엽' : (n.includes('활엽') ? '활엽' : null) };
      }
      return null;
    };

    const out = [];
    let section = null;
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i];
      const formCell = cForm >= 0 ? r[cForm] : '';
      const nameRaw = cName >= 0 ? r[cName] : '';
      const name = norm(nameRaw);
      const fm = FORM_MAP(formCell);
      if (fm) section = fm; // 섹션 헤더 → 현재 성상 갱신
      if (!name || SUBTOTAL.has(name) || SUBTOTAL.has(String(nameRaw).trim())) continue; // 소계/합계/빈행
      const spec = cSpec >= 0 ? String(r[cSpec] ?? '').trim() : '';
      const unit = cUnit >= 0 ? String(r[cUnit] ?? '').trim() : '';
      const qty = numOrNull(cQtyT >= 0 ? r[cQtyT] : null) ?? numOrNull(cQtyM >= 0 ? r[cQtyM] : null);
      if (!spec && qty == null) continue; // 데이터 없는 잔여행
      out.push({
        구분: cDiv >= 0 ? String(r[cDiv] ?? '').trim() : '',
        성상: section ? (section.ever ? section.ever + (section.form === '관목' ? '관목' : (section.leaf ? section.leaf + '교목' : '교목')) : section.form) : '',
        section,
        수목명: String(nameRaw).trim(),
        규격: spec,
        단위: unit,
        수량: qty,
        유지관리기준: cStd >= 0 ? String(r[cStd] ?? '').trim() : '',
        row: i + 1,
      });
    }
    logs.push(`수목 추출: ${out.length}개 항목 (행 ${hi + 2}~)`);
    return out;
  }

  // 일위대가 시트 → 단가 맵
  function parsePriceSheet(sheet, logs) {
    const rows = sheet.rows;
    let hi = -1, cSpec = -1;
    for (let i = 0; i < rows.length; i++) {
      const set = rows[i].map(norm);
      const s = set.indexOf('규격');
      if (s >= 0 && (set.includes('단위') || set.includes('합계'))) { hi = i; cSpec = s; break; }
    }
    if (hi < 0) { logs.push('⚠ 일위대가 헤더를 찾지 못함'); return {}; }
    const H = rows[hi].map(norm);
    const cName = cSpec - 1;
    const cUnit = H.indexOf('단위') >= 0 ? H.indexOf('단위') : cSpec + 1;
    const cSum = H.indexOf('합계') >= 0 ? H.indexOf('합계') : cSpec + 3;
    const cLabor = H.indexOf('노무비'), cMat = H.indexOf('재료비'), cExp = H.indexOf('경비');
    const map = {};
    let n = 0;
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = String(r[cName] ?? '').trim();
      if (!name) continue;
      const spec = String(r[cSpec] ?? '').trim();
      const sum = numOrNull(r[cSum]);
      const rec = { name, spec, unit: String(r[cUnit] ?? '').trim(), sum,
        labor: numOrNull(r[cLabor]), mat: numOrNull(r[cMat]), exp: numOrNull(r[cExp]) };
      map[norm(name) + '|' + norm(spec)] = rec;
      (map.__byName__ = map.__byName__ || {})[norm(name)] =
        (map.__byName__[norm(name)] || []).concat(rec);
      n++;
    }
    logs.push(`일위대가 단가: ${n}개 항목`);
    return map;
  }

  // 유지관리 작업표(우측 표) 파싱 — 시트별 컬럼 위치 자동 탐지
  function parseMaintSchedule(sheet, logs) {
    const rows = sheet.rows;
    let hi = -1, c = -1;
    for (let i = 0; i < rows.length; i++) {
      const set = rows[i].map(norm);
      for (let j = 10; j < set.length; j++) { // 우측(컬럼 10+)의 '규격'
        if (set[j] === '규격') { hi = i; c = j; break; }
      }
      if (hi >= 0) break;
    }
    if (hi < 0) return [];
    const cName = c - 1, cFreq = c + 1, cUnit = c + 2, cQty = c + 3;
    const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = String(r[cName] ?? '').trim();
      const qty = numOrNull(r[cQty]);
      if (!name || qty == null) continue;
      out.push({ name, spec: String(r[c] ?? '').trim(),
        freq: numOrNull(r[cFreq]), unit: String(r[cUnit] ?? '').trim(), qty, sheet: sheet.name });
    }
    return out;
  }

  // 녹지 면적 시트 (참고용 — 현재 산정에는 미사용, 향후 면적기반 확장 대비)
  function parseAreaSheet(sheet, logs) {
    return [];
  }

  // ---- 3.2 classifySpeciesTool : 수종 분류 + 규격 파싱 ---------------------
  function classify(rec) {
    const dict = SPECIES_DICT[norm(rec.수목명)] || SPECIES_DICT[rec.수목명] || null;
    const sec = rec.section || {};
    const form = sec.form || (dict && dict.form) || guessFormFromSpec(rec.규격);
    const ever = sec.ever || (dict && dict.ever) || null;
    const leaf = sec.leaf || (dict && dict.leaf) || null;
    const spec = parseSpec(rec.규격);
    // 계수 키
    let coeffKey = null, coeffGroup = 'trees';
    if (form === '관목') { coeffGroup = 'shrubs'; coeffKey = '관목'; }
    else {
      const e = ever || '상록';
      const l = leaf || (e === '상록' ? '침엽' : '활엽');
      coeffKey = e + l; // 상록침엽 등
    }
    return {
      sci: dict ? dict.sci : null,
      form: form || '교목', ever, leaf,
      spec, coeffGroup, coeffKey,
      assumedForm: !sec.form && !(dict && dict.form),
      assumedEver: !sec.ever && !(dict && dict.ever),
    };
  }
  function guessFormFromSpec(s) {
    const t = parseSpec(s);
    if (t.type === 'H' && (t.unitHint === 'm2' || (t.mid != null && t.mid < 1.5))) return '관목';
    if (t.type === 'B' || t.type === 'R') return '교목';
    return null;
  }
  // 규격 파싱: B21~31 / B11~21 / H0.5 / W1.0 / R15 등 → {type, min, max, mid}
  function parseSpec(s) {
    const raw = String(s || '').trim();
    const m = raw.match(/([BHWR])\s*([\d.]+)\s*(?:[~\-－]\s*([\d.]+))?/i);
    if (!m) return { type: null, raw, min: null, max: null, mid: null };
    const type = m[1].toUpperCase();
    const min = parseFloat(m[2]);
    const max = m[3] != null ? parseFloat(m[3]) : null;
    const mid = max != null ? (min + max) / 2 : min;
    return { type, raw, min, max, mid };
  }

  // ---- 3.3 searchCoefficientTool : 수종·규격별 계수 조회 -------------------
  function searchCoefficient(cls) {
    const group = STATE.coeff[cls.coeffGroup] || {};
    const c = group[cls.coeffKey];
    if (!c) return { ok: false, key: cls.coeffKey, reason: '계수 없음 — 확인 필요' };
    return { ok: true, key: cls.coeffKey, group: cls.coeffGroup, coeff: c };
  }

  // ---- 3.4 calculateCO2Tool : 흡수량 계산 ---------------------------------
  function calcRecord(rec, cls, coeffRes) {
    const out = { co2Annual: null, co2Stock: null, biomass: null, perUnit: null, method: null, note: [] };
    if (!coeffRes.ok || !coeffRes.coeff) { out.note.push('계수 없음'); return out; }
    const c = coeffRes.coeff;
    if (cls.coeffGroup === 'trees') {
      const D = cls.spec.mid;
      if (D == null || cls.spec.type !== 'B') {
        // 흉고직경(B) 아닌 경우(R 근원직경 등) → 환산 가정 필요
        if (cls.spec.type === 'R' && cls.spec.mid != null) {
          // 근원직경→흉고직경 근사 (R ≈ 1.2~1.3·B). 가정.
          out.note.push('근원직경(R)→흉고직경 환산 가정(B≈R/1.25)');
        } else { out.note.push('흉고직경(B) 규격 아님 — 계산 불가'); return out; }
      }
      const Dbh = cls.spec.type === 'R' ? cls.spec.mid / 1.25 : D;
      const W = (d) => c.a * Math.pow(d, c.b);          // 지상부 바이오매스 kg
      const total = (d) => W(d) * (1 + (c.R || 0)) * (c.CF || 0.5); // 탄소 kg
      const stockC = total(Dbh);
      const annualC = total(Dbh + (c.g || 0)) - total(Dbh);
      out.biomass = round2(W(Dbh) * (1 + (c.R || 0)));
      out.perUnit = round2(annualC * CO2_C);            // kgCO₂/주/년
      out.co2Annual = round2(annualC * CO2_C * (rec.수량 || 0));
      out.co2Stock = round2(stockC * CO2_C * (rec.수량 || 0));
      out.Dbh = Dbh;
      out.method = `상대생장식 W=${c.a}·D^${c.b}, R=${c.R}, CF=${c.CF}, g=${c.g}cm/yr`;
    } else {
      // 관목: 면적기반
      if (c.areaCoeff == null) { out.note.push('관목 면적계수 미입력 — 확인 필요'); return out; }
      out.perUnit = round2(c.areaCoeff);
      out.co2Annual = round2(c.areaCoeff * (rec.수량 || 0));
      out.co2Stock = null;
      out.method = `면적기반 ${c.areaCoeff} kgCO₂/㎡/yr`;
    }
    return out;
  }

  // 전체 계산 파이프라인
  function runCalculations() {
    const recs = STATE.parsed.treeRecords;
    const lines = recs.map((rec) => {
      const cls = classify(rec);
      const coeffRes = searchCoefficient(cls);
      const calc = calcRecord(rec, cls, coeffRes);
      return { rec, cls, coeffRes, calc };
    });
    STATE.classified = lines;

    // 수종별 집계
    const bySpeciesMap = {};
    for (const L of lines) {
      const key = L.rec.수목명 + ' (' + (L.cls.spec.raw || L.rec.규격) + ')';
      const k = bySpeciesMap[key] || (bySpeciesMap[key] = {
        name: L.rec.수목명, spec: L.cls.spec.raw || L.rec.규격, unit: L.rec.단위,
        form: L.cls.form, ever: L.cls.ever, leaf: L.cls.leaf, sci: L.cls.sci,
        qty: 0, co2Annual: 0, co2Stock: 0, anyAnnual: false, anyStock: false,
        coeffOk: L.coeffRes.ok, note: new Set(),
      });
      k.qty += L.rec.수량 || 0;
      if (L.calc.co2Annual != null) { k.co2Annual += L.calc.co2Annual; k.anyAnnual = true; }
      if (L.calc.co2Stock != null) { k.co2Stock += L.calc.co2Stock; k.anyStock = true; }
      L.calc.note.forEach((n) => k.note.add(n));
      if (!L.coeffRes.ok) k.coeffOk = false;
    }
    // 사용 가능한 흡수량을 산출하지 못한 수종(예: 관목 면적계수 미입력)은
    // 0이 아니라 null(확인필요)로 표기하고 합계에서 제외(미확정).
    const bySpecies = Object.values(bySpeciesMap).map((s) => ({
      ...s, note: [...s.note],
      co2Annual: s.anyAnnual ? round2(s.co2Annual) : null,
      co2Stock: s.anyStock ? round2(s.co2Stock) : null,
      hasCoeff: s.coeffOk && s.anyAnnual,
    }));

    const confirmed = bySpecies.filter((s) => s.hasCoeff && s.co2Annual != null);
    const pending = bySpecies.filter((s) => !s.hasCoeff || s.co2Annual == null);
    const totalAnnualKg = confirmed.reduce((a, s) => a + (s.co2Annual || 0), 0);
    const totalStockKg = confirmed.reduce((a, s) => a + (s.co2Stock || 0), 0);
    const treeCount = bySpecies.filter((s) => s.form === '교목').reduce((a, s) => a + s.qty, 0);

    STATE.results = {
      bySpecies, confirmed, pending,
      totals: {
        annual_tCO2: round2(totalAnnualKg / 1000),
        annual_kgCO2: round2(totalAnnualKg),
        stock_tCO2: round2(totalStockKg / 1000),
        treeCount, speciesCount: bySpecies.length,
      },
    };
    return STATE.results;
  }

  // ---- 3.5 computeCost : 순공사비 산출 (일위대가 매칭) ---------------------
  function computeCost() {
    const up = STATE.parsed.unitPrices || {};
    const lines = [];
    let total = 0, matched = 0, unmatched = 0;
    for (const m of STATE.parsed.maintRows) {
      let rec = up[norm(m.name) + '|' + norm(m.spec)];
      if (!rec) { // 작업명만으로 근사 매칭(규격 부분일치)
        const cands = (up.__byName__ || {})[norm(m.name)] || [];
        rec = cands.find((r) => norm(m.spec) && (norm(r.spec).includes(norm(m.spec)) || norm(m.spec).includes(norm(r.spec)))) || cands[0];
      }
      const unitPrice = rec ? rec.sum : null;
      const cost = unitPrice != null ? unitPrice * m.qty : null;
      if (cost != null) { total += cost; matched++; } else unmatched++;
      lines.push({ ...m, unitPrice, cost, matchedSpec: rec ? rec.spec : null });
    }
    lines.sort((a, b) => (b.cost || 0) - (a.cost || 0));
    STATE.cost = { lines, total: round2(total), matched, unmatched };
    return STATE.cost;
  }

  // ---- 3.6 validateDataTool : 이상치(±2σ) + 누락 점검 ----------------------
  function validate() {
    const warnings = [], missing = [], outliers = [];
    const lines = STATE.classified || [];
    for (const L of lines) {
      if (L.rec.수량 == null) missing.push(`${L.rec.수목명}(${L.rec.규격}): 수량 누락`);
      if (!L.cls.spec.type) missing.push(`${L.rec.수목명}: 규격 패턴(B/H/W/R) 인식 불가 — "${L.rec.규격}"`);
      if (!L.coeffRes.ok) warnings.push(`${L.rec.수목명}(${L.cls.coeffKey}): 탄소계수 없음 — 확인 필요`);
      if (L.cls.assumedForm) warnings.push(`${L.rec.수목명}: 성상 미상 → '${L.cls.form}'로 가정`);
      L.calc.note.forEach((n) => warnings.push(`${L.rec.수목명}: ${n}`));
    }
    // 이상치: 주당 흡수량(perUnit) ±2σ
    const vals = lines.filter((L) => L.calc.perUnit != null && L.cls.coeffGroup === 'trees')
      .map((L) => ({ L, v: L.calc.perUnit }));
    if (vals.length >= 3) {
      const mean = vals.reduce((a, x) => a + x.v, 0) / vals.length;
      const sd = Math.sqrt(vals.reduce((a, x) => a + (x.v - mean) ** 2, 0) / vals.length);
      for (const x of vals) {
        if (sd > 0 && Math.abs(x.v - mean) > 2 * sd) {
          outliers.push({ name: x.L.rec.수목명, spec: x.L.rec.규격, perUnit: x.v,
            z: round2((x.v - mean) / sd) });
        }
      }
    }
    STATE.validation = { warnings: [...new Set(warnings)], missing: [...new Set(missing)], outliers,
      stats: { records: lines.length } };
    return STATE.validation;
  }

  // ---- 3.7 generateReport : 명세 출력 스키마 ------------------------------
  function buildReport() {
    const R = STATE.results, C = STATE.cost, V = STATE.validation;
    const assumptions = [];
    const usedKeys = new Set(STATE.classified.map((L) => L.cls.coeffGroup + '/' + L.cls.coeffKey));
    usedKeys.forEach((k) => {
      const [g, key] = k.split('/');
      const c = (STATE.coeff[g] || {})[key];
      if (!c) { assumptions.push({ 항목: key, 내용: '계수 없음 — 사용자 확인 필요', 출처: '—', 신뢰도: 'none' }); return; }
      if (g === 'trees') assumptions.push({ 항목: `${key} 상대생장식`,
        내용: `W=${c.a}·D^${c.b}, 뿌리비R=${c.R}, 탄소율CF=${c.CF}, 연간직경생장 g=${c.g}cm/yr`,
        출처: c.src, 신뢰도: c.conf });
      else assumptions.push({ 항목: `${key} 면적계수`,
        내용: c.areaCoeff == null ? '미입력' : `${c.areaCoeff} kgCO₂/㎡/yr`, 출처: c.src, 신뢰도: c.conf });
    });
    assumptions.push({ 항목: '연간 흡수량 정의', 내용: '증분법: [W(D+g)−W(D)]·(1+R)·CF·44/12. 직경생장 가정 적용',
      출처: 'IPCC 2006 GPG 방식', 신뢰도: 'method' });
    assumptions.push({ 항목: 'CO₂ 환산', 내용: '탄소×44/12 = 3.6667', 출처: '분자량비(확정)', 신뢰도: 'high' });
    assumptions.push({ 항목: '순공사비', 내용: '일위대가 단가 × 유지관리수량 합. 횟수는 수량에 반영된 것으로 가정',
      출처: '내역서 일위대가 시트', 신뢰도: 'high' });

    STATE.report = {
      summary: {
        total_tCO2_per_year: R.totals.annual_tCO2,
        total_tCO2_stock: R.totals.stock_tCO2,
        tree_count: R.totals.treeCount,
        species_count: R.totals.speciesCount,
        net_cost_KRW: C ? C.total : null,
        base_year: STATE.baseYear,
      },
      by_species: R.bySpecies.map((s) => ({
        수목명: s.name, 규격: s.spec, 성상: (s.ever || '') + (s.leaf || '') + s.form,
        수량: s.qty, 단위: s.unit,
        '연간흡수_kgCO2/yr': s.co2Annual, '저장량_kgCO2': s.co2Stock,
        계수확인: s.hasCoeff ? 'OK' : '확인필요',
      })),
      assumptions,
      warnings: (V ? V.warnings.concat(V.missing.map((m) => '누락: ' + m)) : [])
        .concat(R.pending.length ? [`미확정(계수없음) 수종 ${R.pending.length}건은 합계 제외`] : []),
    };
    return STATE.report;
  }

  // 익명화된 요약/보고서 — GPT 도구 반환·JSON 출력에 사용(원시 수치 노출 방지).
  function maskedSummary() {
    if (!STATE.report) return {};
    return { ...STATE.report.summary, net_cost_KRW: MASK, tree_count: MASK };
  }
  function maskedReport() {
    if (!STATE.report) return {};
    return {
      ...STATE.report,
      summary: maskedSummary(),
      by_species: STATE.report.by_species.map((s) => ({ ...s, 수량: MASK })),
    };
  }

  // ===========================================================================
  // 4. 차트 (SVG, 무의존)
  // ===========================================================================
  const PALETTE = ['#4a8bff', '#7a4aff', '#22d3a0', '#ffb84a', '#ff6b9d', '#4ad6ff', '#b388ff', '#ff8a8a'];
  function donutSVG(data, opts) {
    const items = data.filter((d) => d.value > 0);
    const total = items.reduce((a, d) => a + d.value, 0) || 1;
    const cx = 70, cy = 70, r = 52, sw = 22;
    let acc = 0;
    const C = 2 * Math.PI * r;
    const segs = items.map((d, i) => {
      const frac = d.value / total;
      const dash = `${(frac * C).toFixed(2)} ${(C - frac * C).toFixed(2)}`;
      const off = (-acc * C).toFixed(2);
      acc += frac;
      return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${d.color || PALETTE[i % PALETTE.length]}" stroke-width="${sw}" stroke-dasharray="${dash}" stroke-dashoffset="${off}" transform="rotate(-90 ${cx} ${cy})"></circle>`;
    }).join('');
    const legend = items.map((d, i) =>
      `<div class="ca-leg"><span class="ca-dot" style="background:${d.color || PALETTE[i % PALETTE.length]}"></span>${esc(d.label)} <b>${fmt(d.value)}</b>${d.unit || ''}</div>`).join('');
    return `<div class="ca-chartwrap"><svg width="140" height="140" viewBox="0 0 140 140">${segs}
      <text x="${cx}" y="${cy - 2}" text-anchor="middle" font-size="15" font-weight="700" fill="#fff">${esc(opts && opts.center || '')}</text>
      <text x="${cx}" y="${cy + 14}" text-anchor="middle" font-size="9" fill="#8892a6">${esc(opts && opts.sub || '')}</text></svg>
      <div class="ca-legend">${legend}</div></div>`;
  }
  function hbarSVG(data, opts) {
    const max = Math.max(...data.map((d) => d.value), 1);
    const rows = data.map((d, i) => {
      const w = Math.max(2, (d.value / max) * 100);
      return `<div class="ca-bar-row"><div class="ca-bar-lbl">${esc(d.label)}</div>
        <div class="ca-bar-track"><div class="ca-bar-fill" style="width:${w}%;background:${d.color || PALETTE[i % PALETTE.length]}"></div></div>
        <div class="ca-bar-val">${fmt(d.value)}${opts && opts.unit || ''}</div></div>`;
    }).join('');
    return `<div class="ca-hbar">${rows}</div>`;
  }

  // ===========================================================================
  // 5. AI 에이전트 — OpenAI GPT 함수호출(Function Calling) 기반
  // ===========================================================================
  //  진짜 에이전트 루프: 사용자 메시지 → GPT가 도구(tool) 선택·호출 → 로컬에서 실행
  //  (표·차트·PDF·테이블창 렌더 등) → 결과를 GPT에 반환 → 최종 답변. LangChain 없이 직접 구현.
  const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : v; } catch (e) { return d; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch (e) {} };
  const j = (o) => JSON.stringify(o);
  const AGENT_HISTORY = []; // 대화 맥락(assistant/tool/user)

  // ---- OpenAI 함수호출 스키마 ----
  const TOOLS = [
    { type: 'function', function: { name: 'run_full_analysis',
      description: '엑셀 수목 데이터로 전체 파이프라인(추출→분류→DBH 상대생장식 CO₂ 산정→순공사비→검증→보고서)을 실행하고 요약·표·차트를 채팅에 렌더한다. "분석/산정/전체"류 요청에 사용.',
      parameters: { type: 'object', properties: {}, required: [] } } },
    { type: 'function', function: { name: 'get_co2_summary',
      description: '연간 CO₂ 흡수량·탄소저장량·교목수·수종수·순공사비 요약 KPI를 렌더하고 반환.',
      parameters: { type: 'object', properties: {}, required: [] } } },
    { type: 'function', function: { name: 'filter_inventory',
      description: '수목 인벤토리를 조건으로 필터링해 표·막대차트로 렌더하고 행을 반환. 예: 교목만, 관목만, 상록수, 소나무, 흉고직경 B21 이상.',
      parameters: { type: 'object', properties: {
        form: { type: 'string', enum: ['교목', '관목'], description: '성상' },
        evergreen: { type: 'boolean', description: 'true면 상록만' },
        species: { type: 'string', description: '수목명 부분일치(예: 소나무)' },
        minDBH: { type: 'number', description: '흉고직경(cm) 이상' },
        maxDBH: { type: 'number', description: '흉고직경(cm) 이하' } }, required: [] } } },
    { type: 'function', function: { name: 'get_cost_breakdown',
      description: '순공사비(일위대가 단가×유지관리수량) 내역 표와 합계를 렌더하고 반환.',
      parameters: { type: 'object', properties: {}, required: [] } } },
    { type: 'function', function: { name: 'validate_data',
      description: '결측·규격오류·±2σ 이상치·계수 누락 등 검증 결과를 렌더하고 반환.',
      parameters: { type: 'object', properties: {}, required: [] } } },
    { type: 'function', function: { name: 'list_coefficients',
      description: '현재 적용 중인 탄소계수(상대생장식 a,b,R,CF,g, 관목 면적계수)와 출처·신뢰도를 반환.',
      parameters: { type: 'object', properties: {}, required: [] } } },
    { type: 'function', function: { name: 'update_coefficient',
      description: '탄소계수를 변경한다. 가정 변경이므로 사용자 승인(HITL) 필요 — 승인 카드를 띄우고 승인 시에만 적용·재계산된다.',
      parameters: { type: 'object', properties: {
        group: { type: 'string', enum: ['trees', 'shrubs'] },
        key: { type: 'string', description: '계수 키(상록침엽/상록활엽/낙엽침엽/낙엽활엽/관목)' },
        a: { type: 'number' }, b: { type: 'number' }, R: { type: 'number' }, CF: { type: 'number' }, g: { type: 'number' },
        areaCoeff: { type: 'number', description: '관목 면적계수 kgCO₂/㎡/yr' } }, required: ['group', 'key'] } } },
    { type: 'function', function: { name: 'show_table',
      description: '임의의 데이터를 별도의 떠 있는 "테이블 창"으로 띄운다. 비교표·커스텀 표를 보여줄 때 사용.',
      parameters: { type: 'object', properties: {
        title: { type: 'string' },
        columns: { type: 'array', items: { type: 'string' } },
        rows: { type: 'array', items: { type: 'array', items: { type: ['string', 'number'] } } } }, required: ['title', 'columns', 'rows'] } } },
    { type: 'function', function: { name: 'generate_pdf_report',
      description: '최종 PDF 보고서를 생성한다. 최종 확정이므로 사용자 승인(HITL) 카드를 띄우고 승인 시 보고서 창을 연다.',
      parameters: { type: 'object', properties: {}, required: [] } } },
    { type: 'function', function: { name: 'get_inventory',
      description: '추출된 원시 수목 레코드(수목명/규격/단위/수량/성상)를 반환.',
      parameters: { type: 'object', properties: {}, required: [] } } },
  ];

  // ---- 도구 실행기 ----
  async function execTool(name, args) {
    args = args || {};
    switch (name) {
      case 'run_full_analysis': agentRenderAnalysis(); return j(maskedSummary());
      case 'get_co2_summary': runFullAnalysisIfNeeded(); buildReport(); addBot([kpiBlock()]); return j(maskedSummary());
      case 'get_inventory': return j(STATE.parsed ? STATE.parsed.treeRecords.map((r) => ({ 수목명: r.수목명, 규격: r.규격, 단위: r.단위, 수량: MASK, 성상: r.성상 })) : []);
      case 'filter_inventory': return agentFilter(args);
      case 'get_cost_breakdown':
        runFullAnalysisIfNeeded();
        if (!STATE.cost) return j({ error: '단가 데이터 없음' });
        addBot([`**순공사비 산출** 💰`, costBlock()]);
        return j({ total_KRW: MASK, matched: STATE.cost.matched, unmatched: STATE.cost.unmatched,
          note: '순공사비·금액·수량은 익명화(***)되었고 단가는 비공개입니다.',
          lines: STATE.cost.lines.slice(0, 10).map((l) => ({ 작업: l.name, 규격: l.spec, 수량: MASK, 금액: MASK })) });
      case 'validate_data': runFullAnalysisIfNeeded(); addBot([`**검증 결과** 🔍`, validationBlock()]); return j(STATE.validation);
      case 'list_coefficients': return j(STATE.coeff);
      case 'update_coefficient': return agentUpdateCoeff(args);
      case 'show_table': showTableWindow(args.title || '표', args.columns || [], args.rows || []); return j({ ok: true, shown: true });
      case 'generate_pdf_report':
        runFullAnalysisIfNeeded();
        askApproval('최종 보고서 확정 (HITL)', '보고서를 확정하고 PDF 생성 창을 열까요?',
          [{ t: '✅ 확정·생성', go: true, value: 'ok', fn: () => generatePDF() },
           { t: '취소', value: 'cancel', fn: () => addBot('보고서 생성을 취소했습니다.') }]);
        return j({ status: 'pending_user_approval', note: '승인 카드를 표시함 — 사용자가 [확정·생성]을 눌러야 PDF가 열립니다.' });
      default: return j({ error: '알 수 없는 도구: ' + name });
    }
  }

  function kpiBlock() {
    const R = STATE.results;
    return { type: 'kpis', items: [
      { v: fmt(R.totals.annual_tCO2), u: 'tCO₂/yr', k: '연간 흡수량(확정)' },
      { v: fmt(R.totals.stock_tCO2), u: 'tCO₂', k: '탄소 저장량' },
      { v: MASK, u: '주', k: '교목 수량' },
      { v: fmtInt(R.totals.speciesCount), u: '종', k: '수종 수' },
    ] };
  }
  function agentRenderAnalysis() {
    runCalculations(); validate(); computeCost(); buildReport();
    addBot([`**연간 CO₂ 흡수량 산정 완료** ✅`, kpiBlock(), co2BySpeciesTable()]);
    addBot([{ type: 'html', html: '<div class="ca-note">수종별 연간 흡수량 비중</div>' }, co2DonutBlock(), co2BarBlock()]);
    if (STATE.cost) addBot([`**순공사비** 💰`, costBlock()]);
    const V = STATE.validation;
    if (V.warnings.length || V.outliers.length || V.missing.length) addBot([`**검증** 🔍`, validationBlock()]);
  }
  function agentFilter(args) {
    runFullAnalysisIfNeeded();
    if (!STATE.classified) return j({ error: '데이터 없음' });
    let lines = STATE.classified.slice();
    if (args.form) lines = lines.filter((L) => L.cls.form === args.form);
    if (args.evergreen) lines = lines.filter((L) => L.cls.ever === '상록');
    if (args.species) { const q = norm(args.species); lines = lines.filter((L) => norm(L.rec.수목명).includes(q)); }
    if (args.minDBH != null) lines = lines.filter((L) => L.cls.spec.type === 'B' && L.cls.spec.mid != null && L.cls.spec.mid >= args.minDBH);
    if (args.maxDBH != null) lines = lines.filter((L) => L.cls.spec.type === 'B' && L.cls.spec.mid != null && L.cls.spec.mid <= args.maxDBH);
    if (!lines.length) { addBot('해당 조건의 수목이 없습니다.'); return j({ count: 0, rows: [] }); }
    const totA = lines.reduce((a, L) => a + (L.calc.co2Annual || 0), 0);
    addBot([
      `**필터 결과** — ${lines.length}항목, 연간 ${fmt(totA)} kgCO₂/yr`,
      { type: 'table', head: ['수목명', '규격', '수량', '연간 kgCO₂/yr'],
        rows: lines.map((L) => [L.rec.수목명, L.rec.규격, MASK, L.calc.co2Annual == null ? '확인필요' : fmt(L.calc.co2Annual)]),
        foot: ['합계', '', '', fmt(totA)] },
      { type: 'html', html: hbarSVG(lines.filter((L) => L.calc.co2Annual).map((L) => ({ label: L.rec.수목명 + ' ' + L.rec.규격, value: L.calc.co2Annual })), {}) },
    ]);
    return j({ count: lines.length, total_annual_kgCO2: round2(totA),
      rows: lines.map((L) => ({ 수목명: L.rec.수목명, 규격: L.rec.규격, 수량: MASK, 단위: L.rec.단위, 연간kgCO2: L.calc.co2Annual })) });
  }
  function agentUpdateCoeff(args) {
    const grp = args.group, key = args.key;
    if (!STATE.coeff[grp] || !STATE.coeff[grp][key]) return j({ error: '계수 키 없음: ' + grp + '/' + key });
    const cur = STATE.coeff[grp][key];
    const changes = {}; ['a', 'b', 'R', 'CF', 'g', 'areaCoeff'].forEach((f) => { if (args[f] != null) changes[f] = args[f]; });
    if (!Object.keys(changes).length) return j({ error: '변경할 필드 없음' });
    const desc = Object.entries(changes).map(([k, v]) => `${k}=${v}`).join(', ');
    askApproval('탄소계수 변경 (HITL)', `**${grp}/${key}** 계수를 다음으로 변경할까요?\n\`${desc}\`\n(가정 변경은 승인이 필요합니다)`,
      [{ t: '✅ 승인·적용', go: true, value: 'ok', fn: () => {
          Object.assign(cur, changes); cur.conf = 'user'; cur.src = '사용자 입력(확정)'; STATE.coeffApproved = true;
          runCalculations(); validate(); computeCost(); buildReport();
          addBot(['✅ 계수를 적용하고 재계산했습니다.', kpiBlock()]);
        } },
       { t: '거부', value: 'no', fn: () => addBot('계수 변경을 취소했습니다.') }]);
    return j({ status: 'pending_user_approval', proposed: changes });
  }

  // ---- OpenAI 호출 + 에이전트 루프 ----
  async function openaiChat(messages, tools) {
    const ctrl = new AbortController();
    const to = setTimeout(() => ctrl.abort(), 60000);
    let res;
    try {
      res = await fetch(STATE.ai.baseUrl.replace(/\/$/, '') + '/chat/completions', {
        method: 'POST', signal: ctrl.signal,
        headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + STATE.ai.apiKey },
        body: JSON.stringify({ model: STATE.ai.model, messages, tools, tool_choice: 'auto' }),
      });
    } finally { clearTimeout(to); }
    if (!res.ok) { let t = ''; try { t = await res.text(); } catch (e) {} throw new Error(res.status + ' ' + t.slice(0, 220)); }
    const data = await res.json();
    if (!data.choices || !data.choices[0]) throw new Error('빈 응답');
    return data.choices[0].message;
  }
  async function agentTurn(userText) {
    if (!STATE.ai.apiKey) return promptForKey();
    AGENT_HISTORY.push({ role: 'user', content: userText });
    const sys = { role: 'system', content: systemPrompt() };
    STATE.ai.busy = true;
    try {
      for (let i = 0; i < 6; i++) {
        const t = typing();
        let msg;
        try { msg = await openaiChat([sys, ...AGENT_HISTORY], TOOLS); } finally { t.remove(); }
        AGENT_HISTORY.push(msg);
        const calls = msg.tool_calls || [];
        if (msg.content && calls.length) addBot([msg.content]);
        if (calls.length) {
          for (const tc of calls) {
            let a = {}; try { a = JSON.parse(tc.function.arguments || '{}'); } catch (e) {}
            addToolNote(tc.function.name);
            let result; try { result = await execTool(tc.function.name, a); } catch (e) { result = j({ error: String(e.message || e) }); }
            AGENT_HISTORY.push({ role: 'tool', tool_call_id: tc.id, content: typeof result === 'string' ? result : j(result) });
          }
          continue;
        }
        if (msg.content) addBot([msg.content]);
        return;
      }
      addBot('(도구 호출 반복 한도에 도달했습니다.)');
    } catch (e) {
      addBot('⚠ OpenAI 호출 오류: ' + esc(String(e.message || e)) + '\nAPI 키·모델·네트워크를 확인하세요. (헤더 ⚙︎ 설정)');
    } finally { STATE.ai.busy = false; }
  }
  function systemPrompt() {
    const sum = STATE.report ? j(maskedSummary()) : '(아직 분석 전 — 필요 시 run_full_analysis)';
    const sheets = STATE.parsed ? STATE.parsed.sheetNames.join(', ') : '(미로드)';
    const n = STATE.parsed ? STATE.parsed.treeRecords.length : 0;
    return [
      '당신은 "AI챗봇" — 조경 유지관리 내역서(엑셀) 기반 연간 CO₂ 흡수량·순공사비 산정 AI 에이전트입니다.',
      '한국어로 간결하고 정확하게 답합니다. 모든 수치는 소수점 2자리.',
      `[데이터] 시트: ${sheets} · 수목 ${n}개 항목 추출됨.`,
      '[도구 사용 규칙]',
      '- 표/데이터/차트는 반드시 도구로 렌더하라(filter_inventory, get_cost_breakdown, get_co2_summary, show_table 등). 답변 본문에 마크다운 표를 직접 쓰지 마라.',
      '- 분석/산정/전체 요청 → run_full_analysis. PDF/보고서 요청 → generate_pdf_report.',
      '- 계수 변경(update_coefficient)·PDF 확정(generate_pdf_report)은 사용자 승인(HITL)이 필요하다. 도구가 승인 카드를 띄우므로 "승인 대기 중"임을 사용자에게 알려라.',
      '- 산정에 쓰는 탄소계수는 "검증 전 기본 가정값"이다. 임의로 단정하지 말고, 공식 산정에는 국립산림과학원 계수 입력이 필요함을 적절히 안내하라. 데이터에 없는 값은 추정 금지("자료 없음").',
      '[익명화 규칙 — 반드시 준수]',
      '- 순공사비(총액·금액), 수목의 수량, 녹지/관목의 면적은 익명화 대상이다. 표·PDF·문장 어디서든 항상 "***"로만 표기하고, 실제 숫자를 추정·역산·노출하지 마라.',
      '- 단가(단위가격)는 비공개다. 표 컬럼으로 만들지 말고, 단가 값을 언급하지도 마라.',
      '- 도구가 돌려주는 값에 위 항목이 "***"로 와 있으면 그대로 "***"로 답하라.',
      `[현재 산정 요약] ${sum}`,
    ].join('\n');
  }

  // ---- 떠 있는 "테이블 창" ----
  function showTableWindow(title, columns, rows) {
    const win = el('div', 'ca-win');
    win.innerHTML = `<div class="ca-win-head"><span>📊 ${esc(title)}</span><button class="ca-win-x">×</button></div><div class="ca-win-body"></div>`;
    win.querySelector('.ca-win-body').appendChild(buildTable({ head: columns, rows: rows.map((r) => (Array.isArray(r) ? r : [r])) }));
    document.body.appendChild(win);
    win.style.left = Math.max(16, (window.innerWidth - 460) / 2) + 'px';
    win.style.top = '80px';
    win.querySelector('.ca-win-x').onclick = () => win.remove();
    ['wheel', 'mousedown', 'mousemove', 'mouseup', 'touchstart', 'touchmove', 'contextmenu'].forEach((ev) =>
      win.addEventListener(ev, (e) => e.stopPropagation()));
    dragify(win, win.querySelector('.ca-win-head'));
    return win;
  }
  function dragify(win, handle) {
    let sx, sy, ox, oy, drag = false;
    handle.addEventListener('mousedown', (e) => { drag = true; sx = e.clientX; sy = e.clientY; ox = parseInt(win.style.left) || 0; oy = parseInt(win.style.top) || 0; e.preventDefault(); });
    window.addEventListener('mousemove', (e) => { if (!drag) return; win.style.left = (ox + e.clientX - sx) + 'px'; win.style.top = Math.max(0, oy + e.clientY - sy) + 'px'; });
    window.addEventListener('mouseup', () => { drag = false; });
  }

  // ---- 설정(API 키·모델) ----
  function openSettings() {
    const wrap = el('div', 'ca-set');
    wrap.innerHTML =
      `<label>OpenAI API Key<input id="ca-key" type="password" placeholder="sk-..." value="${esc(STATE.ai.apiKey)}"></label>` +
      `<label>모델 (model)<input id="ca-model" value="${esc(STATE.ai.model)}" placeholder="gpt-4o-mini"></label>` +
      `<label>Base URL<input id="ca-base" value="${esc(STATE.ai.baseUrl)}"></label>` +
      `<div class="ca-note">🔒 키는 이 브라우저(localStorage)에만 저장되고 OpenAI로만 전송됩니다. 공용 PC 주의. 모델 예: gpt-4o-mini, gpt-4o, gpt-4.1-mini.</div>`;
    const m = addBot(['**⚙︎ OpenAI 설정**', { type: 'html', html: '' }]);
    m.querySelector('.ca-bubble').appendChild(wrap);
    const a = el('div', 'ca-actions');
    const save = el('button', 'ca-chip go', '💾 저장');
    save.onclick = () => {
      STATE.ai.apiKey = $id('ca-key').value.trim();
      STATE.ai.model = $id('ca-model').value.trim() || 'gpt-4o-mini';
      STATE.ai.baseUrl = $id('ca-base').value.trim() || 'https://api.openai.com/v1';
      lsSet('ca_openai_key', STATE.ai.apiKey); lsSet('ca_openai_model', STATE.ai.model); lsSet('ca_openai_base', STATE.ai.baseUrl);
      updateAiBadge();
      addBot(STATE.ai.apiKey ? '✅ 저장 완료. 이제 자유롭게 질문하세요 — GPT 에이전트가 도구를 사용해 응답합니다.' : '키가 비어 있습니다. 빠른 작업 버튼은 키 없이도 동작합니다.');
    };
    a.appendChild(save); m.querySelector('.ca-bubble').appendChild(a);
    if (!panel.classList.contains('open')) togglePanel(true);
  }
  function promptForKey() {
    addBot(['🔑 자유 질의는 **OpenAI API 키**가 필요합니다. 설정에서 키를 입력하세요. (빠른 작업 버튼은 키 없이도 동작합니다.)',
      { type: 'actions', items: [{ t: '⚙︎ OpenAI 설정 열기', go: true, fn: () => openSettings() }] }]);
  }
  function updateAiBadge() {
    const b = $id('ca-llm-badge'); if (!b) return;
    if (STATE.ai.apiKey) { b.className = 'ca-llm on'; b.textContent = 'GPT'; b.title = STATE.ai.model; }
    else { b.className = 'ca-llm off'; b.textContent = 'API 키 필요'; b.title = '⚙︎에서 OpenAI 키 입력'; }
  }
  function addToolNote(name) {
    const labels = { run_full_analysis: '전체 분석 실행', get_co2_summary: 'CO₂ 요약', filter_inventory: '인벤토리 필터',
      get_cost_breakdown: '순공사비 계산', validate_data: '데이터 검증', list_coefficients: '계수 조회',
      update_coefficient: '계수 변경(승인 필요)', show_table: '테이블 창 표시', generate_pdf_report: 'PDF 보고서(승인 필요)', get_inventory: '수목 데이터 조회' };
    const m = el('div', 'ca-msg bot');
    m.innerHTML = `<div class="ca-tool">🔧 도구 실행 · <b>${esc(labels[name] || name)}</b></div>`;
    log.appendChild(m); scrollDown();
  }

  // ===========================================================================
  // 6. 챗봇 UI
  // ===========================================================================
  let panel, log, inputEl;
  function injectStyles() {
    const css = `
    #ca-launch{position:fixed;right:16px;bottom:16px;z-index:320;display:flex;align-items:center;gap:8px;
      background:linear-gradient(135deg,#22a06b,#4a8bff);color:#fff;border:none;border-radius:24px;
      padding:11px 16px;font:600 13px/1 -apple-system,"Pretendard","Noto Sans KR",sans-serif;cursor:pointer;
      box-shadow:0 8px 24px rgba(0,0,0,.45);transition:transform .15s,box-shadow .15s;}
    #ca-launch:hover{transform:translateY(-2px);box-shadow:0 12px 30px rgba(0,0,0,.55);}
    #ca-launch .ca-pulse{width:8px;height:8px;border-radius:50%;background:#7CFFB2;box-shadow:0 0 0 0 #7CFFB2;animation:caPulse 2s infinite;}
    @keyframes caPulse{0%{box-shadow:0 0 0 0 rgba(124,255,178,.6)}70%{box-shadow:0 0 0 7px transparent}100%{box-shadow:0 0 0 0 transparent}}
    #ca-panel{position:fixed;top:0;right:0;height:100%;width:440px;max-width:94vw;z-index:330;
      background:rgba(13,15,22,.94);backdrop-filter:blur(18px);-webkit-backdrop-filter:blur(18px);
      border-left:1px solid rgba(255,255,255,.09);box-shadow:-12px 0 40px rgba(0,0,0,.5);
      display:flex;flex-direction:column;transform:translateX(100%);transition:transform .28s cubic-bezier(.4,0,.2,1);
      font-family:-apple-system,"Pretendard","Noto Sans KR",sans-serif;color:#e7ebf2;}
    #ca-panel.open{transform:translateX(0);}
    .ca-head{display:flex;align-items:center;gap:10px;padding:14px 16px;border-bottom:1px solid rgba(255,255,255,.08);}
    .ca-head .ca-ic{width:30px;height:30px;border-radius:8px;display:grid;place-items:center;background:linear-gradient(135deg,#22a06b,#4a8bff);font-size:16px;}
    .ca-head h3{font-size:14px;font-weight:700;margin:0;}
    .ca-head .ca-sub{font-size:10.5px;color:#8892a6;margin-top:1px;}
    .ca-head .ca-x{margin-left:auto;background:none;border:none;color:#8892a6;font-size:20px;cursor:pointer;line-height:1;padding:4px 8px;border-radius:6px;}
    .ca-head .ca-x:hover{background:rgba(255,255,255,.08);color:#fff;}
    .ca-llm{font-size:9.5px;padding:2px 7px;border-radius:10px;border:1px solid;}
    .ca-llm.on{color:#7CFFB2;border-color:rgba(124,255,178,.4);background:rgba(124,255,178,.08);}
    .ca-llm.off{color:#ffb84a;border-color:rgba(255,184,74,.4);background:rgba(255,184,74,.08);}
    #ca-log{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:12px;scroll-behavior:smooth;}
    #ca-log::-webkit-scrollbar{width:8px}#ca-log::-webkit-scrollbar-thumb{background:rgba(255,255,255,.12);border-radius:4px}
    .ca-msg{max-width:92%;font-size:12.8px;line-height:1.6;}
    .ca-msg.user{align-self:flex-end;background:linear-gradient(135deg,#3a6fd8,#4a8bff);color:#fff;padding:9px 13px;border-radius:14px 14px 4px 14px;}
    .ca-msg.bot{align-self:flex-start;width:100%;}
    .ca-bubble{background:rgba(255,255,255,.045);border:1px solid rgba(255,255,255,.07);border-radius:4px 14px 14px 14px;padding:12px 14px;}
    .ca-msg.bot b{color:#fff;} .ca-msg.bot code{background:rgba(255,255,255,.1);padding:1px 5px;border-radius:4px;font-size:11.5px;}
    .ca-typing{display:inline-flex;gap:4px;padding:4px 0}.ca-typing i{width:7px;height:7px;border-radius:50%;background:#5b6680;animation:caBlink 1.2s infinite}
    .ca-typing i:nth-child(2){animation-delay:.2s}.ca-typing i:nth-child(3){animation-delay:.4s}
    @keyframes caBlink{0%,60%,100%{opacity:.3}30%{opacity:1}}
    table.ca-tbl{width:100%;border-collapse:collapse;margin:8px 0;font-size:11.3px;}
    table.ca-tbl th,table.ca-tbl td{border:1px solid rgba(255,255,255,.09);padding:5px 7px;text-align:right;}
    table.ca-tbl th{background:rgba(74,139,255,.16);color:#cfe0ff;font-weight:600;text-align:center;}
    table.ca-tbl td:first-child,table.ca-tbl th:first-child{text-align:left;}
    table.ca-tbl tr:nth-child(even) td{background:rgba(255,255,255,.025);}
    table.ca-tbl tfoot td{font-weight:700;background:rgba(34,160,107,.14);color:#9cffd0;}
    .ca-kpis{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:6px 0;}
    .ca-kpi{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);border-radius:10px;padding:10px 12px;}
    .ca-kpi .v{font-size:19px;font-weight:800;color:#fff;font-variant-numeric:tabular-nums;}
    .ca-kpi .v small{font-size:11px;font-weight:600;color:#8892a6;margin-left:3px;}
    .ca-kpi .k{font-size:10.5px;color:#8892a6;margin-top:2px;}
    .ca-chartwrap{display:flex;gap:12px;align-items:center;margin:8px 0;flex-wrap:wrap;}
    .ca-legend{display:flex;flex-direction:column;gap:4px;font-size:11px;color:#c4ccda;}
    .ca-leg{display:flex;align-items:center;gap:6px}.ca-dot{width:9px;height:9px;border-radius:2px;display:inline-block}
    .ca-hbar{display:flex;flex-direction:column;gap:7px;margin:8px 0;}
    .ca-bar-row{display:grid;grid-template-columns:90px 1fr auto;gap:8px;align-items:center;font-size:11px;}
    .ca-bar-lbl{color:#c4ccda;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
    .ca-bar-track{background:rgba(255,255,255,.07);border-radius:5px;height:16px;overflow:hidden;}
    .ca-bar-fill{height:100%;border-radius:5px;transition:width .5s;}
    .ca-bar-val{color:#fff;font-weight:600;font-variant-numeric:tabular-nums;}
    .ca-actions{display:flex;flex-wrap:wrap;gap:6px;margin-top:4px;}
    .ca-chip{background:rgba(74,139,255,.12);border:1px solid rgba(74,139,255,.3);color:#bcd4ff;border-radius:14px;
      padding:6px 11px;font-size:11.5px;cursor:pointer;transition:background .15s;}
    .ca-chip:hover{background:rgba(74,139,255,.25);}
    .ca-chip.warn{background:rgba(255,184,74,.1);border-color:rgba(255,184,74,.35);color:#ffd699;}
    .ca-chip.go{background:linear-gradient(135deg,#22a06b,#2bbd7e);border:none;color:#fff;font-weight:600;}
    .ca-hitl{border:1px solid rgba(255,184,74,.4);background:rgba(255,184,74,.07);border-radius:12px;padding:12px;}
    .ca-hitl .h{font-weight:700;color:#ffd699;font-size:12.5px;margin-bottom:6px;display:flex;align-items:center;gap:6px;}
    .ca-badge{font-size:9px;padding:1px 6px;border-radius:8px;font-weight:700;}
    .ca-badge.high{background:rgba(124,255,178,.15);color:#7CFFB2}.ca-badge.method{background:rgba(74,139,255,.15);color:#9cc0ff}
    .ca-badge.low{background:rgba(255,184,74,.15);color:#ffc874}.ca-badge.none{background:rgba(255,138,138,.15);color:#ff9d9d}
    .ca-badge.user{background:rgba(124,255,178,.2);color:#7CFFB2}
    .ca-note{font-size:10.5px;color:#8892a6;margin-top:6px;}
    .ca-foot{padding:10px 12px;border-top:1px solid rgba(255,255,255,.08);}
    .ca-quick{display:flex;gap:6px;overflow-x:auto;padding-bottom:8px;}.ca-quick::-webkit-scrollbar{height:0}
    .ca-inrow{display:flex;gap:8px;align-items:flex-end;}
    #ca-input{flex:1;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:10px;
      color:#fff;padding:9px 12px;font-size:12.5px;resize:none;max-height:90px;font-family:inherit;}
    #ca-input:focus{outline:none;border-color:rgba(74,139,255,.6);}
    #ca-send{background:linear-gradient(135deg,#22a06b,#4a8bff);border:none;color:#fff;border-radius:10px;width:40px;height:38px;cursor:pointer;font-size:16px;}
    .ca-coed{display:flex;flex-direction:column;gap:9px;}
    .ca-coed .grp{border:1px solid rgba(255,255,255,.08);border-radius:9px;padding:9px;}
    .ca-coed .grp h5{font-size:11.5px;margin:0 0 6px;color:#cfe0ff;display:flex;justify-content:space-between;align-items:center;}
    .ca-coed label{display:grid;grid-template-columns:64px 1fr;gap:6px;align-items:center;font-size:11px;margin:3px 0;color:#aab3c5;}
    .ca-coed input{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);border-radius:6px;color:#fff;padding:4px 7px;font-size:11.5px;}
    .ca-coed .src{font-size:10px;color:#8892a6;margin-top:4px;}
    .ca-set{display:flex;flex-direction:column;gap:8px;}
    .ca-set label{display:flex;flex-direction:column;gap:3px;font-size:11px;color:#aab3c5;}
    .ca-set input{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);border-radius:7px;color:#fff;padding:7px 9px;font-size:12px;font-family:inherit;}
    .ca-tool{font-size:10.5px;color:#8aa0c0;background:rgba(74,139,255,.08);border:1px solid rgba(74,139,255,.18);border-radius:8px;padding:5px 9px;display:inline-block;}
    #ca-settings{margin-left:0;}
    .ca-win{position:fixed;z-index:340;width:440px;max-width:92vw;max-height:72vh;display:flex;flex-direction:column;overflow:hidden;
      background:rgba(17,20,29,.97);backdrop-filter:blur(18px);-webkit-backdrop-filter:blur(18px);border:1px solid rgba(255,255,255,.12);
      border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,.6);font-family:-apple-system,"Pretendard","Noto Sans KR",sans-serif;color:#e7ebf2;}
    .ca-win-head{display:flex;align-items:center;justify-content:space-between;padding:9px 12px;border-bottom:1px solid rgba(255,255,255,.1);
      font-size:12.5px;font-weight:700;color:#cfe0ff;background:rgba(74,139,255,.1);cursor:move;user-select:none;}
    .ca-win-x{background:none;border:none;color:#8892a6;font-size:18px;cursor:pointer;line-height:1;}
    .ca-win-x:hover{color:#fff;}
    .ca-win-body{padding:10px 12px;overflow:auto;}
    .ca-win-body table.ca-tbl{font-size:11.5px;}
    `;
    document.head.appendChild(el('style', null, css));
  }

  function buildUI() {
    const launch = el('button', null,
      `<span class="ca-pulse"></span>🤖 AI챗봇`);
    launch.id = 'ca-launch';
    launch.onclick = () => togglePanel(true);
    document.body.appendChild(launch);

    panel = el('div');
    panel.id = 'ca-panel';
    panel.innerHTML = `
      <div class="ca-head">
        <div class="ca-ic">🤖</div>
        <div><h3>AI챗봇</h3><div class="ca-sub">조경 내역서 CO₂·순공사비 AI 에이전트</div></div>
        <span class="ca-llm off" id="ca-llm-badge" style="margin-left:auto">API 키 필요</span>
        <button class="ca-x" id="ca-settings" title="OpenAI 설정">⚙︎</button>
        <button class="ca-x" id="ca-close" title="닫기">×</button>
      </div>
      <div id="ca-log"></div>
      <div class="ca-foot">
        <div class="ca-quick" id="ca-quick"></div>
        <div class="ca-inrow">
          <textarea id="ca-input" rows="1" placeholder="무엇이든 물어보세요 — 예: 전체 분석 / 교목만 / B21 이상 / 순공사비 / PDF 보고서 만들어줘"></textarea>
          <button id="ca-send">➤</button>
        </div>
      </div>`;
    document.body.appendChild(panel);
    log = $id('ca-log');
    inputEl = $id('ca-input');
    $id('ca-close').onclick = () => togglePanel(false);
    $id('ca-settings').onclick = () => openSettings();
    $id('ca-send').onclick = submit;
    inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); submit(); }
    });
    inputEl.addEventListener('input', () => { inputEl.style.height = 'auto'; inputEl.style.height = Math.min(90, inputEl.scrollHeight) + 'px'; });
    // 패널 내부 이벤트가 3D 캔버스로 전파되지 않도록
    ['wheel', 'mousedown', 'mousemove', 'mouseup', 'touchstart', 'touchmove', 'contextmenu'].forEach((ev) =>
      panel.addEventListener(ev, (e) => e.stopPropagation()));
    renderQuick(DEFAULT_QUICK);
  }

  function togglePanel(open) {
    panel.classList.toggle('open', open);
    $id('ca-launch').style.display = open ? 'none' : 'flex';
  }

  const DEFAULT_QUICK = [
    { t: '🚀 전체 분석 실행', go: true, q: '전체 분석' },
    { t: '🌲 교목만', q: '교목만 보여줘' },
    { t: '🌿 관목만', q: '관목만 보여줘' },
    { t: '💰 순공사비', q: '순공사비' },
    { t: '⚙️ 계수 보기·수정', q: '계수' },
    { t: '✅ 검증', q: '검증' },
    { t: '📄 PDF 보고서', q: 'PDF 보고서' },
  ];
  function renderQuick(items) {
    const q = $id('ca-quick');
    q.innerHTML = '';
    items.forEach((it) => {
      const c = el('button', 'ca-chip' + (it.go ? ' go' : '') + (it.warn ? ' warn' : ''), it.t);
      c.onclick = () => { if (it.fn) it.fn(); else handleUser(it.q, true); };
      q.appendChild(c);
    });
  }

  // ---- 메시지 렌더 ----------------------------------------------------------
  function addUser(text) {
    const m = el('div', 'ca-msg user', esc(text));
    log.appendChild(m); scrollDown();
  }
  function addBot(blocks) {
    const m = el('div', 'ca-msg bot');
    const b = el('div', 'ca-bubble');
    b.appendChild(renderBlocks(blocks));
    m.appendChild(b); log.appendChild(m); scrollDown();
    return m;
  }
  function typing() {
    const m = el('div', 'ca-msg bot');
    m.innerHTML = `<div class="ca-bubble"><div class="ca-typing"><i></i><i></i><i></i></div></div>`;
    log.appendChild(m); scrollDown();
    return m;
  }
  function renderBlocks(blocks) {
    const frag = document.createDocumentFragment();
    (Array.isArray(blocks) ? blocks : [blocks]).forEach((blk) => {
      if (blk == null) return;
      if (typeof blk === 'string') { frag.appendChild(el('div', null, mdInline(blk))); return; }
      if (blk.type === 'html') { frag.appendChild(el('div', null, blk.html)); return; }
      if (blk.type === 'table') { frag.appendChild(buildTable(blk)); return; }
      if (blk.type === 'kpis') { frag.appendChild(buildKpis(blk.items)); return; }
      if (blk.type === 'actions') {
        const a = el('div', 'ca-actions');
        blk.items.forEach((it) => {
          const c = el('button', 'ca-chip' + (it.go ? ' go' : '') + (it.warn ? ' warn' : ''), it.t);
          c.onclick = () => { if (it.fn) it.fn(); else handleUser(it.q, true); };
          a.appendChild(c);
        });
        frag.appendChild(a); return;
      }
      if (blk.type === 'hitl') { frag.appendChild(buildHITL(blk)); return; }
    });
    return frag;
  }
  function mdInline(s) {
    return esc(s).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').replace(/`(.+?)`/g, '<code>$1</code>').replace(/\n/g, '<br>');
  }
  function buildKpis(items) {
    const g = el('div', 'ca-kpis');
    items.forEach((k) => g.appendChild(el('div', 'ca-kpi',
      `<div class="v">${k.v}<small>${k.u || ''}</small></div><div class="k">${esc(k.k)}</div>`)));
    return g;
  }
  function buildTable(blk) {
    const t = el('table', 'ca-tbl');
    let h = '<thead><tr>' + blk.head.map((c) => `<th>${esc(c)}</th>`).join('') + '</tr></thead><tbody>';
    h += blk.rows.map((r) => '<tr>' + r.map((c) => `<td>${c == null ? '—' : (typeof c === 'object' ? c.html : esc(c))}</td>`).join('') + '</tr>').join('');
    h += '</tbody>';
    if (blk.foot) h += '<tfoot><tr>' + blk.foot.map((c) => `<td>${esc(c)}</td>`).join('') + '</tr></tfoot>';
    t.innerHTML = h;
    return t;
  }
  function buildHITL(blk) {
    const d = el('div', 'ca-hitl');
    d.innerHTML = `<div class="h">⚠ ${esc(blk.title)}</div><div>${mdInline(blk.body)}</div>`;
    const a = el('div', 'ca-actions');
    blk.options.forEach((o) => {
      const c = el('button', 'ca-chip' + (o.go ? ' go' : '') + (o.warn ? ' warn' : ''), o.t);
      c.onclick = () => { Array.from(a.children).forEach((x) => (x.disabled = true)); o.fn(); };
      a.appendChild(c);
    });
    d.appendChild(a);
    return d;
  }
  function scrollDown() { log.scrollTop = log.scrollHeight; }
  function askApproval(title, body, options) {
    return new Promise((resolve) => {
      addBot([{ type: 'hitl', title, body, options: options.map((o) => ({ ...o, fn: () => { o.fn && o.fn(); resolve(o.value); } })) }]);
    });
  }

  // ===========================================================================
  // 7. 파이프라인 & 인텐트 라우팅
  // ===========================================================================
  function ensureLoaded() {
    if (!STATE.parsed) {
      addBot('엑셀 데이터가 아직 로드되지 않았습니다. 잠시 후 다시 시도하거나 파일을 업로드하세요.');
      return false;
    }
    return true;
  }

  async function runFullAnalysis() {
    if (!ensureLoaded()) return;
    const t = typing();
    await sleep(250);
    runCalculations();
    validate();
    computeCost();
    t.remove();

    // 추출 요약
    const recs = STATE.parsed.treeRecords;
    addBot([
      `**1단계 · 데이터 추출 완료** 📋\n시트 \`${esc(STATE.parsed.sheetNames.join(', '))}\`에서 수목 **${recs.length}개 항목**을 추출했습니다.`,
      { type: 'table', head: ['수목명', '규격', '단위', '수량'],
        rows: recs.map((r) => [r.수목명, r.규격, r.단위, MASK]) },
    ]);
    await sleep(150);

    // 분류 요약
    const forms = {};
    STATE.classified.forEach((L) => {
      const k = (L.cls.ever || '') + (L.cls.leaf || '') + L.cls.form;
      (forms[k] = forms[k] || new Set()).add(L.rec.수목명);
    });
    addBot(`**2단계 · 수종 분류 완료** 🌲\n` +
      Object.entries(forms).map(([k, v]) => `· ${k}: ${v.size}종 (${[...v].join(', ')})`).join('\n'));
    await sleep(150);

    // HITL 게이트 1 — 계수 확인
    const lowConf = collectUsedCoeffs().filter((c) => c.conf === 'low' || c.conf === 'none');
    if (lowConf.length && !STATE.coeffApproved) {
      addBot([
        `**3단계 · 탄소계수 적용 (RAG)** 🔎\nDBH 상대생장식 계수를 적용합니다. 다만 아래 계수는 **검증 전 기본값/미입력**입니다 (명세 규칙: 임의추정 금지 → 사용자 확인 필요).`,
        { type: 'html', html: lowConf.map((c) => `<div class="ca-note">· <b>${esc(c.key)}</b> <span class="ca-badge ${c.conf}">${c.conf === 'none' ? '자료없음' : '검증필요'}</span> — ${esc(c.src)}</div>`).join('') },
      ]);
      const choice = await askApproval('탄소계수 확정 (HITL)',
        '기본 가정값으로 계속 진행할까요, 아니면 검증된 계수를 직접 입력할까요?',
        [
          { t: '✏️ 계수 직접 수정', warn: true, value: 'edit', fn: () => openCoeffEditor() },
          { t: '기본값으로 진행', go: true, value: 'proceed', fn: () => { STATE.coeffApproved = true; } },
        ]);
      if (choice === 'edit') return; // 편집기에서 저장 후 사용자가 다시 실행
    }

    // 계산 결과
    const R = STATE.results;
    addBot([
      `**4단계 · 연간 CO₂ 흡수량 산정 완료** ✅`,
      { type: 'kpis', items: [
        { v: fmt(R.totals.annual_tCO2), u: 'tCO₂/yr', k: '연간 흡수량(확정)' },
        { v: fmt(R.totals.stock_tCO2), u: 'tCO₂', k: '탄소 저장량' },
        { v: MASK, u: '주', k: '교목 수량' },
        { v: fmtInt(R.totals.speciesCount), u: '종', k: '수종 수' },
      ] },
      co2BySpeciesTable(),
    ]);
    await sleep(120);
    addBot([{ type: 'html', html: '<div class="ca-note">수종별 연간 흡수량 비중</div>' }, co2DonutBlock(), co2BarBlock()]);

    // 검증 + 이상치 HITL
    const V = STATE.validation;
    if (V.warnings.length || V.missing.length || V.outliers.length) {
      addBot([`**5단계 · 검증** 🔍`, validationBlock()]);
    }
    if (V.outliers.length) {
      await askApproval('이상치 처리 (HITL)',
        `±2σ 이상치 **${V.outliers.length}건**이 감지되었습니다: ` +
        V.outliers.map((o) => `${o.name}(z=${o.z})`).join(', ') + '\n합계에서 제외할까요?',
        [
          { t: '제외', warn: true, value: 'exclude', fn: () => addBot('이상치를 제외 표시했습니다. (현재 데모 데이터에선 합계 영향 미미)') },
          { t: '포함 유지', go: true, value: 'keep', fn: () => addBot('이상치를 포함하여 유지합니다.') },
        ]);
    }

    // 순공사비
    if (STATE.cost) {
      addBot([`**+ 순공사비 산출** 💰 (일위대가 단가 × 유지관리수량)`, costBlock()]);
    }

    // 최종 보고서 HITL
    buildReport();
    addBot([
      `모든 단계가 끝났습니다. 최종 보고서를 확정할까요?`,
      { type: 'actions', items: [
        { t: '📄 PDF 보고서 생성', go: true, q: 'PDF 보고서' },
        { t: '{ } JSON 결과', q: 'JSON' },
        { t: '🌲 교목만', q: '교목만' }, { t: '🌿 관목만', q: '관목만' },
      ] },
    ]);
  }

  function collectUsedCoeffs() {
    const seen = {}, out = [];
    (STATE.classified || []).forEach((L) => {
      const k = L.cls.coeffGroup + '/' + L.cls.coeffKey;
      if (seen[k]) return; seen[k] = 1;
      const c = (STATE.coeff[L.cls.coeffGroup] || {})[L.cls.coeffKey];
      out.push({ key: L.cls.coeffKey, group: L.cls.coeffGroup, conf: c ? c.conf : 'none', src: c ? c.src : '계수 없음' });
    });
    return out;
  }

  // ---- 결과 블록들 ----------------------------------------------------------
  function co2BySpeciesTable(filter) {
    let list = STATE.results.bySpecies;
    if (filter === '교목') list = list.filter((s) => s.form === '교목');
    if (filter === '관목') list = list.filter((s) => s.form === '관목');
    const rows = list.map((s) => [
      s.name + (s.hasCoeff ? '' : ' ⚠'),
      s.spec, MASK,
      s.co2Annual == null ? '확인필요' : fmt(s.co2Annual),
      s.co2Stock == null ? '—' : fmt(s.co2Stock),
    ]);
    const totA = list.reduce((a, s) => a + (s.co2Annual || 0), 0);
    return { type: 'table',
      head: ['수목명', '규격', '수량', '연간 kgCO₂/yr', '저장 kgCO₂'],
      rows,
      foot: ['합계', '', '', fmt(totA), ''] };
  }
  function co2DonutBlock() {
    const data = STATE.results.bySpecies
      .filter((s) => s.co2Annual)
      .sort((a, b) => b.co2Annual - a.co2Annual)
      .map((s) => ({ label: s.name, value: round2(s.co2Annual), unit: ' kg' }));
    return { type: 'html', html: donutSVG(data, { center: fmt(STATE.results.totals.annual_tCO2), sub: 'tCO₂/yr' }) };
  }
  function co2BarBlock() {
    const data = STATE.results.bySpecies.filter((s) => s.co2Annual)
      .sort((a, b) => b.co2Annual - a.co2Annual)
      .map((s) => ({ label: s.name, value: s.co2Annual }));
    return { type: 'html', html: '<div class="ca-note" style="margin-top:8px">수종별 연간 흡수량 (kgCO₂/yr)</div>' + hbarSVG(data, { unit: '' }) };
  }
  function validationBlock() {
    const V = STATE.validation;
    let h = '';
    if (V.outliers.length) h += `<div class="ca-note">🔺 이상치(±2σ): ${V.outliers.map((o) => `${esc(o.name)}(z=${o.z})`).join(', ')}</div>`;
    if (V.missing.length) h += V.missing.map((m) => `<div class="ca-note">❗ ${esc(m)}</div>`).join('');
    if (V.warnings.length) h += V.warnings.slice(0, 12).map((w) => `<div class="ca-note">⚠ ${esc(w)}</div>`).join('');
    if (!h) h = '<div class="ca-note">이상 없음 ✅</div>';
    return { type: 'html', html: h };
  }
  function costBlock() {
    const C = STATE.cost;
    // 단가 컬럼은 비공개(제거), 수량·금액·순공사비 합계는 익명화(***).
    const rows = C.lines.slice(0, 12).map((l) => [
      l.name, l.spec, MASK, MASK,
    ]);
    return { type: 'table',
      head: ['작업명', '규격', '수량', '금액(원)'],
      rows,
      foot: ['순공사비 합계', '', '', MASK] };
  }

  // ---- 계수 편집기 (HITL) ---------------------------------------------------
  function openCoeffEditor() {
    const wrap = el('div', 'ca-coed');
    const groups = [];
    Object.entries(STATE.coeff.trees).forEach(([key, c]) => {
      groups.push(coeffGroupHTML('trees', key, c, [
        ['a', '계수 a'], ['b', '지수 b'], ['R', '뿌리비 R'], ['CF', '탄소율 CF'], ['g', '직경생장 g(cm/yr)']]));
    });
    Object.entries(STATE.coeff.shrubs).forEach(([key, c]) => {
      groups.push(coeffGroupHTML('shrubs', key, c, [['areaCoeff', 'kgCO₂/㎡/yr'], ['CF', '탄소율 CF']]));
    });
    wrap.innerHTML = groups.join('');
    const m = addBot([
      `**⚙️ 탄소계수 편집기** — 검증된 국립산림과학원 값으로 교체하세요. 저장 시 '사용자 확정'으로 표시되고 재계산됩니다.`,
      { type: 'html', html: '' },
    ]);
    m.querySelector('.ca-bubble').appendChild(wrap);
    const a = el('div', 'ca-actions');
    const save = el('button', 'ca-chip go', '💾 저장 후 재계산');
    save.onclick = () => {
      wrap.querySelectorAll('input[data-grp]').forEach((inp) => {
        const v = inp.value.trim();
        const num = v === '' ? null : Number(v);
        STATE.coeff[inp.dataset.grp][inp.dataset.key][inp.dataset.f] = (v === '' ? null : (isFinite(num) ? num : v));
      });
      // 사용자 수정 → 신뢰도 갱신
      Object.values(STATE.coeff.trees).forEach((c) => { c.conf = 'user'; c.src = '사용자 입력(확정)'; });
      Object.entries(STATE.coeff.shrubs).forEach(([k, c]) => { if (c.areaCoeff != null) { c.conf = 'user'; c.src = '사용자 입력(확정)'; } });
      STATE.coeffApproved = true;
      addBot('✅ 계수를 저장했습니다. 다시 분석을 실행합니다…');
      runFullAnalysis();
    };
    a.appendChild(save);
    m.querySelector('.ca-bubble').appendChild(a);
    if (!panel.classList.contains('open')) togglePanel(true);
  }
  function coeffGroupHTML(grp, key, c, fields) {
    const inputs = fields.map(([f, lbl]) =>
      `<label>${esc(lbl)}<input data-grp="${grp}" data-key="${esc(key)}" data-f="${f}" value="${c[f] == null ? '' : c[f]}"></label>`).join('');
    return `<div class="grp"><h5>${esc(key)} <span class="ca-badge ${c.conf}">${badgeText(c.conf)}</span></h5>${inputs}<div class="src">출처: ${esc(c.src)}</div></div>`;
  }
  function badgeText(conf) { return { high: '확정', method: '방법론', low: '검증필요', none: '자료없음', user: '사용자확정' }[conf] || conf; }

  // ---- JSON 결과 ------------------------------------------------------------
  function showJSON() {
    if (!STATE.report) buildReport();
    addBot([{ type: 'html', html: `<div class="ca-note">명세 출력 스키마 (JSON · 순공사비/수량/면적 익명화, 단가 비공개)</div><pre style="white-space:pre-wrap;font-size:10.5px;background:rgba(0,0,0,.3);padding:10px;border-radius:8px;overflow:auto;max-height:320px;color:#cfe0ff">${esc(JSON.stringify(maskedReport(), null, 2))}</pre>` }]);
  }

  // ---- PDF 보고서 -----------------------------------------------------------
  function generatePDF() {
    if (!STATE.results) { addBot('먼저 "전체 분석"을 실행하세요.'); return; }
    buildReport();
    const html = reportHTML(STATE.report, STATE.results, STATE.cost, STATE.validation);
    addBot([
      '📄 보고서를 생성했습니다. 새 창에서 **"PDF로 저장(인쇄 → 대상: PDF로 저장)"** 하면 다운로드됩니다. (한글 완벽 지원)',
      { type: 'actions', items: [
        { t: '🖨 PDF로 저장 / 인쇄', go: true, fn: () => openReportWindow(html, true) },
        { t: '👁 미리보기', fn: () => openReportWindow(html, false) },
      ] },
    ]);
  }
  function openReportWindow(html, autoPrint) {
    const w = window.open('', '_blank', 'width=900,height=1000');
    if (!w) { addBot('팝업이 차단되었습니다. 브라우저 팝업 차단을 해제한 뒤 다시 시도하세요.'); return null; }
    w.document.write('<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>CO₂ 산정 보고서</title>' +
      '<style>@page{margin:14mm}body{margin:0;background:#fff;}</style></head><body>' + html + '</body></html>');
    w.document.close();
    if (autoPrint) {
      const go = () => { try { w.focus(); w.print(); } catch (e) {} };
      w.onload = go; setTimeout(go, 600);
    }
    return w;
  }
  function reportHTML(R, res, C, V) {
    const s = R.summary;
    const styles = `font-family:'Pretendard','Malgun Gothic',sans-serif;color:#111;line-height:1.5;`;
    const th = 'border:1px solid #bbb;padding:6px 9px;background:#eef3ff;text-align:center;font-size:12px;';
    const td = 'border:1px solid #ccc;padding:6px 9px;font-size:12px;';
    const tdr = td + 'text-align:right;';
    let h = `<div style="${styles}">
      <div style="display:flex;justify-content:space-between;align-items:flex-end;border-bottom:3px solid #22a06b;padding-bottom:10px;">
        <div><h1 style="margin:0;font-size:22px;color:#15463a;">조경 연간 CO₂ 흡수량 · 순공사비 산정 보고서</h1>
        <div style="color:#555;font-size:12px;margin-top:4px;">주암조절지댐 물문화관 · 기준연도 ${s.base_year} · 산정방식: DBH 상대생장식(증분법)</div></div>
        <div style="text-align:right;font-size:11px;color:#888;">발행일 ${new Date().toLocaleDateString('ko-KR')}</div>
      </div>
      <h2 style="font-size:15px;color:#15463a;margin:18px 0 8px;">1. 요약</h2>
      <table style="border-collapse:collapse;width:100%;">
        <tr><th style="${th}">연간 CO₂ 흡수량</th><th style="${th}">탄소 저장량</th><th style="${th}">교목 수량</th><th style="${th}">수종 수</th><th style="${th}">순공사비</th></tr>
        <tr>
          <td style="${tdr}"><b>${fmt(s.total_tCO2_per_year)}</b> tCO₂/yr</td>
          <td style="${tdr}">${fmt(s.total_tCO2_stock)} tCO₂</td>
          <td style="${tdr}">${MASK} 주</td>
          <td style="${tdr}">${fmtInt(s.species_count)} 종</td>
          <td style="${tdr}">${MASK}</td>
        </tr>
      </table>
      <h2 style="font-size:15px;color:#15463a;margin:18px 0 8px;">2. 수종별 산정 내역</h2>
      <table style="border-collapse:collapse;width:100%;">
        <tr><th style="${th}">수목명</th><th style="${th}">성상</th><th style="${th}">규격</th><th style="${th}">수량</th><th style="${th}">연간 흡수 kgCO₂/yr</th><th style="${th}">저장 kgCO₂</th><th style="${th}">계수</th></tr>`;
    R.by_species.forEach((r) => {
      h += `<tr><td style="${td}">${esc(r.수목명)}</td><td style="${td}">${esc(r.성상)}</td><td style="${td}">${esc(r.규격)}</td>
        <td style="${tdr}">${MASK}</td><td style="${tdr}">${r['연간흡수_kgCO2/yr'] == null ? '확인필요' : fmt(r['연간흡수_kgCO2/yr'])}</td>
        <td style="${tdr}">${r['저장량_kgCO2'] == null ? '—' : fmt(r['저장량_kgCO2'])}</td><td style="${td}">${esc(r.계수확인)}</td></tr>`;
    });
    h += `<tr><td style="${td}" colspan="4"><b>합계(확정)</b></td><td style="${tdr}"><b>${fmt(res.totals.annual_kgCO2)}</b></td><td style="${tdr}">${fmt(res.totals.stock_tCO2 * 1000)}</td><td style="${td}"></td></tr></table>`;

    h += `<h2 style="font-size:15px;color:#15463a;margin:18px 0 8px;">3. 산정 가정 및 출처</h2><ul style="font-size:11.5px;color:#333;">`;
    R.assumptions.forEach((a) => { h += `<li><b>${esc(a.항목)}</b>: ${esc(a.내용)} <span style="color:#888;">[출처: ${esc(a.출처)} · 신뢰도: ${esc(a.신뢰도)}]</span></li>`; });
    h += `</ul>`;
    if (R.warnings.length) {
      h += `<h2 style="font-size:15px;color:#a05;margin:18px 0 8px;">4. 경고 / 확인 필요</h2><ul style="font-size:11.5px;color:#a05;">`;
      R.warnings.forEach((w) => { h += `<li>${esc(w)}</li>`; });
      h += `</ul>`;
    }
    h += `<div style="margin-top:24px;border-top:1px solid #ccc;padding-top:8px;font-size:10px;color:#999;">
      ※ 본 보고서의 탄소계수는 ${R.assumptions.some(a=>a.신뢰도==='user') ? '사용자 확정값' : '검증 전 기본 가정값'}이며, 공식 산정에는 국립산림과학원 1차 자료 대조가 필요합니다.
      자동 생성: 탄소흡수 산정 에이전트.</div></div>`;
    return h;
  }
  // ---- 필터/질의 ------------------------------------------------------------
  function showFiltered(kind, label) {
    if (!STATE.results) { runFullAnalysisIfNeeded(); }
    if (!STATE.results) { addBot('먼저 "전체 분석"을 실행하세요.'); return; }
    let list = STATE.results.bySpecies;
    if (kind === '교목') list = list.filter((s) => s.form === '교목');
    else if (kind === '관목') list = list.filter((s) => s.form === '관목');
    else if (kind === 'evergreen') list = list.filter((s) => s.ever === '상록');
    if (!list.length) { addBot(`${label}에 해당하는 수목이 없습니다.`); return; }
    const totA = list.reduce((a, s) => a + (s.co2Annual || 0), 0);
    addBot([
      `**${label}** — ${list.length}종, 연간 ${fmt(round2(totA / 1000))} tCO₂/yr`,
      { type: 'table', head: ['수목명', '규격', '수량', '연간 kgCO₂/yr', '저장 kgCO₂'],
        rows: list.map((s) => [s.name, s.spec, MASK, s.co2Annual == null ? '확인필요' : fmt(s.co2Annual), s.co2Stock == null ? '—' : fmt(s.co2Stock)]),
        foot: ['합계', '', '', fmt(totA), ''] },
      { type: 'html', html: hbarSVG(list.filter((s) => s.co2Annual).map((s) => ({ label: s.name, value: s.co2Annual })), {}) },
    ]);
  }
  function showSpecFilter(text) {
    if (!STATE.results) { addBot('먼저 "전체 분석"을 실행하세요.'); return; }
    const m = text.match(/B\s*(\d+)/i);
    if (!m) return false;
    const thr = parseInt(m[1], 10);
    const ge = /이상|초과|over|≥|>=/.test(text) || !/이하|미만/.test(text);
    const list = STATE.classified.filter((L) => L.cls.spec.type === 'B' && L.cls.spec.mid != null &&
      (ge ? L.cls.spec.mid >= thr : L.cls.spec.mid <= thr));
    if (!list.length) { addBot(`B${thr} ${ge ? '이상' : '이하'} 규격의 수목이 없습니다. (보유 규격: ${[...new Set(STATE.classified.map((L) => L.rec.규격))].join(', ')})`); return true; }
    const totA = list.reduce((a, L) => a + (L.calc.co2Annual || 0), 0);
    addBot([
      `**규격 B${thr} ${ge ? '이상' : '이하'}** — ${list.length}항목, 연간 ${fmt(totA)} kgCO₂/yr`,
      { type: 'table', head: ['수목명', '규격', '수량', '연간 kgCO₂/yr'],
        rows: list.map((L) => [L.rec.수목명, L.rec.규격, MASK, L.calc.co2Annual == null ? '확인필요' : fmt(L.calc.co2Annual)]),
        foot: ['합계', '', '', fmt(totA)] },
    ]);
    return true;
  }
  function runFullAnalysisIfNeeded() { if (!STATE.results && STATE.parsed) { runCalculations(); validate(); computeCost(); } }

  // ---- 라우팅 --------------------------------------------------------------
  // 입력창 자유질의 → (키 있으면) GPT 에이전트가 도구를 호출. 버튼/키없음 → 결정론적 인텐트.
  async function handleUser(text, fromButton) {
    addUser(text);
    if (!fromButton && STATE.ai.apiKey) return agentTurn(text);
    return runIntent(text);
  }
  async function submit() {
    const v = inputEl.value.trim();
    if (!v) return;
    inputEl.value = ''; inputEl.style.height = 'auto';
    await handleUser(v, false);
  }
  // 결정론적 인텐트(빠른 작업 버튼 / 키 없을 때) — LLM 없이 즉시 동작
  async function runIntent(text) {
    const t = norm(text);
    if (/전체분석|분석실행|분석시작|시작|전체|산정/.test(t)) return runFullAnalysis();
    if (/설정|api키|apikey|키설정|openai|지피티|gpt/.test(t)) return openSettings();
    if (/계수|가정|상대생장|편집/.test(t)) { runFullAnalysisIfNeeded(); return openCoeffEditor(); }
    if (/순공사비|공사비|비용|단가|금액/.test(t)) { runFullAnalysisIfNeeded(); return STATE.cost ? addBot([`**순공사비 산출** 💰`, costBlock(), { type: 'html', html: `<div class="ca-note">매칭 ${STATE.cost.matched}건 · 미매칭 ${STATE.cost.unmatched}건. 순공사비·금액·수량은 익명화(***)되며 단가는 비공개입니다.</div>` }]) : addBot('단가 데이터를 찾지 못했습니다.'); }
    if (/pdf|보고서|리포트|출력|다운로드/.test(t)) { runFullAnalysisIfNeeded(); return generatePDF(); }
    if (/json|스키마/.test(t)) { runFullAnalysisIfNeeded(); return showJSON(); }
    if (/검증|이상치|누락/.test(t)) { runFullAnalysisIfNeeded(); return STATE.validation ? addBot([`**검증 결과** 🔍`, validationBlock()]) : addBot('먼저 분석을 실행하세요.'); }
    if (/교목/.test(t)) { runFullAnalysisIfNeeded(); return showFiltered('교목', '교목(상록·낙엽)'); }
    if (/관목/.test(t)) { runFullAnalysisIfNeeded(); return showFiltered('관목', '관목'); }
    if (/상록/.test(t)) { runFullAnalysisIfNeeded(); return showFiltered('evergreen', '상록수'); }
    if (/B\s*\d+/i.test(text)) { runFullAnalysisIfNeeded(); if (showSpecFilter(text)) return; }
    const sp = Object.keys(SPECIES_DICT).find((n) => t.includes(norm(n)));
    if (sp && STATE.results) {
      const found = STATE.results.bySpecies.filter((s) => norm(s.name).includes(norm(sp)));
      if (found.length) return addBot([`**${sp}**`, { type: 'table', head: ['수목명', '규격', '수량', '연간 kgCO₂/yr', '저장 kgCO₂'], rows: found.map((s) => [s.name, s.spec, MASK, fmt(s.co2Annual), fmt(s.co2Stock)]) }]);
    }
    if (/도움|help|사용법|\?/.test(t)) return greet();
    // 매칭 실패 → 키 있으면 GPT 에이전트에 위임, 없으면 안내
    if (STATE.ai.apiKey) return agentTurn(text);
    return addBot([
      '자유로운 질문은 **OpenAI 키 설정** 후 GPT 에이전트가 처리합니다. 아래 작업은 키 없이도 동작해요:',
      { type: 'actions', items: DEFAULT_QUICK.map((q) => ({ t: q.t, q: q.q, go: q.go })).concat([{ t: '⚙︎ OpenAI 설정', fn: () => openSettings() }]) },
    ]);
  }

  function greet() {
    const conn = STATE.ai.apiKey
      ? `OpenAI \`${esc(STATE.ai.model)}\` 연결됨 — 자유롭게 질문하세요. 제가 도구를 사용해 분석·표·차트·PDF를 만들어 드립니다.`
      : 'OpenAI API 키 미설정 — 헤더 ⚙︎에서 키를 넣으면 GPT 에이전트가 활성화됩니다. (아래 빠른 작업 버튼은 키 없이도 동작)';
    addBot([
      `안녕하세요! 저는 **AI챗봇** 🤖 — 조경 유지관리 내역서 기반 **연간 CO₂ 흡수량·순공사비 산정 AI 에이전트**입니다.\n` +
      `엑셀에서 수목 데이터를 추출해 DBH 상대생장식으로 CO₂를 산정하고, 표·차트·PDF 보고서를 만들어 드립니다.` +
      (STATE.parsed ? `\n\n📂 \`${esc(STATE.parsed.sheetNames.join(', '))}\` 로드 — 수목 ${STATE.parsed.treeRecords.length}개 항목 인식.` : '') +
      `\n\n${conn}`,
      { type: 'actions', items: [{ t: '🚀 전체 분석', go: true, q: '전체 분석' }, { t: '⚙︎ OpenAI 설정', fn: () => openSettings() }, { t: '💰 순공사비', q: '순공사비' }] },
    ]);
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ===========================================================================
  // 8. 초기화
  // ===========================================================================
  async function loadEmbeddedXlsx() {
    if (typeof XLSX === 'undefined') { console.warn('[carbon-agent] SheetJS(XLSX) 미로드'); return false; }
    try {
      if (typeof window.__INVENTORY_XLSX_B64__ === 'string' && window.__INVENTORY_XLSX_B64__.length) {
        const wb = XLSX.read(window.__INVENTORY_XLSX_B64__, { type: 'base64' });
        STATE.workbook = wb;
        parseExcel(wb);
        return true;
      }
    } catch (e) { console.error('[carbon-agent] 내장 xlsx 파싱 실패', e); }
    return false;
  }
  // 파일 업로드(드롭/선택) 지원
  function setupFileDrop() {
    window.addEventListener('dragover', (e) => { e.preventDefault(); });
    window.addEventListener('drop', async (e) => {
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (!f || !/\.xlsx?$/.test(f.name)) return;
      e.preventDefault();
      const buf = await f.arrayBuffer();
      STATE.workbook = XLSX.read(buf, { type: 'array' });
      parseExcel(STATE.workbook);
      STATE.results = STATE.classified = STATE.cost = STATE.validation = STATE.report = null;
      STATE.coeffApproved = false;
      togglePanel(true);
      addBot(`📂 새 파일 로드: **${esc(f.name)}** — 수목 ${STATE.parsed.treeRecords.length}개 항목. "전체 분석"을 실행하세요.`);
    });
  }

  async function init() {
    injectStyles();
    // localStorage에서 OpenAI 설정 복원
    STATE.ai.apiKey = lsGet('ca_openai_key', '') || '';
    STATE.ai.model = lsGet('ca_openai_model', 'gpt-4o-mini') || 'gpt-4o-mini';
    STATE.ai.baseUrl = lsGet('ca_openai_base', 'https://api.openai.com/v1') || 'https://api.openai.com/v1';
    buildUI();
    setupFileDrop();
    await loadEmbeddedXlsx();
    updateAiBadge();
    greet();
    console.log('%c[AI챗봇] 준비 완료 · OpenAI 함수호출 에이전트', 'color:#22a06b;font-weight:bold');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();

  // 디버그용 전역 노출
  window.CarbonAgent = { STATE, parseExcel, runCalculations, computeCost, validate, buildReport,
    openCoeffEditor, agentTurn, execTool, openSettings, showTableWindow, TOOLS };
})();
