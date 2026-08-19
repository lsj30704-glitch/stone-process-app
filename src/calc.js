// 엑셀 ③달성률현황 / 내부 달성률현황 / ④만회계획·대응 시트의 수식을 그대로 옮긴 계산 로직.
import { THRESHOLDS } from "./data";

export function statusBadge(rate, th = THRESHOLDS) {
  if (rate === "" || rate === null || rate === undefined) return { label: "–", color: "#999" };
  if (rate >= th.normal) return { label: "✅ 정상", color: "#2e7d32" };
  if (rate >= th.caution) return { label: "⚠️ 주의", color: "#e8a700" };
  return { label: "🚨 위험", color: "#d32f2f" };
}

export function fmtPct(rate) {
  if (rate === "" || rate === null || rate === undefined || isNaN(rate)) return "-";
  return (rate * 100).toFixed(1) + "%";
}

export function fmtNum(n, digits = 1) {
  if (n === "" || n === null || n === undefined || isNaN(n)) return "-";
  return Number(n).toLocaleString("ko-KR", { maximumFractionDigits: digits, minimumFractionDigits: 0 });
}

export function daysBetween(a, b) {
  const d1 = new Date(a);
  const d2 = new Date(b);
  return Math.round((d2 - d1) / 86400000);
}

// ---------- 외부(아파트) ----------

// ②일일실적입력 행 1건의 자동계산 필드 (H=투입인원계, I=계획시공량, K=일일달성률)
export function calcRowExternal(entry, buildings) {
  const totalWorkers = (entry.masonry || 0) + (entry.caulking || 0) + (entry.truss || 0) + (entry.scaffold || 0);
  const b = buildings.find((x) => x.dong === entry.dong);
  const planned = b ? b.dailyPlan : "";
  const actual = entry.actual;
  const rate = planned !== "" && planned > 0 && actual !== "" && actual !== null && actual !== undefined
    ? actual / planned
    : "";
  return { ...entry, totalWorkers, planned, rate };
}

// ③달성률현황
export function calcExternalDashboard(buildings, logs, th = THRESHOLDS) {
  const rows = logs.map((l) => calcRowExternal(l, buildings));
  const totalPlanArea = buildings.reduce((s, b) => s + b.totalArea, 0);
  const cumActual = rows.reduce((s, r) => s + (Number(r.actual) || 0), 0);
  const cumPlan = rows.reduce((s, r) => s + (r.dong ? Number(r.planned) || 0 : 0), 0);
  const overallRate = totalPlanArea > 0 ? cumActual / totalPlanArea : "";
  const periodRate = cumPlan > 0 ? cumActual / cumPlan : "";
  const totalWorkers = rows.reduce((s, r) => s + (Number(r.totalWorkers) || 0), 0);
  const perWorker = totalWorkers > 0 ? cumActual / totalWorkers : "-";
  const disasterDays = rows.filter((r) => r.disaster === "Y (천재지변)").length;
  const totalDays = rows.filter((r) => r.date).length;

  const byBuilding = buildings.map((b) => {
    const buildingActual = rows.filter((r) => r.dong === b.dong).reduce((s, r) => s + (Number(r.actual) || 0), 0);
    const rate = b.totalArea > 0 ? buildingActual / b.totalArea : 0;
    const remain = Math.max(0, b.totalArea - buildingActual);
    return {
      dong: b.dong,
      planArea: b.totalArea,
      cumActual: buildingActual,
      rate,
      remain,
      endDate: b.endDate,
      status: statusBadge(rate, th),
    };
  });

  return { totalPlanArea, cumActual, cumPlan, overallRate, periodRate, totalWorkers, perWorker, disasterDays, totalDays, byBuilding };
}

// ---------- 내부(세대) ----------

// ②일일실적입력(내부) 행 1건의 자동계산 필드 (G=투입인원계, H=계획시공량(세대)=G*3, J=일일달성률)
export function calcRowInternal(entry) {
  const totalWorkers = (entry.masonry || 0) + (entry.caulking || 0) + (entry.truss || 0);
  const planned = totalWorkers * 3;
  const actual = entry.actual;
  const rate = planned > 0 && actual !== "" && actual !== null && actual !== undefined
    ? actual / planned
    : "";
  return { ...entry, totalWorkers, planned, rate };
}

// 내부 달성률현황
export function calcInternalDashboard(buildingsInternal, logs, th = THRESHOLDS) {
  const rows = logs.map((l) => calcRowInternal(l));
  const totalUnits = buildingsInternal.reduce((s, b) => s + b.totalUnits, 0);
  const optionUnits = buildingsInternal.reduce((s, b) => s + b.optionUnits, 0);
  const normalUnits = buildingsInternal.reduce((s, b) => s + b.normalUnits, 0);
  const cumActual = rows.reduce((s, r) => s + (Number(r.actual) || 0), 0);
  const cumPlan = rows.reduce((s, r) => s + (r.dong ? Number(r.planned) || 0 : 0), 0);
  const overallRate = totalUnits > 0 ? cumActual / totalUnits : "";
  const periodRate = cumPlan > 0 ? cumActual / cumPlan : "";
  const totalWorkers = rows.reduce((s, r) => s + (Number(r.totalWorkers) || 0), 0);
  const totalDays = rows.filter((r) => r.date).length;

  const byBuilding = buildingsInternal.map((b) => {
    const buildingActual = rows.filter((r) => r.dong === b.dong).reduce((s, r) => s + (Number(r.actual) || 0), 0);
    const rate = b.totalUnits > 0 ? buildingActual / b.totalUnits : 0;
    const remain = Math.max(0, b.totalUnits - buildingActual);
    return {
      dong: b.dong,
      totalUnits: b.totalUnits,
      optionUnits: b.optionUnits,
      cumActual: buildingActual,
      rate,
      remain,
      status: statusBadge(rate, th),
    };
  });

  return { totalUnits, optionUnits, normalUnits, cumActual, cumPlan, overallRate, periodRate, totalWorkers, totalDays, byBuilding };
}

// ---------- 엑셀 달성률현황 시트 병합 (하이브리드) ----------
//
// 표시값 = ③달성률현황 시트의 값
//        + 그 시트 수식 범위(coveredMaxRow) 밖에 있는 ②시트 실적
//        + 아직 엑셀에 올리지 않은 앱 대기 입력분
//
// 이렇게 하는 이유:
//  1) 엑셀 화면 숫자와 앱 숫자가 같아야 대조가 된다 → 시트 값을 기준으로 삼음
//  2) 그런데 시트의 SUMIF 범위가 C5:C114처럼 고정이라 115행 이후 실적이 엑셀에서도 빠진다
//     → 범위 밖 행을 앱이 보충해서 더함
//  3) 앱에서 방금 입력한(아직 업로드 전) 건도 더해야 현장에서 바로 쓸 수 있다
//
// 각 동에 대해 근거를 basis 객체로 남겨 화면에 "532 (엑셀) + 7.7 (추가)" 형태로 표시합니다.

function sumActual(list) {
  return (list || []).reduce((s, x) => s + (Number(x?.actual) || 0), 0);
}

/**
 * @param {object} dash        기존 calcExternalDashboard/calcInternalDashboard 결과 (폴백 겸 보조지표용)
 * @param {object|null} achv   graphSync.readAchievement 결과
 * @param {array} logs         ②시트에서 읽은 실적 행 (row 번호 포함)
 * @param {array} pending      아직 업로드 안 된 앱 입력분
 * @param {object} th          임계값
 * @param {string} planKey     "planArea"(외부) | "totalUnits"(내부)
 */
export function mergeAchievement(dash, achv, logs, pending, th = THRESHOLDS, planKey = "planArea") {
  if (!achv || !achv.byDong?.length) {
    // 시트를 못 읽었으면 기존 재계산 결과를 그대로 사용 (동작 보장)
    return { ...dash, source: "recalc" };
  }

  const covered = achv.coveredMaxRow;
  const pendingList = pending || [];

  const byBuilding = achv.byDong.map((row) => {
    const sheetVal = Number(row.cumActual) || 0;

    // 시트 수식 범위 밖의 행 (row 번호를 모르면 보충하지 않음 — 중복 합산 방지)
    const extraRows = covered
      ? (logs || []).filter((l) => l.dong === row.dong && Number(l.row) > covered)
      : [];
    const extra = sumActual(extraRows);

    const pendingRows = pendingList.filter((p) => p.dong === row.dong);
    const pendingSum = sumActual(pendingRows);

    const cumActual = sheetVal + extra + pendingSum;
    const plan = Number(row[planKey] ?? row.planArea ?? row.totalUnits) || 0;
    const rate = plan > 0 ? cumActual / plan : 0;

    return {
      dong: row.dong,
      planArea: plan,
      totalUnits: row.totalUnits,
      optionUnits: row.optionUnits,
      cumActual,
      rate,
      remain: Math.max(0, plan - cumActual),
      endDate: row.endDate,
      // 시트가 계산해 둔 상태 문구가 있으면 쓰되, 보충분이 있으면 다시 판정
      status: extra || pendingSum ? statusBadge(rate, th) : (row.statusText ? { label: row.statusText, color: statusBadge(rate, th).color } : statusBadge(rate, th)),
      basis: {
        sheet: sheetVal,
        extra,
        pending: pendingSum,
        extraRowNums: extraRows.map((l) => l.row),
        pendingCount: pendingRows.length,
      },
    };
  });

  const totalPlanArea = byBuilding.reduce((s, b) => s + (Number(b.planArea) || 0), 0);
  const cumActual = byBuilding.reduce((s, b) => s + b.cumActual, 0);
  const sheetCum = byBuilding.reduce((s, b) => s + b.basis.sheet, 0);
  const addedCum = cumActual - sheetCum;

  const overallRate = totalPlanArea > 0 ? cumActual / totalPlanArea : "";

  // [주의] 상단 "전체 현황 요약" 지표는 시트 요약값(C6~C14)을 쓰지 않고 앱 재계산값을 유지합니다.
  // 이 엑셀의 요약 수식들이 ②시트 5:114 범위로 고정돼 있어 최신 행이 빠지고,
  // 특히 총 투입인원은 SUM(D5:D114)로 "석공(D열)"만 세고 코킹·트러스·비계(E~G)가 빠져 있습니다.
  // 동별 표는 사용자가 엑셀 화면과 대조해야 하므로 시트 값을 기준으로 쓰지만,
  // 요약 지표까지 틀린 값을 그대로 옮길 이유는 없습니다.
  // 시트 요약값은 sheetSummary로 남겨 두어 대조·디버깅에 쓸 수 있게 합니다.
  return {
    ...dash,
    source: "sheet",
    sheetName: achv.sheetName,
    coveredMaxRow: covered,
    byBuilding,
    totalPlanArea,
    cumActual,
    overallRate,
    sheetSummary: achv.summary || {},
    basis: { sheetCum, addedCum },
  };
}

// ---------- ④만회계획·대응 ----------

export function calcRecoveryPlan(dong, buildings, logs, today = new Date(), th = THRESHOLDS) {
  const b = buildings.find((x) => x.dong === dong);
  if (!b) return null;
  const dongLogs = logs.filter((l) => l.dong === dong);
  const cumActual = dongLogs.reduce((s, l) => s + (Number(l.actual) || 0), 0);
  // 실행 일일생산성(엑셀 ④만회계획 C16)과 동일 기준:
  //   1인당 생산성 = 누적실적 ÷ SUMIFS(투입인원계, 실적>0)
  // 즉 "실적이 발생한 날"의 투입인원(man-day)만 분모로 삼는다.
  // (자재반입·천재지변 등 실적 0인 날의 인원은 제외 → 실제로 시공한 팀의 1인당 시공량)
  const cumWorkers = dongLogs.reduce(
    (s, l) =>
      Number(l.actual) > 0
        ? s + (Number(l.masonry) || 0) + (Number(l.caulking) || 0) + (Number(l.truss) || 0) + (Number(l.scaffold) || 0)
        : s,
    0
  );
  const planArea = b.totalArea;
  const currentRate = planArea > 0 ? cumActual / planArea : 0;
  const remainArea = Math.max(0, planArea - cumActual);
  const endDate = b.endDate;
  const remainDays = Math.max(0, daysBetween(today, endDate));
  const availableDays = Math.max(1, Math.round(remainDays * 0.85));
  const baseWorkers = b.baseWorkers;
  // 1인당 일일생산성은 동마다 다름 — 실제 누적 실적(㎡) ÷ 실제 누적 투입인원(man-day).
  // 아직 투입 실적이 없는 동은 엑셀 ①기준정보의 기준값(1인당 일일 시공량)으로 폴백.
  const hasActual = cumWorkers > 0 && cumActual > 0;
  const productivity = hasActual ? cumActual / cumWorkers : th.productivityPerWorker;
  const productivitySource = hasActual ? "actual" : "base";
  const currentCapacity = baseWorkers * productivity;
  const neededDaily = remainArea === 0 ? "완료" : remainArea / availableDays;
  const extraWorkers = neededDaily === "완료" ? 0 : Math.max(0, Math.ceil(neededDaily / productivity) - baseWorkers);
  let verdict;
  if (neededDaily === "완료") verdict = "✅ 계획 수량 달성 완료";
  else if (neededDaily <= currentCapacity) verdict = "✅ 현인원으로 만회 가능";
  else if (extraWorkers <= 3) verdict = "⚠️ 소폭 증원 필요";
  else verdict = "🚨 대폭 증원 또는 공기연장 검토";

  return {
    dong, planArea, cumActual, cumWorkers, currentRate, remainArea, endDate, remainDays,
    availableDays, baseWorkers, productivity, productivitySource, currentCapacity, neededDaily, extraWorkers, verdict,
  };
}

export const RECOVERY_CHECKLIST = [
  "잔여 물량 및 공기 재산출 완료 여부",
  "추가 투입 인원 확보 계획 수립 (협력사 협의)",
  "자재 수급 일정 재조정 (석재 납기 확인)",
  "장비 추가 투입 가능 여부 확인 (양중, 리프트 등)",
  "만회계획서 발주처 제출 및 서명 확인",
  "주간·월간 공정회의에 만회 계획 반영",
  "인접 공종(방수, 코킹) 일정 재조율",
  "만회 기간 중 안전 관리 강화 계획 수립",
];

export const DISASTER_MANUAL = [
  { step: "STEP 1", title: "당일 즉시 조치", detail: "기상특보 공문 수신 즉시 현장 일지에 기록 · 출력인원 대피 및 안전 확보 조치 · 일일실적입력에 천재지변여부=Y, 사유코드 입력" },
  { step: "STEP 2", title: "증빙 자료 확보", detail: "기상청 특보 캡처/인쇄 보관 · 현장 사진(피해상황) 촬영 후 저장 · 공사일지·안전일지 상세 기록(감리 서명 필수)" },
  { step: "STEP 3", title: "발주처·감리 보고", detail: "당일 또는 익일 오전 서면 통보 · 불가항력 조항 확인 후 공기연장 신청서 제출 준비 · 인정 시 달성률 분모에서 해당일 제외" },
  { step: "STEP 4", title: "복구 계획 수립", detail: "만회계획 자동 산출로 잔여량·필요인원 확인 · 다음날부터 만회 일일 생산량 목표 설정 · 필요시 2교대/주말 특근 계획" },
  { step: "STEP 5", title: "만회 실행 & 모니터링", detail: "매일 실적 입력 · 달성률현황 대시보드로 만회 진행 확인 · 달성률 80% 미만 지속 시 공기연장/추가인원 협의" },
];

// ---------- 발주현황 (⑥ 석재발주 시트) ----------
// 구분(gubun)별로 표현 단위를 다르게 집계합니다.
//   · 두겁          → 길이 M
//   · 창대 + 창주위 → 길이 M (한 묶음)  ※ 저층부의 "창틀·창대"도 창호 둘레재라 여기에 포함
//   · 그 외(벽체·P석·연마·버너 등) → 면적 ㎡
// 새 행이 밑에 추가돼도 구분 이름만 같은 규칙이면 자동으로 같은 칸에 합산됩니다.
const ORDER_WINDOW_RE = /(창주위|창대)/; // 창대석·창주위석·창틀·창대 포함

export function orderMetric(gubun) {
  const g = String(gubun || "");
  if (g.includes("두겁")) return "cope";    // M
  if (ORDER_WINDOW_RE.test(g)) return "window"; // M (창대+창주위)
  return "area";                             // ㎡
}

// 동 이름 통합: "104동(추가)"·"104동(1차)"→"104동", "101동 호이스트"·"112동(저층부)"→기본 동번호.
// 숫자+동 패턴을 뽑아 기본 동으로 합치고, 게스트하우스처럼 숫자 없는 건 그대로 둡니다.
export function normalizeDong(dong) {
  const s = String(dong || "").trim();
  const m = s.match(/(\d+)\s*동/);
  return m ? `${m[1]}동` : s;
}

export function calcOrderStatus(rows) {
  const byDongMap = new Map();
  for (const r of rows || []) {
    if (!r || !r.dong) continue;
    const dong = normalizeDong(r.dong);
    if (!byDongMap.has(dong)) byDongMap.set(dong, new Map());
    const stoneMap = byDongMap.get(dong);
    const cur = stoneMap.get(r.stone) || { stone: r.stone, area: 0, window: 0, cope: 0 };
    const metric = orderMetric(r.gubun);
    if (metric === "cope") cur.cope += Number(r.m) || 0;
    else if (metric === "window") cur.window += Number(r.m) || 0;
    else cur.area += Number(r.m2) || 0;
    stoneMap.set(r.stone, cur);
  }
  const byDong = [];
  let grand = { area: 0, window: 0, cope: 0 };
  for (const [dong, stoneMap] of byDongMap) {
    const stones = [...stoneMap.values()];
    const total = stones.reduce(
      (s, x) => ({ area: s.area + x.area, window: s.window + x.window, cope: s.cope + x.cope }),
      { area: 0, window: 0, cope: 0 }
    );
    byDong.push({ dong, stones, total });
    grand = { area: grand.area + total.area, window: grand.window + total.window, cope: grand.cope + total.cope };
  }
  // 동 순서: 숫자(동번호) 오름차순 → 숫자 없는 이름(게스트하우스 등)은 한글순으로 뒤에
  byDong.sort((a, b) => {
    const na = parseInt(a.dong, 10);
    const nb = parseInt(b.dong, 10);
    const aNum = !isNaN(na);
    const bNum = !isNaN(nb);
    if (aNum && bNum) return na - nb;
    if (aNum) return -1;
    if (bNum) return 1;
    return a.dong.localeCompare(b.dong, "ko");
  });
  return { byDong, grand, dongCount: byDong.length };
}

// ============================================================================
// 공기(工期) 관리 — 착수지연 / 천재지변 손실 / 인력 소요 곡선 / 공기연장 클레임
// ============================================================================
//
// [설계 전제 — 사용자 확정 사항]
//  1. 부분영향(P)은 실적 기반으로 손실일을 자동 산정한다.
//     그날 실적이 "정상 생산성 × 투입인원"의 몇 %인지 보고 모자란 만큼을 손실로 친다.
//  2. 천재지변(Y)은 조업 불가이므로 1.0일 고정.
//  3. 천재지변 행에 동을 적지 않으면, 직전 조업일에 작업하던 동을 그대로 물려받는다.
//     (비 오는 날 동까지 고르지 않아도 되게)
//  4. 착수 지연 = 계획착수일 → 첫 실적일. 아직 착수 못 한 동은 오늘까지로 계속 늘어난다.
//     이 지연일수만큼 완료예정일이 뒤로 밀린다.
//  5. 완료된 동(누적 ≥ 계획, 또는 체크리스트에서 수동 완료)은 앞으로 시공량이 0이므로
//     인력 소요 곡선과 잔여 계산에서 완전히 제외한다.
//
// [생산성에 대한 주의]
//  인당 생산성 편차가 큰 것은 정상이다. 벽체 판재는 많이 나오고, 창대석·창주위석 같은
//  작은 돌, 상부 두겁석, 그리고 작업 첫날은 적게 나온다. 한 동을 끝내면 이 조합이 섞여
//  동 단위 평균으로 수렴하므로, 생산성은 항상 "동 단위 누적 평균"으로만 쓴다.
//  개별 일자 생산성으로 이상치 경고 같은 것을 띄우지 않는다.

const DAY_MS = 86400000;

export function toDate(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (!v) return null;
  const d = new Date(`${String(v).slice(0, 10)}T00:00:00`);
  return isNaN(d) ? null : d;
}
export function isoOf(d) {
  if (!d) return "";
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}
function addDays(d, n) {
  return new Date(d.getTime() + n * DAY_MS);
}
function diffDays(a, b) {
  return Math.round((toDate(b) - toDate(a)) / DAY_MS);
}
// 일요일만 휴무로 본다 (현장 관행: 토요일 조업)
function isWorkday(d) {
  return d.getDay() !== 0;
}
function clamp(v, lo, hi) {
  return Math.max(lo, Math.min(hi, v));
}
function workersOf(l) {
  return (Number(l.masonry) || 0) + (Number(l.caulking) || 0) + (Number(l.truss) || 0) + (Number(l.scaffold) || 0);
}
function disasterKind(l) {
  const s = String(l?.disaster ?? "").trim().toUpperCase();
  if (s.startsWith("Y")) return "Y";
  if (s.startsWith("P")) return "P";
  if (s.startsWith("N")) return "N";
  return "";
}

// ---------- 1) 천재지변 행의 동 상속 ----------
// 동이 비어 있는 행에, 직전 "실적이 발생한 날"에 작업하던 동 목록을 물려준다.
// 반환: 원본 로그에 dongs[] (실제 적용 대상 동 배열)를 붙인 새 배열.
export function expandLogDongs(logs) {
  const rows = (logs || []).map((l) => ({ ...l, _d: toDate(l.date) })).filter((l) => l._d);
  rows.sort((a, b) => a._d - b._d || (a.row || 0) - (b.row || 0));

  const byDate = new Map();
  for (const r of rows) {
    const k = isoOf(r._d);
    if (!byDate.has(k)) byDate.set(k, []);
    byDate.get(k).push(r);
  }

  let lastActive = [];
  for (const k of [...byDate.keys()].sort()) {
    const group = byDate.get(k);
    for (const r of group) {
      const dong = String(r.dong ?? "").trim();
      r.dongs = dong ? [dong] : [...lastActive];
      r.inherited = !dong && r.dongs.length > 0;
    }
    // 그날 실제로 시공이 있었던 동만 "직전 작업동"으로 기억
    const worked = [...new Set(group.filter((r) => Number(r.actual) > 0).flatMap((r) => r.dongs))];
    if (worked.length) lastActive = worked;
  }
  return rows;
}

// ---------- 2) 동별 정상 생산성 (손실일 산정 기준) ----------
// 정상(N)일의 인당 시공량을 동별로 낸다. 그 동에 정상일이 없으면 전체 정상일 평균,
// 그것도 없으면 그 동의 전체 평균, 마지막으로 기준정보의 1인당 일일시공량을 쓴다.
export function normalProductivity(logs, th = THRESHOLDS) {
  const perDong = {};
  let gAct = 0, gPpl = 0;
  const anyDong = {};
  for (const l of logs || []) {
    const act = Number(l.actual) || 0;
    const ppl = workersOf(l);
    if (act <= 0 || ppl <= 0) continue;
    const dongs = l.dongs?.length ? l.dongs : (l.dong ? [l.dong] : []);
    for (const d of dongs) {
      anyDong[d] = anyDong[d] || { act: 0, ppl: 0 };
      anyDong[d].act += act / dongs.length;
      anyDong[d].ppl += ppl / dongs.length;
      if (disasterKind(l) === "N") {
        perDong[d] = perDong[d] || { act: 0, ppl: 0 };
        perDong[d].act += act / dongs.length;
        perDong[d].ppl += ppl / dongs.length;
        gAct += act / dongs.length;
        gPpl += ppl / dongs.length;
      }
    }
  }
  const globalNormal = gPpl > 0 ? gAct / gPpl : null;
  const base = th?.productivityPerWorker || 10;
  const get = (dong) => {
    const p = perDong[dong];
    if (p && p.ppl > 0) return { value: p.act / p.ppl, source: "dong-normal" };
    if (globalNormal) return { value: globalNormal, source: "global-normal" };
    const a = anyDong[dong];
    if (a && a.ppl > 0) return { value: a.act / a.ppl, source: "dong-all" };
    return { value: base, source: "base" };
  };
  return { get, globalNormal, base };
}

// ---------- 3) 일자별 공기 손실 산정 ----------
// Y = 1.0일, N = 0, P = 1 - (실적 / (인원 × 정상생산성)) 을 0~1로 자른 값.
// 인원이 0인 날(아예 못 나온 날)은 1.0일.
export function calcDayLoss(log, normProd) {
  const kind = disasterKind(log);
  if (kind === "Y") return 1;
  if (kind !== "P") return 0;
  const ppl = workersOf(log);
  if (ppl <= 0) return 1;
  const expected = ppl * (normProd || 10);
  if (expected <= 0) return 0;
  const ratio = (Number(log.actual) || 0) / expected;
  return clamp(1 - ratio, 0, 1);
}

// ---------- 4) 동별 공기 현황 ----------
// buildings: [{dong, totalArea, startDate, endDate, baseWorkers, ...}]
// checklist: [{dong, gates:{}, done, reason}]
export function calcScheduleStatus(buildings, logs, checklist = [], th = THRESHOLDS, today = new Date()) {
  const T = toDate(today);
  const rows = expandLogDongs(logs);
  const np = normalProductivity(rows, th);
  const cl = new Map((checklist || []).map((c) => [c.dong, c]));

  const acc = {};
  const ensure = (d) => (acc[d] = acc[d] || { act: 0, ppl: 0, first: null, firstSeen: null, last: null, lossBefore: 0, lossAfter: 0, lossRows: [] });

  // 1차: 동별 첫 실적일(first)과 첫 등장일(firstSeen).
  // firstSeen은 실적이 0이어도 그 동에 인원이 배치돼 로그에 처음 잡힌 날 = 실질 착수일.
  for (const l of rows) {
    for (const d of l.dongs || []) {
      const a = ensure(d);
      if (!a.firstSeen || l._d < a.firstSeen) a.firstSeen = l._d;
    }
    const act = Number(l.actual) || 0;
    if (act <= 0) continue;
    for (const d of l.dongs || []) {
      const a = ensure(d);
      if (!a.first || l._d < a.first) a.first = l._d;
      if (!a.last || l._d > a.last) a.last = l._d;
      a.act += act / (l.dongs.length || 1);
      a.ppl += workersOf(l) / (l.dongs.length || 1);
    }
  }
  // 2차: 손실일. 경계는 firstSeen(실질 착수일).
  // 착수지연도 firstSeen까지만 세므로 두 값이 겹치지 않는다 → 이중계산 없음.
  for (const l of rows) {
    const loss = calcDayLoss(l, np.get((l.dongs || [])[0] || "").value);
    if (loss <= 0) continue;
    for (const d of l.dongs || []) {
      const a = ensure(d);
      const share = loss / (l.dongs.length || 1);
      if (a.firstSeen && l._d >= a.firstSeen) a.lossAfter += share;
      else a.lossBefore += share;
      a.lossRows.push({
        date: isoOf(l._d), dong: d, loss: share, kind: disasterKind(l),
        reason: l.reason || "", note: l.note || "", memo: l.memo || "",
        workers: workersOf(l), actual: Number(l.actual) || 0, inherited: !!l.inherited,
      });
    }
  }

  return (buildings || []).map((b) => {
    const a = acc[b.dong] || ensure(b.dong);
    const plan = Number(b.totalArea) || 0;
    const cum = a.act;
    const remain = Math.max(0, plan - cum);
    const manualDone = !!cl.get(b.dong)?.done;
    const autoDone = plan > 0 && cum >= plan;
    const done = manualDone || autoDone;

    const plannedStart = toDate(b.startDate);
    const plannedEnd = toDate(b.endDate);
    // 착수 지연: 계획착수일 → 첫 실적일. 미착수면 오늘까지(계속 늘어남).
    // 기준일 = 첫 실적일과 첫 등장일 중 빠른 쪽. 실적이 아직 없어도 인원이 들어간 날이
    // 있으면 그날 착수한 것으로 본다(그 뒤는 손실일로 잡히므로 중복되지 않는다).
    let startDelay = 0;
    let startBasis = "on-time";
    const actualStart = a.first && a.firstSeen ? (a.first < a.firstSeen ? a.first : a.firstSeen) : (a.first || a.firstSeen);
    if (plannedStart && !done) {
      if (actualStart) {
        startDelay = Math.max(0, diffDays(plannedStart, actualStart));
        startBasis = startDelay > 0 ? "late-start" : "on-time";
      } else if (T > plannedStart) {
        startDelay = diffDays(plannedStart, T);
        startBasis = "not-started";
      }
    }
    const adjustedEnd = plannedEnd ? addDays(plannedEnd, startDelay) : null;
    const prod = np.get(b.dong);

    const gate = cl.get(b.dong);
    const gatesOk = gate ? GATE_KEYS.every((k) => gate.gates?.[k]) : false;
    const gatesMissing = gate ? GATE_KEYS.filter((k) => !gate.gates?.[k]) : GATE_KEYS.slice();

    return {
      dong: b.dong,
      planArea: plan,
      cumActual: cum,
      remain: done ? 0 : remain,
      done, doneBy: manualDone ? "manual" : autoDone ? "auto" : null,
      plannedStart: isoOf(plannedStart),
      plannedEnd: isoOf(plannedEnd),
      adjustedEnd: isoOf(adjustedEnd),
      firstActual: isoOf(a.first),
      firstSeen: isoOf(a.firstSeen),
      actualStart: isoOf(actualStart),
      lastActual: isoOf(a.last),
      startDelay, startBasis,
      lossAfter: a.lossAfter, lossBefore: a.lossBefore,
      lossRows: a.lossRows,
      productivity: prod.value, productivitySource: prod.source,
      baseWorkers: Number(b.baseWorkers) || 0,
      gate: gate || null, gatesOk, gatesMissing,
      delayReason: gate?.reason || "",
    };
  });
}

export const GATE_KEYS = ["scaffold", "window", "frame", "material", "shopdwg"];
export const GATE_LABELS = {
  scaffold: "비계 설치",
  window: "창호 취부",
  frame: "골조·먹매김",
  material: "자재 반입",
  shopdwg: "샵도면 승인",
};

// ---------- 5) 인력 소요 곡선 ----------
// 완료된 동은 앞으로 시공량이 0이므로 제외한다.
// 남은 동만, 조정 완료예정일(= 원 완료예정일 + 착수지연)까지 균등 배분해서
// 날짜별 필요 인원을 쌓아 올린다. 일요일은 제외.
export function calcManpowerCurve(scheduleRows, logs, today = new Date(), horizonDays = 120) {
  const T = toDate(today);
  const perDate = new Map(); // iso → { total, byDong: {dong: n} }
  const add = (d, dong, n) => {
    const k = isoOf(d);
    if (!perDate.has(k)) perDate.set(k, { date: k, total: 0, byDong: {} });
    const e = perDate.get(k);
    e.total += n;
    e.byDong[dong] = (e.byDong[dong] || 0) + n;
  };

  const active = (scheduleRows || []).filter((r) => !r.done && r.remain > 0);
  const overdue = [];

  for (const r of active) {
    const end = toDate(r.adjustedEnd) || toDate(r.plannedEnd);
    const plannedStart = toDate(r.plannedStart);
    const start = plannedStart && plannedStart > T ? plannedStart : T;
    const prod = r.productivity > 0 ? r.productivity : 10;

    const days = [];
    if (end) {
      for (let d = new Date(start); d <= end && days.length < horizonDays; d = addDays(d, 1)) {
        if (isWorkday(d)) days.push(new Date(d));
      }
    }
    if (!days.length) {
      // 조정 완료예정일이 이미 지났거나 산출 불가 → 공기 초과. 오늘 하루에 몰아 표시.
      overdue.push({ dong: r.dong, remain: r.remain, end: r.adjustedEnd || r.plannedEnd });
      if (isWorkday(T)) add(T, r.dong, r.remain / prod);
      continue;
    }
    const perDay = r.remain / days.length / prod;
    for (const d of days) add(d, r.dong, perDay);
  }

  const series = [...perDate.values()].sort((a, b) => (a.date < b.date ? -1 : 1));
  const peak = series.reduce((m, s) => (s.total > m.total ? s : m), { total: 0, date: "" });

  // 가용 인원 기준선: 최근 30일 실투입 중 최대값 (없으면 0)
  const recent = (logs || [])
    .map((l) => ({ d: toDate(l.date), w: workersOf(l) }))
    .filter((x) => x.d && diffDays(x.d, T) <= 30 && diffDays(x.d, T) >= 0);
  const byDay = new Map();
  for (const x of recent) {
    const k = isoOf(x.d);
    byDay.set(k, (byDay.get(k) || 0) + x.w);
  }
  const capacity = byDay.size ? Math.max(...byDay.values()) : 0;
  const avgRecent = byDay.size ? [...byDay.values()].reduce((a, b) => a + b, 0) / byDay.size : 0;

  const shortDays = series.filter((s) => capacity > 0 && s.total > capacity);

  return {
    series, peak, capacity, avgRecent, overdue,
    dongs: active.map((r) => r.dong),
    shortDays: shortDays.length,
    firstShortDate: shortDays[0]?.date || "",
  };
}

// ---------- 6) 공기연장 클레임 집계 ----------
// 이미 발생한 손실이므로 완료된 동의 과거 손실도 포함한다(청구 근거이기 때문).
// 다만 착수 전 손실(lossBefore)은 착수지연에 이미 반영돼 있으므로 별도로 구분해 표시한다.
export function calcClaimSummary(scheduleRows, logs) {
  const rows = expandLogDongs(logs);
  const all = [];
  for (const r of scheduleRows || []) all.push(...r.lossRows);

  const byDong = (scheduleRows || []).map((r) => ({
    dong: r.dong,
    startDelay: r.startDelay,
    delayReason: r.delayReason,
    lossAfter: r.lossAfter,
    lossBefore: r.lossBefore,
    claimDays: r.startDelay + r.lossAfter,
    done: r.done,
    plannedEnd: r.plannedEnd,
    adjustedEnd: r.adjustedEnd,
  })).sort((a, b) => b.claimDays - a.claimDays);

  const byMonth = {};
  for (const l of all) {
    const m = l.date.slice(0, 7);
    if (!byMonth[m]) byMonth[m] = { month: m, loss: 0, y: 0, p: 0, days: new Set() };
    byMonth[m].loss += l.loss;
    if (l.kind === "Y") byMonth[m].y += 1;
    if (l.kind === "P") byMonth[m].p += 1;
    byMonth[m].days.add(l.date);
  }
  const months = Object.values(byMonth)
    .map((m) => ({ ...m, days: m.days.size }))
    .sort((a, b) => (a.month < b.month ? -1 : 1));

  const counts = { Y: 0, P: 0, N: 0 };
  for (const l of rows) {
    const k = disasterKind(l);
    if (counts[k] !== undefined) counts[k] += 1;
  }

  const totalStartDelay = byDong.reduce((s, d) => s + d.startDelay, 0);
  const totalLoss = byDong.reduce((s, d) => s + d.lossAfter, 0);

  // 근거 리스트: 손실이 큰 순 → 날짜 순
  const evidence = all
    .filter((l) => l.loss > 0.01)
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : b.loss - a.loss));

  return {
    byDong, months, counts, evidence,
    totalStartDelay, totalLoss,
    totalClaimDays: totalStartDelay + totalLoss,
    totalRows: rows.length,
  };
}
