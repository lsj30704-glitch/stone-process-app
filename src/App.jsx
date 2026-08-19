import { useEffect, useMemo, useRef, useState } from "react";
import {
  BUILDINGS_EXTERNAL, BUILDINGS_INTERNAL, SEED_LOGS_EXTERNAL, SEED_LOGS_INTERNAL,
  DISASTER_OPTIONS, REASON_CODES, THRESHOLDS, SEED_ORDER_ROWS,
} from "./data";
import {
  calcRowExternal, calcRowInternal, calcExternalDashboard, calcInternalDashboard,
  calcRecoveryPlan, calcOrderStatus, mergeAchievement, fmtPct, fmtNum, RECOVERY_CHECKLIST, DISASTER_MANUAL,
  calcScheduleStatus, calcManpowerCurve, calcClaimSummary, GATE_KEYS, GATE_LABELS,
} from "./calc";
import { Storage } from "./storage";
import * as Graph from "./graphSync";

function today() {
  return new Date().toISOString().slice(0, 10);
}

const EXTRA_INTERNAL_DONG = "게스트하우스";

// ===== 공사 범위(scope) 정의 =====
// 면적(m²)형 공사 범위 — 외부/호이스트/부대시설은 동일한 시트 구조라 같은 계산 로직을 공유.
// 새 면적형 범위가 생기면 엑셀에 시트(②일일실적입력(○○))와 기준정보 블록(▶ ○○ 수량)을 추가하고
// 여기에 한 줄만 더 넣으면 앱 화면에 자동으로 추가됩니다.
const AREA_SCOPES = [
  { key: "external", label: "외부(아파트)", sheet: "②일일실적입력", planLabel: "동별 시공 계획 수량", planRange: "①기준정보!$B$6:$G$13" },
  { key: "hoist", label: "호이스트", sheet: "②일일실적입력(호이스트)", planLabel: "호이스트 시공 계획 수량", planRange: "①기준정보!$B$18:$G$25" },
  { key: "facility", label: "부대시설", sheet: "②일일실적입력(부대시설)", planLabel: "부대시설 계획 수량", planRange: "①기준정보!$B$30:$G$37" },
];
const AREA_KEYS = AREA_SCOPES.map((s) => s.key);
const AREA_STORAGE = {
  external: { b: Storage.KEYS.buildingsExternal, l: Storage.KEYS.logsExternal, p: Storage.KEYS.pendingExternal },
  hoist: { b: Storage.KEYS.buildingsHoist, l: Storage.KEYS.logsHoist, p: Storage.KEYS.pendingHoist },
  facility: { b: Storage.KEYS.buildingsFacility, l: Storage.KEYS.logsFacility, p: Storage.KEYS.pendingFacility },
};
const SCOPE_TABS = [...AREA_SCOPES.map((s) => ({ key: s.key, label: s.label })), { key: "internal", label: "내부(세대)" }];

// 투입인원 입력 칸 기본 구성 (동기화 전, 또는 헤더를 못 읽을 때 사용).
// 동기화하면 각 범위 시트의 헤더(4행)에서 실제 라벨/개수를 읽어 자동으로 대체됩니다.
const DEFAULT_AREA_FIELDS = [
  { key: "masonry", label: "석공" },
  { key: "caulking", label: "코킹" },
  { key: "truss", label: "트러스" },
  { key: "scaffold", label: "비계" },
];
function defaultAreaFieldsMap() {
  return Object.fromEntries(AREA_SCOPES.map((s) => [s.key, DEFAULT_AREA_FIELDS]));
}

function emptyAreaForm(dongList) {
  return { date: today(), dong: dongList?.[0] || "", masonry: "", caulking: "", truss: "", scaffold: "", actual: "", disaster: "N (정상)", reason: "", note: "", memo: "" };
}
function emptyInternalForm(dongList) {
  return { date: today(), dong: dongList?.[0] || "", masonry: "", caulking: "", truss: "", actual: "", disaster: "N (정상)", reason: "", note: "", memo: "" };
}

// 현장별 면적형 범위 데이터(맵)를 로드. kind: "b"(기준정보) | "l"(실적) | "p"(대기)
function loadAreaMap(kind, siteId, isDefault) {
  const m = {};
  for (const s of AREA_SCOPES) {
    const base = AREA_STORAGE[s.key][kind];
    let fallback = [];
    if (s.key === "external" && isDefault) {
      if (kind === "b") fallback = BUILDINGS_EXTERNAL;
      else if (kind === "l") fallback = SEED_LOGS_EXTERNAL;
    }
    m[s.key] = Storage.get(Storage.siteKey(base, siteId), fallback);
  }
  return m;
}

export default function App() {
  const [tab, setTab] = useState("dash");
  const [inputScope, setInputScope] = useState("external");
  const [dashScope, setDashScope] = useState("external");
  const [recoveryScope, setRecoveryScope] = useState("external");

  const [sites, setSites] = useState(() => Storage.getSites());
  const [activeSiteId, setActiveSiteId] = useState(() => Storage.getActiveSiteId());
  const activeSite = useMemo(() => sites.find((s) => s.id === activeSiteId) || sites[0] || null, [sites, activeSiteId]);
  const currentSiteIdRef = useRef(activeSiteId);

  const initialIsDefault = !!Storage.getSites().find((s) => s.id === activeSiteId)?.isDefault;

  const [areaB, setAreaB] = useState(() => loadAreaMap("b", activeSiteId, initialIsDefault));
  const [areaL, setAreaL] = useState(() => loadAreaMap("l", activeSiteId, initialIsDefault));
  const [areaP, setAreaP] = useState(() => loadAreaMap("p", activeSiteId, initialIsDefault));

  const [buildingsInternal, setBuildingsInternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.buildingsInternal, activeSiteId), initialIsDefault ? BUILDINGS_INTERNAL : []));
  const [logsInternal, setLogsInternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.logsInternal, activeSiteId), initialIsDefault ? SEED_LOGS_INTERNAL : []));
  const [pendingInternal, setPendingInternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.pendingInternal, activeSiteId), []));
  const [checklist, setChecklist] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.checklist, activeSiteId), {}));
  const [orderRows, setOrderRows] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.orderRows, activeSiteId), initialIsDefault ? SEED_ORDER_ROWS : []));
  const [thresholds, setThresholds] = useState(() => ({ ...THRESHOLDS, ...Storage.get(Storage.siteKey(Storage.KEYS.thresholds, activeSiteId), {}) }));
  // 엑셀 ③달성률현황 / 내부 달성률현황 시트에서 읽어온 동별 달성률 표 (없으면 null → 기존 재계산으로 폴백)
  const [achvExternal, setAchvExternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.achvExternal, activeSiteId), null));
  const [achvInternal, setAchvInternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.achvInternal, activeSiteId), null));
  // 착수 선행공정 체크리스트 (엑셀 ⑦착수체크리스트 시트와 동기화)
  const [gateList, setGateList] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.checklist2, activeSiteId), []));
  const [checklistDirty, setChecklistDirty] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.checklistDirty, activeSiteId), false));

  const dongListsArea = useMemo(() => Object.fromEntries(AREA_SCOPES.map((s) => [s.key, (areaB[s.key] || []).map((b) => b.dong)])), [areaB]);
  const dongListInternal = useMemo(() => [...buildingsInternal.map((b) => b.dong), EXTRA_INTERNAL_DONG], [buildingsInternal]);

  // 범위별 투입인원 입력 칸 구성 (동기화 시 엑셀 헤더에서 읽어 채움)
  const [areaFields, setAreaFields] = useState(() => defaultAreaFieldsMap());
  const [areaForms, setAreaForms] = useState(() => Object.fromEntries(AREA_SCOPES.map((s) => [s.key, emptyAreaForm([])])));
  const [formInternal, setFormInternal] = useState(() => emptyInternalForm([]));
  const [recoveryDong, setRecoveryDong] = useState("");

  const [account, setAccount] = useState(null);
  const [justLoggedIn, setJustLoggedIn] = useState(false);
  const [sync, setSync] = useState({ state: "idle", message: "", lastSyncedAt: Storage.get(Storage.siteKey(Storage.KEYS.fileMeta, activeSiteId), {})?.lastSyncedAt || null });
  const wbRef = useRef(null);
  const itemIdRef = useRef(null);

  // 현장 전환 시 해당 현장 데이터 재로드
  useEffect(() => {
    if (!activeSiteId) return;
    currentSiteIdRef.current = activeSiteId;
    const site = Storage.getSites().find((s) => s.id === activeSiteId);
    const isDef = !!site?.isDefault;
    setAreaB(loadAreaMap("b", activeSiteId, isDef));
    setAreaL(loadAreaMap("l", activeSiteId, isDef));
    setAreaP(loadAreaMap("p", activeSiteId, isDef));
    setAreaFields(defaultAreaFieldsMap());
    setBuildingsInternal(Storage.get(Storage.siteKey(Storage.KEYS.buildingsInternal, activeSiteId), isDef ? BUILDINGS_INTERNAL : []));
    setLogsInternal(Storage.get(Storage.siteKey(Storage.KEYS.logsInternal, activeSiteId), isDef ? SEED_LOGS_INTERNAL : []));
    setPendingInternal(Storage.get(Storage.siteKey(Storage.KEYS.pendingInternal, activeSiteId), []));
    setChecklist(Storage.get(Storage.siteKey(Storage.KEYS.checklist, activeSiteId), {}));
    setOrderRows(Storage.get(Storage.siteKey(Storage.KEYS.orderRows, activeSiteId), isDef ? SEED_ORDER_ROWS : []));
    setThresholds({ ...THRESHOLDS, ...Storage.get(Storage.siteKey(Storage.KEYS.thresholds, activeSiteId), {}) });
    setAchvExternal(Storage.get(Storage.siteKey(Storage.KEYS.achvExternal, activeSiteId), null));
    setAchvInternal(Storage.get(Storage.siteKey(Storage.KEYS.achvInternal, activeSiteId), null));
    setGateList(Storage.get(Storage.siteKey(Storage.KEYS.checklist2, activeSiteId), []));
    setChecklistDirty(Storage.get(Storage.siteKey(Storage.KEYS.checklistDirty, activeSiteId), false));
    setSync({ state: "idle", message: "", lastSyncedAt: Storage.get(Storage.siteKey(Storage.KEYS.fileMeta, activeSiteId), {})?.lastSyncedAt || null });
    wbRef.current = null;
    itemIdRef.current = null;
  }, [activeSiteId]);

  useEffect(() => { for (const s of AREA_SCOPES) Storage.set(Storage.siteKey(AREA_STORAGE[s.key].b, currentSiteIdRef.current), areaB[s.key]); }, [areaB]);
  useEffect(() => { for (const s of AREA_SCOPES) Storage.set(Storage.siteKey(AREA_STORAGE[s.key].l, currentSiteIdRef.current), areaL[s.key]); }, [areaL]);
  useEffect(() => { for (const s of AREA_SCOPES) Storage.set(Storage.siteKey(AREA_STORAGE[s.key].p, currentSiteIdRef.current), areaP[s.key]); }, [areaP]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.buildingsInternal, currentSiteIdRef.current), buildingsInternal); }, [buildingsInternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.logsInternal, currentSiteIdRef.current), logsInternal); }, [logsInternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.pendingInternal, currentSiteIdRef.current), pendingInternal); }, [pendingInternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.checklist, currentSiteIdRef.current), checklist); }, [checklist]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.orderRows, currentSiteIdRef.current), orderRows); }, [orderRows]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.thresholds, currentSiteIdRef.current), thresholds); }, [thresholds]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.achvExternal, currentSiteIdRef.current), achvExternal); }, [achvExternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.achvInternal, currentSiteIdRef.current), achvInternal); }, [achvInternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.checklist2, currentSiteIdRef.current), gateList); }, [gateList]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.checklistDirty, currentSiteIdRef.current), checklistDirty); }, [checklistDirty]);

  useEffect(() => {
    setAreaForms((prev) => {
      let changed = false;
      const next = { ...prev };
      for (const s of AREA_SCOPES) {
        const dl = dongListsArea[s.key] || [];
        if (dl.length && !dl.includes(prev[s.key].dong)) { next[s.key] = { ...prev[s.key], dong: dl[0] }; changed = true; }
      }
      return changed ? next : prev;
    });
  }, [dongListsArea]);

  useEffect(() => {
    const dl = dongListsArea[recoveryScope] || [];
    if (dl.length && !dl.includes(recoveryDong)) setRecoveryDong(dl[0]);
  }, [recoveryScope, dongListsArea]);

  useEffect(() => {
    if (dongListInternal.length && !dongListInternal.includes(formInternal.dong)) {
      setFormInternal((f) => ({ ...f, dong: dongListInternal[0] }));
    }
  }, [dongListInternal]);

  useEffect(() => {
    if (!Graph.isConfigured()) return;
    Graph.initMsal().then((result) => {
      const acc = result?.account || Graph.getActiveAccount();
      if (acc) {
        setAccount(acc);
        setJustLoggedIn(true); // 로그인돼 있으면 앱 열 때 자동 동기화 → 묵은 화면(이전 저장값) 방지
      }
    }).catch(() => {});
  }, []);

  useEffect(() => {
    if (justLoggedIn && account) {
      setJustLoggedIn(false);
      runSync();
    }
  }, [justLoggedIn, account]);

  // 외부(아파트)는 엑셀 ③달성률현황 시트 값을 기준으로 삼고, 시트 수식 범위 밖의 최신 실적과
  // 아직 업로드 안 된 앱 입력분을 더해서 보여줍니다(하이브리드). 호이스트/부대시설은 전용 달성률
  // 시트가 없으므로 기존 재계산을 그대로 씁니다.
  const areaDash = useMemo(() => Object.fromEntries(AREA_SCOPES.map((s) => {
    const base = calcExternalDashboard(areaB[s.key] || [], areaL[s.key] || [], thresholds);
    if (s.key !== "external") return [s.key, { ...base, source: "recalc" }];
    return [s.key, mergeAchievement(base, achvExternal, areaL[s.key] || [], areaP[s.key] || [], thresholds, "planArea")];
  })), [areaB, areaL, areaP, thresholds, achvExternal]);
  const dashInternal = useMemo(() => {
    const base = calcInternalDashboard(buildingsInternal, logsInternal, thresholds);
    return mergeAchievement(base, achvInternal, logsInternal, pendingInternal, thresholds, "totalUnits");
  }, [buildingsInternal, logsInternal, pendingInternal, thresholds, achvInternal]);
  const recovery = useMemo(() => calcRecoveryPlan(recoveryDong, areaB[recoveryScope] || [], areaL[recoveryScope] || [], new Date(), thresholds), [recoveryDong, recoveryScope, areaB, areaL, thresholds]);
  const orderDash = useMemo(() => calcOrderStatus(orderRows), [orderRows]);

  // 공기 관리 — 착수지연/공기손실/인력소요/클레임. 외부(아파트) 기준.
  const schedule = useMemo(
    () => calcScheduleStatus(areaB.external || [], areaL.external || [], gateList, thresholds, new Date()),
    [areaB, areaL, gateList, thresholds]
  );
  const manpower = useMemo(() => calcManpowerCurve(schedule, areaL.external || [], new Date()), [schedule, areaL]);
  const claim = useMemo(() => calcClaimSummary(schedule, areaL.external || []), [schedule, areaL]);

  const pendingCount = AREA_KEYS.reduce((s, k) => s + (areaP[k]?.length || 0), 0) + pendingInternal.length;

  async function doLogin() {
    try {
      setSync((s) => ({ ...s, state: "syncing", message: "Microsoft 로그인 화면으로 이동 중..." }));
      await Graph.login();
    } catch (e) {
      setSync((s) => ({ ...s, state: "error", message: String(e.message || e) }));
    }
  }

  function doLogout() {
    Graph.logout();
    setAccount(null);
    setSync((s) => ({ ...s, state: "idle", message: "로그아웃됨" }));
  }

  async function runSync() {
    if (!account || !activeSite) return;
    setSync((s) => ({ ...s, state: "syncing", message: "OneDrive에서 받아오는 중..." }));
    try {
      const result = await Graph.syncDown(activeSite);
      const wb = result.wb;
      const itemId = result.itemId;

      let anyPending = false;
      const appendedKeys = new Set();
      for (const s of AREA_SCOPES) {
        if (!wb.getWorksheet(s.sheet)) continue;
        const list = areaP[s.key] || [];
        for (const entry of list) Graph.appendAreaRow(wb, s.sheet, s.planRange, entry);
        appendedKeys.add(s.key);
        if (list.length) anyPending = true;
      }
      for (const entry of pendingInternal) Graph.appendInternalRow(wb, entry);
      if (pendingInternal.length) anyPending = true;

      // 체크리스트를 앱에서 고쳤으면 엑셀 ⑦착수체크리스트 시트에 반영 후 업로드
      if (checklistDirty && gateList.length) {
        Graph.writeChecklist(wb, gateList);
        anyPending = true;
      }

      if (anyPending) {
        setSync((s) => ({ ...s, message: "변경사항 업로드 중..." }));
        await Graph.syncUp(wb, itemId);
      }

      wbRef.current = wb;
      itemIdRef.current = itemId;

      const newB = {};
      const newL = {};
      const newFields = {};
      for (const s of AREA_SCOPES) {
        const b = Graph.readAreaBuildings(wb, s.planLabel);
        // 파일이 진실원천 — 읽은 결과를 그대로 반영(빈 값이면 비움). 묵은 데이터가 남지 않게 함.
        newB[s.key] = b;
        newL[s.key] = Graph.readAreaLogs(wb, s.sheet).map((l) => calcRowExternal(l, b));
        const f = Graph.readAreaFields(wb, s.sheet);
        newFields[s.key] = f.length ? f : DEFAULT_AREA_FIELDS;
      }
      setAreaB(newB);
      setAreaL(newL);
      setAreaFields(newFields);

      const sheetChecklist = Graph.readChecklist(wb);
      if (sheetChecklist.length || !checklistDirty) setGateList(sheetChecklist);
      setChecklistDirty(false);

      setAchvExternal(result.achvExternal ?? Graph.readAchievement(wb, Graph.SHEET_ACHV_EXTERNAL));
      setAchvInternal(result.achvInternal ?? Graph.readAchievement(wb, Graph.SHEET_ACHV_INTERNAL));

      setOrderRows(result.orderRows || Graph.readOrderStatus(wb));
      setThresholds({ ...THRESHOLDS, ...(result.thresholds || Graph.readThresholds(wb)) });

      const intB = Graph.readBuildings(wb).internal;
      const intLogs = Graph.readInternalLogs(wb).map((l) => calcRowInternal(l));
      setLogsInternal(intLogs);
      setBuildingsInternal(intB); // 항상 갱신 — 이전에 잘못 저장된 묵은 내부 데이터 제거

      setAreaP((prev) => Object.fromEntries(AREA_KEYS.map((k) => [k, appendedKeys.has(k) ? [] : (prev[k] || [])])));
      setPendingInternal([]);

      const now = new Date().toISOString();
      setSync({ state: "done", message: "동기화 완료", lastSyncedAt: now });
      const fileMetaKey = Storage.siteKey(Storage.KEYS.fileMeta, activeSite.id);
      const meta = Storage.get(fileMetaKey, {});
      Storage.set(fileMetaKey, { ...meta, lastSyncedAt: now });
    } catch (e) {
      setSync((s) => ({ ...s, state: "error", message: String(e.message || e) }));
    }
  }

  function selectSite(id) {
    Storage.setActiveSiteId(id);
    setActiveSiteId(id);
  }
  function handleAddSite(name, filePath) {
    const site = Storage.addSite({ name, filePath });
    setSites(Storage.getSites());
    selectSite(site.id);
  }
  function handleUpdateSite(id, name, filePath) {
    Storage.updateSite(id, { name, filePath });
    setSites(Storage.getSites());
    if (id === activeSiteId) { wbRef.current = null; itemIdRef.current = null; }
  }
  function handleDeleteSite(id) {
    Storage.removeSite(id);
    setSites(Storage.getSites());
    if (id === activeSiteId) setActiveSiteId(Storage.getActiveSiteId());
  }

  function saveArea(key) {
    const form = areaForms[key];
    // 동은 비워둘 수 있다(천재지변·조업불가일). 날짜만 있으면 저장한다.
    // 빈 동은 계산 시 직전 작업 동을 물려받는다. — calc.expandLogDongs
    if (!form || !form.date) return;
    const entry = {
      id: `local-${Date.now()}`,
      date: form.date,
      dong: form.dong,
      masonry: Number(form.masonry) || 0,
      caulking: Number(form.caulking) || 0,
      truss: Number(form.truss) || 0,
      scaffold: Number(form.scaffold) || 0,
      actual: form.actual === "" ? 0 : Number(form.actual),
      disaster: form.disaster,
      reason: form.reason,
      note: form.note,
      memo: form.memo,
    };
    setAreaL((prev) => ({ ...prev, [key]: [...(prev[key] || []), calcRowExternal(entry, areaB[key] || [])] }));
    setAreaP((prev) => ({ ...prev, [key]: [...(prev[key] || []), entry] }));
    setAreaForms((prev) => ({ ...prev, [key]: emptyAreaForm(dongListsArea[key]) }));
    if (account) setTimeout(runSync, 50);
  }

  function saveInternal() {
    if (!formInternal.dong) return;
    const entry = {
      id: `local-${Date.now()}`,
      date: formInternal.date,
      dong: formInternal.dong,
      masonry: Number(formInternal.masonry) || 0,
      caulking: Number(formInternal.caulking) || 0,
      truss: Number(formInternal.truss) || 0,
      actual: formInternal.actual === "" ? 0 : Number(formInternal.actual),
      disaster: formInternal.disaster,
      reason: formInternal.reason,
      note: formInternal.note,
      memo: formInternal.memo,
    };
    setLogsInternal((prev) => [...prev, calcRowInternal(entry)]);
    setPendingInternal((prev) => [...prev, entry]);
    setFormInternal(emptyInternalForm(dongListInternal));
    if (account) setTimeout(runSync, 50);
  }

  function toggleCheck(idx) {
    setChecklist((prev) => ({ ...prev, [idx]: !prev[idx] }));
  }

  const setAreaForm = (key, f) => setAreaForms((prev) => ({ ...prev, [key]: f }));

  return (
    <div className="app">
      <div className="header">
        <h1>🪨 석공사 공정관리</h1>
        <select className="sitepicker" value={activeSiteId || ""} onChange={(e) => selectSite(e.target.value)}>
          {sites.map((s) => (<option key={s.id} value={s.id}>🏗 {s.name}</option>))}
        </select>
        <div className="sub">
          {account ? (
            <span className="pill dot-ok">🟢 {account.username}</span>
          ) : (
            <span className="pill">🔌 OneDrive 미연결</span>
          )}
          {pendingCount > 0 && <span className="pill dot-warn">대기 {pendingCount}건</span>}
          {sync.state === "syncing" && <span className="pill">동기화 중…</span>}
        </div>
      </div>

      <div className="content">
        {tab === "input" && (
          <InputTab
            scope={inputScope} setScope={setInputScope}
            areaForms={areaForms} setAreaForm={setAreaForm} saveArea={saveArea} areaFields={areaFields}
            formInternal={formInternal} setFormInternal={setFormInternal} saveInternal={saveInternal}
            areaL={areaL} logsInternal={logsInternal}
            dongListsArea={dongListsArea} dongListInternal={dongListInternal}
          />
        )}
        {tab === "dash" && (
          <DashTab scope={dashScope} setScope={setDashScope} areaDash={areaDash} dashInternal={dashInternal} orderDash={orderDash} />
        )}
        {tab === "recovery" && (
          <RecoveryTab
            recoveryScope={recoveryScope} setRecoveryScope={setRecoveryScope}
            recoveryDong={recoveryDong} setRecoveryDong={setRecoveryDong} recovery={recovery}
            checklist={checklist} toggleCheck={toggleCheck} dongList={dongListsArea[recoveryScope] || []}
            thresholds={thresholds}
          />
        )}
        {tab === "schedule" && (
          <ScheduleTab
            schedule={schedule} manpower={manpower} claim={claim}
            checklist={gateList}
            onChangeChecklist={(next) => { setGateList(next); setChecklistDirty(true); }}
          />
        )}
        {tab === "base" && (
          <BaseTab areaB={areaB} buildingsInternal={buildingsInternal} thresholds={thresholds} />
        )}
        {tab === "sync" && (
          <SyncTab
            account={account} sync={sync} pendingCount={pendingCount}
            doLogin={doLogin} doLogout={doLogout} runSync={runSync}
            configured={Graph.isConfigured()}
            sites={sites} activeSiteId={activeSiteId} activeSite={activeSite}
            onSelectSite={selectSite} onAddSite={handleAddSite}
            onUpdateSite={handleUpdateSite} onDeleteSite={handleDeleteSite}
          />
        )}
        <div style={{ textAlign: "center", fontSize: 11, color: "#9ca3af", padding: "18px 8px 4px", lineHeight: 1.5 }}>
          만든이 폭풍간지 이상준 01045166010 무분별 사용시 법적소송을 당할수있음.<br />
          법적소송을 감당할수 있으면 승인받지 말고 쓸것
        </div>
      </div>

      <div className="tabbar">
        <TabBtn active={tab === "dash"} onClick={() => setTab("dash")} icon="📊" label="현황" />
        <TabBtn active={tab === "recovery"} onClick={() => setTab("recovery")} icon="🔄" label="만회계획" />
        <TabBtn active={tab === "schedule"} onClick={() => setTab("schedule")} icon="⏱️" label="공기관리" />
        <TabBtn active={tab === "base"} onClick={() => setTab("base")} icon="📋" label="기준정보" />
        <TabBtn active={tab === "input"} onClick={() => setTab("input")} icon="✏️" label="입력" />
        <TabBtn active={tab === "sync"} onClick={() => setTab("sync")} icon="☁️" label="동기화" />
      </div>
    </div>
  );
}

// ============================================================================
// 공기관리 탭 — 인력 소요 곡선 / 착수 체크리스트 / 공기연장 클레임
// ============================================================================

// dataviz 스킬의 검증된 기본 카테고리 팔레트 (slot 1~8, light).
// 5개 슬롯 조합은 validate_palette.js 전 항목 통과(최악 인접 CVD ΔE 9.1 / 일반시야 19.6).
// 대비 경고가 걸리는 슬롯이 있어 "표로 보기"를 항상 함께 제공한다(relief rule).
const SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"];

const SCHEDULE_VIEWS = [
  { key: "manpower", label: "인력 소요" },
  { key: "gate", label: "착수 체크" },
  { key: "claim", label: "공기연장" },
];

function ScheduleTab({ schedule, manpower, claim, checklist, onChangeChecklist }) {
  const [view, setView] = useState("manpower");
  return (
    <div>
      <div className="toggle2">
        {SCHEDULE_VIEWS.map((v) => (
          <button key={v.key} className={view === v.key ? "active" : ""} onClick={() => setView(v.key)}>{v.label}</button>
        ))}
      </div>
      {view === "manpower" && <ManpowerCard schedule={schedule} manpower={manpower} />}
      {view === "gate" && <GateCard schedule={schedule} checklist={checklist} onChange={onChangeChecklist} />}
      {view === "claim" && <ClaimCard claim={claim} />}
    </div>
  );
}

// ---------- 인력 소요 곡선 ----------
function ManpowerCard({ schedule, manpower }) {
  const { series, peak, capacity, avgRecent, shortDays, firstShortDate, dongs, overdue } = manpower;
  const [showTable, setShowTable] = useState(false);
  const colorOf = (dong) => SERIES_COLORS[dongs.indexOf(dong) % SERIES_COLORS.length];

  if (!series.length) {
    return (
      <div className="card">
        <h2>인력 소요 곡선</h2>
        <p className="meta">진행 중인 동이 없습니다. 모든 동이 완료되었거나 기준정보가 비어 있습니다.</p>
      </div>
    );
  }

  const gap = capacity > 0 ? peak.total - capacity : 0;

  return (
    <>
      <div className="card">
        <h2>인력 소요 요약</h2>
        <div className="statgrid">
          <div className="stat"><div className="label">피크 필요 인원</div><div className="value">{fmtNum(peak.total, 0)}명</div></div>
          <div className="stat"><div className="label">피크 시점</div><div className="value" style={{ fontSize: 15 }}>{peak.date}</div></div>
          <div className="stat"><div className="label">가용 인원(최근 30일 최대)</div><div className="value">{fmtNum(capacity, 0)}명</div></div>
          <div className="stat"><div className="label">최근 30일 평균 투입</div><div className="value">{fmtNum(avgRecent, 1)}명</div></div>
        </div>
        {gap > 0 && (
          <p className="meta" style={{ marginTop: 10, color: "#d32f2f", fontWeight: 700 }}>
            🚨 피크에 {fmtNum(gap, 0)}명 부족 · 인원 부족 예상일 {shortDays}일 (최초 {firstShortDate})
          </p>
        )}
        {overdue.length > 0 && (
          <p className="meta" style={{ marginTop: 6, color: "#d32f2f" }}>
            ⛔ 조정 완료예정일이 이미 지난 동: {overdue.map((o) => `${o.dong}(${o.end})`).join(", ")}
          </p>
        )}
        <p className="meta" style={{ marginTop: 8 }}>
          완료된 동은 제외했습니다. 각 동의 잔여 물량을 조정 완료예정일(원 완료예정일 + 착수지연)까지
          균등 배분하고, 동별 실적 생산성으로 나눠 필요 인원을 낸 값입니다. 일요일은 제외했습니다.
        </p>
      </div>

      <div className="card">
        <h2>날짜별 필요 인원</h2>
        <ManpowerChart series={series} dongs={dongs} capacity={capacity} colorOf={colorOf} />
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginTop: 10 }}>
          {dongs.map((d) => (
            <span key={d} style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11.5, color: "var(--sub)" }}>
              <span style={{ width: 10, height: 10, borderRadius: 3, background: colorOf(d) }} />
              {d}
            </span>
          ))}
          {capacity > 0 && (
            <span style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11.5, color: "var(--sub)" }}>
              <span style={{ width: 14, height: 0, borderTop: "2px dashed #52514e" }} />
              가용 {fmtNum(capacity, 0)}명
            </span>
          )}
        </div>
        <button
          onClick={() => setShowTable((v) => !v)}
          style={{ marginTop: 12, border: "1px solid var(--line)", background: "none", borderRadius: 8, padding: "6px 10px", fontSize: 12, color: "var(--sub)" }}
        >
          {showTable ? "표 닫기" : "표로 보기"}
        </button>
        {showTable && (
          <div style={{ overflowX: "auto", marginTop: 10 }}>
            <table className="dashtable">
              <thead>
                <tr><th>날짜</th>{dongs.map((d) => <th key={d}>{d}</th>)}<th>합계</th></tr>
              </thead>
              <tbody>
                {series.map((s) => (
                  <tr key={s.date}>
                    <td className="dong" style={{ whiteSpace: "nowrap" }}>{s.date.slice(5)}</td>
                    {dongs.map((d) => <td key={d}>{s.byDong[d] ? fmtNum(s.byDong[d], 1) : "-"}</td>)}
                    <td style={{ fontWeight: 700, color: capacity > 0 && s.total > capacity ? "#d32f2f" : "inherit" }}>{fmtNum(s.total, 1)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="card">
        <h2>동별 잔여 · 공기</h2>
        <div style={{ overflowX: "auto" }}>
          <table className="dashtable">
            <thead>
              <tr><th>동</th><th>잔여(m²)</th><th>착수지연</th><th>원 완료예정</th><th>조정 완료예정</th><th>생산성</th></tr>
            </thead>
            <tbody>
              {schedule.map((r) => (
                <tr key={r.dong} style={{ opacity: r.done ? 0.45 : 1 }}>
                  <td className="dong">{r.dong}{r.done && <span style={{ fontSize: 10, color: "var(--sub)" }}> 완료</span>}</td>
                  <td>{r.done ? "0" : fmtNum(r.remain, 0)}</td>
                  <td style={{ color: r.startDelay > 0 ? "#d32f2f" : "inherit", whiteSpace: "nowrap" }}>
                    {r.startDelay > 0 ? `${r.startDelay}일` : "-"}
                    {r.startBasis === "not-started" && <span style={{ fontSize: 10, display: "block" }}>미착수</span>}
                  </td>
                  <td style={{ whiteSpace: "nowrap" }}>{r.plannedEnd || "-"}</td>
                  <td style={{ whiteSpace: "nowrap", fontWeight: r.startDelay > 0 ? 700 : 400 }}>{r.adjustedEnd || "-"}</td>
                  <td>{fmtNum(r.productivity, 1)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="meta" style={{ marginTop: 8 }}>
          생산성은 동 단위 누적 평균입니다. 벽체 판재는 많이 나오고 창대석·창주위석·두겁석이나 첫날은 적게 나오므로,
          하루 단위로 보면 편차가 크지만 동을 끝내면 평균으로 수렴합니다. 그래서 일자별 생산성으로는 판정하지 않습니다.
        </p>
      </div>
    </>
  );
}

// 누적 막대 + 가용 인원 기준선
function ManpowerChart({ series, dongs, capacity, colorOf }) {
  const [hover, setHover] = useState(null);
  const W = 320, H = 190, padL = 26, padR = 6, padT = 10, padB = 34;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const maxVal = Math.max(capacity, ...series.map((s) => s.total)) * 1.12 || 1;
  const bw = Math.max(2, Math.min(18, plotW / series.length - 2));
  const x = (i) => padL + (plotW / series.length) * (i + 0.5) - bw / 2;
  const y = (v) => padT + plotH - (v / maxVal) * plotH;
  const ticks = [0, maxVal / 2, maxVal].map((v) => Math.round(v));

  return (
    <div style={{ position: "relative" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }} role="img" aria-label="날짜별 필요 인원 누적 막대 차트">
        {ticks.map((t) => (
          <g key={t}>
            <line x1={padL} x2={W - padR} y1={y(t)} y2={y(t)} stroke="#e5e7eb" strokeWidth="1" />
            <text x={padL - 4} y={y(t) + 3} textAnchor="end" fontSize="8" fill="#8b8b86">{t}</text>
          </g>
        ))}
        {series.map((s, i) => {
          let acc = 0;
          return (
            <g key={s.date} onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(null)}>
              <rect x={x(i) - 1} y={padT} width={bw + 2} height={plotH} fill="transparent" />
              {dongs.map((d) => {
                const v = s.byDong[d];
                if (!v) return null;
                const h = (v / maxVal) * plotH;
                const yy = padT + plotH - acc - h;
                acc += h;
                return <rect key={d} x={x(i)} y={yy} width={bw} height={Math.max(0, h - 2)} fill={colorOf(d)} rx="1.5" />;
              })}
            </g>
          );
        })}
        {capacity > 0 && (
          <>
            <line x1={padL} x2={W - padR} y1={y(capacity)} y2={y(capacity)} stroke="#52514e" strokeWidth="2" strokeDasharray="5 3" />
            <text x={W - padR} y={y(capacity) - 4} textAnchor="end" fontSize="8.5" fill="#52514e" fontWeight="700">가용 {Math.round(capacity)}명</text>
          </>
        )}
        {series.map((s, i) =>
          i % Math.ceil(series.length / 6) === 0 ? (
            <text key={s.date} x={x(i) + bw / 2} y={H - padB + 12} textAnchor="middle" fontSize="8" fill="#8b8b86">{s.date.slice(5)}</text>
          ) : null
        )}
        <line x1={padL} x2={W - padR} y1={padT + plotH} y2={padT + plotH} stroke="#d1d5db" strokeWidth="1" />
      </svg>
      {hover !== null && series[hover] && (
        <div style={{ position: "absolute", top: 4, left: "50%", transform: "translateX(-50%)", background: "#111", color: "#fff", borderRadius: 8, padding: "6px 9px", fontSize: 11, lineHeight: 1.5, pointerEvents: "none", whiteSpace: "nowrap", zIndex: 5 }}>
          <b>{series[hover].date}</b> · 합계 {fmtNum(series[hover].total, 1)}명
          {dongs.filter((d) => series[hover].byDong[d]).map((d) => (
            <div key={d}>{d} {fmtNum(series[hover].byDong[d], 1)}명</div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- 착수 선행공정 체크리스트 ----------
function GateCard({ schedule, checklist, onChange }) {
  const map = new Map((checklist || []).map((c) => [c.dong, c]));
  const rowOf = (dong) => map.get(dong) || { dong, gates: {}, done: false, reason: "" };

  function update(dong, patch) {
    const cur = rowOf(dong);
    const next = { ...cur, ...patch, gates: { ...cur.gates, ...(patch.gates || {}) } };
    const list = (checklist || []).filter((c) => c.dong !== dong);
    onChange([...list, next].sort((a, b) => String(a.dong).localeCompare(String(b.dong), "ko")));
  }

  return (
    <>
      <div className="card">
        <h2>착수 선행공정 체크</h2>
        <p className="meta">
          동당 한 번만 체크하면 됩니다. 매일 입력하는 항목이 아닙니다.
          지연사유에 적은 내용은 그대로 공기연장 근거로 넘어갑니다.
        </p>
      </div>
      {[...schedule].sort((a, b) => (a.done === b.done ? 0 : a.done ? 1 : -1)).map((r) => {
        const row = rowOf(r.dong);
        const missing = GATE_KEYS.filter((k) => !row.gates?.[k]);
        return (
          <div className="card" key={r.dong} style={{ opacity: r.done ? 0.5 : 1 }}>
            <h2 style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span>
                {r.dong}
                {r.done && <span style={{ fontSize: 11, color: "var(--sub)", fontWeight: 400 }}> · 완료</span>}
                {!r.done && r.startDelay > 0 && (
                  <span style={{ fontSize: 11, color: "#d32f2f", fontWeight: 700 }}> · 착수 {r.startDelay}일 지연</span>
                )}
              </span>
              <label style={{ fontSize: 11.5, fontWeight: 400, color: "var(--sub)", display: "inline-flex", alignItems: "center", gap: 4 }}>
                <input type="checkbox" checked={!!row.done} onChange={(e) => update(r.dong, { done: e.target.checked })} />
                완료 처리
              </label>
            </h2>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>
              {GATE_KEYS.map((k) => (
                <label key={k} style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 12, border: "1px solid var(--line)", borderRadius: 8, padding: "5px 8px", background: row.gates?.[k] ? "#eef7ee" : "none" }}>
                  <input type="checkbox" checked={!!row.gates?.[k]} onChange={(e) => update(r.dong, { gates: { [k]: e.target.checked } })} />
                  {GATE_LABELS[k]}
                </label>
              ))}
            </div>
            {!r.done && !r.actualStart && missing.length > 0 && (
              <p className="meta" style={{ color: "#e8a700", marginBottom: 6 }}>
                미충족: {missing.map((k) => GATE_LABELS[k]).join(", ")}
              </p>
            )}
            <Field label="지연사유 (어떤 선행공정 때문인지)">
              <input
                type="text"
                value={row.reason || ""}
                placeholder="예: 비계 해체 지연 / 창호 취부 미완료"
                onChange={(e) => update(r.dong, { reason: e.target.value })}
              />
            </Field>
            <p className="meta" style={{ marginTop: 6 }}>
              계획착수 {r.plannedStart || "-"} → 실제착수 {r.actualStart || "미착수"} ·
              완료예정 {r.plannedEnd || "-"} → <b>{r.adjustedEnd || "-"}</b>
            </p>
          </div>
        );
      })}
    </>
  );
}

// ---------- 공기연장 클레임 집계 ----------
function ClaimCard({ claim }) {
  const [showAll, setShowAll] = useState(false);
  const ev = showAll ? claim.evidence : claim.evidence.slice(0, 20);
  return (
    <>
      <div className="card">
        <h2>공기연장 청구 가능일수</h2>
        <div className="statgrid">
          <div className="stat"><div className="label">합계</div><div className="value">{fmtNum(claim.totalClaimDays, 1)}일</div></div>
          <div className="stat"><div className="label">착수지연</div><div className="value">{fmtNum(claim.totalStartDelay, 0)}일</div></div>
          <div className="stat"><div className="label">천재지변·부분영향 손실</div><div className="value">{fmtNum(claim.totalLoss, 1)}일</div></div>
          <div className="stat"><div className="label">정상 조업일</div><div className="value">{claim.counts.N} / {claim.totalRows}일</div></div>
        </div>
        <p className="meta" style={{ marginTop: 10 }}>
          천재지변(Y) {claim.counts.Y}건은 1일로, 부분영향(P) {claim.counts.P}건은 그날 실적이
          정상 생산성 대비 몇 %였는지로 손실일을 자동 산정했습니다.
          (예: 정상의 40%만 했으면 0.6일 손실)
        </p>
        <p className="meta" style={{ marginTop: 6 }}>
          완료된 동의 과거 손실도 포함합니다. 이미 발생한 청구 근거이기 때문입니다.
        </p>
      </div>

      <div className="card">
        <h2>동별</h2>
        <div style={{ overflowX: "auto" }}>
          <table className="dashtable">
            <thead><tr><th>동</th><th>착수지연</th><th>손실일</th><th>합계</th><th>지연사유</th></tr></thead>
            <tbody>
              {claim.byDong.map((d) => (
                <tr key={d.dong} style={{ opacity: d.done ? 0.55 : 1 }}>
                  <td className="dong">{d.dong}{d.done && <span style={{ fontSize: 10, color: "var(--sub)" }}> 완료</span>}</td>
                  <td>{d.startDelay ? `${d.startDelay}일` : "-"}</td>
                  <td>{fmtNum(d.lossAfter, 1)}일</td>
                  <td style={{ fontWeight: 700 }}>{fmtNum(d.claimDays, 1)}일</td>
                  <td style={{ fontSize: 11, textAlign: "left" }}>{d.delayReason || "-"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card">
        <h2>월별</h2>
        <table className="dashtable">
          <thead><tr><th>월</th><th>손실일</th><th>천재지변</th><th>부분영향</th><th>해당일수</th></tr></thead>
          <tbody>
            {claim.months.map((m) => (
              <tr key={m.month}>
                <td className="dong">{m.month}</td>
                <td style={{ fontWeight: 700 }}>{fmtNum(m.loss, 1)}일</td>
                <td>{m.y}건</td>
                <td>{m.p}건</td>
                <td>{m.days}일</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="card">
        <h2>근거 리스트 ({claim.evidence.length}건)</h2>
        <div style={{ overflowX: "auto" }}>
          <table className="dashtable">
            <thead><tr><th>날짜</th><th>동</th><th>구분</th><th>손실</th><th>인원</th><th>실적</th><th>사유</th></tr></thead>
            <tbody>
              {ev.map((e, i) => (
                <tr key={`${e.date}-${e.dong}-${i}`}>
                  <td style={{ whiteSpace: "nowrap" }}>{e.date.slice(5)}</td>
                  <td className="dong">{e.dong}{e.inherited && <span title="직전 작업동에서 자동 지정" style={{ fontSize: 9, color: "var(--sub)" }}> 자동</span>}</td>
                  <td>{e.kind}</td>
                  <td style={{ fontWeight: 700 }}>{e.loss.toFixed(2)}</td>
                  <td>{e.workers}</td>
                  <td>{fmtNum(e.actual, 1)}</td>
                  <td style={{ fontSize: 10.5, textAlign: "left" }}>{[e.reason, e.note].filter(Boolean).join(" / ") || "-"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {claim.evidence.length > 20 && (
          <button onClick={() => setShowAll((v) => !v)} style={{ marginTop: 10, border: "1px solid var(--line)", background: "none", borderRadius: 8, padding: "6px 10px", fontSize: 12, color: "var(--sub)" }}>
            {showAll ? "접기" : `전체 ${claim.evidence.length}건 보기`}
          </button>
        )}
      </div>
    </>
  );
}

function TabBtn({ active, onClick, icon, label }) {
  return (
    <button className={active ? "active" : ""} onClick={onClick}>
      <span className="ico">{icon}</span>
      <span>{label}</span>
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="field">
      <span className="lbl">{label}</span>
      {children}
    </label>
  );
}

function ScopeToggle({ scope, setScope }) {
  return (
    <div className="toggle2" style={{ flexWrap: "wrap" }}>
      {SCOPE_TABS.map((s) => (
        <button key={s.key} className={scope === s.key ? "active" : ""} onClick={() => setScope(s.key)}>{s.label}</button>
      ))}
    </div>
  );
}

function RecentList({ logs, unit }) {
  return (
    <div className="card">
      <h2>최근 입력 내역</h2>
      <div className="entrylist">
        {logs.slice(-8).reverse().map((l) => (
          <div key={l.id} className="entryitem">
            <div>
              <div>{l.dong} · 실적 {fmtNum(l.actual)}{unit}</div>
              <div className="meta">{l.date} · 달성률 {fmtPct(l.rate)}</div>
            </div>
            {l.disaster && l.disaster !== "N (정상)" && <span className="pill dot-warn">{l.disaster}</span>}
          </div>
        ))}
        {logs.length === 0 && <div className="meta">입력된 실적이 없습니다.</div>}
      </div>
    </div>
  );
}

// 투입인원 필드를 2개씩 한 줄로 묶음
function pairRows(fields) {
  const rows = [];
  fields.forEach((f, i) => {
    if (i % 2 === 0) rows.push([f]);
    else rows[rows.length - 1].push(f);
  });
  return rows;
}

function AreaInputCard({ scope, form, setForm, onSave, dongList, fields }) {
  const workerFields = fields && fields.length ? fields : DEFAULT_AREA_FIELDS;
  return (
    <div className="card">
      <h2>일일 실적 입력 · {scope.label}</h2>
      <Field label="날짜">
        <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
      </Field>
      <Field label="해당 동(구역)">
        {/* 비 오는 날처럼 조업을 못 한 날은 동을 고르지 않아도 됩니다.
            빈 값으로 두면 직전에 작업하던 동이 자동으로 적용됩니다. */}
        <select value={form.dong} onChange={(e) => setForm({ ...form, dong: e.target.value })}>
          <option value="">(자동) 직전 작업 동 — 천재지변·조업불가일용</option>
          {dongList.length === 0 && <option value="">(기준정보 없음 — 동기화 필요)</option>}
          {dongList.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
      </Field>
      {!form.dong && (
        <p className="meta" style={{ marginTop: -4, marginBottom: 10 }}>
          동을 비워두면 직전에 작업하던 동이 자동으로 적용됩니다. 공기연장 집계에도 그 동으로 잡힙니다.
        </p>
      )}
      {pairRows(workerFields).map((pair, ri) => (
        <div className="row" key={ri}>
          {pair.map((f) => (
            <Field key={f.key} label={f.label + "(명)"}>
              <input type="number" min="0" value={form[f.key] ?? ""} onChange={(e) => setForm({ ...form, [f.key]: e.target.value })} />
            </Field>
          ))}
        </div>
      ))}
      <Field label="실제시공량(m²)">
        <input type="number" step="0.01" min="0" value={form.actual} onChange={(e) => setForm({ ...form, actual: e.target.value })} />
      </Field>
      <div className="row">
        <Field label="천재지변 여부">
          <select value={form.disaster} onChange={(e) => setForm({ ...form, disaster: e.target.value })}>
            {DISASTER_OPTIONS.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="사유코드">
          <select value={form.reason} onChange={(e) => setForm({ ...form, reason: e.target.value })}>
            <option value="">-</option>
            {REASON_CODES.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
        </Field>
      </div>
      <Field label="특기사항·사유">
        <textarea value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
      </Field>
      <Field label="비고">
        <textarea value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })} />
      </Field>
      <button className="btn btn-primary" onClick={onSave}>저장</button>
    </div>
  );
}

function InternalInputCard({ form, setForm, onSave, dongList }) {
  return (
    <div className="card">
      <h2>일일 실적 입력 · 내부(세대)</h2>
      <Field label="날짜">
        <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
      </Field>
      <Field label="해당 동(구역)">
        <select value={form.dong} onChange={(e) => setForm({ ...form, dong: e.target.value })}>
          {dongList.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
      </Field>
      <div className="row">
        <Field label="석공(명)"><input type="number" min="0" value={form.masonry} onChange={(e) => setForm({ ...form, masonry: e.target.value })} /></Field>
        <Field label="코킹(명)"><input type="number" min="0" value={form.caulking} onChange={(e) => setForm({ ...form, caulking: e.target.value })} /></Field>
      </div>
      <Field label="트러스(명)"><input type="number" min="0" value={form.truss} onChange={(e) => setForm({ ...form, truss: e.target.value })} /></Field>
      <Field label="실제시공량(세대)">
        <input type="number" step="0.5" min="0" value={form.actual} onChange={(e) => setForm({ ...form, actual: e.target.value })} />
      </Field>
      <div className="row">
        <Field label="천재지변 여부">
          <select value={form.disaster} onChange={(e) => setForm({ ...form, disaster: e.target.value })}>
            {DISASTER_OPTIONS.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="사유코드">
          <select value={form.reason} onChange={(e) => setForm({ ...form, reason: e.target.value })}>
            <option value="">-</option>
            {REASON_CODES.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
        </Field>
      </div>
      <Field label="특기사항·사유">
        <textarea value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
      </Field>
      <Field label="비고">
        <textarea value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })} />
      </Field>
      <button className="btn btn-primary" onClick={onSave}>저장</button>
    </div>
  );
}

function InputTab({ scope, setScope, areaForms, setAreaForm, saveArea, areaFields, formInternal, setFormInternal, saveInternal, areaL, logsInternal, dongListsArea, dongListInternal }) {
  const areaScope = AREA_SCOPES.find((s) => s.key === scope);
  return (
    <div>
      <ScopeToggle scope={scope} setScope={setScope} />
      {areaScope ? (
        <AreaInputCard
          scope={areaScope}
          form={areaForms[scope]}
          setForm={(f) => setAreaForm(scope, f)}
          onSave={() => saveArea(scope)}
          dongList={dongListsArea[scope] || []}
          fields={areaFields[scope] || DEFAULT_AREA_FIELDS}
        />
      ) : (
        <InternalInputCard form={formInternal} setForm={setFormInternal} onSave={saveInternal} dongList={dongListInternal} />
      )}
      <RecentList logs={areaScope ? (areaL[scope] || []) : logsInternal} unit={areaScope ? "m²" : "세대"} />
    </div>
  );
}

// 이 화면 숫자가 어디서 온 것인지 한 줄로 밝혀 줍니다.
// "엑셀 ③달성률현황 기준" 인지, 앱이 직접 재계산한 값인지 헷갈리지 않도록.
function SourceNote({ dash }) {
  if (dash.source !== "sheet") {
    return <p className="meta">앱이 ②일일실적입력에서 직접 재계산한 값입니다. (달성률현황 시트를 읽지 못함)</p>;
  }
  const added = Number(dash.basis?.addedCum) || 0;
  return (
    <p className="meta">
      엑셀 「{dash.sheetName}」 시트 기준
      {dash.coveredMaxRow ? ` (시트 수식 범위 ~${dash.coveredMaxRow}행)` : ""}
      {added > 0 && ` + 그 이후 실적 ${fmtNum(added)} m² 보충 합산`}
    </p>
  );
}

// 동별 실적 숫자 밑에 "532 엑셀 + 7.7 추가" 형태로 근거를 작게 표시
function Basis({ basis }) {
  if (!basis) return null;
  const extra = Number(basis.extra) || 0;
  const pending = Number(basis.pending) || 0;
  if (!extra && !pending) return null;
  return (
    <div className="meta" style={{ fontSize: "0.75em", lineHeight: 1.3 }}>
      {fmtNum(basis.sheet)} 엑셀
      {extra > 0 && ` + ${fmtNum(extra)} 추가행`}
      {pending > 0 && ` + ${fmtNum(pending)} 미동기화`}
    </div>
  );
}

function AreaDashCard({ scope, dash }) {
  return (
    <>
      <div className="card">
        <h2>{scope.label} · 전체 현황 요약</h2>
        <div className="statgrid">
          <div className="stat"><div className="label">전체 계획 수량</div><div className="value">{fmtNum(dash.totalPlanArea)} m²</div></div>
          <div className="stat"><div className="label">누적 실제시공량</div><div className="value">{fmtNum(dash.cumActual)} m²</div></div>
          <div className="stat"><div className="label">전체 달성률</div><div className="value">{fmtPct(dash.overallRate)}</div></div>
          <div className="stat"><div className="label">기간 달성률</div><div className="value">{fmtPct(dash.periodRate)}</div></div>
          <div className="stat"><div className="label">총 투입인원(연인원)</div><div className="value">{fmtNum(dash.totalWorkers, 0)}명</div></div>
          <div className="stat"><div className="label">1인당 평균 시공량</div><div className="value">{fmtNum(dash.perWorker)} m²</div></div>
          <div className="stat"><div className="label">천재지변 발생일수</div><div className="value">{dash.disasterDays}일</div></div>
          <div className="stat"><div className="label">작업 총 일수</div><div className="value">{dash.totalDays}일</div></div>
        </div>
      </div>
      <div className="card">
        <h2>동별 달성률 현황</h2>
        <SourceNote dash={dash} />
        <table className="dashtable">
          <thead><tr><th>동</th><th>계획(m²)</th><th>실적(m²)</th><th>달성률</th><th>상태</th></tr></thead>
          <tbody>
            {dash.byBuilding.map((b) => (
              <tr key={b.dong}>
                <td className="dong">{b.dong}</td>
                <td>{fmtNum(b.planArea)}</td>
                <td>{fmtNum(b.cumActual)}<Basis basis={b.basis} /></td>
                <td>{fmtPct(b.rate)}</td>
                <td style={{ color: b.status.color }}>{b.status.label}</td>
              </tr>
            ))}
            {dash.byBuilding.length === 0 && <tr><td colSpan="5" className="meta">기준정보가 없습니다 — 동기화하세요.</td></tr>}
          </tbody>
        </table>
      </div>
    </>
  );
}

const DASH_TABS = [...SCOPE_TABS, { key: "order", label: "📦 발주현황" }];
function DashScopeToggle({ scope, setScope }) {
  return (
    <div className="toggle2" style={{ flexWrap: "wrap" }}>
      {DASH_TABS.map((s) => (
        <button key={s.key} className={scope === s.key ? "active" : ""} onClick={() => setScope(s.key)}>{s.label}</button>
      ))}
    </div>
  );
}

function OrderDashCard({ order }) {
  const v = (n) => (n ? fmtNum(n) : "-");
  if (!order || !order.byDong || !order.byDong.length) {
    return (
      <div className="card">
        <h2>📦 발주현황</h2>
        <p className="meta">발주 데이터가 없습니다 — 동기화하면 엑셀 ⑥ 석재발주 시트에서 불러옵니다.</p>
      </div>
    );
  }
  return (
    <>
      <div className="card">
        <h2>📦 발주현황 · 전체 요약</h2>
        <div className="statgrid">
          <div className="stat"><div className="label">발주 동/현장</div><div className="value">{order.dongCount}곳</div></div>
          <div className="stat"><div className="label">면적 (벽체 등)</div><div className="value">{v(order.grand.area)} ㎡</div></div>
          <div className="stat"><div className="label">창대·창주위</div><div className="value">{v(order.grand.window)} m</div></div>
          <div className="stat"><div className="label">두겁</div><div className="value">{v(order.grand.cope)} m</div></div>
        </div>
      </div>
      {order.byDong.map((d) => (
        <div className="card" key={d.dong}>
          <h2>{d.dong} <span style={{ fontSize: 13, fontWeight: 400, color: "#6b7280" }}>· 발주 소계</span></h2>
          <table className="dashtable">
            <thead><tr><th>석종</th><th>면적(㎡)</th><th>창대·창주위(m)</th><th>두겁(m)</th></tr></thead>
            <tbody>
              {d.stones.map((st) => (
                <tr key={st.stone}>
                  <td className="dong">{st.stone}</td>
                  <td>{v(st.area)}</td>
                  <td>{v(st.window)}</td>
                  <td>{v(st.cope)}</td>
                </tr>
              ))}
              <tr style={{ fontWeight: 700, background: "#eef2fb" }}>
                <td className="dong">동 합계</td>
                <td>{v(d.total.area)}</td>
                <td>{v(d.total.window)}</td>
                <td>{v(d.total.cope)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      ))}
    </>
  );
}

function DashTab({ scope, setScope, areaDash, dashInternal, orderDash }) {
  const areaScope = AREA_SCOPES.find((s) => s.key === scope);
  return (
    <div>
      <DashScopeToggle scope={scope} setScope={setScope} />
      {scope === "order" ? (
        <OrderDashCard order={orderDash} />
      ) : areaScope ? (
        <AreaDashCard scope={areaScope} dash={areaDash[scope]} />
      ) : (
        <>
          <div className="card">
            <h2>전체 내부(세대) 현황 요약</h2>
            <div className="statgrid">
              <div className="stat"><div className="label">전체 세대수</div><div className="value">{dashInternal.totalUnits}세대</div></div>
              <div className="stat"><div className="label">옵션세대수</div><div className="value">{dashInternal.optionUnits}세대</div></div>
              <div className="stat"><div className="label">누적 실제시공량</div><div className="value">{fmtNum(dashInternal.cumActual)}세대</div></div>
              <div className="stat"><div className="label">전체 달성률</div><div className="value">{fmtPct(dashInternal.overallRate)}</div></div>
              <div className="stat"><div className="label">기간 달성률</div><div className="value">{fmtPct(dashInternal.periodRate)}</div></div>
              <div className="stat"><div className="label">총 투입인원(연인원)</div><div className="value">{fmtNum(dashInternal.totalWorkers, 0)}명</div></div>
            </div>
          </div>
          <div className="card">
            <h2>동별 내부(세대) 달성률 현황</h2>
            <SourceNote dash={dashInternal} />
            <table className="dashtable">
              <thead><tr><th>동</th><th>전체세대</th><th>실적</th><th>달성률</th><th>상태</th></tr></thead>
              <tbody>
                {dashInternal.byBuilding.map((b) => (
                  <tr key={b.dong}>
                    <td className="dong">{b.dong}</td>
                    <td>{b.totalUnits}</td>
                    <td>{fmtNum(b.cumActual)}<Basis basis={b.basis} /></td>
                    <td>{fmtPct(b.rate)}</td>
                    <td style={{ color: b.status.color }}>{b.status.label}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}

function RecoveryTab({ recoveryScope, setRecoveryScope, recoveryDong, setRecoveryDong, recovery, checklist, toggleCheck, dongList, thresholds }) {
  return (
    <div>
      <div className="card">
        <h2>만회계획 자동 산출</h2>
        {recovery && (
          <div className="banner info" style={{ marginBottom: 10 }}>
            1인당 일일생산성: {fmtNum(recovery.productivity)}㎡/명 — {recovery.productivitySource === "actual"
              ? `실행 생산성 (누적실적 ${fmtNum(recovery.cumActual)}㎡ ÷ 실작업 투입 ${fmtNum(recovery.cumWorkers, 0)}명, 동별)`
              : `실적 없음 → 엑셀 ①기준정보 기준값 ${thresholds.productivityPerWorker}㎡`}
          </div>
        )}
        <Field label="공사 범위">
          <select value={recoveryScope} onChange={(e) => setRecoveryScope(e.target.value)}>
            {AREA_SCOPES.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
          </select>
        </Field>
        <Field label="대상 동 선택">
          <select value={recoveryDong} onChange={(e) => setRecoveryDong(e.target.value)}>
            {dongList.length === 0 && <option value="">(기준정보 없음)</option>}
            {dongList.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
        </Field>
        {recovery && (
          <div className="statgrid">
            <div className="stat"><div className="label">계획 수량</div><div className="value">{fmtNum(recovery.planArea)} m²</div></div>
            <div className="stat"><div className="label">누적 실적</div><div className="value">{fmtNum(recovery.cumActual)} m²</div></div>
            <div className="stat"><div className="label">현재 달성률</div><div className="value">{fmtPct(recovery.currentRate)}</div></div>
            <div className="stat"><div className="label">잔여 계획량</div><div className="value">{fmtNum(recovery.remainArea)} m²</div></div>
            <div className="stat"><div className="label">계획 완료일</div><div className="value">{recovery.endDate}</div></div>
            <div className="stat"><div className="label">잔여 일수</div><div className="value">{recovery.remainDays}일</div></div>
            <div className="stat"><div className="label">실 가용 작업일</div><div className="value">{recovery.availableDays}일</div></div>
            <div className="stat"><div className="label">기준 투입인원</div><div className="value">{recovery.baseWorkers}명</div></div>
            <div className="stat"><div className="label">현 인원 일일생산가능량</div><div className="value">{fmtNum(recovery.currentCapacity)} m²</div></div>
            <div className="stat"><div className="label">1인당 일일생산성</div><div className="value">{fmtNum(recovery.productivity)} m²</div></div>
            <div className="stat"><div className="label">만회 필요 일일생산량</div><div className="value">{typeof recovery.neededDaily === "number" ? fmtNum(recovery.neededDaily) + " m²" : recovery.neededDaily}</div></div>
            <div className="stat"><div className="label">만회 필요 추가인원</div><div className="value">{recovery.extraWorkers}명</div></div>
          </div>
        )}
        {recovery && <div className="banner info" style={{ marginTop: 10 }}>{recovery.verdict}</div>}
        {!recovery && <div className="meta">선택한 범위·동의 기준정보가 없습니다. 동기화 후 다시 시도하세요.</div>}
      </div>

      <div className="card">
        <h2>천재지변 발생 시 대응 매뉴얼</h2>
        {DISASTER_MANUAL.map((m) => (
          <div key={m.step} className="manual-step">
            <div className="step">{m.step}</div>
            <div className="title">{m.title}</div>
            <div className="detail">{m.detail.replaceAll(" · ", "\n· ")}</div>
          </div>
        ))}
      </div>

      <div className="card">
        <h2>공정 지연 만회계획 체크리스트</h2>
        {RECOVERY_CHECKLIST.map((item, idx) => (
          <label key={idx} className="checklist-item">
            <input type="checkbox" checked={!!checklist[idx]} onChange={() => toggleCheck(idx)} />
            <span style={{ textDecoration: checklist[idx] ? "line-through" : "none", color: checklist[idx] ? "#9aa0a8" : "inherit" }}>{item}</span>
          </label>
        ))}
      </div>
    </div>
  );
}

function BaseTab({ areaB, buildingsInternal, thresholds }) {
  return (
    <div>
      {AREA_SCOPES.map((s) => (
        <div className="card" key={s.key}>
          <h2>{s.label} · 시공 계획 수량</h2>
          <table className="dashtable">
            <thead><tr><th>동</th><th>전체수량(m²)</th><th>시작일</th><th>완료예정</th><th>기준인원</th></tr></thead>
            <tbody>
              {(areaB[s.key] || []).map((b) => (
                <tr key={b.dong}>
                  <td className="dong">{b.dong}</td>
                  <td>{fmtNum(b.totalArea)}</td>
                  <td>{b.startDate}</td>
                  <td>{b.endDate}</td>
                  <td>{b.baseWorkers}명</td>
                </tr>
              ))}
              {(areaB[s.key] || []).length === 0 && <tr><td colSpan="5" className="meta">기준정보 없음 — 동기화하세요.</td></tr>}
            </tbody>
          </table>
        </div>
      ))}
      <div className="card">
        <h2>내부(세대) 기준정보</h2>
        <table className="dashtable">
          <thead><tr><th>동</th><th>전체세대</th><th>옵션세대</th><th>일반세대</th></tr></thead>
          <tbody>
            {buildingsInternal.map((b) => (
              <tr key={b.dong}>
                <td className="dong">{b.dong}</td>
                <td>{b.totalUnits}</td>
                <td>{b.optionUnits}</td>
                <td>{b.normalUnits}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="card">
        <h3>달성률 경보 기준</h3>
        <div className="statgrid">
          <div className="stat"><div className="label">정상 기준</div><div className="value">{fmtPct(thresholds.normal)}</div></div>
          <div className="stat"><div className="label">주의 기준</div><div className="value">{fmtPct(thresholds.caution)}</div></div>
          <div className="stat"><div className="label">1인당 일일생산성(기준)</div><div className="value">{thresholds.productivityPerWorker} m²</div></div>
          <div className="stat"><div className="label">만회 허용일수</div><div className="value">{thresholds.recoveryDays}일</div></div>
        </div>
      </div>
    </div>
  );
}

function SyncTab({
  account, sync, pendingCount, doLogin, doLogout, runSync, configured,
  sites, activeSiteId, activeSite, onSelectSite, onAddSite, onUpdateSite, onDeleteSite,
}) {
  return (
    <div>
      {!configured && (
        <div className="banner warn">
          Microsoft 앱 등록(Client ID)이 아직 설정되지 않았습니다. Azure 앱 등록 가이드를 먼저 진행해주세요. 등록 후 .env의 VITE_MSAL_CLIENT_ID 값을 채우면 OneDrive 연결이 가능합니다.
        </div>
      )}

      <SiteManager
        sites={sites} activeSiteId={activeSiteId}
        onSelect={onSelectSite} onAdd={onAddSite} onUpdate={onUpdateSite} onDelete={onDeleteSite}
      />

      <div className="card">
        <h2>OneDrive 연동</h2>
        <p style={{ fontSize: 13, color: "#6b7280", lineHeight: 1.5 }}>
          현재 선택된 현장(<b>{activeSite?.name}</b>)은 OneDrive의 "{activeSite?.filePath}" 파일과 동기화합니다(데스크탑 파일이 아니라 OneDrive에 저장된 파일 기준입니다). 모바일에서 입력 후 저장하면 자동으로 이 파일에 새 행이 추가됩니다. Microsoft 계정 연결은 모든 현장에 공통으로 적용되며, 현장마다 OneDrive 파일 경로만 다르게 지정하면 됩니다.
        </p>
        {account ? (
          <>
            <div className="banner info">{account.username} 로 연결됨</div>
            <button className="btn btn-ghost" style={{ width: "100%", marginBottom: 8 }} onClick={runSync}>지금 동기화</button>
            <button className="btn btn-ghost" style={{ width: "100%" }} onClick={doLogout}>연결 해제</button>
          </>
        ) : (
          <button className="btn btn-primary" disabled={!configured} onClick={doLogin}>Microsoft 계정으로 OneDrive 연결</button>
        )}
        <div style={{ marginTop: 10, fontSize: 12, color: "#6b7280" }}>
          상태: {sync.message || "대기 중"}<br />
          대기 중인 미동기화 입력: {pendingCount}건<br />
          마지막 동기화: {sync.lastSyncedAt ? new Date(sync.lastSyncedAt).toLocaleString("ko-KR") : "없음"}
        </div>
      </div>
      <div className="card">
        <h3>알아두실 점</h3>
        <p style={{ fontSize: 12.5, color: "#6b7280", lineHeight: 1.6 }}>
          개인 OneDrive 계정은 엑셀 셀 단위 API를 지원하지 않아, 동기화 시 파일 전체를 받아 수정 후 다시 업로드합니다(서식·색상·메모는 보존됩니다).
          동기화 중에는 PC에서 같은 파일을 열어두지 않는 것을 권장합니다(동시 저장 시 충돌 가능).
          공사 범위(외부·호이스트·부대시설 등)는 엑셀 ①기준정보의 "▶ ○○ 수량" 블록과 "②일일실적입력(○○)" 시트로 관리되며, 같은 형식이면 앱이 자동으로 읽어 화면에 표시합니다. 투입인원 칸 이름도 각 시트의 헤더(4행)를 그대로 따라갑니다.
        </p>
      </div>
    </div>
  );
}

function SiteManager({ sites, activeSiteId, onSelect, onAdd, onUpdate, onDelete }) {
  const [editingId, setEditingId] = useState(null);
  const [name, setName] = useState("");
  const [filePath, setFilePath] = useState("");

  function startAdd() { setEditingId("new"); setName(""); setFilePath(""); }
  function startEdit(site) { setEditingId(site.id); setName(site.name); setFilePath(site.filePath); }
  function cancelForm() { setEditingId(null); setName(""); setFilePath(""); }
  function submitForm() {
    const trimmedName = name.trim();
    const trimmedPath = filePath.trim();
    if (!trimmedName || !trimmedPath) return;
    if (editingId === "new") onAdd(trimmedName, trimmedPath);
    else if (editingId) onUpdate(editingId, trimmedName, trimmedPath);
    cancelForm();
  }

  return (
    <div className="card">
      <h2>현장 관리</h2>
      <div className="sitelist">
        {sites.map((s) => (
          <div key={s.id} className={`siteitem${s.id === activeSiteId ? " active" : ""}`}>
            <div className="siteinfo" onClick={() => onSelect(s.id)}>
              <div className="sitename">{s.id === activeSiteId ? "🟢 " : ""}{s.name}</div>
              <div className="sitepath">{s.filePath}</div>
            </div>
            <div className="siteactions">
              <button className="btn btn-ghost btn-sm" onClick={() => startEdit(s)}>수정</button>
              <button
                className="btn btn-ghost btn-sm"
                disabled={sites.length <= 1}
                onClick={() => {
                  if (window.confirm(`"${s.name}" 현장을 삭제할까요?\n이 기기에 저장된 해당 현장의 입력 데이터가 삭제됩니다. (OneDrive의 엑셀 파일 자체는 삭제되지 않습니다)`)) {
                    onDelete(s.id);
                  }
                }}
              >
                삭제
              </button>
            </div>
          </div>
        ))}
      </div>

      {editingId ? (
        <div style={{ background: "#fafbff", border: "1px solid var(--line)", borderRadius: 10, padding: 12 }}>
          <Field label="현장명">
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="예: ○○아파트 신축공사" />
          </Field>
          <Field label="OneDrive 파일 경로 (루트 기준)">
            <input value={filePath} onChange={(e) => setFilePath(e.target.value)} placeholder="예: 2. 근무지_현장정리/4. ○○현장/석공사_일일공정관리.xlsx" />
          </Field>
          <p style={{ fontSize: 11.5, color: "#9aa0a8", margin: "0 0 10px" }}>
            OneDrive 앱에서 해당 엑셀 파일의 경로(위치)를 그대로 입력하세요. 파일 형식·시트 구조(①기준정보, ②일일실적입력)는 기존과 동일해야 합니다.
          </p>
          <div className="row">
            <button className="btn btn-ghost" onClick={cancelForm}>취소</button>
            <button className="btn btn-primary" onClick={submitForm}>{editingId === "new" ? "추가" : "저장"}</button>
          </div>
        </div>
      ) : (
        <button className="btn btn-ghost" style={{ width: "100%" }} onClick={startAdd}>+ 새 현장 추가</button>
      )}
    </div>
  );
}
