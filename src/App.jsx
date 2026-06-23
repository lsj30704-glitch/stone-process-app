import { useEffect, useMemo, useRef, useState } from "react";
import {
  BUILDINGS_EXTERNAL, BUILDINGS_INTERNAL, SEED_LOGS_EXTERNAL, SEED_LOGS_INTERNAL,
  DISASTER_OPTIONS, REASON_CODES, THRESHOLDS,
} from "./data";
import {
  calcRowExternal, calcRowInternal, calcExternalDashboard, calcInternalDashboard,
  calcRecoveryPlan, fmtPct, fmtNum, RECOVERY_CHECKLIST, DISASTER_MANUAL,
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
  const [tab, setTab] = useState("input");
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

  const areaDash = useMemo(() => Object.fromEntries(AREA_SCOPES.map((s) => [s.key, calcExternalDashboard(areaB[s.key] || [], areaL[s.key] || [])])), [areaB, areaL]);
  const dashInternal = useMemo(() => calcInternalDashboard(buildingsInternal, logsInternal), [buildingsInternal, logsInternal]);
  const recovery = useMemo(() => calcRecoveryPlan(recoveryDong, areaB[recoveryScope] || [], areaL[recoveryScope] || [], new Date()), [recoveryDong, recoveryScope, areaB, areaL]);

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
    if (!form || !form.dong) return;
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
          <DashTab scope={dashScope} setScope={setDashScope} areaDash={areaDash} dashInternal={dashInternal} />
        )}
        {tab === "recovery" && (
          <RecoveryTab
            recoveryScope={recoveryScope} setRecoveryScope={setRecoveryScope}
            recoveryDong={recoveryDong} setRecoveryDong={setRecoveryDong} recovery={recovery}
            checklist={checklist} toggleCheck={toggleCheck} dongList={dongListsArea[recoveryScope] || []}
          />
        )}
        {tab === "base" && (
          <BaseTab areaB={areaB} buildingsInternal={buildingsInternal} />
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
        <TabBtn active={tab === "input"} onClick={() => setTab("input")} icon="✏️" label="입력" />
        <TabBtn active={tab === "dash"} onClick={() => setTab("dash")} icon="📊" label="현황" />
        <TabBtn active={tab === "recovery"} onClick={() => setTab("recovery")} icon="🔄" label="만회계획" />
        <TabBtn active={tab === "base"} onClick={() => setTab("base")} icon="📋" label="기준정보" />
        <TabBtn active={tab === "sync"} onClick={() => setTab("sync")} icon="☁️" label="동기화" />
      </div>
    </div>
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
        <select value={form.dong} onChange={(e) => setForm({ ...form, dong: e.target.value })}>
          {dongList.length === 0 && <option value="">(기준정보 없음 — 동기화 필요)</option>}
          {dongList.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
      </Field>
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
        <table className="dashtable">
          <thead><tr><th>동</th><th>계획(m²)</th><th>실적(m²)</th><th>달성률</th><th>상태</th></tr></thead>
          <tbody>
            {dash.byBuilding.map((b) => (
              <tr key={b.dong}>
                <td className="dong">{b.dong}</td>
                <td>{fmtNum(b.planArea)}</td>
                <td>{fmtNum(b.cumActual)}</td>
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

function DashTab({ scope, setScope, areaDash, dashInternal }) {
  const areaScope = AREA_SCOPES.find((s) => s.key === scope);
  return (
    <div>
      <ScopeToggle scope={scope} setScope={setScope} />
      {areaScope ? (
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
            <table className="dashtable">
              <thead><tr><th>동</th><th>전체세대</th><th>실적</th><th>달성률</th><th>상태</th></tr></thead>
              <tbody>
                {dashInternal.byBuilding.map((b) => (
                  <tr key={b.dong}>
                    <td className="dong">{b.dong}</td>
                    <td>{b.totalUnits}</td>
                    <td>{fmtNum(b.cumActual)}</td>
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

function RecoveryTab({ recoveryScope, setRecoveryScope, recoveryDong, setRecoveryDong, recovery, checklist, toggleCheck, dongList }) {
  return (
    <div>
      <div className="card">
        <h2>만회계획 자동 산출</h2>
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

function BaseTab({ areaB, buildingsInternal }) {
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
          <div className="stat"><div className="label">정상 기준</div><div className="value">{fmtPct(THRESHOLDS.normal)}</div></div>
          <div className="stat"><div className="label">주의 기준</div><div className="value">{fmtPct(THRESHOLDS.caution)}</div></div>
          <div className="stat"><div className="label">1인당 일일생산성</div><div className="value">{THRESHOLDS.productivityPerWorker} m²</div></div>
          <div className="stat"><div className="label">만회 허용일수</div><div className="value">{THRESHOLDS.recoveryDays}일</div></div>
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
