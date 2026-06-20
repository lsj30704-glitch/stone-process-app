import { useEffect, useMemo, useRef, useState } from "react";
import {
  BUILDINGS_EXTERNAL, BUILDINGS_INTERNAL, SEED_LOGS_EXTERNAL, SEED_LOGS_INTERNAL,
  DISASTER_OPTIONS, REASON_CODES, THRESHOLDS,
} from "./data";
import {
  calcRowExternal, calcRowInternal, calcExternalDashboard, calcInternalDashboard,
  calcRecoveryPlan, statusBadge, fmtPct, fmtNum, RECOVERY_CHECKLIST, DISASTER_MANUAL,
} from "./calc";
import { Storage } from "./storage";
import * as Graph from "./graphSync";

function today() {
  return new Date().toISOString().slice(0, 10);
}

// 내부(세대) 동 목록에는 기준정보(엑셀)에 세대수가 없는 특수 구역도 선택할 수 있게 항상 추가해줌
const EXTRA_INTERNAL_DONG = "게스트하우스";

function emptyExternalForm(dongList) {
  return { date: today(), dong: dongList?.[0] || "", masonry: "", caulking: "", truss: "", scaffold: "", actual: "", disaster: "N (정상)", reason: "", note: "", memo: "" };
}
function emptyInternalForm(dongList) {
  return { date: today(), dong: dongList?.[0] || "", masonry: "", caulking: "", truss: "", actual: "", disaster: "N (정상)", reason: "", note: "", memo: "" };
}

export default function App() {
  const [tab, setTab] = useState("input");
  const [inputMode, setInputMode] = useState("external");
  const [dashMode, setDashMode] = useState("external");

  // 현장(site)이 여러 개일 수 있음 — 각 현장은 자기만의 OneDrive 파일(경로)과 데이터를 가짐
  const [sites, setSites] = useState(() => Storage.getSites());
  const [activeSiteId, setActiveSiteId] = useState(() => Storage.getActiveSiteId());
  const activeSite = useMemo(() => sites.find((s) => s.id === activeSiteId) || sites[0] || null, [sites, activeSiteId]);
  // 저장 effect들이 "지금 어떤 현장 데이터를 쓰고 있는지"를 항상 최신으로 참조하기 위한 ref.
  // (state인 activeSiteId를 저장 effect의 deps에 넣으면, 현장 전환 시 아직 로드되지 않은
  //  이전 현장의 데이터가 새 현장 키에 잘못 저장되는 경합이 생길 수 있어 ref로 분리함)
  const currentSiteIdRef = useRef(activeSiteId);

  const [buildingsExternal, setBuildingsExternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.buildingsExternal, activeSiteId), BUILDINGS_EXTERNAL));
  const [buildingsInternal, setBuildingsInternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.buildingsInternal, activeSiteId), BUILDINGS_INTERNAL));
  const [logsExternal, setLogsExternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.logsExternal, activeSiteId), SEED_LOGS_EXTERNAL));
  const [logsInternal, setLogsInternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.logsInternal, activeSiteId), SEED_LOGS_INTERNAL));
  const [pendingExternal, setPendingExternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.pendingExternal, activeSiteId), []));
  const [pendingInternal, setPendingInternal] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.pendingInternal, activeSiteId), []));
  const [checklist, setChecklist] = useState(() => Storage.get(Storage.siteKey(Storage.KEYS.checklist, activeSiteId), {}));

  // 동 목록은 고정 상수가 아니라 기준정보(엑셀 ①기준정보, 동기화로 갱신됨)에서 파생.
  // 현장이 바뀌어 엑셀의 기준정보(동 목록/계획수량 등)가 바뀌면 동기화 시 여기도 자동으로 갱신됨.
  const dongListExternal = useMemo(() => buildingsExternal.map((b) => b.dong), [buildingsExternal]);
  const dongListInternal = useMemo(() => [...buildingsInternal.map((b) => b.dong), EXTRA_INTERNAL_DONG], [buildingsInternal]);

  const [formExternal, setFormExternal] = useState(() => emptyExternalForm(dongListExternal));
  const [formInternal, setFormInternal] = useState(() => emptyInternalForm(dongListInternal));
  const [recoveryDong, setRecoveryDong] = useState(() => dongListExternal[0] || "");

  const [account, setAccount] = useState(null);
  const [justLoggedIn, setJustLoggedIn] = useState(false);
  const [sync, setSync] = useState({ state: "idle", message: "", lastSyncedAt: Storage.get(Storage.siteKey(Storage.KEYS.fileMeta, activeSiteId), {})?.lastSyncedAt || null });
  const wbRef = useRef(null);
  const itemIdRef = useRef(null);

  // 현장을 전환하면 해당 현장의 데이터를 다시 불러옴. (마이그레이션으로 생성된 "기본 현장"만
  // 기존 시드 데이터를 기본값으로 쓰고, 새로 추가한 현장은 빈 상태에서 시작해 동기화로 채워짐)
  useEffect(() => {
    if (!activeSiteId) return;
    currentSiteIdRef.current = activeSiteId;
    const site = Storage.getSites().find((s) => s.id === activeSiteId);
    const isDefaultSite = !!site?.isDefault;
    setBuildingsExternal(Storage.get(Storage.siteKey(Storage.KEYS.buildingsExternal, activeSiteId), isDefaultSite ? BUILDINGS_EXTERNAL : []));
    setBuildingsInternal(Storage.get(Storage.siteKey(Storage.KEYS.buildingsInternal, activeSiteId), isDefaultSite ? BUILDINGS_INTERNAL : []));
    setLogsExternal(Storage.get(Storage.siteKey(Storage.KEYS.logsExternal, activeSiteId), isDefaultSite ? SEED_LOGS_EXTERNAL : []));
    setLogsInternal(Storage.get(Storage.siteKey(Storage.KEYS.logsInternal, activeSiteId), isDefaultSite ? SEED_LOGS_INTERNAL : []));
    setPendingExternal(Storage.get(Storage.siteKey(Storage.KEYS.pendingExternal, activeSiteId), []));
    setPendingInternal(Storage.get(Storage.siteKey(Storage.KEYS.pendingInternal, activeSiteId), []));
    setChecklist(Storage.get(Storage.siteKey(Storage.KEYS.checklist, activeSiteId), {}));
    setSync({ state: "idle", message: "", lastSyncedAt: Storage.get(Storage.siteKey(Storage.KEYS.fileMeta, activeSiteId), {})?.lastSyncedAt || null });
    wbRef.current = null;
    itemIdRef.current = null;
  }, [activeSiteId]);

  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.buildingsExternal, currentSiteIdRef.current), buildingsExternal); }, [buildingsExternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.buildingsInternal, currentSiteIdRef.current), buildingsInternal); }, [buildingsInternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.logsExternal, currentSiteIdRef.current), logsExternal); }, [logsExternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.logsInternal, currentSiteIdRef.current), logsInternal); }, [logsInternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.pendingExternal, currentSiteIdRef.current), pendingExternal); }, [pendingExternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.pendingInternal, currentSiteIdRef.current), pendingInternal); }, [pendingInternal]);
  useEffect(() => { Storage.set(Storage.siteKey(Storage.KEYS.checklist, currentSiteIdRef.current), checklist); }, [checklist]);

  // 기준정보 동기화로 동 목록이 바뀌어 현재 선택값이 더 이상 유효하지 않으면 새 목록의 첫 항목으로 보정
  useEffect(() => {
    if (dongListExternal.length && !dongListExternal.includes(formExternal.dong)) {
      setFormExternal((f) => ({ ...f, dong: dongListExternal[0] }));
    }
    if (dongListExternal.length && !dongListExternal.includes(recoveryDong)) {
      setRecoveryDong(dongListExternal[0]);
    }
  }, [dongListExternal]);

  useEffect(() => {
    if (dongListInternal.length && !dongListInternal.includes(formInternal.dong)) {
      setFormInternal((f) => ({ ...f, dong: dongListInternal[0] }));
    }
  }, [dongListInternal]);

  useEffect(() => {
    if (!Graph.isConfigured()) return;
    Graph.initMsal().then((result) => {
      const acc = Graph.getActiveAccount();
      if (acc) setAccount(acc);
      // 리다이렉트 로그인에서 막 돌아온 경우(result.account 존재) 자동으로 동기화 1회 실행
      if (result?.account) setJustLoggedIn(true);
    }).catch(() => {});
  }, []);

  // account가 막 설정된 직후(리다이렉트 로그인 복귀) 동기화를 한 번 트리거.
  // useEffect로 분리한 이유: mount 시점 effect의 클로저는 account를 항상 null로 캡처하므로
  // setAccount 직후 곧바로 runSync()를 호출하면 stale closure 때문에 동작하지 않음.
  useEffect(() => {
    if (justLoggedIn && account) {
      setJustLoggedIn(false);
      runSync();
    }
  }, [justLoggedIn, account]);

  const dashExternal = useMemo(() => calcExternalDashboard(buildingsExternal, logsExternal), [buildingsExternal, logsExternal]);
  const dashInternal = useMemo(() => calcInternalDashboard(buildingsInternal, logsInternal), [buildingsInternal, logsInternal]);
  const recovery = useMemo(() => calcRecoveryPlan(recoveryDong, buildingsExternal, logsExternal, new Date()), [recoveryDong, buildingsExternal, logsExternal]);

  const pendingCount = pendingExternal.length + pendingInternal.length;

  async function doLogin() {
    try {
      setSync((s) => ({ ...s, state: "syncing", message: "Microsoft 로그인 화면으로 이동 중..." }));
      // loginRedirect는 페이지를 이동시키므로 이 함수는 보통 끝까지 실행되지 않습니다.
      // 로그인 후 돌아오면 위쪽의 justLoggedIn 효과가 계정 설정과 동기화를 처리합니다.
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
      let wb = result.wb;
      const itemId = result.itemId;

      // 보류 중인 입력을 워크북에 반영
      for (const entry of pendingExternal) Graph.appendExternalRow(wb, entry);
      for (const entry of pendingInternal) Graph.appendInternalRow(wb, entry);

      if (pendingExternal.length || pendingInternal.length) {
        setSync((s) => ({ ...s, message: "변경사항 업로드 중..." }));
        await Graph.syncUp(wb, itemId);
      }

      wbRef.current = wb;
      itemIdRef.current = itemId;

      const mergedExternal = Graph.readExternalLogs(wb).map((l) => calcRowExternal(l, result.buildings.external.length ? result.buildings.external : buildingsExternal));
      const mergedInternal = Graph.readInternalLogs(wb).map((l) => calcRowInternal(l));

      setLogsExternal(mergedExternal);
      setLogsInternal(mergedInternal);
      if (result.buildings.external.length) setBuildingsExternal(result.buildings.external);
      if (result.buildings.internal.length) setBuildingsInternal(result.buildings.internal);
      setPendingExternal([]);
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

  // ---- 현장 관리 ----
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
    if (id === activeSiteId) {
      // 경로가 바뀌었으니 캐시된 워크북/itemId를 비워 다음 동기화 때 새 경로로 다시 찾게 함
      wbRef.current = null;
      itemIdRef.current = null;
    }
  }

  function handleDeleteSite(id) {
    Storage.removeSite(id);
    setSites(Storage.getSites());
    if (id === activeSiteId) setActiveSiteId(Storage.getActiveSiteId());
  }

  function saveExternal() {
    if (!formExternal.dong) return;
    const entry = {
      id: `local-${Date.now()}`,
      date: formExternal.date,
      dong: formExternal.dong,
      masonry: Number(formExternal.masonry) || 0,
      caulking: Number(formExternal.caulking) || 0,
      truss: Number(formExternal.truss) || 0,
      scaffold: Number(formExternal.scaffold) || 0,
      actual: formExternal.actual === "" ? 0 : Number(formExternal.actual),
      disaster: formExternal.disaster,
      reason: formExternal.reason,
      note: formExternal.note,
      memo: formExternal.memo,
    };
    setLogsExternal((prev) => [...prev, calcRowExternal(entry, buildingsExternal)]);
    setPendingExternal((prev) => [...prev, entry]);
    setFormExternal(emptyExternalForm(dongListExternal));
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

  return (
    <div className="app">
      <div className="header">
        <h1>🪨 석공사 공정관리</h1>
        <select
          className="sitepicker"
          value={activeSiteId || ""}
          onChange={(e) => selectSite(e.target.value)}
        >
          {sites.map((s) => (
            <option key={s.id} value={s.id}>🏗 {s.name}</option>
          ))}
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
            mode={inputMode} setMode={setInputMode}
            formExternal={formExternal} setFormExternal={setFormExternal}
            formInternal={formInternal} setFormInternal={setFormInternal}
            saveExternal={saveExternal} saveInternal={saveInternal}
            logsExternal={logsExternal} logsInternal={logsInternal}
            dongListExternal={dongListExternal} dongListInternal={dongListInternal}
          />
        )}
        {tab === "dash" && (
          <DashTab mode={dashMode} setMode={setDashMode} dashExternal={dashExternal} dashInternal={dashInternal} />
        )}
        {tab === "recovery" && (
          <RecoveryTab
            recoveryDong={recoveryDong} setRecoveryDong={setRecoveryDong} recovery={recovery}
            checklist={checklist} toggleCheck={toggleCheck} dongListExternal={dongListExternal}
          />
        )}
        {tab === "base" && (
          <BaseTab buildingsExternal={buildingsExternal} buildingsInternal={buildingsInternal} />
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

function InputTab({ mode, setMode, formExternal, setFormExternal, formInternal, setFormInternal, saveExternal, saveInternal, logsExternal, logsInternal, dongListExternal, dongListInternal }) {
  return (
    <div>
      <div className="toggle2">
        <button className={mode === "external" ? "active" : ""} onClick={() => setMode("external")}>외부(아파트)</button>
        <button className={mode === "internal" ? "active" : ""} onClick={() => setMode("internal")}>내부(세대)</button>
      </div>

      {mode === "external" ? (
        <div className="card">
          <h2>일일 실적 입력 · 외부</h2>
          <Field label="날짜">
            <input type="date" value={formExternal.date} onChange={(e) => setFormExternal({ ...formExternal, date: e.target.value })} />
          </Field>
          <Field label="해당 동(구역)">
            <select value={formExternal.dong} onChange={(e) => setFormExternal({ ...formExternal, dong: e.target.value })}>
              {dongListExternal.map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
          </Field>
          <div className="row">
            <Field label="석공(명)"><input type="number" min="0" value={formExternal.masonry} onChange={(e) => setFormExternal({ ...formExternal, masonry: e.target.value })} /></Field>
            <Field label="코킹(명)"><input type="number" min="0" value={formExternal.caulking} onChange={(e) => setFormExternal({ ...formExternal, caulking: e.target.value })} /></Field>
          </div>
          <div className="row">
            <Field label="트러스(명)"><input type="number" min="0" value={formExternal.truss} onChange={(e) => setFormExternal({ ...formExternal, truss: e.target.value })} /></Field>
            <Field label="비계(명)"><input type="number" min="0" value={formExternal.scaffold} onChange={(e) => setFormExternal({ ...formExternal, scaffold: e.target.value })} /></Field>
          </div>
          <Field label="실제시공량(m²)">
            <input type="number" step="0.01" min="0" value={formExternal.actual} onChange={(e) => setFormExternal({ ...formExternal, actual: e.target.value })} />
          </Field>
          <div className="row">
            <Field label="천재지변 여부">
              <select value={formExternal.disaster} onChange={(e) => setFormExternal({ ...formExternal, disaster: e.target.value })}>
                {DISASTER_OPTIONS.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </Field>
            <Field label="사유코드">
              <select value={formExternal.reason} onChange={(e) => setFormExternal({ ...formExternal, reason: e.target.value })}>
                <option value="">-</option>
                {REASON_CODES.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </Field>
          </div>
          <Field label="특기사항·사유">
            <textarea value={formExternal.note} onChange={(e) => setFormExternal({ ...formExternal, note: e.target.value })} />
          </Field>
          <Field label="비고">
            <textarea value={formExternal.memo} onChange={(e) => setFormExternal({ ...formExternal, memo: e.target.value })} />
          </Field>
          <button className="btn btn-primary" onClick={saveExternal}>저장</button>
        </div>
      ) : (
        <div className="card">
          <h2>일일 실적 입력 · 내부(세대)</h2>
          <Field label="날짜">
            <input type="date" value={formInternal.date} onChange={(e) => setFormInternal({ ...formInternal, date: e.target.value })} />
          </Field>
          <Field label="해당 동(구역)">
            <select value={formInternal.dong} onChange={(e) => setFormInternal({ ...formInternal, dong: e.target.value })}>
              {dongListInternal.map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
          </Field>
          <div className="row">
            <Field label="석공(명)"><input type="number" min="0" value={formInternal.masonry} onChange={(e) => setFormInternal({ ...formInternal, masonry: e.target.value })} /></Field>
            <Field label="코킹(명)"><input type="number" min="0" value={formInternal.caulking} onChange={(e) => setFormInternal({ ...formInternal, caulking: e.target.value })} /></Field>
          </div>
          <Field label="트러스(명)"><input type="number" min="0" value={formInternal.truss} onChange={(e) => setFormInternal({ ...formInternal, truss: e.target.value })} /></Field>
          <Field label="실제시공량(세대)">
            <input type="number" step="0.5" min="0" value={formInternal.actual} onChange={(e) => setFormInternal({ ...formInternal, actual: e.target.value })} />
          </Field>
          <div className="row">
            <Field label="천재지변 여부">
              <select value={formInternal.disaster} onChange={(e) => setFormInternal({ ...formInternal, disaster: e.target.value })}>
                {DISASTER_OPTIONS.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </Field>
            <Field label="사유코드">
              <select value={formInternal.reason} onChange={(e) => setFormInternal({ ...formInternal, reason: e.target.value })}>
                <option value="">-</option>
                {REASON_CODES.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </Field>
          </div>
          <Field label="특기사항·사유">
            <textarea value={formInternal.note} onChange={(e) => setFormInternal({ ...formInternal, note: e.target.value })} />
          </Field>
          <Field label="비고">
            <textarea value={formInternal.memo} onChange={(e) => setFormInternal({ ...formInternal, memo: e.target.value })} />
          </Field>
          <button className="btn btn-primary" onClick={saveInternal}>저장</button>
        </div>
      )}

      <div className="card">
        <h2>최근 입력 내역</h2>
        <div className="entrylist">
          {(mode === "external" ? logsExternal : logsInternal).slice(-8).reverse().map((l) => (
            <div key={l.id} className="entryitem">
              <div>
                <div>{l.dong} · 실적 {fmtNum(l.actual)}{mode === "external" ? "m²" : "세대"}</div>
                <div className="meta">{l.date} · 달성률 {fmtPct(l.rate)}</div>
              </div>
              {l.disaster && l.disaster !== "N (정상)" && <span className="pill dot-warn">{l.disaster}</span>}
            </div>
          ))}
          {(mode === "external" ? logsExternal : logsInternal).length === 0 && <div className="meta">입력된 실적이 없습니다.</div>}
        </div>
      </div>
    </div>
  );
}

function DashTab({ mode, setMode, dashExternal, dashInternal }) {
  return (
    <div>
      <div className="toggle2">
        <button className={mode === "external" ? "active" : ""} onClick={() => setMode("external")}>외부(아파트)</button>
        <button className={mode === "internal" ? "active" : ""} onClick={() => setMode("internal")}>내부(세대)</button>
      </div>

      {mode === "external" ? (
        <>
          <div className="card">
            <h2>전체 공사 현황 요약</h2>
            <div className="statgrid">
              <div className="stat"><div className="label">전체 계획 수량</div><div className="value">{fmtNum(dashExternal.totalPlanArea)} m²</div></div>
              <div className="stat"><div className="label">누적 실제시공량</div><div className="value">{fmtNum(dashExternal.cumActual)} m²</div></div>
              <div className="stat"><div className="label">전체 달성률</div><div className="value">{fmtPct(dashExternal.overallRate)}</div></div>
              <div className="stat"><div className="label">기간 달성률</div><div className="value">{fmtPct(dashExternal.periodRate)}</div></div>
              <div className="stat"><div className="label">총 투입인원(연인원)</div><div className="value">{fmtNum(dashExternal.totalWorkers, 0)}명</div></div>
              <div className="stat"><div className="label">1인당 평균 시공량</div><div className="value">{fmtNum(dashExternal.perWorker)} m²</div></div>
              <div className="stat"><div className="label">천재지변 발생일수</div><div className="value">{dashExternal.disasterDays}일</div></div>
              <div className="stat"><div className="label">작업 총 일수</div><div className="value">{dashExternal.totalDays}일</div></div>
            </div>
          </div>
          <div className="card">
            <h2>동별 달성률 현황</h2>
            <table className="dashtable">
              <thead><tr><th>동</th><th>계획(m²)</th><th>실적(m²)</th><th>달성률</th><th>상태</th></tr></thead>
              <tbody>
                {dashExternal.byBuilding.map((b) => (
                  <tr key={b.dong}>
                    <td className="dong">{b.dong}</td>
                    <td>{fmtNum(b.planArea)}</td>
                    <td>{fmtNum(b.cumActual)}</td>
                    <td>{fmtPct(b.rate)}</td>
                    <td style={{ color: b.status.color }}>{b.status.label}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
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

function RecoveryTab({ recoveryDong, setRecoveryDong, recovery, checklist, toggleCheck, dongListExternal }) {
  return (
    <div>
      <div className="card">
        <h2>만회계획 자동 산출</h2>
        <Field label="대상 동 선택">
          <select value={recoveryDong} onChange={(e) => setRecoveryDong(e.target.value)}>
            {dongListExternal.map((d) => <option key={d} value={d}>{d}</option>)}
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
        {recovery && (
          <div className="banner info" style={{ marginTop: 10 }}>{recovery.verdict}</div>
        )}
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

function BaseTab({ buildingsExternal, buildingsInternal }) {
  return (
    <div>
      <div className="card">
        <h2>동별 시공 계획 수량 (외부)</h2>
        <table className="dashtable">
          <thead><tr><th>동</th><th>전체수량(m²)</th><th>시작일</th><th>완료예정</th><th>기준인원</th></tr></thead>
          <tbody>
            {buildingsExternal.map((b) => (
              <tr key={b.dong}>
                <td className="dong">{b.dong}</td>
                <td>{fmtNum(b.totalArea)}</td>
                <td>{b.startDate}</td>
                <td>{b.endDate}</td>
                <td>{b.baseWorkers}명</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
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
          개인 OneDrive 계정은 엑셀 셀 단위 API를 지원하지 않아, 동기화 시 파일 전체를 받아 수정 후 다시 업로드합니다.
          동기화 중에는 PC에서 OneDrive 동기화 중인 같은 파일을 열어두지 않는 것을 권장합니다(동시 저장 시 충돌 가능).
          서식·색상 등 일부 디자인은 동기화 과정에서 미세하게 달라질 수 있습니다. 현장을 추가할 때도 같은 형식(①기준정보·②일일실적입력 시트 구조)의 엑셀 파일이어야 합니다.
        </p>
      </div>
    </div>
  );
}

function SiteManager({ sites, activeSiteId, onSelect, onAdd, onUpdate, onDelete }) {
  const [editingId, setEditingId] = useState(null); // "new" | site.id | null
  const [name, setName] = useState("");
  const [filePath, setFilePath] = useState("");

  function startAdd() {
    setEditingId("new");
    setName("");
    setFilePath("");
  }
  function startEdit(site) {
    setEditingId(site.id);
    setName(site.name);
    setFilePath(site.filePath);
  }
  function cancelForm() {
    setEditingId(null);
    setName("");
    setFilePath("");
  }
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
