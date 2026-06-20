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

  const [buildingsExternal, setBuildingsExternal] = useState(() => Storage.get(Storage.KEYS.buildingsExternal, BUILDINGS_EXTERNAL));
  const [buildingsInternal, setBuildingsInternal] = useState(() => Storage.get(Storage.KEYS.buildingsInternal, BUILDINGS_INTERNAL));
  const [logsExternal, setLogsExternal] = useState(() => Storage.get(Storage.KEYS.logsExternal, SEED_LOGS_EXTERNAL));
  const [logsInternal, setLogsInternal] = useState(() => Storage.get(Storage.KEYS.logsInternal, SEED_LOGS_INTERNAL));
  const [pendingExternal, setPendingExternal] = useState(() => Storage.get(Storage.KEYS.pendingExternal, []));
  const [pendingInternal, setPendingInternal] = useState(() => Storage.get(Storage.KEYS.pendingInternal, []));
  const [checklist, setChecklist] = useState(() => Storage.get(Storage.KEYS.checklist, {}));

  // 동 목록은 고정 상수가 아니라 기준정보(엑셀 ①기준정보, 동기화로 갱신됨)에서 파생.
  // 현장이 바뀌어 엑셀의 기준정보(동 목록/계획수량 등)가 바뀌면 동기화 시 여기도 자동으로 갱신됨.
  const dongListExternal = useMemo(() => buildingsExternal.map((b) => b.dong), [buildingsExternal]);
  const dongListInternal = useMemo(() => [...buildingsInternal.map((b) => b.dong), EXTRA_INTERNAL_DONG], [buildingsInternal]);

  const [formExternal, setFormExternal] = useState(() => emptyExternalForm(dongListExternal));
  const [formInternal, setFormInternal] = useState(() => emptyInternalForm(dongListInternal));
  const [recoveryDong, setRecoveryDong] = useState(() => dongListExternal[0] || "");

  const [account, setAccount] = useState(null);
  const [sync, setSync] = useState({ state: "idle", message: "", lastSyncedAt: Storage.get(Storage.KEYS.fileMeta, {})?.lastSyncedAt || null });
  const wbRef = useRef(null);
  const itemIdRef = useRef(null);

  useEffect(() => { Storage.set(Storage.KEYS.buildingsExternal, buildingsExternal); }, [buildingsExternal]);
  useEffect(() => { Storage.set(Storage.KEYS.buildingsInternal, buildingsInternal); }, [buildingsInternal]);
  useEffect(() => { Storage.set(Storage.KEYS.logsExternal, logsExternal); }, [logsExternal]);
  useEffect(() => { Storage.set(Storage.KEYS.logsInternal, logsInternal); }, [logsInternal]);
  useEffect(() => { Storage.set(Storage.KEYS.pendingExternal, pendingExternal); }, [pendingExternal]);
  useEffect(() => { Storage.set(Storage.KEYS.pendingInternal, pendingInternal); }, [pendingInternal]);
  useEffect(() => { Storage.set(Storage.KEYS.checklist, checklist); }, [checklist]);

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
    Graph.initMsal().then(() => {
      const acc = Graph.getActiveAccount();
      if (acc) setAccount(acc);
    }).catch(() => {});
  }, []);

  const dashExternal = useMemo(() => calcExternalDashboard(buildingsExternal, logsExternal), [buildingsExternal, logsExternal]);
  const dashInternal = useMemo(() => calcInternalDashboard(buildingsInternal, logsInternal), [buildingsInternal, logsInternal]);
  const recovery = useMemo(() => calcRecoveryPlan(recoveryDong, buildingsExternal, logsExternal, new Date()), [recoveryDong, buildingsExternal, logsExternal]);

  const pendingCount = pendingExternal.length + pendingInternal.length;

  async function doLogin() {
    try {
      setSync((s) => ({ ...s, state: "syncing", message: "로그인 중..." }));
      const acc = await Graph.login();
      setAccount(acc);
      await runSync();
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
    if (!account) return;
    setSync((s) => ({ ...s, state: "syncing", message: "OneDrive에서 받아오는 중..." }));
    try {
      const result = await Graph.syncDown();
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
      const meta = Storage.get(Storage.KEYS.fileMeta, {});
      Storage.set(Storage.KEYS.fileMeta, { ...meta, lastSyncedAt: now });
    } catch (e) {
      setSync((s) => ({ ...s, state: "error", message: String(e.message || e) }));
    }
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
          />
        )}
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

function SyncTab({ account, sync, pendingCount, doLogin, doLogout, runSync, configured }) {
  return (
    <div>
      {!configured && (
        <div className="banner warn">
          Microsoft 앱 등록(Client ID)이 아직 설정되지 않았습니다. Azure 앱 등록 가이드를 먼저 진행해주세요. 등록 후 .env의 VITE_MSAL_CLIENT_ID 값을 채우면 OneDrive 연결이 가능합니다.
        </div>
      )}
      <div className="card">
        <h2>OneDrive 연동</h2>
        <p style={{ fontSize: 13, color: "#6b7280", lineHeight: 1.5 }}>
          데스크탑의 석공사_일일공정관리_lsj.xlsx 파일과 동기화합니다. 모바일에서 입력 후 저장하면 자동으로 이 파일에 새 행이 추가됩니다.
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
          동기화 중에는 데스크탑에서 같은 파일을 열어두지 않는 것을 권장합니다(동시 저장 시 충돌 가능).
          서식·색상 등 일부 디자인은 동기화 과정에서 미세하게 달라질 수 있습니다.
        </p>
      </div>
    </div>
  );
}
