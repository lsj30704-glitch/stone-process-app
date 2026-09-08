import { PublicClientApplication } from "@azure/msal-browser";
import ExcelJS from "exceljs";
import { Storage } from "./storage";

// ---- MSAL 설정 ----
// VITE_MSAL_CLIENT_ID는 Azure Portal에서 앱 등록 후 발급되는 "애플리케이션(클라이언트) ID"를
// .env(또는 Vercel 환경변수)에 넣어야 합니다. 등록 전에는 로그인 버튼이 비활성화됩니다.
const CLIENT_ID = import.meta.env.VITE_MSAL_CLIENT_ID || "";
// 어떤 OneDrive 파일과 동기화할지는 더 이상 전역 상수가 아니라 현장(site)별로 다릅니다.
// 각 함수는 { id, name, filePath, fileName } 형태의 site 객체를 받아 처리합니다.
// (filePath: OneDrive 루트 기준 경로, fileName: 경로의 마지막 파일명 — 검색 폴백용)

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/common", // 개인 + 회사/학교 계정 모두 허용
    redirectUri: typeof window !== "undefined" ? window.location.origin : undefined,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

const SCOPES = ["Files.ReadWrite"];

let msalInstance = null;
function getMsal() {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication(msalConfig);
  }
  return msalInstance;
}

export function isConfigured() {
  return !!CLIENT_ID;
}

export async function initMsal() {
  const inst = getMsal();
  await inst.initialize();
  // 모바일 인앱 브라우저(카카오톡/네이버 등)에서는 팝업이 "block_nested_popups"로
  // 차단되므로 로그인은 리다이렉트 방식을 사용합니다. 리다이렉트로 로그인 후
  // 돌아왔을 때 여기서 결과를 처리하고 활성 계정으로 설정합니다.
  const result = await inst.handleRedirectPromise().catch(() => null);
  if (result?.account) inst.setActiveAccount(result.account);
  return result;
}

export function getActiveAccount() {
  const inst = getMsal();
  const accounts = inst.getAllAccounts();
  return accounts && accounts.length ? accounts[0] : null;
}

export async function login() {
  const inst = getMsal();
  // 팝업 대신 리다이렉트: 페이지가 Microsoft 로그인 화면으로 이동했다가 돌아오면
  // initMsal()의 handleRedirectPromise가 로그인 결과를 처리합니다.
  await inst.loginRedirect({ scopes: SCOPES });
}

export function logout() {
  const inst = getMsal();
  const account = getActiveAccount();
  if (account) inst.logoutRedirect({ account });
}

export async function getToken() {
  const inst = getMsal();
  const account = getActiveAccount();
  if (!account) throw new Error("NOT_LOGGED_IN");
  try {
    const res = await inst.acquireTokenSilent({ scopes: SCOPES, account });
    return res.accessToken;
  } catch {
    // 자동(무인) 토큰 갱신 실패 시에도 팝업 대신 리다이렉트로 재인증합니다.
    await inst.acquireTokenRedirect({ scopes: SCOPES, account });
    throw new Error("REDIRECTING_FOR_TOKEN");
  }
}

// ---- Graph 파일 탐색/다운로드/업로드 ----
// 주의: Microsoft Graph의 Excel "workbook" REST API(셀 단위 읽기/쓰기)는
// 개인(Consumer) OneDrive 계정에서는 지원되지 않습니다.
// 그래서 파일 전체를 받아(content) 브라우저에서 수정한 뒤
// 파일 전체를 다시 업로드(content)하는 방식으로 동작합니다.
//
// [중요] 예전에는 이 수정을 SheetJS(xlsx)로 했는데, SheetJS 무료판은 쓰기 시
// 셀 채움색·글꼴 등 서식을 보존하지 못하고 줄바꿈을 "_x000D_"로 깨뜨립니다.
// 그래서 서식·메모·데이터유효성·조건부서식을 모두 보존하는 ExcelJS로 교체했습니다.
// 우리는 "값만" 입력 행에 채우고, 기존 템플릿 행의 서식은 그대로 두므로 색이 깨지지 않습니다.

const GRAPH = "https://graph.microsoft.com/v1.0";

async function graphFetch(url, token, opts = {}) {
  const res = await fetch(url, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, ...(opts.headers || {}) },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph API 오류 (${res.status}): ${text}`);
  }
  return res;
}

export async function findFileId(token, site) {
  const fileMetaKey = Storage.siteKey(Storage.KEYS.fileMeta, site.id);

  // 캐시된 itemId가 있어도 무조건 신뢰하지 않고, 실제로 같은 파일을 가리키는지 먼저 확인합니다.
  // (예전에 잘못 캐시된 값이 영구히 남아 동기화가 계속 실패하는 문제를 방지)
  const cached = Storage.get(fileMetaKey, null);
  if (cached?.itemId) {
    try {
      const res = await fetch(`${GRAPH}/me/drive/items/${cached.itemId}`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (res.ok) {
        const meta = await res.json();
        if (meta.name === site.fileName) return cached.itemId;
      }
    } catch {
      // 확인 실패 시 무시하고 아래에서 다시 찾음
    }
  }

  // 1) OneDrive 내 정확한 경로로 직접 조회 (검색 인덱싱 지연과 무관하게 항상 정확)
  try {
    const encodedPath = site.filePath.split("/").map(encodeURIComponent).join("/");
    const res = await graphFetch(`${GRAPH}/me/drive/root:/${encodedPath}`, token);
    const json = await res.json();
    if (json?.id) {
      Storage.set(fileMetaKey, { itemId: json.id, name: json.name, webUrl: json.webUrl });
      return json.id;
    }
  } catch {
    // 경로가 바뀌었거나 조회 실패 시 아래 검색으로 폴백
  }

  // 2) 파일명으로 전체 OneDrive 검색 (폴더 구조가 바뀐 경우의 대비책)
  const q = encodeURIComponent(site.fileName);
  const res = await graphFetch(`${GRAPH}/me/drive/root/search(q='${q}')`, token);
  const json = await res.json();
  const match = (json.value || []).find((it) => it.name === site.fileName) || (json.value || [])[0];
  if (!match) throw new Error(`OneDrive에서 "${site.fileName}" 파일을 찾을 수 없습니다. (경로: ${site.filePath})`);
  Storage.set(fileMetaKey, { itemId: match.id, name: match.name, webUrl: match.webUrl });
  return match.id;
}

export async function downloadWorkbookArrayBuffer(token, itemId) {
  const res = await graphFetch(`${GRAPH}/me/drive/items/${itemId}/content`, token);
  return await res.arrayBuffer();
}

export async function uploadWorkbookArrayBuffer(token, itemId, arrayBuffer) {
  await graphFetch(`${GRAPH}/me/drive/items/${itemId}/content`, token, {
    method: "PUT",
    headers: { "Content-Type": "application/octet-stream" },
    body: arrayBuffer,
  });
}

// ---- ExcelJS 워크북 로드/직렬화 ----

export async function parseWorkbook(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  return wb;
}

// ExcelJS는 데이터 유효성(드롭다운)을 셀 단위로 들고 있다가 저장 시 범위로 묶는데,
// 이때 같은 범위가 중복(C5:C104 + C10:C104처럼 겹침)으로 출력되어 Excel이
// "파일을 복구해야 합니다" 경고를 띄울 수 있습니다. 저장 직전에 열·규칙별로 하나의
// 연속 범위로 정리해 중복을 제거합니다.
function normalizeDataValidations(ws) {
  const dv = ws.dataValidations;
  if (!dv || !dv.model) return;
  const model = dv.model;
  const addrs = Object.keys(model);
  if (!addrs.length) return;

  const groups = {};
  for (const addr of addrs) {
    if (addr.indexOf(":") !== -1) {
      // 이미 범위 형태면 그대로 보존
      groups[`${addr}__range`] = { rangeKey: addr, rule: model[addr] };
      continue;
    }
    const col = addr.replace(/[0-9]/g, "");
    const rowNum = parseInt(addr.replace(/[^0-9]/g, ""), 10);
    const sig = `${col}||${JSON.stringify(model[addr])}`;
    if (!groups[sig]) groups[sig] = { col, rows: [], rule: model[addr] };
    groups[sig].rows.push(rowNum);
  }

  const next = {};
  for (const key in groups) {
    const g = groups[key];
    if (g.rangeKey) {
      next[g.rangeKey] = g.rule;
      continue;
    }
    g.rows.sort((a, b) => a - b);
    const lo = g.rows[0];
    const hi = g.rows[g.rows.length - 1];
    next[`${g.col}${lo}:${g.col}${hi}`] = g.rule;
  }
  dv.model = next;
}

export async function serializeWorkbook(wb) {
  for (const ws of wb.worksheets) normalizeDataValidations(ws);
  // ArrayBuffer 반환 (Graph 업로드 body로 그대로 사용)
  return await wb.xlsx.writeBuffer();
}

// ---- 시트 ↔ 앱 데이터 변환 ----

const SHEET_EXTERNAL = "②일일실적입력";
const SHEET_INTERNAL = "②일일실적입력(내부)";
const SHEET_BASE = "①기준정보";
export const SHEET_ACHV_EXTERNAL = "③달성률현황";
export const SHEET_ACHV_INTERNAL = "내부 달성률현황";

// 일일실적입력 시트의 입력 가능 행 범위.
// [변경 이유] 예전에는 5~104행으로 하드코딩돼 있어서 105행부터의 실적이 통째로 무시됐고,
// 시트가 104행까지 차면 앱에서 새 실적 저장 자체가 불가능했습니다.
// 이제 시트의 실제 마지막 행을 기준으로 동적으로 잡습니다.
const LOG_START_ROW = 5;
const LOG_HARD_MAX_ROW = 2000; // 무한 스캔 방지용 상한

// B열(날짜)에 값이 있는 마지막 행. 중간 빈 행이 있어도 끝까지 훑어 마지막을 찾습니다.
function lastDataRow(ws, col = "B", start = LOG_START_ROW) {
  const scanEnd = Math.min(Math.max(ws.rowCount || start, start), LOG_HARD_MAX_ROW);
  let last = start - 1;
  for (let r = start; r <= scanEnd; r++) {
    const v = rawCell(ws, `${col}${r}`);
    if (v !== null && v !== "") last = r;
  }
  return last;
}

// ExcelJS 셀 값을 단순 원시값(문자열/숫자/Date/null)으로 정규화.
// - 수식 셀: { formula, result } → result
// - 하이퍼링크: { text, hyperlink } → text
// - 서식있는 텍스트: { richText:[...] } → 텍스트 합치기
function rawCell(ws, addr) {
  const v = ws.getCell(addr).value;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    if ("result" in v) {
      const r = v.result;
      // 수식 결과가 에러 객체({error})인 경우 빈값 처리
      return r && typeof r === "object" ? "" : r ?? "";
    }
    if ("text" in v) return v.text ?? "";
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
    return "";
  }
  return v;
}

function excelDateToISO(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    // 엑셀 일련번호(1900 체계) → ISO 날짜
    const ms = Math.round((v - 25569) * 86400 * 1000);
    return new Date(ms).toISOString().slice(0, 10);
  }
  return v || "";
}

export function readExternalLogs(wb) {
  const ws = wb.getWorksheet(SHEET_EXTERNAL);
  if (!ws) return [];
  const out = [];
  const endRow = lastDataRow(ws);
  for (let r = LOG_START_ROW; r <= endRow; r++) {
    const b = rawCell(ws, `B${r}`);
    if (b === null || b === "") continue;
    out.push({
      id: `row-${r}`,
      row: r,
      date: excelDateToISO(b),
      dong: rawCell(ws, `C${r}`) ?? "",
      masonry: rawCell(ws, `D${r}`) ?? 0,
      caulking: rawCell(ws, `E${r}`) ?? 0,
      truss: rawCell(ws, `F${r}`) ?? 0,
      scaffold: rawCell(ws, `G${r}`) ?? 0,
      actual: rawCell(ws, `J${r}`) ?? "",
      disaster: rawCell(ws, `L${r}`) ?? "",
      reason: rawCell(ws, `M${r}`) ?? "",
      note: rawCell(ws, `N${r}`) ?? "",
      memo: rawCell(ws, `O${r}`) ?? "",
    });
  }
  return out;
}

export function readInternalLogs(wb) {
  const ws = wb.getWorksheet(SHEET_INTERNAL);
  if (!ws) return [];
  const out = [];
  const endRow = lastDataRow(ws);
  for (let r = LOG_START_ROW; r <= endRow; r++) {
    const b = rawCell(ws, `B${r}`);
    if (b === null || b === "") continue;
    out.push({
      id: `row-${r}`,
      row: r,
      date: excelDateToISO(b),
      dong: rawCell(ws, `C${r}`) ?? "",
      masonry: rawCell(ws, `D${r}`) ?? 0,
      caulking: rawCell(ws, `E${r}`) ?? 0,
      truss: rawCell(ws, `F${r}`) ?? 0,
      actual: rawCell(ws, `I${r}`) ?? "",
      disaster: rawCell(ws, `K${r}`) ?? "",
      reason: rawCell(ws, `L${r}`) ?? "",
      note: rawCell(ws, `M${r}`) ?? "",
      memo: rawCell(ws, `N${r}`) ?? "",
    });
  }
  return out;
}

// ①기준정보 시트는 사용자가 중간에 공사 범위 블록(호이스트·부대시설 등)을 추가/이동할 수 있어,
// 고정 행 위치로 읽으면 행이 밀릴 때 엉뚱한 데이터를 읽게 됩니다. 그래서 "▶ ○○ 수량/기준정보"
// 헤더 라벨로 블록을 찾아 그 안의 데이터 행만 읽습니다. (행이 밀려도 안전)
function findBlockDataRows(ws, headerIncludes, maxRow = 300) {
  let headerRow = null;
  for (let r = 1; r <= maxRow; r++) {
    const b = rawCell(ws, `B${r}`);
    if (typeof b === "string" && b.includes(headerIncludes)) {
      headerRow = r;
      break;
    }
  }
  if (!headerRow) return [];
  const rows = [];
  // 헤더 다음 행은 컬럼 제목 행 → 그 다음부터가 데이터. 다음 블록(▶) 또는 "합 계" 행에서 종료.
  for (let r = headerRow + 2; r <= maxRow; r++) {
    const b = rawCell(ws, `B${r}`);
    const bs = b == null ? "" : String(b).trim();
    if (bs.startsWith("▶")) break;
    if (bs.replace(/\s/g, "").startsWith("합계")) break;
    if (bs === "") continue; // 동(구역)명이 빈 보조행은 건너뜀
    rows.push(r);
  }
  return rows;
}

export function readBuildings(wb) {
  const ws = wb.getWorksheet(SHEET_BASE);
  if (!ws) return { external: [], internal: [] };

  const external = findBlockDataRows(ws, "동별 시공 계획 수량").map((r) => ({
    dong: rawCell(ws, `B${r}`),
    totalArea: rawCell(ws, `C${r}`) ?? 0,
    startDate: excelDateToISO(rawCell(ws, `D${r}`)),
    endDate: excelDateToISO(rawCell(ws, `E${r}`)),
    workDays: rawCell(ws, `F${r}`) ?? 0,
    dailyPlan: rawCell(ws, `G${r}`) ?? 0,
    baseWorkers: rawCell(ws, `H${r}`) ?? 0,
    hoist: rawCell(ws, `I${r}`) ?? 0,
    note: rawCell(ws, `J${r}`) ?? "",
  }));

  const internal = findBlockDataRows(ws, "내부(세대) 기준정보").map((r) => {
    const totalUnits = rawCell(ws, `C${r}`) ?? 0;
    const optionUnits = rawCell(ws, `D${r}`) ?? 0;
    return { dong: rawCell(ws, `B${r}`), totalUnits, optionUnits, normalUnits: totalUnits - optionUnits };
  });

  return { external, internal };
}

// ①기준정보의 달성률 경보/만회 설정값을 라벨(부분일치)로 찾아 읽습니다.
// 위쪽에 블록이 추가돼 행이 밀려도 안전하도록 셀 위치를 고정하지 않고 라벨 텍스트로 검색합니다.
// 엑셀 구조상 라벨은 C열, 값은 바로 오른쪽 D열. %값(95 등)은 /100 해서 소수로 저장합니다.
// 못 찾은 값은 결과 객체에서 생략 → 앱이 기본 THRESHOLDS로 채웁니다.
export function readThresholds(wb) {
  const ws = wb.getWorksheet(SHEET_BASE);
  const out = {};
  if (!ws) return out;
  const specs = [
    { match: "정상기준", key: "normal", pct: true },
    { match: "주의기준", key: "caution", pct: true },
    { match: "위험기준", key: "danger", pct: true },
    { match: "천재지변면책", key: "disasterExemption", pct: true },
    { match: "만회허용일수", key: "recoveryDays", pct: false },
    { match: "1인당일일시공량", key: "productivityPerWorker", pct: false },
  ];
  const maxRow = ws.rowCount || 300;
  for (let r = 1; r <= maxRow; r++) {
    const label = rawCell(ws, `C${r}`);
    if (typeof label !== "string" || !label.trim()) continue;
    const norm = label.replace(/\s/g, "");
    for (const sp of specs) {
      if (out[sp.key] !== undefined) continue;
      if (norm.includes(sp.match)) {
        const v = Number(rawCell(ws, `D${r}`));
        if (isFinite(v)) out[sp.key] = sp.pct ? v / 100 : v;
      }
    }
  }
  return out;
}

// ===== 공사 범위(scope) 일반화 — 외부/호이스트/부대시설처럼 면적(m²)형 시트 공용 =====
// 면적형 일일실적 시트(외부와 동일한 컬럼 구조)를 시트명만 받아 일반적으로 읽음.
export function readAreaLogs(wb, sheetName) {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) return [];
  const out = [];
  const endRow = lastDataRow(ws);
  for (let r = LOG_START_ROW; r <= endRow; r++) {
    const b = rawCell(ws, `B${r}`);
    if (b === null || b === "") continue;
    out.push({
      id: `row-${r}`,
      row: r,
      date: excelDateToISO(b),
      dong: rawCell(ws, `C${r}`) ?? "",
      masonry: rawCell(ws, `D${r}`) ?? 0,
      caulking: rawCell(ws, `E${r}`) ?? 0,
      truss: rawCell(ws, `F${r}`) ?? 0,
      scaffold: rawCell(ws, `G${r}`) ?? 0,
      actual: rawCell(ws, `J${r}`) ?? "",
      disaster: rawCell(ws, `L${r}`) ?? "",
      reason: rawCell(ws, `M${r}`) ?? "",
      note: rawCell(ws, `N${r}`) ?? "",
      memo: rawCell(ws, `O${r}`) ?? "",
    });
  }
  return out;
}

// 면적형 시트의 투입인원 슬롯(D~G 4칸)과 그 헤더(4행) 라벨.
// 사용자가 엑셀 헤더 텍스트만 바꾸면 앱 입력폼의 칸 이름/개수가 자동으로 따라갑니다.
const AREA_FIELD_SLOTS = [
  { key: "masonry", col: "D" },
  { key: "caulking", col: "E" },
  { key: "truss", col: "F" },
  { key: "scaffold", col: "G" },
];

function cleanFieldLabel(s) {
  return String(s ?? "")
    .replace(/_x000D_/g, "")
    .replace(/[\r\n]+/g, " ")
    .replace(/투입\s*인원/g, "")
    .replace(/\(\s*명\s*\)/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// 헤더(기본 4행)에서 투입인원 컬럼(D~G) 라벨을 읽어 동적 입력필드 구성을 만든다.
// 라벨이 빈 슬롯은 제외 → 범위별로 공정(인원) 항목 수·이름이 다를 수 있음.
export function readAreaFields(wb, sheetName, headerRow = 4) {
  const ws = wb.getWorksheet(sheetName);
  const out = [];
  if (!ws) return out;
  for (const slot of AREA_FIELD_SLOTS) {
    const label = cleanFieldLabel(rawCell(ws, `${slot.col}${headerRow}`));
    if (label) out.push({ key: slot.key, col: slot.col, label });
  }
  return out;
}

// 면적형 기준정보 블록을 "▶ ○○ 수량" 라벨로 찾아 읽음 (외부/호이스트/부대시설 공용).
export function readAreaBuildings(wb, headerIncludes) {
  const ws = wb.getWorksheet(SHEET_BASE);
  if (!ws) return [];
  return findBlockDataRows(ws, headerIncludes).map((r) => ({
    dong: rawCell(ws, `B${r}`),
    totalArea: rawCell(ws, `C${r}`) ?? 0,
    startDate: excelDateToISO(rawCell(ws, `D${r}`)),
    endDate: excelDateToISO(rawCell(ws, `E${r}`)),
    workDays: rawCell(ws, `F${r}`) ?? 0,
    dailyPlan: rawCell(ws, `G${r}`) ?? 0,
    baseWorkers: rawCell(ws, `H${r}`) ?? 0,
    hoist: rawCell(ws, `I${r}`) ?? 0,
    note: rawCell(ws, `J${r}`) ?? "",
  }));
}

// 면적형 시트에 새 행 추가 (시트명 + 계획수량 참조범위를 받아 일반화).
// planRange 예: "①기준정보!$B$18:$G$25"
export function appendAreaRow(wb, sheetName, planRange, entry) {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) throw new Error(`시트를 찾을 수 없습니다: ${sheetName}`);
  const row = findNextEmptyRow(ws);
  if (!row) throw new Error(`${sheetName} 시트에 새 행을 추가할 수 없습니다 (행 상한 초과).`);
  const templateRow = row > 5 ? row - 1 : 5;

  ws.getCell(`A${row}`).value = row - 4;
  setDateCell(ws, `B${row}`, entry.date, templateRow);
  ws.getCell(`C${row}`).value = entry.dong;
  ws.getCell(`D${row}`).value = Number(entry.masonry) || 0;
  ws.getCell(`E${row}`).value = Number(entry.caulking) || 0;
  ws.getCell(`F${row}`).value = Number(entry.truss) || 0;
  ws.getCell(`G${row}`).value = Number(entry.scaffold) || 0;
  ws.getCell(`H${row}`).value = { formula: `SUM(D${row}:G${row})` };
  // 목표 시공량(I열)은 시트마다 규칙이 다릅니다.
  //   ②일일실적입력 = H*11 / ②일일실적입력(부대시설) = D*9 / ②일일실적입력(호이스트) = VLOOKUP 계획일평균
  // 그래서 바로 윗 행(템플릿)의 수식을 그대로 물려받고, 행 번호만 바꿔 씁니다.
  // 못 읽으면 기존처럼 기준정보 계획일평균 VLOOKUP으로 폴백합니다.
  const inherited = shiftFormulaRow(cellFormulaOf(ws, `I${templateRow}`), templateRow, row);
  ws.getCell(`I${row}`).value = {
    formula: inherited || `IFERROR(IF(C${row}="","",VLOOKUP(C${row},${planRange},6,0)),"")`,
  };
  ws.getCell(`J${row}`).value = entry.actual === "" || entry.actual === null ? 0 : Number(entry.actual);
  ws.getCell(`K${row}`).value = { formula: `IFERROR(IF(OR(J${row}="",I${row}="",I${row}=0),"",J${row}/I${row}),"")` };
  ws.getCell(`L${row}`).value = entry.disaster || "";
  ws.getCell(`M${row}`).value = entry.reason || "";
  ws.getCell(`N${row}`).value = entry.note || "";
  ws.getCell(`O${row}`).value = entry.memo || "";
  return row;
}

// 셀의 원본 수식 문자열(없으면 null)
function cellFormulaOf(ws, addr) {
  const v = ws.getCell(addr).value;
  if (v && typeof v === "object" && typeof v.formula === "string") return v.formula;
  return null;
}

// 상대참조 행 번호만 옮긴다: "D23*9" (23→24) → "D24*9".
// 절대참조($B$6)는 열 문자 뒤에 "$"가 붙어 패턴에 걸리지 않고, 숫자 리터럴(*9, 11)도
// 앞에 열 문자가 없어 걸리지 않는다. (구형 사파리 호환 위해 lookbehind 미사용)
function shiftFormulaRow(formula, fromRow, toRow) {
  if (!formula || fromRow === toRow) return formula;
  const re = new RegExp(`([A-Z]{1,3})${fromRow}(?![0-9])`, "g");
  return formula.replace(re, (m, col, offset, str) => {
    const prev = offset > 0 ? str[offset - 1] : "";
    // 앞 글자가 영문/숫자/$ 이면 셀 주소가 아니거나 절대참조 → 그대로 둔다
    if (/[A-Za-z0-9$]/.test(prev)) return m;
    return `${col}${toRow}`;
  });
}

// 새 실적을 쓸 행 = 마지막 데이터 행의 바로 다음 행.
// [변경 이유] 예전에는 5~104행만 훑어서, 시트가 104행까지 차면 중간 빈 행에 덮어쓰거나
// "입력 가능한 행이 모두 채워졌습니다" 오류로 저장이 아예 막혔습니다.
function findNextEmptyRow(ws, startRow = LOG_START_ROW) {
  const last = lastDataRow(ws, "B", startRow);
  const next = Math.max(last + 1, startRow);
  return next <= LOG_HARD_MAX_ROW ? next : null;
}

// 입력 행은 템플릿(5~104행)에 이미 색·글꼴·표시형식이 들어 있으므로 "값만" 채웁니다.
// (ExcelJS는 값을 설정해도 기존 셀 스타일을 유지하므로 색이 깨지지 않습니다.)
// 날짜는 사용자의 로컬 타임존(예: KST)에서 하루 밀리지 않도록 UTC 자정으로 만듭니다.
function setDateCell(ws, addr, isoDate, templateRow) {
  const cell = ws.getCell(addr);
  cell.value = new Date(`${isoDate}T00:00:00Z`);
  // 혹시 표시형식이 비어 있으면 바로 윗 템플릿 행의 형식을 복사 (안전장치)
  const col = addr.replace(/[0-9]/g, "");
  const tmpl = ws.getCell(`${col}${templateRow}`);
  if (!cell.numFmt && tmpl.numFmt) cell.numFmt = tmpl.numFmt;
}

// 새 외부 실적 행을 워크북에 추가 (값 + 기존과 동일한 수식 패턴)
export function appendExternalRow(wb, entry) {
  const ws = wb.getWorksheet(SHEET_EXTERNAL);
  if (!ws) throw new Error(`시트를 찾을 수 없습니다: ${SHEET_EXTERNAL}`);
  const row = findNextEmptyRow(ws);
  if (!row) throw new Error("②일일실적입력 시트에 새 행을 추가할 수 없습니다 (행 상한 초과).");
  const templateRow = row > 5 ? row - 1 : 5;

  ws.getCell(`A${row}`).value = row - 4;
  setDateCell(ws, `B${row}`, entry.date, templateRow);
  ws.getCell(`C${row}`).value = entry.dong;
  ws.getCell(`D${row}`).value = Number(entry.masonry) || 0;
  ws.getCell(`E${row}`).value = Number(entry.caulking) || 0;
  ws.getCell(`F${row}`).value = Number(entry.truss) || 0;
  ws.getCell(`G${row}`).value = Number(entry.scaffold) || 0;
  ws.getCell(`H${row}`).value = { formula: `SUM(D${row}:G${row})` };
  ws.getCell(`I${row}`).value = { formula: `IFERROR(IF(C${row}="","",VLOOKUP(C${row},①기준정보!$B$6:$G$13,6,0)),"")` };
  ws.getCell(`J${row}`).value = entry.actual === "" || entry.actual === null ? 0 : Number(entry.actual);
  ws.getCell(`K${row}`).value = { formula: `IFERROR(IF(OR(J${row}="",I${row}="",I${row}=0),"",J${row}/I${row}),"")` };
  ws.getCell(`L${row}`).value = entry.disaster || "";
  ws.getCell(`M${row}`).value = entry.reason || "";
  ws.getCell(`N${row}`).value = entry.note || "";
  ws.getCell(`O${row}`).value = entry.memo || "";

  return row;
}

// 새 내부(세대) 실적 행을 워크북에 추가
export function appendInternalRow(wb, entry) {
  const ws = wb.getWorksheet(SHEET_INTERNAL);
  if (!ws) throw new Error(`시트를 찾을 수 없습니다: ${SHEET_INTERNAL}`);
  const row = findNextEmptyRow(ws);
  if (!row) throw new Error("②일일실적입력(내부) 시트에 새 행을 추가할 수 없습니다 (행 상한 초과).");
  const templateRow = row > 5 ? row - 1 : 5;

  ws.getCell(`A${row}`).value = row - 4;
  setDateCell(ws, `B${row}`, entry.date, templateRow);
  ws.getCell(`C${row}`).value = entry.dong;
  ws.getCell(`D${row}`).value = Number(entry.masonry) || 0;
  ws.getCell(`E${row}`).value = Number(entry.caulking) || 0;
  ws.getCell(`F${row}`).value = Number(entry.truss) || 0;
  ws.getCell(`G${row}`).value = { formula: `SUM(D${row}:F${row})` };
  ws.getCell(`H${row}`).value = { formula: `G${row}*3` };
  ws.getCell(`I${row}`).value = entry.actual === "" || entry.actual === null ? 0 : Number(entry.actual);
  ws.getCell(`J${row}`).value = { formula: `IFERROR(IF(OR(I${row}="",H${row}="",H${row}=0),"",I${row}/H${row}),"")` };
  ws.getCell(`K${row}`).value = entry.disaster || "";
  ws.getCell(`L${row}`).value = entry.reason || "";
  ws.getCell(`M${row}`).value = entry.note || "";
  ws.getCell(`N${row}`).value = entry.memo || "";

  return row;
}

// ===== ③달성률현황 / 내부 달성률현황 시트 읽기 =====
// 엑셀이 이미 계산해 둔 "동별 달성률 현황" 표를 그대로 읽어옵니다.
// 앱이 ②시트에서 자체 재계산하던 값과 엑셀 화면 숫자가 어긋나던 문제를 없애기 위함입니다.
//
// [핵심 주의] 이 시트의 값은 "엑셀에서 마지막으로 재계산·저장된 시점"의 값입니다.
// ExcelJS는 수식을 계산하지 않고 파일에 저장된 캐시 결과만 읽습니다.
// 게다가 이 시트의 SUMIF 범위는 ②일일실적입력!C5:C114 처럼 행이 고정돼 있어
// 그 범위를 넘어선 최신 실적은 엑셀 숫자에도 빠져 있습니다.
// 그래서 수식에서 참조 범위의 끝 행(coveredMaxRow)까지 같이 뽑아내,
// 앱이 그 뒤 행들을 보충 합산할 수 있게 합니다. (calc.js의 mergeAchievement)

// 셀 주소에서 열 문자만 (예: "D21" → "D")
function colLetter(n) {
  let s = "";
  let x = n;
  while (x > 0) {
    const m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

// 수식 문자열에서 참조된 범위들의 "끝 행 번호" 중 최대값을 뽑음.
// 예) 'IFERROR(SUMIF(②일일실적입력!C5:C114,"104동",②일일실적입력!J5:J114),0)' → 114
function formulaMaxRefRow(formula) {
  if (typeof formula !== "string") return null;
  let max = null;
  const re = /\$?[A-Z]{1,3}\$?(\d+)\s*:\s*\$?[A-Z]{1,3}\$?(\d+)/g;
  let m;
  while ((m = re.exec(formula)) !== null) {
    const end = parseInt(m[2], 10);
    if (isFinite(end) && (max === null || end > max)) max = end;
  }
  return max;
}

// 셀의 원본 수식(있으면)을 반환
function cellFormula(ws, addr) {
  const v = ws.getCell(addr).value;
  if (v && typeof v === "object" && typeof v.formula === "string") return v.formula;
  return null;
}

// 헤더 행에서 라벨로 열을 찾아 매핑. 열 위치가 바뀌어도 따라갑니다.
const ACHV_COL_SPECS = [
  { key: "dong", test: (s) => s.includes("동") && (s.includes("구역") || s === "동") },
  { key: "planArea", test: (s) => s.includes("계획수량") },
  { key: "totalUnits", test: (s) => s.includes("전체세대수") },
  { key: "optionUnits", test: (s) => s.includes("옵션세대수") },
  { key: "cumActual", test: (s) => s.includes("누적실적") },
  { key: "rate", test: (s) => s.includes("달성률") },
  { key: "remain", test: (s) => s.includes("잔여") },
  { key: "endDate", test: (s) => s.includes("완료예정") },
  { key: "status", test: (s) => s.includes("상태") },
];

function mapAchvColumns(ws, headerRow, maxCol = 14) {
  const map = {};
  for (let c = 1; c <= maxCol; c++) {
    const raw = rawCell(ws, `${colLetter(c)}${headerRow}`);
    if (typeof raw !== "string" || !raw.trim()) continue;
    const norm = raw.replace(/\s/g, "");
    for (const spec of ACHV_COL_SPECS) {
      if (map[spec.key] !== undefined) continue;
      if (spec.test(norm)) map[spec.key] = colLetter(c);
    }
  }
  return map;
}

// "▶ ... 동별 ... 현황" 블록의 헤더 행을 찾음 (B열 스캔)
function findAchvHeaderRow(ws, maxRow) {
  for (let r = 1; r <= maxRow; r++) {
    const b = rawCell(ws, `B${r}`);
    if (typeof b !== "string") continue;
    const s = b.replace(/\s/g, "");
    if (s.startsWith("▶") && s.includes("동별") && s.includes("달성률")) return r + 1; // 다음 행이 컬럼 제목 행
  }
  // 폴백: "동(구역)"이 들어간 행을 직접 찾음
  for (let r = 1; r <= maxRow; r++) {
    const b = rawCell(ws, `B${r}`);
    if (typeof b === "string" && b.replace(/\s/g, "").includes("동(구역)")) return r;
  }
  return null;
}

// 상단 "▶ 전체 ... 현황 요약" 블록: B=항목명, C=값 을 라벨→값 맵으로 읽음
const ACHV_SUMMARY_SPECS = [
  { key: "totalPlanArea", match: "전체계획수량" },
  { key: "totalUnits", match: "전체세대수" },
  { key: "optionUnits", match: "옵션세대수" },
  { key: "normalUnits", match: "일반세대수" },
  { key: "cumActual", match: "누적실제시공량" },
  { key: "cumPlan", match: "누적계획량" },
  { key: "overallRate", match: "전체달성률" },
  { key: "periodRate", match: "기간달성률" },
  { key: "totalWorkers", match: "총투입인원" },
  { key: "perWorker", match: "1인당평균시공량" },
  { key: "disasterDays", match: "천재지변발생일수" },
  { key: "totalDays", match: "작업총일수" },
];

function readAchvSummary(ws, maxRow) {
  const out = {};
  const covered = {};
  for (let r = 1; r <= maxRow; r++) {
    const label = rawCell(ws, `B${r}`);
    if (typeof label !== "string" || !label.trim()) continue;
    const norm = label.replace(/\s/g, "");
    for (const sp of ACHV_SUMMARY_SPECS) {
      if (out[sp.key] !== undefined) continue;
      if (norm.includes(sp.match)) {
        const v = rawCell(ws, `C${r}`);
        const n = Number(v);
        out[sp.key] = isFinite(n) && v !== "" && v !== null ? n : v;
        const mr = formulaMaxRefRow(cellFormula(ws, `C${r}`));
        if (mr) covered[sp.key] = mr;
      }
    }
  }
  return { summary: out, summaryCovered: covered };
}

/**
 * 달성률현황 시트를 읽어 { byDong, summary, coveredMaxRow, sheetName } 반환.
 * 시트가 없거나 표를 못 찾으면 null → 호출측에서 기존 재계산 방식으로 폴백.
 */
export function readAchievement(wb, sheetName) {
  let ws = wb.getWorksheet(sheetName);
  if (!ws) {
    // 시트명이 바뀐 경우를 대비해 "달성률"이 들어간 시트로 폴백 (내부/외부 구분 유지)
    const wantInternal = String(sheetName).includes("내부");
    ws = wb.worksheets.find(
      (w) => w?.name?.includes("달성률") && String(w.name).includes("내부") === wantInternal
    );
  }
  if (!ws) return null;

  const maxRow = Math.min(ws.rowCount || 100, 500);
  const headerRow = findAchvHeaderRow(ws, maxRow);
  if (!headerRow) return null;

  const cols = mapAchvColumns(ws, headerRow);
  if (!cols.dong || !cols.cumActual) return null;

  const byDong = [];
  let coveredMaxRow = null;
  for (let r = headerRow + 1; r <= maxRow; r++) {
    const dongRaw = rawCell(ws, `${cols.dong}${r}`);
    const dong = dongRaw == null ? "" : String(dongRaw).trim();
    if (!dong) continue;
    if (dong.replace(/\s/g, "").startsWith("합계")) break; // 합계 행에서 종료
    if (dong.startsWith("▶")) break;                       // 다음 블록에서 종료

    const num = (key) => {
      if (!cols[key]) return null;
      const v = rawCell(ws, `${cols[key]}${r}`);
      const n = Number(v);
      return isFinite(n) && v !== "" && v !== null ? n : null;
    };

    // 누적실적 셀의 수식에서 참조 범위 끝 행을 뽑아둠 (보충 합산 기준)
    const mr = formulaMaxRefRow(cellFormula(ws, `${cols.cumActual}${r}`));
    if (mr && (coveredMaxRow === null || mr > coveredMaxRow)) coveredMaxRow = mr;

    byDong.push({
      dong,
      planArea: num("planArea"),
      totalUnits: num("totalUnits"),
      optionUnits: num("optionUnits"),
      cumActual: num("cumActual") ?? 0,
      rate: num("rate"),
      remain: num("remain"),
      endDate: excelDateToISO(rawCell(ws, cols.endDate ? `${cols.endDate}${r}` : `Z${r}`)),
      statusText: cols.status ? String(rawCell(ws, `${cols.status}${r}`) ?? "") : "",
    });
  }

  if (!byDong.length) return null;

  const { summary, summaryCovered } = readAchvSummary(ws, headerRow - 1);

  return { sheetName: ws.name, byDong, summary, summaryCovered, coveredMaxRow };
}

// ===== ⑥ 석재발주 시트(발주현황) 읽기 =====
// 시트 이름에 "발주"가 들어간 시트를 찾아, "동  석종  구분 소계" 형태의 말단(leaf) 행만
// {dong, stone, gubun, ea, m2, m} 목록으로 반환합니다. 동별 합계·석종 소계는 앱(calc)에서
// 이 말단 행들을 직접 재집계하므로, 사용자가 같은 형식으로 행을 추가하면 자동으로 반영됩니다.
const ORDER_SHEET_HINT = "발주";
const KNOWN_STONES = ["보니브라운", "스틸그레이", "마천석", "포천석", "블랑코머핀"];

function orderSheetCandidates(wb) {
  const named = wb.worksheets.filter((ws) => ws && ws.name && ws.name.includes(ORDER_SHEET_HINT));
  return named.length ? named : wb.worksheets;
}

// 라벨에서 {dong, stone, gubun} 추출. 말단 소계 행만 객체를 반환하고,
// 석종 소계(2토막)·동 합계·총계·제목 행 등은 null.
function parseOrderLabel(label, stoneSet) {
  const s = String(label == null ? "" : label).trim();
  if (!s.endsWith("소계")) return null;          // "소계"로 끝나는 행만 (합계/총계/제목 제외)
  const body = s.slice(0, -2).trim();            // 끝의 "소계" 제거
  // 1) 생성기 기본 포맷: 동·석종·구분 사이가 2칸 이상 공백
  const parts = body.split(/\s{2,}/).map((x) => x.trim()).filter(Boolean);
  if (parts.length >= 3) return { dong: parts[0], stone: parts[1], gubun: parts.slice(2).join(" ") };
  // 2) 폴백: 동·석종·구분이 2칸 미만 공백으로 붙어 2토막 이하로 쪼개진 경우
  //    (예: "근생2   스틸그레이 벽체" → ["근생2","스틸그레이 벽체"])에도 버리지 않고
  //    알려진 석종명으로 동/구분 경계를 추정. 석종 소계("스틸그레이 소계")는 석종 뒤가
  //    비어 gubun이 없으므로 여기서 자동 제외됨.
  const scanText = parts.length === 2 ? parts.join(" ") : body;
  for (const st of stoneSet) {
    const i = scanText.indexOf(st);
    if (i > 0) {
      const dong = scanText.slice(0, i).trim();
      const gubun = scanText.slice(i + st.length).trim();
      if (dong && gubun) return { dong, stone: st, gubun };
    }
  }
  return null;
}

function readOrderLeavesFromSheet(ws) {
  const rowsRaw = [];
  const maxRow = ws.rowCount || 0;
  for (let r = 1; r <= maxRow; r++) {
    rowsRaw.push({
      label: rawCell(ws, `A${r}`),
      ea: rawCell(ws, `B${r}`),
      m2: rawCell(ws, `C${r}`),
      m: rawCell(ws, `D${r}`),
    });
  }
  // 1차 패스: 표에 등장하는 석종 집합 수집(폴백 매칭용)
  const stoneSet = new Set(KNOWN_STONES);
  for (const row of rowsRaw) {
    const p = parseOrderLabel(row.label, []);
    if (p) stoneSet.add(p.stone);
  }
  // 2차 패스: 말단(동·석종·구분 소계) 행만 추출
  const out = [];
  for (const row of rowsRaw) {
    const p = parseOrderLabel(row.label, stoneSet);
    if (!p) continue;
    const ea = Number(row.ea);
    if (!isFinite(ea)) continue;
    out.push({ dong: p.dong, stone: p.stone, gubun: p.gubun, ea, m2: Number(row.m2) || 0, m: Number(row.m) || 0 });
  }
  return out;
}

// "발주"가 들어간 시트가 여러 개(⑤ 면별, ⑥ 시트별)일 수 있으므로,
// 말단(구분 소계) 행이 가장 많이 추출되는 시트를 발주현황 원천으로 채택합니다.
// → 시트 이름이 바뀌거나 순서가 달라져도 올바른 시트(⑥)를 자동으로 고릅니다.
export function readOrderStatus(wb) {
  let best = [];
  for (const ws of orderSheetCandidates(wb)) {
    const leaves = readOrderLeavesFromSheet(ws);
    if (leaves.length > best.length) best = leaves;
  }
  return best;
}

// ===== ⑦착수체크리스트 시트 =====
// 동별 "착수 선행공정" 게이트를 기록하는 시트. 앱이 없으면 자동으로 만듭니다.
// 매일 입력하는 게 아니라 동당 한 번만 손대는 표라서 일일 입력 부담이 늘지 않습니다.
// 지연사유(비고)에 적은 내용이 그대로 공기연장 근거 리스트에 실립니다.

export const SHEET_CHECKLIST = "⑦착수체크리스트";

// 선행공정 게이트 항목 (열 순서 = 이 순서)
export const GATE_ITEMS = [
  { key: "scaffold", label: "비계 설치" },
  { key: "window", label: "창호 취부" },
  { key: "frame", label: "골조·먹매김" },
  { key: "material", label: "자재 반입" },
  { key: "shopdwg", label: "샵도면 승인" },
];

const CHECKLIST_HEADERS = [
  "동(구역)",
  ...GATE_ITEMS.map((g) => g.label),
  "완료",
  "지연사유(선행공정)",
];

function truthy(v) {
  const s = String(v ?? "").trim().toUpperCase();
  return s === "Y" || s === "O" || s === "TRUE" || s === "1" || s === "V" || s === "✓";
}

// 헤더 라벨로 열을 찾아 매핑 (사용자가 열을 옮겨도 따라감)
function mapChecklistCols(ws) {
  const map = {};
  for (let c = 1; c <= 20; c++) {
    const raw = rawCell(ws, `${colLetter(c)}1`);
    if (typeof raw !== "string" || !raw.trim()) continue;
    const norm = raw.replace(/\s/g, "");
    if (map.dong === undefined && norm.includes("동")) map.dong = colLetter(c);
    for (const g of GATE_ITEMS) {
      const gl = g.label.replace(/[\s·]/g, "");
      if (map[g.key] === undefined && norm.replace(/·/g, "").includes(gl.slice(0, 2))) map[g.key] = colLetter(c);
    }
    if (map.done === undefined && norm === "완료") map.done = colLetter(c);
    if (map.reason === undefined && norm.includes("지연사유")) map.reason = colLetter(c);
  }
  return map;
}

export function readChecklist(wb) {
  const ws = wb.getWorksheet(SHEET_CHECKLIST);
  if (!ws) return [];
  const cols = mapChecklistCols(ws);
  if (!cols.dong) return [];
  const out = [];
  const maxRow = Math.min(ws.rowCount || 1, 300);
  for (let r = 2; r <= maxRow; r++) {
    const dong = rawCell(ws, `${cols.dong}${r}`);
    if (dong === null || String(dong).trim() === "") continue;
    const row = { dong: String(dong).trim(), gates: {} };
    for (const g of GATE_ITEMS) row.gates[g.key] = cols[g.key] ? truthy(rawCell(ws, `${cols[g.key]}${r}`)) : false;
    row.done = cols.done ? truthy(rawCell(ws, `${cols.done}${r}`)) : false;
    row.reason = cols.reason ? String(rawCell(ws, `${cols.reason}${r}`) ?? "") : "";
    out.push(row);
  }
  return out;
}

// 체크리스트 전체를 시트에 다시 씀 (시트가 없으면 생성).
// 값만 쓰므로 사용자가 시트에 서식을 입혀두면 그대로 유지됩니다.
export function writeChecklist(wb, rows) {
  let ws = wb.getWorksheet(SHEET_CHECKLIST);
  if (!ws) {
    ws = wb.addWorksheet(SHEET_CHECKLIST);
    ws.getColumn(1).width = 14;
    for (let i = 2; i <= CHECKLIST_HEADERS.length; i++) ws.getColumn(i).width = i === CHECKLIST_HEADERS.length ? 40 : 12;
  }
  CHECKLIST_HEADERS.forEach((h, i) => {
    const cell = ws.getCell(`${colLetter(i + 1)}1`);
    cell.value = h;
    cell.font = { bold: true };
  });
  // 기존 데이터 행 비우기 (동 목록이 줄어든 경우 잔여물 제거)
  const maxRow = Math.min(ws.rowCount || 1, 300);
  for (let r = 2; r <= maxRow; r++) {
    for (let c = 1; c <= CHECKLIST_HEADERS.length; c++) ws.getCell(`${colLetter(c)}${r}`).value = null;
  }
  rows.forEach((row, i) => {
    const r = i + 2;
    let c = 1;
    ws.getCell(`${colLetter(c++)}${r}`).value = row.dong;
    for (const g of GATE_ITEMS) ws.getCell(`${colLetter(c++)}${r}`).value = row.gates?.[g.key] ? "Y" : "";
    ws.getCell(`${colLetter(c)}${r}`).value = row.done ? "Y" : "";
    ws.getCell(`${colLetter(c + 1)}${r}`).value = row.reason || "";
  });
  return ws;
}

// 전체 흐름: 로그인 → 다운로드 → 파싱 → 앱 상태 반환 (워크북 객체도 함께 보관해야 재업로드 가능)
// site: { id, name, filePath, fileName } — 어떤 현장의 OneDrive 파일을 동기화할지 지정
export async function syncDown(site) {
  const token = await getToken();
  const itemId = await findFileId(token, site);
  const buf = await downloadWorkbookArrayBuffer(token, itemId);
  const wb = await parseWorkbook(buf);
  const buildings = readBuildings(wb);
  const logsExternal = readExternalLogs(wb);
  const logsInternal = readInternalLogs(wb);
  const orderRows = readOrderStatus(wb);
  const thresholds = readThresholds(wb);
  const achvExternal = readAchievement(wb, SHEET_ACHV_EXTERNAL);
  const achvInternal = readAchievement(wb, SHEET_ACHV_INTERNAL);
  const checklist = readChecklist(wb);
  return { wb, itemId, buildings, logsExternal, logsInternal, orderRows, thresholds, achvExternal, achvInternal, checklist, syncedAt: new Date().toISOString() };
}

export async function syncUp(wb, itemId) {
  const token = await getToken();
  const buf = await serializeWorkbook(wb);
  await uploadWorkbookArrayBuffer(token, itemId, buf);
  return new Date().toISOString();
}
