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
  for (let r = 5; r <= 104; r++) {
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
  for (let r = 5; r <= 104; r++) {
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

// ===== 공사 범위(scope) 일반화 — 외부/호이스트/부대시설처럼 면적(m²)형 시트 공용 =====
// 면적형 일일실적 시트(외부와 동일한 컬럼 구조)를 시트명만 받아 일반적으로 읽음.
export function readAreaLogs(wb, sheetName) {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) return [];
  const out = [];
  for (let r = 5; r <= 104; r++) {
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
  if (!row) throw new Error(`${sheetName} 시트의 입력 가능한 행(5~104)이 모두 채워졌습니다.`);
  const templateRow = row > 5 ? row - 1 : 5;

  ws.getCell(`A${row}`).value = row - 4;
  setDateCell(ws, `B${row}`, entry.date, templateRow);
  ws.getCell(`C${row}`).value = entry.dong;
  ws.getCell(`D${row}`).value = Number(entry.masonry) || 0;
  ws.getCell(`E${row}`).value = Number(entry.caulking) || 0;
  ws.getCell(`F${row}`).value = Number(entry.truss) || 0;
  ws.getCell(`G${row}`).value = Number(entry.scaffold) || 0;
  ws.getCell(`H${row}`).value = { formula: `SUM(D${row}:G${row})` };
  ws.getCell(`I${row}`).value = { formula: `IFERROR(IF(C${row}="","",VLOOKUP(C${row},${planRange},6,0)),"")` };
  ws.getCell(`J${row}`).value = entry.actual === "" || entry.actual === null ? 0 : Number(entry.actual);
  ws.getCell(`K${row}`).value = { formula: `IFERROR(IF(OR(J${row}="",I${row}="",I${row}=0),"",J${row}/I${row}),"")` };
  ws.getCell(`L${row}`).value = entry.disaster || "";
  ws.getCell(`M${row}`).value = entry.reason || "";
  ws.getCell(`N${row}`).value = entry.note || "";
  ws.getCell(`O${row}`).value = entry.memo || "";
  return row;
}

function findNextEmptyRow(ws, startRow = 5, endRow = 104) {
  for (let r = startRow; r <= endRow; r++) {
    const b = rawCell(ws, `B${r}`);
    if (b === null || b === "") return r;
  }
  return null; // 시트가 가득 찼음
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
  if (!row) throw new Error("일일실적입력 시트의 입력 가능한 행(5~104)이 모두 채워졌습니다.");
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
  if (!row) throw new Error("일일실적입력(내부) 시트의 입력 가능한 행(5~104)이 모두 채워졌습니다.");
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
  return { wb, itemId, buildings, logsExternal, logsInternal, syncedAt: new Date().toISOString() };
}

export async function syncUp(wb, itemId) {
  const token = await getToken();
  const buf = await serializeWorkbook(wb);
  await uploadWorkbookArrayBuffer(token, itemId, buf);
  return new Date().toISOString();
}
