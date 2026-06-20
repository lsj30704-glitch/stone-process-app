import { PublicClientApplication } from "@azure/msal-browser";
import * as XLSX from "xlsx";
import { Storage } from "./storage";

// ---- MSAL 설정 ----
// VITE_MSAL_CLIENT_ID는 Azure Portal에서 앱 등록 후 발급되는 "애플리케이션(클라이언트) ID"를
// .env(또는 Vercel 환경변수)에 넣어야 합니다. 등록 전에는 로그인 버튼이 비활성화됩니다.
const CLIENT_ID = import.meta.env.VITE_MSAL_CLIENT_ID || "";
const FILE_NAME = import.meta.env.VITE_GRAPH_FILE_NAME || "석공사_일일공정관리_lsj.xlsx";

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
  // 리다이렉트 로그인 사용 시 결과 처리 (팝업 사용 시에는 큰 영향 없음)
  await inst.handleRedirectPromise().catch(() => null);
  return inst;
}

export function getActiveAccount() {
  const inst = getMsal();
  const accounts = inst.getAllAccounts();
  return accounts && accounts.length ? accounts[0] : null;
}

export async function login() {
  const inst = getMsal();
  const result = await inst.loginPopup({ scopes: SCOPES });
  inst.setActiveAccount(result.account);
  return result.account;
}

export function logout() {
  const inst = getMsal();
  const account = getActiveAccount();
  if (account) inst.logoutPopup({ account });
}

export async function getToken() {
  const inst = getMsal();
  const account = getActiveAccount();
  if (!account) throw new Error("NOT_LOGGED_IN");
  try {
    const res = await inst.acquireTokenSilent({ scopes: SCOPES, account });
    return res.accessToken;
  } catch {
    const res = await inst.acquireTokenPopup({ scopes: SCOPES, account });
    return res.accessToken;
  }
}

// ---- Graph 파일 탐색/다운로드/업로드 ----
// 주의: Microsoft Graph의 Excel "workbook" REST API(셀 단위 읽기/쓰기)는
// 개인(Consumer) OneDrive 계정에서는 지원되지 않습니다.
// 그래서 파일 전체를 받아(content) 브라우저에서 SheetJS로 수정한 뒤
// 파일 전체를 다시 업로드(content)하는 방식으로 동작합니다.

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

export async function findFileId(token) {
  const cached = Storage.get(Storage.KEYS.fileMeta, null);
  if (cached?.itemId) return cached.itemId;
  const q = encodeURIComponent(FILE_NAME);
  const res = await graphFetch(`${GRAPH}/me/drive/root/search(q='${q}')`, token);
  const json = await res.json();
  const match = (json.value || []).find((it) => it.name === FILE_NAME) || (json.value || [])[0];
  if (!match) throw new Error(`OneDrive에서 "${FILE_NAME}" 파일을 찾을 수 없습니다.`);
  Storage.set(Storage.KEYS.fileMeta, { itemId: match.id, name: match.name, webUrl: match.webUrl });
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

export function parseWorkbook(arrayBuffer) {
  return XLSX.read(arrayBuffer, { type: "array", cellFormula: true, cellStyles: true, cellDates: true });
}

export function serializeWorkbook(wb) {
  return XLSX.write(wb, { type: "array", bookType: "xlsx", cellStyles: true });
}

// ---- 시트 ↔ 앱 데이터 변환 ----

const SHEET_EXTERNAL = "②일일실적입력";
const SHEET_INTERNAL = "②일일실적입력(내부)";
const SHEET_BASE = "①기준정보";

function excelDateToISO(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  return v || "";
}

export function readExternalLogs(wb) {
  const ws = wb.Sheets[SHEET_EXTERNAL];
  if (!ws) return [];
  const out = [];
  for (let r = 5; r <= 104; r++) {
    const bCell = ws[`B${r}`];
    if (!bCell || bCell.v === undefined || bCell.v === "") continue;
    out.push({
      id: `row-${r}`,
      row: r,
      date: excelDateToISO(bCell.v),
      dong: ws[`C${r}`]?.v ?? "",
      masonry: ws[`D${r}`]?.v ?? 0,
      caulking: ws[`E${r}`]?.v ?? 0,
      truss: ws[`F${r}`]?.v ?? 0,
      scaffold: ws[`G${r}`]?.v ?? 0,
      actual: ws[`J${r}`]?.v ?? "",
      disaster: ws[`L${r}`]?.v ?? "",
      reason: ws[`M${r}`]?.v ?? "",
      note: ws[`N${r}`]?.v ?? "",
      memo: ws[`O${r}`]?.v ?? "",
    });
  }
  return out;
}

export function readInternalLogs(wb) {
  const ws = wb.Sheets[SHEET_INTERNAL];
  if (!ws) return [];
  const out = [];
  for (let r = 5; r <= 104; r++) {
    const bCell = ws[`B${r}`];
    if (!bCell || bCell.v === undefined || bCell.v === "") continue;
    out.push({
      id: `row-${r}`,
      row: r,
      date: excelDateToISO(bCell.v),
      dong: ws[`C${r}`]?.v ?? "",
      masonry: ws[`D${r}`]?.v ?? 0,
      caulking: ws[`E${r}`]?.v ?? 0,
      truss: ws[`F${r}`]?.v ?? 0,
      actual: ws[`I${r}`]?.v ?? "",
      disaster: ws[`K${r}`]?.v ?? "",
      reason: ws[`L${r}`]?.v ?? "",
      note: ws[`M${r}`]?.v ?? "",
      memo: ws[`N${r}`]?.v ?? "",
    });
  }
  return out;
}

export function readBuildings(wb) {
  const ws = wb.Sheets[SHEET_BASE];
  if (!ws) return { external: [], internal: [] };
  const external = [];
  for (let r = 6; r <= 13; r++) {
    const dong = ws[`B${r}`]?.v;
    if (!dong) continue;
    external.push({
      dong,
      totalArea: ws[`C${r}`]?.v ?? 0,
      startDate: excelDateToISO(ws[`D${r}`]?.v),
      endDate: excelDateToISO(ws[`E${r}`]?.v),
      workDays: ws[`F${r}`]?.v ?? 0,
      dailyPlan: ws[`G${r}`]?.v ?? 0,
      baseWorkers: ws[`H${r}`]?.v ?? 0,
      hoist: ws[`I${r}`]?.v ?? 0,
      note: ws[`J${r}`]?.v ?? "",
    });
  }
  const internal = [];
  for (let r = 45; r <= 52; r++) {
    const dong = ws[`B${r}`]?.v;
    if (!dong) continue;
    const totalUnits = ws[`C${r}`]?.v ?? 0;
    const optionUnits = ws[`D${r}`]?.v ?? 0;
    internal.push({ dong, totalUnits, optionUnits, normalUnits: totalUnits - optionUnits });
  }
  return { external, internal };
}

function findNextEmptyRow(ws, startRow = 5, endRow = 104) {
  for (let r = startRow; r <= endRow; r++) {
    const bCell = ws[`B${r}`];
    if (!bCell || bCell.v === undefined || bCell.v === "") return r;
  }
  return null; // 시트가 가득 찼음
}

function extendRef(ws, row) {
  const ref = XLSX.utils.decode_range(ws["!ref"]);
  if (row > ref.e.r + 1) {
    ref.e.r = row - 1;
    ws["!ref"] = XLSX.utils.encode_range(ref);
  }
}

function copyStyle(ws, fromAddr, toAddr) {
  const from = ws[fromAddr];
  if (from && from.s) {
    if (!ws[toAddr]) ws[toAddr] = {};
    ws[toAddr].s = from.s;
  }
}

// 새 외부 실적 행을 워크북에 추가 (값 + 기존과 동일한 수식 패턴)
export function appendExternalRow(wb, entry) {
  const ws = wb.Sheets[SHEET_EXTERNAL];
  if (!ws) throw new Error(`시트를 찾을 수 없습니다: ${SHEET_EXTERNAL}`);
  const row = findNextEmptyRow(ws);
  if (!row) throw new Error("일일실적입력 시트의 입력 가능한 행(5~104)이 모두 채워졌습니다.");
  const templateRow = row > 5 ? row - 1 : 5;

  ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"].forEach((col) =>
    copyStyle(ws, `${col}${templateRow}`, `${col}${row}`)
  );

  ws[`A${row}`] = { t: "n", v: row - 4 };
  ws[`B${row}`] = { t: "d", v: new Date(entry.date + "T00:00:00"), z: ws[`B${templateRow}`]?.z || "yyyy-mm-dd" };
  ws[`C${row}`] = { t: "s", v: entry.dong };
  ws[`D${row}`] = { t: "n", v: Number(entry.masonry) || 0 };
  ws[`E${row}`] = { t: "n", v: Number(entry.caulking) || 0 };
  ws[`F${row}`] = { t: "n", v: Number(entry.truss) || 0 };
  ws[`G${row}`] = { t: "n", v: Number(entry.scaffold) || 0 };
  ws[`H${row}`] = { t: "n", f: `SUM(D${row}:G${row})` };
  ws[`I${row}`] = { t: "n", f: `IFERROR(IF(C${row}="","",VLOOKUP(C${row},①기준정보!$B$6:$G$13,6,0)),"")` };
  ws[`J${row}`] = { t: "n", v: entry.actual === "" || entry.actual === null ? 0 : Number(entry.actual) };
  ws[`K${row}`] = { t: "n", f: `IFERROR(IF(OR(J${row}="",I${row}="",I${row}=0),"",J${row}/I${row}),"")` };
  ws[`L${row}`] = { t: "s", v: entry.disaster || "" };
  ws[`M${row}`] = { t: "s", v: entry.reason || "" };
  ws[`N${row}`] = { t: "s", v: entry.note || "" };
  ws[`O${row}`] = { t: "s", v: entry.memo || "" };

  extendRef(ws, row);
  return row;
}

// 새 내부(세대) 실적 행을 워크북에 추가
export function appendInternalRow(wb, entry) {
  const ws = wb.Sheets[SHEET_INTERNAL];
  if (!ws) throw new Error(`시트를 찾을 수 없습니다: ${SHEET_INTERNAL}`);
  const row = findNextEmptyRow(ws);
  if (!row) throw new Error("일일실적입력(내부) 시트의 입력 가능한 행(5~104)이 모두 채워졌습니다.");
  const templateRow = row > 5 ? row - 1 : 5;

  ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"].forEach((col) =>
    copyStyle(ws, `${col}${templateRow}`, `${col}${row}`)
  );

  ws[`A${row}`] = { t: "n", v: row - 4 };
  ws[`B${row}`] = { t: "d", v: new Date(entry.date + "T00:00:00"), z: ws[`B${templateRow}`]?.z || "yyyy-mm-dd" };
  ws[`C${row}`] = { t: "s", v: entry.dong };
  ws[`D${row}`] = { t: "n", v: Number(entry.masonry) || 0 };
  ws[`E${row}`] = { t: "n", v: Number(entry.caulking) || 0 };
  ws[`F${row}`] = { t: "n", v: Number(entry.truss) || 0 };
  ws[`G${row}`] = { t: "n", f: `SUM(D${row}:F${row})` };
  ws[`H${row}`] = { t: "n", f: `G${row}*3` };
  ws[`I${row}`] = { t: "n", v: entry.actual === "" || entry.actual === null ? 0 : Number(entry.actual) };
  ws[`J${row}`] = { t: "n", f: `IFERROR(IF(OR(I${row}="",H${row}="",H${row}=0),"",I${row}/H${row}),"")` };
  ws[`K${row}`] = { t: "s", v: entry.disaster || "" };
  ws[`L${row}`] = { t: "s", v: entry.reason || "" };
  ws[`M${row}`] = { t: "s", v: entry.note || "" };
  ws[`N${row}`] = { t: "s", v: entry.memo || "" };

  extendRef(ws, row);
  return row;
}

// 전체 흐름: 로그인 → 다운로드 → 파싱 → 앱 상태 반환 (워크북 객체도 함께 보관해야 재업로드 가능)
export async function syncDown() {
  const token = await getToken();
  const itemId = await findFileId(token);
  const buf = await downloadWorkbookArrayBuffer(token, itemId);
  const wb = parseWorkbook(buf);
  const buildings = readBuildings(wb);
  const logsExternal = readExternalLogs(wb);
  const logsInternal = readInternalLogs(wb);
  return { wb, itemId, buildings, logsExternal, logsInternal, syncedAt: new Date().toISOString() };
}

export async function syncUp(wb, itemId) {
  const token = await getToken();
  const buf = serializeWorkbook(wb);
  await uploadWorkbookArrayBuffer(token, itemId, buf);
  return new Date().toISOString();
}
