// 석공사_일일공정관리_lsj.xlsx 의 ①기준정보 시트를 그대로 옮긴 시드 데이터.
// OneDrive 연동(동기화) 후에는 실제 파일에서 읽은 값으로 대체됩니다.

export const DONG_LIST = [
  "101동", "102동", "103동", "104동", "110동", "111동", "112동", "113동",
];

// 외부(아파트 외벽) 기준정보 — ①기준정보 B6:J13
export const BUILDINGS_EXTERNAL = [
  { dong: "101동", totalArea: 1053.8, startDate: "2026-05-27", endDate: "2026-07-10", workDays: 33, dailyPlan: 31.93, baseWorkers: 6, hoist: 184.2, note: "실적입력 기준(이미 진행중) - 완료예정 동일" },
  { dong: "102동", totalArea: 1267.4, startDate: "2026-06-17", endDate: "2026-07-24", workDays: 17, dailyPlan: 74.55, baseWorkers: 13, hoist: 137.6, note: "실적입력 기준 시작일로 정정 / 완료예정 7/12->7/24 갱신" },
  { dong: "103동", totalArea: 1388.1, startDate: "2026-07-17", endDate: "2026-08-18", workDays: 16, dailyPlan: 86.76, baseWorkers: 13, hoist: 115.9, note: "비계설치 착공일 기준 / 6/26->7/17, 7/16->8/18 갱신" },
  { dong: "104동", totalArea: 2123.4, startDate: "2026-07-11", endDate: "2026-08-29", workDays: 17, dailyPlan: 124.91, baseWorkers: 13, hoist: 236.6, note: "비계설치 착공일 기준 / 7/3->7/11, 8/1->8/29 갱신" },
  { dong: "110동", totalArea: 1331.18, startDate: "2026-07-25", endDate: "2026-08-29", workDays: 17, dailyPlan: 78.30, baseWorkers: 13, hoist: 197.1, note: "비계설치 착공일 기준 / 7/17->7/25, 8/3->8/29 갱신" },
  { dong: "111동", totalArea: 974, startDate: "2026-08-05", endDate: "2026-08-31", workDays: 10, dailyPlan: 97.4, baseWorkers: 13, hoist: 0, note: "비계설치 착공일 기준 / 7/23->8/5, 8/1->8/31 갱신" },
  { dong: "112동", totalArea: 974.2, startDate: "2026-06-12", endDate: "2026-07-16", workDays: 12, dailyPlan: 81.18, baseWorkers: 13, hoist: 26.8, note: "실적입력 기준 시작일로 정정 / 완료예정 7/14->7/16 갱신" },
  { dong: "113동", totalArea: 1695.36, startDate: "2026-06-23", endDate: "2026-08-04", workDays: 17, dailyPlan: 99.73, baseWorkers: 13, hoist: 157.64, note: "비계설치 착공일 기준 / 6/26->6/23, 7/16->8/4 갱신" },
];

// 내부(세대) 기준정보 — ①기준정보 B45:F52
export const BUILDINGS_INTERNAL = [
  { dong: "101동", totalUnits: 131, optionUnits: 36 },
  { dong: "102동", totalUnits: 88, optionUnits: 19 },
  { dong: "103동", totalUnits: 90, optionUnits: 32 },
  { dong: "104동", totalUnits: 115, optionUnits: 51 },
  { dong: "110동", totalUnits: 115, optionUnits: 39 },
  { dong: "111동", totalUnits: 41, optionUnits: 10 },
  { dong: "112동", totalUnits: 44, optionUnits: 13 },
  { dong: "113동", totalUnits: 131, optionUnits: 52 },
].map((b) => ({ ...b, normalUnits: b.totalUnits - b.optionUnits }));

// ①기준정보 C18:D23 — 달성률 경보/만회계획 설정값
export const THRESHOLDS = {
  normal: 0.95,
  caution: 0.8,
  danger: 0.7,
  disasterExemption: 1.0,
  recoveryDays: 5,
  productivityPerWorker: 8.5,
};

export const DISASTER_OPTIONS = ["N (정상)", "Y (천재지변)", "P (부분영향)"];
export const REASON_CODES = ["TY-태풍", "HR-호우", "SN-폭설", "HW-한파", "EQ-지진", "FF-화재", "OT-기타"];

export const SEED_LOGS_EXTERNAL = [
  { id: "seed-1", date: "2026-05-27", dong: "101동", masonry: 2, caulking: 0, truss: 0, scaffold: 0, actual: 0, disaster: "Y (천재지변)", reason: "HR-호우", note: "자재반입", memo: "우천" },
  { id: "seed-2", date: "2026-05-28", dong: "101동", masonry: 6, caulking: 0, truss: 0, scaffold: 0, actual: 0, disaster: "P (부분영향)", reason: "HR-호우", note: "", memo: "앵글밑작업후 퇴근" },
  { id: "seed-3", date: "2026-05-29", dong: "101동", masonry: 0, caulking: 0, truss: 0, scaffold: 0, actual: 0, disaster: "N (정상)", reason: "OT-기타", note: "안전점검 오전 9시 30분퇴근", memo: "안전점검으로 퇴근" },
  { id: "seed-4", date: "2026-05-30", dong: "101동", masonry: 6, caulking: 0, truss: 0, scaffold: 0, actual: 24.07, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-5", date: "2026-06-01", dong: "101동", masonry: 4, caulking: 0, truss: 0, scaffold: 0, actual: 46.11, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-6", date: "2026-06-02", dong: "101동", masonry: 5, caulking: 0, truss: 0, scaffold: 0, actual: 49.22, disaster: "N (정상)", reason: "", note: "석재 색상 잇슈", memo: "1명 안전서류 미제출 퇴근" },
  { id: "seed-7", date: "2026-06-03", dong: "101동", masonry: 6, caulking: 0, truss: 0, scaffold: 0, actual: 48.81, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-8", date: "2026-06-04", dong: "101동", masonry: 6, caulking: 0, truss: 0, scaffold: 0, actual: 80.43, disaster: "P (부분영향)", reason: "HR-호우", note: "15시 우천 철수", memo: "" },
  { id: "seed-9", date: "2026-06-05", dong: "101동", masonry: 6, caulking: 0, truss: 0, scaffold: 0, actual: 46.31, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-10", date: "2026-06-08", dong: "101동", masonry: 7, caulking: 0, truss: 0, scaffold: 0, actual: 60.95, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-11", date: "2026-06-11", dong: "101동", masonry: 3, caulking: 0, truss: 0, scaffold: 0, actual: 35.66, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-12", date: "2026-06-12", dong: "112동", masonry: 4, caulking: 0, truss: 0, scaffold: 0, actual: 85.59, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-13", date: "2026-06-13", dong: "112동", masonry: 3, caulking: 0, truss: 0, scaffold: 0, actual: 43.63, disaster: "N (정상)", reason: "", note: "", memo: "" },
  { id: "seed-14", date: "2026-06-17", dong: "102동", masonry: 1, caulking: 0, truss: 0, scaffold: 0, actual: 0, disaster: "N (정상)", reason: "", note: "자재반입", memo: "" },
];

export const SEED_LOGS_INTERNAL = [
  { id: "seed-i1", date: "2026-06-12", dong: "103동", masonry: 2, caulking: 0, truss: 0, actual: 3, disaster: "N (정상)", reason: "OT-기타", note: "목업", memo: "" },
  { id: "seed-i2", date: "2026-06-13", dong: "110동", masonry: 2, caulking: 0, truss: 0, actual: 3, disaster: "N (정상)", reason: "OT-기타", note: "목업", memo: "" },
  { id: "seed-i3", date: "2026-06-18", dong: "게스트하우스", masonry: 3, caulking: 0, truss: 0, actual: null, disaster: "", reason: "", note: "게스트하우스 습식", memo: "복도,방풍실 1,2" },
  { id: "seed-i4", date: "2026-06-19", dong: "게스트하우스", masonry: 3, caulking: 0, truss: 0, actual: null, disaster: "", reason: "", note: "게스트하우스 습식", memo: "복도,방풍실 1,2" },
];

export const DONG_LIST_INTERNAL = [...DONG_LIST, "게스트하우스"];
