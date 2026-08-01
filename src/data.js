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
  productivityPerWorker: 10,
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

// ⑥ 석재발주_시트별 시트의 말단(동·석종·구분) 발주 수량 — 동기화 전 기본현장에서 보여줄 시드.
// 동기화하면 실제 엑셀 ⑥ 시트에서 읽은 값으로 대체됩니다.
export const SEED_ORDER_ROWS = [
  { dong: "101동(저층부)", stone: "보니브라운", gubun: "벽체", ea: 1703, m2: 830.2, m: 1515.98 },
  { dong: "101동(저층부)", stone: "보니브라운", gubun: "창틀·창대", ea: 607, m2: 122.65, m: 568.03 },
  { dong: "101동(저층부)", stone: "스틸그레이", gubun: "벽체", ea: 494, m2: 240.33, m: 453.55 },
  { dong: "101동(저층부)", stone: "스틸그레이", gubun: "창틀·창대", ea: 42, m2: 14.08, m: 49.14 },
  { dong: "101동(저층부)", stone: "스틸그레이", gubun: "두겁", ea: 159, m2: 21.04, m: 146.09 },
  { dong: "112동(저층부)", stone: "보니브라운", gubun: "벽체", ea: 1530, m2: 708.93, m: 1292.38 },
  { dong: "112동(저층부)", stone: "보니브라운", gubun: "창틀·창대", ea: 438, m2: 90.23, m: 423.57 },
  { dong: "112동(저층부)", stone: "스틸그레이", gubun: "벽체", ea: 578, m2: 268.26, m: 497.28 },
  { dong: "112동(저층부)", stone: "스틸그레이", gubun: "창틀·창대", ea: 26, m2: 8.58, m: 29.93 },
  { dong: "112동(저층부)", stone: "스틸그레이", gubun: "두겁", ea: 116, m2: 15.77, m: 109.49 },
  { dong: "102동", stone: "보니브라운", gubun: "벽체", ea: 2143, m2: 950.48, m: 1835.97 },
  { dong: "102동", stone: "보니브라운", gubun: "창주위", ea: 480, m2: 82.1, m: 449.31 },
  { dong: "102동", stone: "보니브라운", gubun: "창대", ea: 160, m2: 35.33, m: 147.35 },
  { dong: "102동", stone: "보니브라운", gubun: "P석", ea: 58, m2: 13.6, m: 45.41 },
  { dong: "102동", stone: "스틸그레이", gubun: "벽체", ea: 515, m2: 262.4, m: 497.12 },
  { dong: "102동", stone: "스틸그레이", gubun: "두겁", ea: 172, m2: 19.73, m: 159.99 },
  { dong: "102동", stone: "스틸그레이", gubun: "창주위", ea: 39, m2: 7.13, m: 45.92 },
  { dong: "102동", stone: "스틸그레이", gubun: "창대", ea: 10, m2: 2.36, m: 11.58 },
  { dong: "게스트하우스", stone: "마천석", gubun: "연마", ea: 276, m2: 84.19, m: 166.81 },
  { dong: "게스트하우스", stone: "보니브라운", gubun: "벽체", ea: 169, m2: 80.6, m: 158.35 },
  { dong: "게스트하우스", stone: "보니브라운", gubun: "두겁", ea: 20, m2: 9.5, m: 20.95 },
  { dong: "게스트하우스", stone: "포천석", gubun: "버너", ea: 92, m2: 30.3, m: 54.11 },
  { dong: "113동", stone: "보니브라운", gubun: "벽체", ea: 2701, m2: 1249.71, m: 2375.09 },
  { dong: "113동", stone: "보니브라운", gubun: "창주위", ea: 589, m2: 108.16, m: 573.43 },
  { dong: "113동", stone: "보니브라운", gubun: "창대", ea: 161, m2: 39.88, m: 166.67 },
  { dong: "113동", stone: "보니브라운", gubun: "P석", ea: 32, m2: 8.79, m: 29.34 },
  { dong: "113동", stone: "보니브라운", gubun: "pw", ea: 28, m2: 12.18, m: 24.83 },
  { dong: "113동", stone: "스틸그레이", gubun: "벽체", ea: 737, m2: 399.42, m: 727.03 },
  { dong: "113동", stone: "스틸그레이", gubun: "두겁", ea: 188, m2: 26.19, m: 181.26 },
  { dong: "113동", stone: "스틸그레이", gubun: "창주위", ea: 47, m2: 8.37, m: 56.53 },
  { dong: "113동", stone: "스틸그레이", gubun: "창대", ea: 15, m2: 3.56, m: 17.45 },
  { dong: "103동", stone: "보니브라운", gubun: "벽체", ea: 2121, m2: 990.56, m: 1935.1 },
  { dong: "103동", stone: "보니브라운", gubun: "창주위", ea: 564, m2: 99.86, m: 544.6 },
  { dong: "103동", stone: "보니브라운", gubun: "창대", ea: 193, m2: 43.48, m: 181.29 },
  { dong: "103동", stone: "보니브라운", gubun: "P석", ea: 10, m2: 2.04, m: 6.82 },
  { dong: "103동", stone: "보니브라운", gubun: "pw", ea: 6, m2: 3.24, m: 6.65 },
  { dong: "103동", stone: "스틸그레이", gubun: "벽체", ea: 612, m2: 323.59, m: 600.74 },
  { dong: "103동", stone: "스틸그레이", gubun: "두겁", ea: 182, m2: 24.9, m: 171.34 },
  { dong: "103동", stone: "스틸그레이", gubun: "창주위", ea: 37, m2: 6.59, m: 44.5 },
  { dong: "103동", stone: "스틸그레이", gubun: "창대", ea: 10, m2: 2.39, m: 11.72 },
  { dong: "104동(추가)", stone: "보니브라운", gubun: "벽체", ea: 90, m2: 28.0, m: 76.71 },
  { dong: "104동(1차)", stone: "보니브라운", gubun: "벽체", ea: 3390, m2: 1695.8, m: 3276.95 },
  { dong: "104동(1차)", stone: "보니브라운", gubun: "창주위", ea: 587, m2: 104.43, m: 576.37 },
  { dong: "104동(1차)", stone: "보니브라운", gubun: "창대", ea: 211, m2: 50.74, m: 211.75 },
  { dong: "104동(1차)", stone: "보니브라운", gubun: "P석", ea: 110, m2: 30.44, m: 101.64 },
  { dong: "104동(1차)", stone: "보니브라운", gubun: "pw", ea: 80, m2: 36.53, m: 74.93 },
  { dong: "104동(1차)", stone: "스틸그레이", gubun: "벽체", ea: 731, m2: 376.47, m: 690.32 },
  { dong: "104동(1차)", stone: "스틸그레이", gubun: "두겁", ea: 184, m2: 25.55, m: 176.87 },
  { dong: "104동(1차)", stone: "스틸그레이", gubun: "창주위", ea: 46, m2: 8.14, m: 55.0 },
  { dong: "104동(1차)", stone: "스틸그레이", gubun: "창대", ea: 15, m2: 3.63, m: 17.8 },
  { dong: "101동 호이스트", stone: "보니브라운", gubun: "벽체", ea: 147, m2: 76.61, m: 143.87 },
  { dong: "101동 호이스트", stone: "보니브라운", gubun: "창주위", ea: 63, m2: 11.02, m: 60.21 },
  { dong: "101동 호이스트", stone: "보니브라운", gubun: "창대", ea: 24, m2: 5.96, m: 24.83 },
  { dong: "101동 호이스트", stone: "스틸그레이", gubun: "두겁", ea: 13, m2: 1.86, m: 12.87 },
  { dong: "112동 호이스트", stone: "보니브라운", gubun: "벽체", ea: 36, m2: 19.85, m: 37.3 },
  { dong: "112동 호이스트", stone: "보니브라운", gubun: "창주위", ea: 21, m2: 4.11, m: 22.47 },
  { dong: "112동 호이스트", stone: "보니브라운", gubun: "창대", ea: 9, m2: 2.34, m: 9.74 },
  { dong: "112동 호이스트", stone: "스틸그레이", gubun: "두겁", ea: 4, m2: 0.61, m: 4.24 },
  { dong: "110동", stone: "보니브라운", gubun: "벽체", ea: 2900, m2: 1313.04, m: 2584.19 },
  { dong: "110동", stone: "보니브라운", gubun: "창주위", ea: 419, m2: 75.43, m: 409.07 },
  { dong: "110동", stone: "보니브라운", gubun: "창대", ea: 147, m2: 36.14, m: 149.66 },
  { dong: "110동", stone: "보니브라운", gubun: "P석", ea: 98, m2: 22.25, m: 90.35 },
  { dong: "110동", stone: "스틸그레이", gubun: "벽체", ea: 567, m2: 289.49, m: 555.48 },
  { dong: "110동", stone: "스틸그레이", gubun: "두겁", ea: 182, m2: 25.56, m: 177.48 },
  { dong: "110동", stone: "스틸그레이", gubun: "창주위", ea: 37, m2: 6.44, m: 43.54 },
  { dong: "110동", stone: "스틸그레이", gubun: "창대", ea: 14, m2: 3.31, m: 16.21 },
  { dong: "111동", stone: "보니브라운", gubun: "벽체", ea: 1433, m2: 647.41, m: 1316.17 },
  { dong: "111동", stone: "보니브라운", gubun: "창주위", ea: 335, m2: 59.32, m: 326.13 },
  { dong: "111동", stone: "보니브라운", gubun: "창대", ea: 109, m2: 25.04, m: 104.9 },
  { dong: "111동", stone: "보니브라운", gubun: "P석", ea: 8, m2: 2.05, m: 6.39 },
  { dong: "111동", stone: "스틸그레이", gubun: "벽체", ea: 476, m2: 215.96, m: 408.76 },
  { dong: "111동", stone: "스틸그레이", gubun: "두겁", ea: 120, m2: 16.05, m: 111.45 },
  { dong: "111동", stone: "스틸그레이", gubun: "창주위", ea: 9, m2: 1.6, m: 10.82 },
  { dong: "111동", stone: "스틸그레이", gubun: "창대", ea: 3, m2: 0.71, m: 3.47 },
  // ⑥ 근생2 발주 시드 (동기화 전 표시용 — 실제 엑셀 ⑥석재발주_시트별 값. 동기화하면 최신값으로 대체됨)
  { dong: "근생2", stone: "스틸그레이", gubun: "벽체", ea: 160, m2: 64.24, m: 5.2 },
  { dong: "근생2", stone: "스틸그레이", gubun: "두겁석", ea: 88, m2: 51.33, m: 6.93 },
  { dong: "근생2", stone: "스틸그레이", gubun: "창주위", ea: 6, m2: 0.88, m: 0.05 },
  { dong: "근생2", stone: "스틸그레이", gubun: "창대", ea: 65, m2: 27.84, m: 2.25 },
  { dong: "근생2", stone: "포천석", gubun: "벽체", ea: 358, m2: 159.75, m: 12.94 },
  { dong: "근생2", stone: "포천석", gubun: "두겁석", ea: 53, m2: 30.02, m: 4.05 },
  { dong: "근생2", stone: "포천석", gubun: "창주위", ea: 20, m2: 4.32, m: 0.23 },
  { dong: "근생2", stone: "포천석", gubun: "창대", ea: 4, m2: 1.08, m: 0.09 },];
