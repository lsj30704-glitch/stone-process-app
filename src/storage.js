// localStorage 기반 로컬 저장 + 동기화 대기열 관리

const KEYS = {
  buildingsExternal: "spm_buildings_external_v1",
  buildingsInternal: "spm_buildings_internal_v1",
  logsExternal: "spm_logs_external_v1",
  logsInternal: "spm_logs_internal_v1",
  fileMeta: "spm_file_meta_v1", // OneDrive 파일 id, 마지막 동기화 시각 등
  pendingExternal: "spm_pending_external_v1", // 아직 업로드 못한 외부 입력
  pendingInternal: "spm_pending_internal_v1",
  checklist: "spm_recovery_checklist_v1",
};

function get(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}
function set(key, val) {
  try {
    localStorage.setItem(key, JSON.stringify(val));
  } catch {
    /* ignore */
  }
}

export const Storage = {
  KEYS,
  get,
  set,
};
