// localStorage 기반 로컬 저장 + 동기화 대기열 관리 (현장별로 데이터 분리)

const SITES_KEY = "spm_sites_v1";
const ACTIVE_SITE_KEY = "spm_active_site_v1";

// 현장(site)별로 분리 저장되는 데이터 종류. 실제 key는 siteKey()로 현장 id를 붙여서 사용합니다.
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
function remove(key) {
  try {
    localStorage.removeItem(key);
  } catch {
    /* ignore */
  }
}

// 현장(siteId)별로 분리된 storage key를 만듭니다.
function siteKey(base, siteId) {
  return siteId ? `${base}__${siteId}` : base;
}

function defaultFilePath() {
  return import.meta.env.VITE_GRAPH_FILE_PATH || "2. 근무지_현장정리/3. 은진산업/석공사_일일공정관리_lsj.xlsx";
}

function fileNameFromPath(path) {
  const parts = String(path || "").split("/").filter(Boolean);
  return parts[parts.length - 1] || path;
}

// 현장 구분이 없던 예전 버전의 데이터(키에 __siteId 접미사가 없던 시절)를
// "기본 현장" 한 곳으로 1회만 이전합니다. 이미 현장 목록이 있으면 아무 것도 하지 않습니다.
function migrateLegacyToFirstSite() {
  const existingSites = get(SITES_KEY, null);
  if (existingSites && existingSites.length) return existingSites;

  const filePath = defaultFilePath();
  const site = {
    id: `site-${Date.now()}`,
    name: "기본 현장",
    filePath,
    fileName: fileNameFromPath(filePath),
    isDefault: true, // 마이그레이션으로 생성된 기본 현장임을 표시 (시드 데이터 fallback 판단용)
    createdAt: new Date().toISOString(),
  };

  // 과거(현장 구분 없이 쓰던) 키들의 값을 그대로 기본 현장의 키로 복사
  Object.values(KEYS).forEach((base) => {
    const legacy = get(base, undefined);
    if (legacy !== undefined) set(siteKey(base, site.id), legacy);
  });

  set(SITES_KEY, [site]);
  set(ACTIVE_SITE_KEY, site.id);
  return [site];
}

function getSites() {
  return migrateLegacyToFirstSite();
}

function setSites(sites) {
  set(SITES_KEY, sites);
}

function getActiveSiteId() {
  const sites = getSites();
  let activeId = get(ACTIVE_SITE_KEY, null);
  if (!activeId || !sites.find((s) => s.id === activeId)) {
    activeId = sites[0]?.id || null;
    set(ACTIVE_SITE_KEY, activeId);
  }
  return activeId;
}

function setActiveSiteId(id) {
  set(ACTIVE_SITE_KEY, id);
}

function addSite({ name, filePath }) {
  const sites = getSites();
  const site = {
    id: `site-${Date.now()}`,
    name,
    filePath,
    fileName: fileNameFromPath(filePath),
    isDefault: false,
    createdAt: new Date().toISOString(),
  };
  setSites([...sites, site]);
  return site;
}

function updateSite(id, { name, filePath }) {
  const sites = getSites();
  const next = sites.map((s) => (s.id === id ? { ...s, name, filePath, fileName: fileNameFromPath(filePath) } : s));
  setSites(next);
  // 경로가 바뀌면 캐시된 itemId가 더 이상 유효하지 않을 수 있으니 비워서 다음 동기화 때 다시 찾게 함
  remove(siteKey(KEYS.fileMeta, id));
}

function removeSite(id) {
  const sites = getSites();
  if (sites.length <= 1) return sites; // 최소 1개 현장은 항상 유지
  const next = sites.filter((s) => s.id !== id);
  setSites(next);
  Object.values(KEYS).forEach((base) => remove(siteKey(base, id)));
  if (get(ACTIVE_SITE_KEY, null) === id) setActiveSiteId(next[0]?.id || null);
  return next;
}

export const Storage = {
  KEYS,
  get,
  set,
  remove,
  siteKey,
  getSites,
  setSites,
  getActiveSiteId,
  setActiveSiteId,
  addSite,
  updateSite,
  removeSite,
};
