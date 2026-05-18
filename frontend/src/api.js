// 휘발성 — 서버에 영속 저장 없음. localStorage 미사용 (Source 9, 10).
//
// API base 결정 규칙:
//   - VITE_API_BASE 가 빈 값이면 same-origin (Vite dev proxy 또는 동일 도메인 deploy)
//   - 값이 있으면 절대경로로 호출 (예: https://blue-nine-api.onrender.com)
const API_BASE = (import.meta.env.VITE_API_BASE || '').replace(/\/$/, '');

function url(path) {
  if (path.startsWith('http')) return path;
  return API_BASE + path;
}

let sessionId = null;
let mode = 'precise';

export function setMode(m) { mode = m; }
export function getMode() { return mode; }
export function getSessionId() { return sessionId; }
export function getApiBase() { return API_BASE; }

async function call(path, options = {}) {
  const opts = {
    ...options,
    headers: {
      'X-Blue-Nine-Mode': mode,
      ...(options.headers || {}),
    },
  };
  const r = await fetch(url(path), opts);
  if (!r.ok) {
    const t = await r.text().catch(() => '');
    throw new Error(`${r.status}: ${t || r.statusText}`);
  }
  if (r.headers.get('content-type')?.includes('application/json')) {
    return r.json();
  }
  return r;
}

export async function ensureSession() {
  if (sessionId) return sessionId;
  const r = await call('/api/session/new', { method: 'POST' });
  sessionId = r.session_id;
  return sessionId;
}

export async function destroySession() {
  if (!sessionId) return;
  const fd = new FormData();
  fd.append('session_id', sessionId);
  // 페이지 unload 중에는 일반 fetch 가 취소될 수 있어 sendBeacon 우선 시도.
  try {
    if (navigator.sendBeacon) {
      navigator.sendBeacon(url('/api/session/destroy'), fd);
    } else {
      await fetch(url('/api/session/destroy'), { method: 'POST', body: fd, keepalive: true });
    }
  } catch {}
  sessionId = null;
}

export async function fetchCategories() {
  return call('/api/categories');
}

export async function fetchMasterInfo() {
  return call('/api/master/info');
}

export async function parseEstimate({ l1, l2, file, billing, client, campaign, versionLabel }) {
  const sid = await ensureSession();
  const fd = new FormData();
  fd.append('session_id', sid);
  fd.append('category_l1', l1);
  fd.append('category_l2', l2);
  fd.append('mode', mode);
  if (client) fd.append('client', client);
  if (campaign) fd.append('campaign', campaign);
  if (versionLabel) fd.append('version_label', versionLabel);
  fd.append('file', file);
  if (billing) fd.append('billing', billing);
  return call('/api/estimate/parse', { method: 'POST', body: fd });
}

export async function updateRow({ estimateId, rowId, patch }) {
  const sid = await ensureSession();
  return call('/api/estimate/update_row', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ session_id: sid, estimate_id: estimateId, row_id: rowId, patch }),
  });
}

export async function exportXlsx(estimateId, filename) {
  const sid = await ensureSession();
  const r = await fetch(url(`/api/estimate/${sid}/${estimateId}/export.xlsx`), {
    method: 'POST',
    headers: { 'X-Blue-Nine-Mode': mode },
  });
  if (!r.ok) {
    const t = await r.text().catch(() => '');
    throw new Error(`Export 실패: ${r.status} ${t}`);
  }
  const blob = await r.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename || `BLUE_NINE_${estimateId}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

export async function monitorSummary() { return call('/api/monitor/summary'); }
export async function monitorLogs(n = 50) { return call(`/api/monitor/logs?limit=${n}`); }
