// 요구사항 4 — Latency / Traffic / Token-Cost 모니터링 패널.
// 2026-05-19: 폴링 60초로 늘리고 🔄 수동 새로고침 버튼 추가.
import React, { useEffect, useState, useCallback } from 'react';
import { monitorSummary, monitorLogs } from '../api.js';

const POLL_INTERVAL_MS = 60_000;        // 60초

export default function MonitorPanel() {
  const [s, setS] = useState(null);
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(false);
  const [lastTick, setLastTick] = useState(null);

  const refresh = useCallback(async () => {
    setLoading(true);
    try {
      const [su, lo] = await Promise.all([monitorSummary(), monitorLogs(8)]);
      setS(su);
      setLogs(lo);
      setLastTick(new Date());
    } catch {
      // 무시 — 다음 tick 시 자동 복구
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    refresh();
    const id = setInterval(refresh, POLL_INTERVAL_MS);
    return () => clearInterval(id);
  }, [refresh]);

  const ts = lastTick
    ? lastTick.toLocaleTimeString('ko-KR', { hour12: false })
    : '—';

  return (
    <div className="monitor">
      <div className="monitor-head">
        <span className="monitor-tick">마지막 갱신: <strong>{ts}</strong></span>
        <button
          className="monitor-refresh"
          onClick={refresh}
          disabled={loading}
          title="수동 새로고침 (자동 폴링은 60초)"
        >
          {loading ? '⏳' : '🔄'} 새로고침
        </button>
      </div>
      <div className="kv"><span className="k">총 호출 수</span><span className="v">{s?.total_calls ?? 0}</span></div>
      <div className="kv"><span className="k">평균 Latency</span><span className="v">{s ? s.avg_latency_ms.toFixed(1) : 0} ms</span></div>
      <div className="kv"><span className="k">P95 Latency</span><span className="v">{s ? s.p95_latency_ms.toFixed(1) : 0} ms</span></div>
      <div className="kv"><span className="k">예상 토큰</span><span className="v">{s?.total_tokens?.toLocaleString() ?? 0}</span></div>
      <div className="kv"><span className="k">예상 과금</span><span className="v">$ {s?.total_cost_usd?.toFixed(6) ?? 0}</span></div>
      <div className="kv"><span className="k">에러율</span><span className="v">{s ? (s.error_rate*100).toFixed(2) : 0}%</span></div>
      <div style={{marginTop:8,fontSize:11,color:'#6b7891'}}>최근 호출</div>
      <div style={{maxHeight:140,overflow:'auto',fontFamily:'monospace',fontSize:10.5}}>
        {logs.slice().reverse().map((l, i) => (
          <div key={i} style={{borderBottom:'1px dotted #dde3ee',padding:'2px 0'}}>
            <span style={{color:l.status>=400?'#e0533d':'#2bb673'}}>{l.status}</span>{' '}
            {l.method} {l.path} · {l.latency_ms}ms · ${l.est_cost_usd}
          </div>
        ))}
      </div>
    </div>
  );
}
