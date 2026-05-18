// 요구사항 4 — Latency / Traffic / Token-Cost 모니터링 패널.
import React, { useEffect, useState } from 'react';
import { monitorSummary, monitorLogs } from '../api.js';

export default function MonitorPanel() {
  const [s, setS] = useState(null);
  const [logs, setLogs] = useState([]);

  useEffect(() => {
    let mounted = true;
    async function tick() {
      try {
        const [su, lo] = await Promise.all([monitorSummary(), monitorLogs(8)]);
        if (mounted) { setS(su); setLogs(lo); }
      } catch {}
    }
    tick();
    const id = setInterval(tick, 4000);
    return () => { mounted = false; clearInterval(id); };
  }, []);

  return (
    <div className="monitor">
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
