// BLUE NINE — App shell. (Source 37: 이름 일관 노출)
import React, { useEffect, useState } from 'react';
import ModeToggle from './components/ModeToggle.jsx';
import CategoryStepper from './components/CategoryStepper.jsx';
import UploadForm from './components/UploadForm.jsx';
import EstimateSheet from './components/EstimateSheet.jsx';
import ExportBar from './components/ExportBar.jsx';
import MonitorPanel from './components/MonitorPanel.jsx';
import { ensureSession, getSessionId, fetchCategories, fetchMasterInfo, destroySession } from './api.js';

export default function App() {
  const [mode, setMode] = useState('precise');
  const [cats, setCats] = useState(null);
  const [pick, setPick] = useState({ l1: null, l2: null });
  const [doc, setDoc] = useState(null);
  const [sessionId, setSid] = useState(null);
  const [master, setMaster] = useState(null);

  useEffect(() => {
    (async () => {
      await ensureSession();
      setSid(getSessionId());
      setCats(await fetchCategories());
      setMaster(await fetchMasterInfo());
    })();
    // Source 10 — 브라우저 종료 시 세션 파기
    const off = () => destroySession();
    window.addEventListener('beforeunload', off);
    return () => window.removeEventListener('beforeunload', off);
  }, []);

  return (
    <div className="app-shell">
      <div className="topbar">
        <div className="logo">BLUE <span className="nine">NINE</span></div>
        <div className="tagline">사내 AE 전용 광고 견적서 효율화 솔루션</div>
        <div className="spacer" />
        <ModeToggle mode={mode} onChange={setMode} />
        <div className="session-pill" title="휘발성 세션 ID — 브라우저 종료 시 자동 파기">
          SID: {sessionId ? sessionId.slice(0, 8) + '...' : '—'}
        </div>
      </div>

      <div className="banner">
        🔒 입력하신 청구 정보는 <b>서버/DB에 저장되지 않으며</b> 브라우저 종료 또는 30분 idle 시 자동 파기됩니다.
        Export 후 종료를 권장합니다. <b>마스터 데이터(제작단가기준집)</b> 항목 수: {master?.item_count ?? 0}
      </div>

      <main className="workspace">
        <aside className="panel">
          <h3>① 카테고리 분기</h3>
          <CategoryStepper cats={cats} l1={pick.l1} l2={pick.l2} onPick={setPick} />
          <hr style={{border:0,borderTop:'1px solid #e5eaf2',margin:'12px 0'}} />
          <h3>② 협력사 견적서 업로드</h3>
          <UploadForm l1={pick.l1} l2={pick.l2} onParsed={setDoc} />
        </aside>

        <section>
          {!doc && (
            <div className="panel" style={{textAlign:'center',padding:'60px 20px'}}>
              <h2 style={{color:'#0b2a4a',margin:0}}>BLUE NINE에 오신 것을 환영합니다</h2>
              <p style={{color:'#6b7891'}}>
                좌측에서 <b>제작비 / 매체비</b> → 세부 카테고리를 선택한 뒤,<br />
                협력사 견적서를 업로드하면 자동으로 검증된 견적서가 생성됩니다.
              </p>
              <div style={{display:'inline-flex',gap:14,marginTop:12,fontSize:12,color:'#6b7891'}}>
                <span><span className="light green" /> 정상 ≥ 99%</span>
                <span><span className="light yellow" /> 주의 ≥ 90%</span>
                <span><span className="light red" /> 위험 &lt; 90%</span>
              </div>
            </div>
          )}
          {doc && <EstimateSheet doc={doc} onUpdated={setDoc} />}
          {doc && <ExportBar doc={doc} />}
        </section>

        <aside className="panel">
          <h3>③ 운영 모니터링</h3>
          <MonitorPanel />
          <hr style={{border:0,borderTop:'1px solid #e5eaf2',margin:'12px 0'}} />
          <h3>마스터 데이터</h3>
          <div className="monitor">
            <div className="kv"><span className="k">버전</span><span className="v">v{master?.version ?? 0}</span></div>
            <div className="kv"><span className="k">항목 수</span><span className="v">{master?.item_count ?? 0}</span></div>
            <div style={{fontSize:11,color:'#6b7891',marginTop:6}}>
              일반 사용자: 읽기 전용. 변경은 Admin 콘솔.
            </div>
          </div>
        </aside>
      </main>

      <footer className="foot">
        <span>BLUE NINE v1.0 Prototype · Rule Book v1.0</span>
        <span>Volatile Session · No DB · Chrome 권장</span>
      </footer>
    </div>
  );
}
