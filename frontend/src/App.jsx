// BLUE NINE — Step-by-step orchestrator.
//
// Flow:
//   Step 1 (대형 2버튼) → 클릭 시 자동 Step 2
//   Step 2 (세부 카테고리) → 클릭 시 자동 Step 3
//   Step 3 (멀티파일 업로드 + 분석 시작) → 완료 시 Step 4
//   Step 4 (결과: 신호등 테이블 + Export)
//
// 각 step 에 '이전 단계로' 버튼 (Step 1 제외).

import React, { useEffect, useState } from 'react';
import ModeToggle from './components/ModeToggle.jsx';
import StepIndicator from './components/StepIndicator.jsx';
import Step1Category from './components/Step1Category.jsx';
import Step2SubCategory from './components/Step2SubCategory.jsx';
import Step3Upload from './components/Step3Upload.jsx';
import EstimateSheet from './components/EstimateSheet.jsx';
import ExportBar from './components/ExportBar.jsx';
import MonitorPanel from './components/MonitorPanel.jsx';
import {
  ensureSession, getSessionId, fetchMasterInfo, destroySession,
} from './api.js';

export default function App() {
  const [currentStep, setCurrentStep] = useState(1);
  const [mode, setMode] = useState('precise');
  const [pick, setPick] = useState({ l1: null, l2: null });
  const [doc, setDoc] = useState(null);
  const [sessionId, setSid] = useState(null);
  const [master, setMaster] = useState(null);
  const [showMonitor, setShowMonitor] = useState(false);

  useEffect(() => {
    (async () => {
      await ensureSession();
      setSid(getSessionId());
      try { setMaster(await fetchMasterInfo()); } catch {}
    })();
    const off = () => destroySession();
    window.addEventListener('beforeunload', off);
    return () => window.removeEventListener('beforeunload', off);
  }, []);

  function reset() {
    setPick({ l1: null, l2: null });
    setDoc(null);
    setCurrentStep(1);
  }

  function pickL1(l1) { setPick({ l1, l2: null }); setCurrentStep(2); }
  function pickL2(l2) { setPick((p) => ({ ...p, l2 })); setCurrentStep(3); }
  function back() { setCurrentStep((s) => Math.max(1, s - 1)); }
  function onAnalyzed(doc) { setDoc(doc); setCurrentStep(4); }
  function jumpTo(n) {
    // 진행 표시기 클릭 시 이전 단계로만 이동 허용
    if (n < currentStep) setCurrentStep(n);
  }

  return (
    <div className="app-shell v2">
      <div className="topbar">
        <div className="logo">BLUE <span className="nine">NINE</span></div>
        <div className="tagline">사내 AE 전용 광고 견적서 효율화 솔루션</div>
        <div className="spacer" />
        <ModeToggle mode={mode} onChange={setMode} />
        <button
          className="monitor-toggle"
          onClick={() => setShowMonitor((v) => !v)}
          title="운영 모니터링 패널 토글"
        >📊</button>
        <div className="session-pill" title="휘발성 세션 ID — 브라우저 종료 시 자동 파기">
          SID: {sessionId ? sessionId.slice(0, 8) + '...' : '—'}
        </div>
      </div>

      <div className="banner">
        🔒 입력 데이터는 <b>서버/DB에 저장되지 않으며</b> 브라우저 종료 또는 30분 idle 시 자동 파기됩니다.
        Export 후 종료를 권장합니다. {master ? `· 마스터: v${master.version} (${master.item_count}건)` : ''}
      </div>

      <StepIndicator current={currentStep} onJump={jumpTo} />

      <main className="step-main">
        {currentStep === 1 && <Step1Category onPick={pickL1} />}
        {currentStep === 2 && (
          <Step2SubCategory l1={pick.l1} onPick={pickL2} onBack={back} />
        )}
        {currentStep === 3 && (
          <Step3Upload
            l1={pick.l1}
            l2={pick.l2}
            mode={mode}
            onBack={back}
            onAnalyzed={onAnalyzed}
          />
        )}
        {currentStep === 4 && doc && (
          <div className="step-screen result">
            <div className="result-actions">
              <button className="step-back" onClick={back}>← 이전 단계로</button>
              <button className="step-restart" onClick={reset}>↺ 새 견적서 시작</button>
            </div>
            <EstimateSheet doc={doc} onUpdated={setDoc} />
            <ExportBar doc={doc} />
          </div>
        )}
      </main>

      {showMonitor && (
        <aside className="monitor-drawer">
          <div className="md-head">
            <strong>운영 모니터링</strong>
            <button onClick={() => setShowMonitor(false)} aria-label="닫기">×</button>
          </div>
          <MonitorPanel />
        </aside>
      )}

      <footer className="foot">
        <span>BLUE NINE v1.0 Prototype · Rule Book v1.0</span>
        <span>Volatile Session · No DB · Chrome 권장</span>
      </footer>
    </div>
  );
}
