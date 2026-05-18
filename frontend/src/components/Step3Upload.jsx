// Step 3 — 다중 파일 업로드 + 'BLUE NINE 분석 및 취합 시작'.
import React, { useRef, useState } from 'react';
import { parseEstimate } from '../api.js';

const ACCEPT = '.xlsx,.xls,.pdf,.csv';

function fmtSize(b) {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / 1024 / 1024).toFixed(2)} MB`;
}

export default function Step3Upload({ l1, l2, mode, onBack, onAnalyzed }) {
  const [files, setFiles] = useState([]);
  const [client, setClient] = useState('');
  const [campaign, setCampaign] = useState('');
  const [versionLabel, setVersionLabel] = useState('초안');
  const [dragOver, setDragOver] = useState(false);
  const [analyzing, setAnalyzing] = useState(false);
  const [err, setErr] = useState('');
  const inputRef = useRef(null);

  function addFiles(incoming) {
    const arr = Array.from(incoming || []);
    if (!arr.length) return;
    // 중복 제거 (name + size 키)
    setFiles((prev) => {
      const seen = new Set(prev.map((f) => `${f.name}__${f.size}`));
      const merged = [...prev];
      for (const f of arr) {
        const k = `${f.name}__${f.size}`;
        if (!seen.has(k)) { merged.push(f); seen.add(k); }
      }
      return merged;
    });
  }

  function removeAt(i) {
    setFiles((prev) => prev.filter((_, idx) => idx !== i));
  }

  function onDrop(e) {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer?.files?.length) addFiles(e.dataTransfer.files);
  }

  async function analyze() {
    if (!files.length) return;
    setAnalyzing(true);
    setErr('');
    try {
      const doc = await parseEstimate({
        l1, l2, files,
        client: client.trim() || null,
        campaign: campaign.trim() || null,
        versionLabel,
      });
      onAnalyzed(doc);
    } catch (e) {
      setErr(e.message || String(e));
    } finally {
      setAnalyzing(false);
    }
  }

  const catLabel = (l1 === 'production' ? '제작비' : '매체비') + ' · ' + l2.toUpperCase();

  return (
    <div className="step-screen">
      <button className="step-back" onClick={onBack} disabled={analyzing}>← 이전 단계로</button>
      <div className="step-hero">
        <h1 className="step-title">파일 업로드 및 분석</h1>
        <p className="step-sub">
          선택: <strong>{catLabel}</strong> · 모드: <strong>{mode.toUpperCase()}</strong>
          {l1 === 'media' && ' · Billing Status 파일은 자동 인식돼 삼각 검증에 사용됩니다'}
        </p>
      </div>

      <div className="meta-row">
        <label>
          <span>광고주</span>
          <input value={client} onChange={(e) => setClient(e.target.value)} placeholder="예) 삼성전자" disabled={analyzing} />
        </label>
        <label>
          <span>캠페인</span>
          <input value={campaign} onChange={(e) => setCampaign(e.target.value)} placeholder="예) 갤럭시 S26 런칭" disabled={analyzing} />
        </label>
        <label>
          <span>버전</span>
          <select value={versionLabel} onChange={(e) => setVersionLabel(e.target.value)} disabled={analyzing}>
            <option>초안</option>
            <option>사전견적</option>
            <option>1차견적</option>
            <option>2차견적</option>
            <option>최종견적</option>
          </select>
        </label>
      </div>

      <div
        className={`dropzone ${dragOver ? 'over' : ''} ${analyzing ? 'disabled' : ''}`}
        onClick={() => !analyzing && inputRef.current?.click()}
        onDragOver={(e) => { e.preventDefault(); if (!analyzing) setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        role="button"
        tabIndex={0}
      >
        <div className="dropzone-icon">📂</div>
        <div className="dropzone-title">
          {dragOver ? '여기에 놓아주세요' : '여기로 파일을 끌어다 놓거나 클릭해서 선택'}
        </div>
        <div className="dropzone-sub">
          여러 개의 견적서를 동시에 올릴 수 있어요 — .xlsx / .xls / .pdf / .csv
        </div>
        <input
          ref={inputRef}
          type="file"
          multiple
          accept={ACCEPT}
          style={{ display: 'none' }}
          onChange={(e) => { addFiles(e.target.files); e.target.value = ''; }}
        />
      </div>

      {files.length > 0 && (
        <div className="file-list">
          <div className="file-list-head">
            <span>업로드 대기 ({files.length}개)</span>
            <button className="link-btn" onClick={() => setFiles([])} disabled={analyzing}>전체 비우기</button>
          </div>
          <ul>
            {files.map((f, i) => {
              const looksLikeBilling = /billing|빌링|집행내역/i.test(f.name);
              return (
                <li key={`${f.name}-${f.size}-${i}`}>
                  <span className="fname">
                    {looksLikeBilling && l1 === 'media' && <span className="tag billing">Billing Status</span>}
                    {f.name}
                  </span>
                  <span className="fsize">{fmtSize(f.size)}</span>
                  <button className="x-btn" onClick={() => removeAt(i)} disabled={analyzing} aria-label="제거">×</button>
                </li>
              );
            })}
          </ul>
        </div>
      )}

      {err && <div className="warning-list" style={{ marginTop: 10 }}>⚠ {err}</div>}

      <div className="primary-cta">
        <button
          className="cta-button"
          disabled={!files.length || analyzing}
          onClick={analyze}
        >
          {analyzing
            ? '🌀 BLUE NINE 분석 중...'
            : `▶ BLUE NINE 분석 및 취합 시작 ${files.length ? `(${files.length}개 파일)` : ''}`}
        </button>
      </div>

      {analyzing && (
        <div className="analyzing-overlay" role="status" aria-live="polite">
          <div className="spinner" />
          <div className="analyzing-text">
            <strong>BLUE NINE</strong> 가 {files.length}개의 파일을 분석 중입니다…
            <div className="analyzing-sub">휴리스틱 파서 → 신호등 검증 → 합계 더블체크</div>
          </div>
        </div>
      )}
    </div>
  );
}
