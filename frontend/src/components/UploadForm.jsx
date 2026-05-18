// 협력사 견적서 업로드 — 휘발성 보관 (Source 9).
import React, { useState } from 'react';
import { parseEstimate } from '../api.js';

export default function UploadForm({ l1, l2, onParsed }) {
  const [client, setClient] = useState('');
  const [campaign, setCampaign] = useState('');
  const [versionLabel, setVersionLabel] = useState('초안');
  const [file, setFile] = useState(null);
  const [billing, setBilling] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState('');

  const canSubmit = !!(l1 && l2 && file && !busy);

  async function submit(e) {
    e.preventDefault();
    if (!canSubmit) return;
    setBusy(true);
    setErr('');
    try {
      const doc = await parseEstimate({ l1, l2, file, billing, client, campaign, versionLabel });
      onParsed(doc);
    } catch (e2) {
      setErr(e2.message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <form onSubmit={submit}>
      <div className="field">
        <label>광고주 (CLIENT)</label>
        <input value={client} onChange={(e) => setClient(e.target.value)} placeholder="예) 삼성전자" />
      </div>
      <div className="field">
        <label>캠페인 / Job Name</label>
        <input value={campaign} onChange={(e) => setCampaign(e.target.value)} placeholder="예) 갤럭시 S26 런칭" />
      </div>
      <div className="field">
        <label>버전</label>
        <select value={versionLabel} onChange={(e) => setVersionLabel(e.target.value)}>
          <option>초안</option><option>사전견적</option><option>1차견적</option>
          <option>2차견적</option><option>최종견적</option>
        </select>
      </div>
      <div className="field">
        <label>협력사 견적서 (.xlsx / .xls)</label>
        <input type="file" accept=".xlsx,.xls" onChange={(e) => setFile(e.target.files[0])} />
      </div>
      {l1 === 'media' && (
        <div className="field">
          <label>Billing Status (선택, 매체비 삼각 검증용)</label>
          <input type="file" accept=".xlsx,.xls" onChange={(e) => setBilling(e.target.files[0])} />
        </div>
      )}
      <button type="submit" className="btn-primary" disabled={!canSubmit}>
        {busy ? '분석 중...' : '견적서 생성'}
      </button>
      {!file && <div style={{fontSize:11,color:'#6b7891',marginTop:6}}>파일 선택 후 [견적서 생성]</div>}
      {err && <div className="warning-list" style={{marginTop:8}}>⚠ {err}</div>}
    </form>
  );
}
