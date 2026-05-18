// Source 24, 25, 28~33 — 견적서 본문 + 3단계 신호등 + AE 수동 수정 인풋.
import React, { useState } from 'react';
import { updateRow } from '../api.js';

const SECTION_ORDER_PRODUCTION = ['정가항목', '외주비', '대행수수료'];
const SECTION_ORDER_MEDIA = ['매체청구액', '매체지급액', '매체수수료'];

function fmt(n) {
  if (n == null || isNaN(n)) return '';
  return Math.round(n).toLocaleString('ko-KR');
}

function Light({ color }) {
  return <span className={`light ${color}`} title={`신호등 ${color}`} />;
}

function EditableNum({ row, field, onSave }) {
  const [v, setV] = useState(String(row[field] ?? ''));
  const editable = row.light !== 'green';   // 노랑/빨강 인 경우 수정 가능 (요구사항 3)
  return (
    <input
      className={`cell-num ${editable ? 'editable' : ''}`}
      readOnly={!editable}
      value={v}
      onChange={(e) => setV(e.target.value.replace(/[^\d.\-]/g, ''))}
      onBlur={() => {
        const parsed = parseFloat(v.replace(/,/g, ''));
        if (!isNaN(parsed) && parsed !== row[field]) onSave(field, parsed);
        else setV(String(row[field] ?? ''));
      }}
    />
  );
}

export default function EstimateSheet({ doc, onUpdated }) {
  const [openRowId, setOpenRowId] = useState(null);

  async function persist(rowId, field, value) {
    try {
      const updated = await updateRow({ estimateId: doc.estimate_id, rowId, patch: { [field]: value } });
      onUpdated(updated);
    } catch (e) {
      alert('업데이트 실패: ' + e.message);
    }
  }

  const order = doc.category_l1 === 'production' ? SECTION_ORDER_PRODUCTION : SECTION_ORDER_MEDIA;
  const grouped = order.map((sec) => ({ sec, rows: doc.rows.filter((r) => r.section === sec) }));
  const orphan = doc.rows.filter((r) => !order.includes(r.section));
  if (orphan.length) grouped.push({ sec: '미분류', rows: orphan });

  return (
    <div className="sheet">
      <header className="sheet-head">
        <div>
          <h2>BLUE NINE 광고 견적서</h2>
          <div style={{fontSize:12,color:'#6b7891'}}>
            {doc.category_l1.toUpperCase()} · {doc.category_l2.toUpperCase()} · {doc.version_label}
          </div>
        </div>
        <div style={{textAlign:'right',fontSize:12}}>
          <div><strong>{doc.client || '광고주 미입력'}</strong></div>
          <div>{doc.campaign || '캠페인 미입력'}</div>
          <div>{doc.issue_date}</div>
        </div>
      </header>

      <div className="meta-grid">
        <div><span className="k">Job No.</span> {doc.job_no || '—'}</div>
        <div><span className="k">처리 모드</span> {doc.mode.toUpperCase()}</div>
        <div><span className="k">신호등</span> <Light color={doc.overall_light} /> {Math.round(doc.overall_confidence*100)}%</div>
      </div>

      {doc.sources?.length > 0 && (
        <div className="source-summary">
          <strong>📥 입력 파일 ({doc.sources.length}개)</strong>
          <ul>
            {doc.sources.map((s, i) => (
              <li key={i} className={s.error ? 'err' : ''}>
                <span className={`role-tag ${s.role}`}>{s.role === 'billing' ? 'Billing' : 'Estimate'}</span>
                <span className="src-name">{s.filename}</span>
                <span className="src-stat">{s.error ? `❌ ${s.error}` : `${s.rows}개 행 · ${(s.size_bytes/1024).toFixed(1)} KB`}</span>
              </li>
            ))}
          </ul>
        </div>
      )}

      {doc.warnings?.length > 0 && (
        <ul className="warning-list">
          {doc.warnings.map((w, i) => <li key={i}>{w}</li>)}
        </ul>
      )}

      <table className="estimate">
        <thead>
          <tr>
            <th style={{width:80}}>섹션</th>
            <th>항목</th>
            <th style={{width:110}}>협력사</th>
            <th style={{width:110}}>단가</th>
            <th style={{width:60}}>수량</th>
            <th style={{width:130}}>금액</th>
            <th style={{width:36}}>신호</th>
            <th style={{width:130}}>원천 파일</th>
          </tr>
        </thead>
        <tbody>
          {grouped.map(({ sec, rows }) => rows.length > 0 && (
            <React.Fragment key={sec}>
              <tr className="section-head"><td colSpan={8}>{sec}</td></tr>
              {rows.map((r) => (
                <React.Fragment key={r.id}>
                  <tr className={`row-${r.light}`} onClick={() => setOpenRowId(openRowId === r.id ? null : r.id)}>
                    <td>{/* hidden — section label above */}</td>
                    <td>{r.item_name}</td>
                    <td>{r.vendor || ''}</td>
                    <td className="num"><EditableNum row={r} field="unit_price" onSave={(f,v)=>persist(r.id,f,v)} /></td>
                    <td className="num"><EditableNum row={r} field="quantity"  onSave={(f,v)=>persist(r.id,f,v)} /></td>
                    <td className="num"><EditableNum row={r} field="amount"    onSave={(f,v)=>persist(r.id,f,v)} /></td>
                    <td style={{textAlign:'center'}}><Light color={r.light} /></td>
                    <td style={{fontSize:11,color:'#6b7891'}} title={r.source_file || ''}>
                      {r.source_file ? r.source_file.length > 16 ? r.source_file.slice(0,14)+'…' : r.source_file : (r.note || '')}
                    </td>
                  </tr>
                  {openRowId === r.id && (
                    <tr>
                      <td colSpan={8}>
                        <div className="reasoning-pop">
                          <strong>AI 판정 근거</strong> · confidence {Math.round(r.confidence*100)}% · {r.reasoning}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </React.Fragment>
          ))}
        </tbody>
      </table>

      <div className="totals">
        <div className="card"><div className="k">(A) 정가합계</div><div className="v">{fmt(doc.sum_jeongga)}</div></div>
        <div className="card"><div className="k">(B) 외주비</div>  <div className="v">{fmt(doc.sum_outsourcing)}</div></div>
        <div className="card"><div className="k">(C) 대행수수료</div><div className="v">{fmt(doc.sum_agency_fee)}</div></div>
        <div className="card"><div className="k">청구금액 (VAT 포함)</div><div className="v">{fmt(doc.sum_with_vat)} ₩</div></div>
      </div>

      {doc.triangle && (
        <div className="warning-list" style={{
          background: doc.triangle.consistent ? '#eaf7f0' : '#fff5f3',
          borderColor: doc.triangle.consistent ? '#bfe5cf' : '#f5c2b8',
          color: doc.triangle.consistent ? '#1f6b46' : '#8a2c1a',
        }}>
          <strong>매체비 삼각 검증 (광고주 청구액 — 매체사 지급액 — 대행수수료)</strong>
          <div>광고주 청구: {fmt(doc.triangle.media_charged_sum)} / Billing: {fmt(doc.triangle.billing_status_sum)} / 지급+수수료: {fmt(doc.triangle.media_paid_plus_fee)}</div>
          <div>Δ = {fmt(doc.triangle.delta)} — {doc.triangle.consistent ? '✅ 일치' : '❌ 불일치'}</div>
        </div>
      )}
    </div>
  );
}
