// Step 2 — Step 1 선택에 따른 세부 카테고리. 항목 클릭 시 즉시 Step 3 로 이동.
import React from 'react';

const SUB = {
  production: [
    { key: 'video', label: '영상', icon: '🎬', template: '영상제작비견적서' },
    { key: 'radio', label: '라디오', icon: '📻', template: '라디오제작비견적서' },
    { key: 'print', label: '인쇄',   icon: '🖨', template: '인쇄제작비견적서' },
    { key: 'btl',   label: 'BTL',    icon: '🎪', template: 'BTL제작비견적서' },
    { key: 'other', label: '기타',   icon: '🧩', template: 'Generic' },
  ],
  media: [
    { key: 'tvc',     label: 'TVC',     icon: '📺', template: 'Billing Status' },
    { key: 'radio',   label: '라디오',   icon: '📻', template: 'Billing Status' },
    { key: 'print',   label: 'PRINT',   icon: '📰', template: 'Billing Status' },
    { key: 'digital', label: '디지털',   icon: '💻', template: 'Billing Status' },
    { key: 'other',   label: '기타',    icon: '🧩', template: 'Billing Status' },
  ],
};

export default function Step2SubCategory({ l1, onPick, onBack }) {
  const subs = SUB[l1] || [];
  const heading = l1 === 'production' ? '제작비' : '매체비';
  return (
    <div className="step-screen">
      <button className="step-back" onClick={onBack}>← 이전 단계로</button>
      <div className="step-hero">
        <h1 className="step-title">{heading} — 세부 카테고리</h1>
        <p className="step-sub">선택한 카테고리에 맞는 검증 템플릿이 자동 매칭됩니다</p>
      </div>
      <div className="medium-choice-grid">
        {subs.map((s) => (
          <button
            key={s.key}
            className="medium-choice"
            onClick={() => onPick(s.key)}
            title={`템플릿: ${s.template}`}
          >
            <div className="medium-icon">{s.icon}</div>
            <div className="medium-label">{s.label}</div>
            <div className="medium-template">{s.template}</div>
          </button>
        ))}
      </div>
    </div>
  );
}
