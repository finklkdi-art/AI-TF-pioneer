// Step 1 — 청구 유형 선택 (제작비 / 매체비). 대형 버튼만 노출.
import React from 'react';

export default function Step1Category({ onPick }) {
  return (
    <div className="step-screen">
      <div className="step-hero">
        <h1 className="step-title">청구하실 견적의 유형을 선택해 주세요</h1>
        <p className="step-sub">선택에 따라 검증 양식과 매칭 로직이 자동으로 결정됩니다 (Source 11)</p>
      </div>
      <div className="big-choice-row">
        <button
          className="big-choice production"
          onClick={() => onPick('production')}
        >
          <div className="big-icon">🎬</div>
          <div className="big-label">제작비</div>
          <div className="big-desc">영상 · 라디오 · 인쇄 · BTL · 기타</div>
        </button>
        <button
          className="big-choice media"
          onClick={() => onPick('media')}
        >
          <div className="big-icon">📺</div>
          <div className="big-label">매체비</div>
          <div className="big-desc">TVC · 라디오 · PRINT · 디지털 · 기타</div>
        </button>
      </div>
    </div>
  );
}
