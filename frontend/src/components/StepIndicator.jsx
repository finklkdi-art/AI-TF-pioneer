// 상단 진행 표시기 (① 청구유형 → ② 세부카테고리 → ③ 파일분석 → ④ 결과)
import React from 'react';

const STEPS = [
  { n: 1, label: '청구 유형' },
  { n: 2, label: '세부 카테고리' },
  { n: 3, label: '파일 업로드 & 분석' },
  { n: 4, label: '결과 확인' },
];

export default function StepIndicator({ current, onJump }) {
  return (
    <ol className="step-indicator">
      {STEPS.map((s) => {
        const state = current === s.n ? 'active' : current > s.n ? 'done' : 'todo';
        const clickable = current > s.n && typeof onJump === 'function';
        return (
          <li key={s.n} className={`step-pill ${state} ${clickable ? 'clickable' : ''}`}
              onClick={() => clickable && onJump(s.n)}>
            <span className="step-num">{state === 'done' ? '✓' : s.n}</span>
            <span className="step-name">{s.label}</span>
          </li>
        );
      })}
    </ol>
  );
}
