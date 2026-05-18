// Source 11, 12, 17 — 사용자가 가장 먼저 선택해야 하는 분기:
//   [Step 1]  제작비 / 매체비
//   [Step 2]  L1 선택값에 따라 동적으로 하위 카테고리 노출
import React from 'react';

export default function CategoryStepper({ cats, l1, l2, onPick }) {
  if (!cats) return null;
  return (
    <>
      <div className="step">
        <div className="step-label">Step 1 — 청구 유형</div>
        <div className="btn-row">
          {cats.step1.map((c) => (
            <button
              key={c.key}
              className={`btn-choice ${l1 === c.key ? 'active' : ''}`}
              onClick={() => onPick({ l1: c.key, l2: null })}
            >{c.label}</button>
          ))}
        </div>
      </div>
      {l1 && (
        <div className="step">
          <div className="step-label">Step 2 — 세부 카테고리</div>
          <div className="btn-row">
            {cats.step2[l1].map((c) => (
              <button
                key={c.key}
                className={`btn-choice ${l2 === c.key ? 'active' : ''}`}
                onClick={() => onPick({ l1, l2: c.key })}
                title={`템플릿: ${c.template}`}
              >{c.label}</button>
            ))}
          </div>
        </div>
      )}
    </>
  );
}
