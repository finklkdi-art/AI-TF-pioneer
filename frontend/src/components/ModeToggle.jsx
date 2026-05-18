// Source 3, 4 — 'Fast' / 'Precise' 토글 (상단 상시 노출).
import React from 'react';
import { setMode } from '../api.js';

export default function ModeToggle({ mode, onChange }) {
  function pick(m) {
    setMode(m);
    onChange(m);
  }
  return (
    <div className="mode-toggle" role="tablist" aria-label="작업 모드">
      <button
        className={mode === 'fast' ? 'active' : ''}
        onClick={() => pick('fast')}
        title="Fast: 이미지 임베딩 스캔 — 표준 양식 견적서를 빠르게 처리. 토큰비용 절감."
      >⚡ Fast</button>
      <button
        className={mode === 'precise' ? 'active' : ''}
        onClick={() => pick('precise')}
        title="Precise: 전체 OCR + 라인 정밀 교차검증. 비정형/수기 메모 견적서용."
      >🎯 Precise</button>
    </div>
  );
}
