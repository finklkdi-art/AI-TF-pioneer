// Source 25 — A4 인쇄 + 살아있는 수식 Excel Export.
import React from 'react';
import { exportXlsx } from '../api.js';

export default function ExportBar({ doc }) {
  if (!doc) return null;
  const blocked = doc.overall_light === 'red';

  async function dl() {
    try {
      await exportXlsx(doc.estimate_id, `BLUE_NINE_${doc.category_l1}_${doc.category_l2}_${doc.version_label}.xlsx`);
    } catch (e) {
      alert('Excel 다운로드 실패: ' + e.message);
    }
  }

  return (
    <div className="export-bar">
      {blocked && (
        <span style={{color:'#c8541a',fontSize:12,alignSelf:'center',marginRight:'auto'}}>
          ⚠ Red 검증 경고 — 그래도 다운로드 가능 (수정 후 재출력 권장)
        </span>
      )}
      <button className="alt" onClick={() => window.print()}>🖨 인쇄 (Chrome A4)</button>
      <button onClick={dl}>⬇ Excel 다운로드</button>
    </div>
  );
}
