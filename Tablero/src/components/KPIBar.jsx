
import React, { useMemo } from 'react'

export default function KPIBar({ data = [], field, threshold = 0, title = '', unit = 'días' }) {
  const { green, red, total, avg } = useMemo(() => {
    let g = 0, r = 0, sum = 0, n = 0
    for (const row of data) {
      const v = Number(row?.[field])
      if (Number.isFinite(v)) {
        n += 1; sum += v
        if (v <= threshold) g += 1
        else r += 1
      }
    }
    return { green: g, red: r, total: g + r, avg: n ? (sum / n) : null }
  }, [data, field, threshold])

  const pctG = total ? (green / total) * 100 : 0
  const pctR = total ? (red / total) * 100 : 0

  return (
    <div>
      <h2>{title}</h2>

      <div className="kpi-bar" role="meter" aria-valuemin={0} aria-valuenow={green} aria-valuemax={total}>
        <div className="kpi-bar__green" style={{ width: `${pctG}%` }} />
        <div className="kpi-bar__red" style={{ width: `${pctR}%` }} />
      </div>

      <div className="kpi-legend">
        <span className="pill">
          <i className="dot green" /> Verde (≤ {threshold} {unit}): <strong>{green}</strong>
        </span>
        <span className="pill">   
          <i className="dot red" /> Rojo ({'>'} {threshold} {unit}): <strong>{red}</strong>
        </span>
        <span>Total: <strong>{total}</strong></span>
      </div>

      {avg !== null && (
        <div className="meta">
          Promedio: <strong>{avg.toFixed(2)}</strong> {unit}
        </div>
      )}
    </div>
  )
}
