
import React, { useEffect, useState } from 'react'
import KPIBar from './components/KPIBar'
import './styles.css'

export default function App() {
  const [rows, setRows] = useState([])
  const [error, setError] = useState(null)
  const [loading, setLoading] = useState(true) // 👈 faltaba

  useEffect(() => {  
    fetch('/api/indicadores/tabla-resultado')
      .then((r) => {
        if (!r.ok) throw new Error('Backend respondió con error');
        return r.json();
      })
      .then((data) => {
        if (!Array.isArray(data)) throw new Error('JSON inesperado')
        setRows(data)
      })
      .catch((e) => setError(String(e)))
      .finally(() => setLoading(false))
  }, [])

  if (loading) return <h2 style={{ padding: 16 }}>Cargando datos...</h2>

  return (
    <>
      <header className="header">
        <h1>REPORTE — Indicadores operativos</h1>
      </header>

      <main className="container">
        <section className="card">
          <KPIBar
            title="TOTAL_DIAS_VENTANILLA (verde ≤ 2 días)"
            data={rows}
            field="TOTAL_DIAS_VENTANILLA"
            threshold={2}
          />
        </section>

        <section className="card">
          <KPIBar
            title="TIEMPO_EN_MESA_DE_CREACION (verde ≤ 4 días)"
            data={rows}
            field="TIEMPO_EN_MESA_DE_CREACION"
            threshold={4}
          />
        </section>

        {error && (
          <pre className="card" style={{ gridColumn: 'span 12', color: 'crimson' }}>
            {error}
          </pre>
        )}
      </main>
    </>
  )
}
