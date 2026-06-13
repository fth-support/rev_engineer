import { useState } from 'react'
import { ChevronLeft, ChevronRight, ArrowRightCircle } from 'lucide-react'
import { ROLE_VARS } from '../../data/diagramData'

// Interactive vertical stepper for flow diagrams. The segmented rail (inspired by the
// "Segmented Bar" + "Icon Steps" patterns in ag-kit-main/uxui/ai-steppers) is clickable;
// Prev/Next walk through; the active step renders its code/lines/branch detail.
const TONE_CLASS = { exit: 'is-exit', error: 'is-error', ok: 'is-ok' }

function Stepper({ steps }) {
  const [active, setActive] = useState(0)
  const step = steps[active]
  const roleVar = (r) => `var(--${ROLE_VARS[r] || 'role-config'})`

  return (
    <div className="stepper">
      <div className="stepper__rail" role="tablist" aria-label="Flow steps">
        {steps.map((s, i) => {
          const state = i < active ? 'done' : i === active ? 'active' : 'upcoming'
          return (
            <button
              key={i}
              type="button"
              role="tab"
              aria-selected={i === active}
              className={`stepper__seg ${state}`}
              onClick={() => setActive(i)}
              style={{ '--seg-color': roleVar(s.role) }}
            >
              <span className="stepper__seg-bar" />
              <span className="stepper__seg-text">
                <span className="stepper__seg-num">{String(i + 1).padStart(2, '0')}</span>
                {s.title}
              </span>
            </button>
          )
        })}
      </div>

      <div className="stepper__panel" style={{ '--seg-color': roleVar(step.role) }}>
        <div className="stepper__panel-head">
          <span className="stepper__panel-num">{String(active + 1).padStart(2, '0')}</span>
          <div>
            <h4 className="stepper__panel-title">{step.title}</h4>
            {step.code && <code className="stepper__panel-code">{step.code}</code>}
          </div>
        </div>

        <ul className="stepper__lines">
          {step.lines.map((line, i) => (
            <li key={i}>{line}</li>
          ))}
        </ul>

        {step.branch && (
          <div className={`stepper__branch ${TONE_CLASS[step.branch.tone] || ''}`}>
            <ArrowRightCircle size={15} />
            <span>{step.branch.label}</span>
          </div>
        )}

        <div className="stepper__controls">
          <button type="button" className="stepper__btn" onClick={() => setActive((a) => Math.max(0, a - 1))} disabled={active === 0}>
            <ChevronLeft size={16} /> Prev
          </button>
          <span className="stepper__counter">{active + 1} / {steps.length}</span>
          <button type="button" className="stepper__btn primary" onClick={() => setActive((a) => Math.min(steps.length - 1, a + 1))} disabled={active === steps.length - 1}>
            Next <ChevronRight size={16} />
          </button>
        </div>
      </div>
    </div>
  )
}

export default Stepper
