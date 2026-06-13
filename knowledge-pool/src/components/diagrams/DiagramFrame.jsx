import { MousePointerClick } from 'lucide-react'
import './Diagrams.css'

// Shared header/wrapper for every interactive diagram.
function DiagramFrame({ icon, title, subtitle, hint, children }) {
  return (
    <div className="diagram-frame">
      <div className="diagram-frame__head">
        <h3 className="diagram-frame__title">{icon}{title}</h3>
        {subtitle && <p className="diagram-frame__sub">{subtitle}</p>}
        {hint && (
          <span className="diagram-frame__hint"><MousePointerClick size={13} />{hint}</span>
        )}
      </div>
      {children}
    </div>
  )
}

export default DiagramFrame
