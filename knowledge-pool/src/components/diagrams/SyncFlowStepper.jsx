import { RefreshCw } from 'lucide-react'
import DiagramFrame from './DiagramFrame'
import Stepper from '../primitives/Stepper'
import { syncFlowSteps } from '../../data/diagramData'

function SyncFlowStepper() {
  return (
    <DiagramFrame
      icon={<RefreshCw size={20} color="var(--accent)" />}
      title="Data Synchronization Flow"
      subtitle="otmForm_Timer() — one 6-phase cycle, every 500ms"
      hint="Click a phase or use Prev / Next to step through"
    >
      <Stepper steps={syncFlowSteps} />
    </DiagramFrame>
  )
}

export default SyncFlowStepper
