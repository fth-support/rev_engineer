import { GitBranch } from 'lucide-react'
import DiagramFrame from './DiagramFrame'
import Stepper from '../primitives/Stepper'
import { txLifecycleSteps } from '../../data/diagramData'

function TxLifecycleStepper() {
  return (
    <DiagramFrame
      icon={<GitBranch size={20} color="var(--accent)" />}
      title="End-to-End Transaction Lifecycle"
      subtitle="POSFront (producer) → Local DB → ServiceTransfer (consumer) → HQ"
      hint="Step through one sale, from cashier to Head Office"
    >
      <Stepper steps={txLifecycleSteps} />
    </DiagramFrame>
  )
}

export default TxLifecycleStepper
