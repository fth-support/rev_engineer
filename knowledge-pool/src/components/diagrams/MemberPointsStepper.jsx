import { CreditCard } from 'lucide-react'
import DiagramFrame from './DiagramFrame'
import Stepper from '../primitives/Stepper'
import { memberPointsSteps } from '../../data/diagramData'

function MemberPointsStepper() {
  return (
    <DiagramFrame
      icon={<CreditCard size={20} color="var(--accent)" />}
      title="Member Points Flow"
      subtitle="W_UPDxUpdatePoint() — accrue loyalty points to the Member DB"
      hint="Click a step or use Prev / Next to walk the flow"
    >
      <Stepper steps={memberPointsSteps} />
    </DiagramFrame>
  )
}

export default MemberPointsStepper
