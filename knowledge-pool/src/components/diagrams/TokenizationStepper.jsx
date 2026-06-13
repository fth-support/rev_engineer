import { ShieldCheck } from 'lucide-react'
import DiagramFrame from './DiagramFrame'
import Stepper from '../primitives/Stepper'
import { tokenizationSteps } from '../../data/diagramData'

function TokenizationStepper() {
  return (
    <DiagramFrame
      icon={<ShieldCheck size={20} color="var(--accent)" />}
      title="Tokenization Flow"
      subtitle="SP_PRCbToken() — guard → cache → SafeNet SOAP → persist"
      hint="Click a step or use Prev / Next; watch the guard & error branches"
    >
      <Stepper steps={tokenizationSteps} />
    </DiagramFrame>
  )
}

export default TokenizationStepper
