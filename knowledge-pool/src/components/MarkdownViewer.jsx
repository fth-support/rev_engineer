import { useState, useEffect } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'

import ArchitectureGraph from './diagrams/ArchitectureGraph'
import TargetArchitectureGraph from './diagrams/TargetArchitectureGraph'
import TxLifecycleStepper from './diagrams/TxLifecycleStepper'
import SyncFlowStepper from './diagrams/SyncFlowStepper'
import ERGraph from './diagrams/ERGraph'
import TokenizationStepper from './diagrams/TokenizationStepper'
import MemberPointsStepper from './diagrams/MemberPointsStepper'

// Maps a ```diagram fenced block (whose body is an id) to an interactive component.
const DIAGRAMS = {
  architecture: ArchitectureGraph,
  target_architecture: TargetArchitectureGraph,
  tx_lifecycle: TxLifecycleStepper,
  sync_flow: SyncFlowStepper,
  er_diagram: ERGraph,
  tokenization: TokenizationStepper,
  member_points: MemberPointsStepper,
}

function DiagramRenderer({ id }) {
  const key = String(id).trim().toLowerCase()
  const Comp = DIAGRAMS[key]
  if (!Comp) return <div style={{ color: 'var(--danger)', padding: '1rem' }}>Unknown diagram: {id}</div>
  return <div className="markdown-diagram"><Comp /></div>
}

function MarkdownViewer({ file }) {
  const [content, setContent] = useState('Loading…')

  useEffect(() => {
    let cancelled = false
    fetch(`./docs/${file}`)
      .then((res) => res.text())
      .then((text) => { if (!cancelled) setContent(text) })
      .catch((err) => { if (!cancelled) setContent(`Error loading document: ${err.message}`) })
    return () => { cancelled = true }
  }, [file])

  return (
    <div className="markdown-body glass" style={{ padding: '2rem 2.25rem' }}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        components={{
          code({ inline, className, children, ...props }) {
            const match = /language-(\w+)/.exec(className || '')
            if (!inline && match && match[1] === 'diagram') {
              return <DiagramRenderer id={children} />
            }
            return <code className={className} {...props}>{children}</code>
          },
        }}
      >
        {content}
      </ReactMarkdown>
    </div>
  )
}

export default MarkdownViewer
