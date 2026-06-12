import { useState, useEffect, useRef } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'

// Import diagram components
import SystemArchitecture from './diagrams/SystemArchitecture'
import SyncFlow from './diagrams/SyncFlow'
import ERDiagram from './diagrams/ERDiagram'
import TokenizationFlow from './diagrams/TokenizationFlow'
import MemberPointsFlow from './diagrams/MemberPointsFlow'
import TargetArchitecture from './diagrams/TargetArchitecture'

const DiagramRenderer = ({ id }) => {
  const normalizedId = String(id).trim().toLowerCase();
  switch (normalizedId) {
    case 'architecture': return <SystemArchitecture />;
    case 'target_architecture': return <TargetArchitecture />;
    case 'sync_flow': return <SyncFlow />;
    case 'er_diagram': return <ERDiagram />;
    case 'tokenization': return <TokenizationFlow />;
    case 'member_points': return <MemberPointsFlow />;
    default: return <div style={{ color: 'red', padding: '1rem' }}>Unknown diagram: {id}</div>;
  }
};

function MarkdownViewer({ file }) {
  const [content, setContent] = useState('Loading...')
  const containerRef = useRef(null)

  useEffect(() => {
    fetch(`./docs/${file}`)
      .then((res) => res.text())
      .then((text) => setContent(text))
      .catch((err) => setContent(`Error loading document: ${err.message}`))
  }, [file])

  return (
    <div className="markdown-body" ref={containerRef}>
      <ReactMarkdown 
        remarkPlugins={[remarkGfm]}
        components={{
          code({node, inline, className, children, ...props}) {
            const match = /language-(\w+)/.exec(className || '')
            if (!inline && match && match[1] === 'diagram') {
              return <DiagramRenderer id={children} />
            }
            return !inline && match ? (
              <code className={className} {...props}>
                {children}
              </code>
            ) : (
              <code className={className} {...props}>
                {children}
              </code>
            )
          }
        }}
      >
        {content}
      </ReactMarkdown>
    </div>
  )
}

export default MarkdownViewer
