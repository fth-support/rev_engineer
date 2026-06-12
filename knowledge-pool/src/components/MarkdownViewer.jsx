import { useState, useEffect, useRef } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import mermaid from 'mermaid'

function MarkdownViewer({ file }) {
  const [content, setContent] = useState('Loading...')
  const containerRef = useRef(null)

  useEffect(() => {
    mermaid.initialize({
      startOnLoad: false,
      theme: 'dark',
      securityLevel: 'loose',
    })
  }, [])

  useEffect(() => {
    fetch(`./docs/${file}`)
      .then((res) => res.text())
      .then((text) => setContent(text))
      .catch((err) => setContent(`Error loading document: ${err.message}`))
  }, [file])

  useEffect(() => {
    // Render mermaid diagrams after content updates
    if (containerRef.current) {
      const elements = containerRef.current.querySelectorAll('.mermaid')
      elements.forEach((el, i) => {
        const id = `mermaid-${Date.now()}-${i}`
        const graphDefinition = el.textContent
        if (el.getAttribute('data-processed')) return
        
        el.setAttribute('data-processed', 'true')
        mermaid.render(id, graphDefinition).then(({ svg }) => {
          el.innerHTML = svg
        }).catch(e => console.error('Mermaid render error', e))
      })
    }
  }, [content])

  return (
    <div className="markdown-body" ref={containerRef}>
      <ReactMarkdown 
        remarkPlugins={[remarkGfm]}
        components={{
          code({node, inline, className, children, ...props}) {
            const match = /language-(\w+)/.exec(className || '')
            if (!inline && match && match[1] === 'mermaid') {
              return <div className="mermaid">{String(children).replace(/\n$/, '')}</div>
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
