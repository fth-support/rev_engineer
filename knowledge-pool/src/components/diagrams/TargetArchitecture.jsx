import React from 'react';
import './Diagrams.css';

const TargetArchitecture = () => {
  return (
    <div className="diagram-wrapper">
      <h1>Revamp Target Architecture</h1>
      <p className="sub">Modern Event-Driven & Microservices Approach</p>

      <div className="diagram-grid" style={{ gridTemplateColumns: '1fr 60px 1fr 60px 1fr', gridTemplateRows: 'auto auto auto', gap: '15px' }}>
        
        {/* Branch */}
        <div className="diagram-box branch" style={{ gridColumn: 1, gridRow: '1/4' }}>
          <div className="box-title">POS Terminal (Branch)</div>
          <div style={{ fontSize: '11px', color: '#888', marginBottom: '4px' }}>Local Storage</div>
          <ul className="db-list">
            <li>Local DB <span>(Sales/Points)</span></li>
          </ul>
          <div style={{ marginTop: '15px' }}>
            <h3 style={{ margin: '10px 0 5px 0' }}>New Agent (Go/.NET)</h3>
            <div className="config-item"><div className="config-icon" style={{ background: '#78909c' }}>E</div> Event-Driven</div>
            <div className="config-item"><div className="config-icon" style={{ background: '#5c6bc0' }}>C</div> Tail Logs / CDC</div>
          </div>
        </div>

        {/* Arrow to MQ */}
        <div className="diagram-arrow" style={{ gridColumn: 2, gridRow: 2 }}>
          →<br /><span className="arrow-label">Publish TLS</span>
        </div>

        {/* Message Broker */}
        <div className="diagram-box" style={{ border: '2px solid #8e24aa', gridColumn: 3, gridRow: 2 }}>
          <div className="box-title" style={{ background: '#8e24aa' }}>Message Broker</div>
          <h3 style={{ textAlign: 'center', margin: '10px 0' }}>RabbitMQ / Kafka</h3>
          <p style={{ fontSize: '11px', textAlign: 'center', color: '#666' }}>Asynchronous Events</p>
        </div>

        {/* Arrow to HQ */}
        <div className="diagram-arrow" style={{ gridColumn: 4, gridRow: 2 }}>
          →<br /><span className="arrow-label">Consume</span>
        </div>

        {/* HQ Microservices */}
        <div className="diagram-box hq" style={{ gridColumn: 5, gridRow: '1/3' }}>
          <div className="box-title">HQ Microservices</div>
          <h3 style={{ textAlign: 'center', marginBottom: '15px' }}>API Gateway (REST/gRPC)</h3>
          
          <div style={{ display: 'flex', gap: '10px' }}>
            <div style={{ background: '#e8f5e9', padding: '10px', borderRadius: '6px', flex: 1, border: '1px solid #c8e6c9', textAlign: 'center' }}>
              <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#2e7d32' }}>Sync Service</span>
            </div>
            <div style={{ background: '#e8f5e9', padding: '10px', borderRadius: '6px', flex: 1, border: '1px solid #c8e6c9', textAlign: 'center' }}>
              <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#2e7d32' }}>Point Service</span>
            </div>
          </div>
        </div>

        {/* Databases & Security */}
        <div className="diagram-box ext" style={{ gridColumn: 5, gridRow: 3, marginTop: '20px' }}>
          <div className="box-title">Databases & Security</div>
          
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
            <ul className="db-list">
              <li>Central DB <span>(Sales)</span></li>
              <li>Member DB <span>(Points)</span></li>
            </ul>
            <div style={{ background: '#fff3e0', padding: '10px', borderRadius: '6px', border: '1px solid #ffe0b2', textAlign: 'center' }}>
              <span style={{ fontSize: '11px', fontWeight: 'bold', color: '#e65100' }}>Modern Token Svc<br/>HSM / Cloud KMS</span>
            </div>
          </div>
        </div>

      </div>
    </div>
  );
};

export default TargetArchitecture;
