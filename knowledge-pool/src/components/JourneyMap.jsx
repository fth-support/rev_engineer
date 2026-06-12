import React from 'react';
import { Search, CalendarCheck, Stethoscope, HeartHandshake, Database, RefreshCw, ShieldCheck, UploadCloud } from 'lucide-react';
import './JourneyMap.css';

const JourneyMap = () => {
  return (
    <div className="journey-map-container">
      <div className="journey-header">
        <div className="persona-info glass-panel">
          <h3>SYSTEM PERSONA: SERVICETRANSFER</h3>
          <p><strong>Scenario:</strong> Offline-first POS captures local sales. Background agent must sync data securely and reliably to Central HO and Member DB.</p>
        </div>
        <div className="expectations glass-panel">
          <h3>EXPECTATIONS</h3>
          <ul>
            <li>No missing data or partial transaction uploads</li>
            <li>No sensitive data (Credit Cards/Mem ID) leaving the branch unencrypted</li>
            <li>Low manual intervention (Automatic recovery and sync)</li>
          </ul>
        </div>
      </div>

      <div className="journey-grid">
        {/* Phase 1 */}
        <div className="phase-column phase-1">
          <div className="phase-header">OFFLINE CAPTURE</div>
          <div className="phase-actions">
            <div className="action-card">
              <Database size={32} />
              <p>POS saves sales transaction to Local Express DB</p>
            </div>
            <div className="thought-bubble">Flag FTStaSentOnOff = '0'</div>
          </div>
          <div className="phase-footer">
            <ul>
              <li>High volume of local inserts</li>
              <li>Network might be offline</li>
            </ul>
          </div>
        </div>

        {/* Phase 2 */}
        <div className="phase-column phase-2">
          <div className="phase-header">BACKGROUND SYNC</div>
          <div className="phase-actions">
            <div className="action-card">
              <RefreshCw size={32} />
              <p>Agent polls local DB every 500ms</p>
            </div>
            <div className="thought-bubble">Wait for HD-First Rule to complete</div>
          </div>
          <div className="phase-footer">
            <ul>
              <li>500ms timer consumes high CPU/IO</li>
              <li>Legacy String concatenation SQL</li>
            </ul>
          </div>
        </div>

        {/* Phase 3 */}
        <div className="phase-column phase-3">
          <div className="phase-header">TOKENIZATION</div>
          <div className="phase-actions">
            <div className="action-card">
              <ShieldCheck size={32} />
              <p>Sensitive cards sent to SafeNet SOAP</p>
            </div>
            <div className="thought-bubble">Convert to FIRST_SIX_TOKEN</div>
          </div>
          <div className="phase-footer">
            <ul>
              <li>Plain-text WS credentials</li>
              <li>Dependent on external SOAP service</li>
            </ul>
          </div>
        </div>

        {/* Phase 4 */}
        <div className="phase-column phase-4">
          <div className="phase-header">CENTRAL DELIVERY</div>
          <div className="phase-actions">
            <div className="action-card">
              <UploadCloud size={32} />
              <p>Data pushed to HQ and Member DB</p>
            </div>
            <div className="thought-bubble">Flag FTStaSentOnOff = '1'</div>
          </div>
          <div className="phase-footer">
            <ul>
              <li>Direct remote DB connection via ADO</li>
              <li>No API Gateway abstraction</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
};

export default JourneyMap;
