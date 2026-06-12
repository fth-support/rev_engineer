import React from 'react';
import './Diagrams.css';

const TokenizationFlow = () => {
  return (
    <div className="diagram-wrapper">
      <h1>Tokenization Flow</h1>
      <p className="sub">ServiceTransfer — SP_PRCbToken() & MSSOAP</p>

      <div className="flow-container">
        
        <div className="phase p1">
          <div className="phase-header"><span>1</span> Input Validation</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#78909c' }}>I</div><div>Input <code>CardNumber</code> (e.g. from TPSTSalRC.FTSrcRefNo1 or TPSTSalePoint.FTSpoMemID)</div></div>
            <div className="decision">
              <strong>Check Length:</strong><br />
              IF <code>Len(CardNumber) &lt; 7</code> OR <code>Len(CardNumber) &gt;= 14</code> → <strong>Skip Tokenization</strong> (Return Original)
            </div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#78909c' }}>V</div><div>Verify Tender Type/Trans Type based on configuration rules.</div></div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p2">
          <div className="phase-header"><span>2</span> Local DB Cache Lookup</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#1565c0' }}>Q</div><div><code>SELECT FTTklToken FROM TPSTTokenLst WHERE FTTklCardCode = [CardNumber]</code></div></div>
            <div className="decision">
              <strong>Check Result:</strong><br />
              IF <code>Record Found</code> → <strong>Return Cached Token</strong> (Fast Path)
            </div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p4">
          <div className="phase-header"><span>3</span> SafeNet Web Service (SOAP)</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#e65100' }}>S</div><div>Initialize <code>MSSOAP.SoapClient30</code> (Timeout 10,000ms)</div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#e65100' }}>C</div><div>Call <code>InsertToken(CardNumber, Format="FIRST_SIX_TOKEN")</code></div></div>
            <div className="decision">
              <strong>Check HTTP Response:</strong><br />
              IF <code>HTTP 200 OK</code> AND <code>Token != ""</code> → Proceed to Save<br />
              IF <code>Error/Timeout</code> → Raise <strong>Error 557</strong> (Stop Sync for this transaction)
            </div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p3">
          <div className="phase-header"><span>4</span> Save & Return</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#2e7d32' }}>I</div><div><code>INSERT INTO TPSTTokenLst (FTTklCardCode, FTTklToken) VALUES (...)</code></div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#2e7d32' }}>R</div><div>Return New Token to caller (to be sent to Central Server)</div></div>
          </div>
        </div>

      </div>
    </div>
  );
};

export default TokenizationFlow;
