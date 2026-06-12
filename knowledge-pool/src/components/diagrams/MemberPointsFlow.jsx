import React from 'react';
import './Diagrams.css';

const MemberPointsFlow = () => {
  return (
    <div className="diagram-wrapper">
      <h1>Member Points Synchronization Flow</h1>
      <p className="sub">ServiceTransfer — W_UPDxUpdatePoint()</p>

      <div className="flow-container">
        
        <div className="phase p1">
          <div className="phase-header"><span>1</span> Query Local Sale Points</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#78909c' }}>Q</div><div><code>SELECT FTSpoMemID, FCSpoPointQty, FTShdDocType FROM TPSTSalePoint</code></div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#78909c' }}>F</div><div>Filter: <code>WHERE FTRemark = '0'</code> (Not yet calculated)</div></div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p2">
          <div className="phase-header"><span>2</span> Tokenize & Aggregate</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#1565c0' }}>T</div><div>Tokenize <code>FTSpoMemID</code> via SafeNet (if not already tokenized)</div></div>
            <div className="decision">
              <strong>Calculate Points (+ / -):</strong><br />
              IF <code>DocType = '1'</code> (Sale) → Points = <code>+FCSpoPointQty</code><br />
              IF <code>DocType = '5'</code> (Return) → Points = <code>-FCSpoPointQty</code>
            </div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#1565c0' }}>A</div><div>Aggregate total points per Member ID (Token) for the batch.</div></div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p5">
          <div className="phase-header"><span>3</span> Update Member DB (HQ)</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#6a1b9a' }}>S</div><div><code>SELECT FNMcaPointQty FROM TCNMMallCard WHERE FTMcaCode = [Token]</code></div></div>
            <div className="grid2" style={{ margin: '8px 0' }}>
              <div className="success-box"><strong>If Exists:</strong><br /><code>UPDATE TCNMMallCard SET FNMcaPointQty = FNMcaPointQty + NewPoints</code></div>
              <div className="success-box"><strong>If Not Exists:</strong><br /><code>INSERT INTO TCNMMallCard (FTMcaCode, FNMcaPointQty) VALUES ([Token], NewPoints)</code></div>
            </div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p3">
          <div className="phase-header"><span>4</span> Write History</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#2e7d32' }}>H</div><div><code>INSERT INTO TPSTBPHis (FTBphMemID, FCBphPointQty) VALUES ([Token], NewPoints)</code></div></div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p6">
          <div className="phase-header"><span>5</span> Mark as Calculated (Local)</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#c62828' }}>U</div><div><code>UPDATE TPSTSalePoint SET FTRemark = '1'</code> in Local DB</div></div>
          </div>
        </div>

      </div>
    </div>
  );
};

export default MemberPointsFlow;
