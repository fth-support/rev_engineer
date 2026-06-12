import React from 'react';
import './Diagrams.css';

const SystemArchitecture = () => {
  return (
    <div className="diagram-wrapper">
      <h1>System Architecture</h1>
      <p className="sub">ServiceTransfer v5.1810.1 — Deployment & Component Overview</p>

      <div className="diagram-grid">
        {/* Branch POS */}
        <div className="diagram-box branch" style={{ gridColumn: 1, gridRow: '1/3' }}>
          <div className="box-title">POS Terminal (Branch)</div>
          <h3>ServiceTransfer.exe <span className="timer-badge">Timer: 500ms</span></h3>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', marginBottom: '12px' }}>
            <div>
              <div style={{ fontSize: '11px', color: '#888', marginBottom: '4px' }}>Core Files</div>
              <ul className="db-list">
                <li>wMain.frm <span>1,763 lines</span></li>
                <li>mMain.bas <span>1,099 lines</span></li>
                <li>Module4.bas <span>SleepEx</span></li>
              </ul>
            </div>
            <div>
              <div style={{ fontSize: '11px', color: '#888', marginBottom: '4px' }}>Dependencies</div>
              <ul className="db-list">
                <li>AdaUT.dll <span>Utility</span></li>
                <li>ADODB 2.7 <span>DB</span></li>
                <li>MSSOAP 3.0 <span>Token</span></li>
              </ul>
            </div>
          </div>
          <div style={{ fontSize: '11px', color: '#888', marginBottom: '4px' }}>SQL Server Express (Local DB)</div>
          <ul className="db-list">
            <li>TPSTSalHD / DT / RC / CD <span>Sales</span></li>
            <li>TPSTSalePoint <span>Points</span></li>
            <li>TPSTDepositHD / DT / RC / CD <span>Deposit</span></li>
            <li>TPSTHoldHD / DT / CD <span>Hold</span></li>
            <li>TPSTVoucherHD / RC <span>Voucher</span></li>
            <li>TPSTCV / TPSTLogDT <span>CV & Log</span></li>
          </ul>
          <div style={{ marginTop: '10px' }}>
            <div style={{ fontSize: '11px', color: '#888', marginBottom: '4px' }}>Local Configuration</div>
            <div className="config-item"><div className="config-icon" style={{ background: '#78909c' }}>A</div> AdaIni.Ada — Access DB</div>
            <div className="config-item"><div className="config-icon" style={{ background: '#5c6bc0' }}>R</div> Registry: AdaPosMall\POSFront</div>
            <div className="config-item"><div className="config-icon" style={{ background: '#26a69a' }}>L</div> LOG\TransferOffline*.Log</div>
          </div>
        </div>

        {/* Arrow */}
        <div className="diagram-arrow" style={{ gridColumn: 2, gridRow: '1/3' }}>
          ⇄<br /><span className="arrow-label">SQLOLEDB.1<br />LAN/WAN</span>
        </div>

        {/* Central Server */}
        <div className="diagram-box hq" style={{ gridColumn: 3, gridRow: 1 }}>
          <div className="box-title">Central Server (HQ)</div>
          <div style={{ fontSize: '11px', color: '#888', marginBottom: '4px' }}>SQL Server (Central DB)</div>
          <ul className="db-list">
            <li>TPSTSal* <span>Sales (copy)</span></li>
            <li>TSysSync <span>Sync Config</span></li>
            <li>TSysConfig <span>System Config</span></li>
            <li>TCNMTerminalMtn <span>Terminal Map</span></li>
            <li>TCNMStorePara <span>Store Params</span></li>
          </ul>
          <div style={{ marginTop: '10px', fontSize: '11px', color: '#888', marginBottom: '4px' }}>SQL Server (Member DB)</div>
          <ul className="db-list">
            <li>TCNMMallCard <span>Member Cards</span></li>
            <li>TPSTBPHis <span>Point History</span></li>
            <li>TPSTTokenLst <span>Token Map</span></li>
          </ul>
        </div>

        {/* SafeNet */}
        <div className="diagram-box ext" style={{ gridColumn: 3, gridRow: 2 }}>
          <div className="box-title">SafeNet Tokenizer</div>
          <div style={{ fontSize: '12px', color: '#333' }}>
            <div className="config-item"><div className="config-icon" style={{ background: '#e65100' }}>S</div> MSSOAP.SoapClient30 (SOAP)</div>
            <div className="config-item"><strong>InsertToken()</strong> — Card → Token</div>
            <div className="config-item"><strong>GetValue()</strong> — Token → Card</div>
            <div className="config-item">Format: FIRST_SIX_TOKEN</div>
            <div className="config-item">Timeout: 10,000ms</div>
          </div>
        </div>
      </div>

      <div className="diagram-section">
        <h2>Main Processing Flow (otmForm_Timer — every 500ms)</h2>
        <div className="flow-step"><div className="step-num">1</div><div className="step-text"><strong>Pre-check:</strong> Disable timer, verify Log file, read Registry</div></div>
        <div className="flow-step"><div className="step-num">2</div><div className="step-text"><strong>Connect:</strong> W_DATbConnectOrChk() to Central Server</div></div>
        <div className="flow-step"><div className="step-num">3</div><div className="step-text"><strong>Store Check:</strong> W_GETbStoreOpen() verifies TsysTerminalSta.FTTmsOpnStore = 'Y'</div></div>
        <div className="flow-step"><div className="step-num">4</div><div className="step-text"><strong>Config:</strong> Fetch TSysSync table list, Terminal Number, Sale Date</div></div>
        <div className="flow-step"><div className="step-num">5</div><div className="step-text"><strong>Sync Loop:</strong> For each table → INSERT/UPDATE to Server with tokenization</div></div>
        <div className="flow-step"><div className="step-num">6</div><div className="step-text"><strong>Points:</strong> W_UPDxUpdatePoint() — UPDATE/INSERT TCNMMallCard, log to TPSTBPHis</div></div>
        <div className="flow-step"><div className="step-num">7</div><div className="step-text"><strong>Cleanup:</strong> Save Registry state, re-enable timer</div></div>
      </div>
    </div>
  );
};

export default SystemArchitecture;
