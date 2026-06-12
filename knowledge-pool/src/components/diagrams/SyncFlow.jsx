import React from 'react';
import './Diagrams.css';

const SyncFlow = () => {
  return (
    <div className="diagram-wrapper">
      <h1>Data Synchronization Flow</h1>
      <p className="sub">ServiceTransfer — otmForm_Timer() / W_DATbInsertData() / W_DATbUpdateData()</p>

      <div className="flow-container">
        
        <div className="phase p1">
          <div className="phase-header"><span>1</span> Pre-check Phase</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#78909c' }}>T</div><div><code>otmForm.Enabled = False</code> — Disable timer to prevent re-entrant execution</div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#78909c' }}>L</div><div>Check/create log file, read Registry (<code>SentSale</code>, <code>SentStartDate</code>)</div></div>
            <div className="decision"><strong>Decision:</strong> IF <code>SentSale != "1"</code> OR <code>Registry("Online") != "1"</code> → Skip to Cleanup</div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p2">
          <div className="phase-header"><span>2</span> Connection & Validation Phase</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#1565c0' }}>S</div><div><code>W_DATbConnectOrChk(ocnOnline)</code> — Health check: <code>SELECT FTTmnCode FROM TCNMTerminalMtn</code></div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#1565c0' }}>O</div><div><code>W_GETbStoreOpen()</code> — Check <code>TsysTerminalSta.FTTmsOpnStore = 'Y'</code></div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#1565c0' }}>L</div><div><code>W_DATbConnectOrChk(ocnOffline)</code> — Verify local DB connection</div></div>
            <div className="decision"><strong>Decision:</strong> Any check fails → Log error, skip to Cleanup</div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p3">
          <div className="phase-header"><span>3</span> Configuration Phase</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#2e7d32' }}>T</div><div><code>MM_GETxTable()</code> — Fetch sync table list from <code>TSysSync WHERE FTSscStaActive='1'</code></div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#2e7d32' }}>#</div><div><code>W_GETtTmnNum()</code> — Map Computer Name → Terminal Number via <code>TCNMTerminalMtn</code></div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#2e7d32' }}>D</div><div><code>SP_GETdSaleDate()</code> — Get operation date from <code>TSysChgDateTime.FDCdtDate</code></div></div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p4">
          <div className="phase-header"><span>4</span> Main Sync Loop (FOR EACH Table × Record)</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#e65100' }}>Q</div><div><code>MM_GETxData()</code> — <code>SELECT * FROM [Table] WHERE FTStaSentOnOff='0' AND FDShdTransDate &lt;= SaleDate</code></div></div>
            
            <div style={{ margin: '12px 0', padding: '12px', background: '#fafafa', borderRadius: '8px', border: '1px solid #e0e0e0', color: '#333' }}>
              <div style={{ fontWeight: 700, fontSize: '13px', marginBottom: '8px', color: '#e65100' }}>For Each Record → W_DATbInsertData()</div>
              
              <div className="phase-step"><div className="step-icon" style={{ background: '#ff8f00' }}>1</div><div><strong>Pre-check (HD only):</strong> Verify all child tables (DT/RC/CD/SalePoint) have <code>FTStaSentOnOff='1'</code></div></div>
              <div className="phase-step"><div className="step-icon" style={{ background: '#ff8f00' }}>2</div><div><strong>Duplicate check:</strong> SELECT from Server → if exists & <code>FTShdStaDoc != '1'</code> → DELETE old</div></div>
              <div className="phase-step"><div className="step-icon" style={{ background: '#ff8f00' }}>3</div><div><strong>Build SQL:</strong> Loop all fields → FT=Text (escape quotes), FC/FN=Numeric, FD=Date format</div></div>
              <div className="phase-step"><div className="step-icon" style={{ background: '#ff8f00' }}>4</div><div><strong>Tokenize:</strong> Check field × table × tender × transtype rules → <code>SP_PRCbToken()</code></div></div>
              <div className="phase-step"><div className="step-icon" style={{ background: '#ff8f00' }}>5</div><div><strong>Execute:</strong> <code>INSERT INTO [Table](...) VALUES(...)</code> on Central Server</div></div>
              
              <div className="grid2" style={{ marginTop: '10px' }}>
                <div className="success-box"><strong>Success →</strong> <code>W_DATxUpdateByScript()</code> — Set <code>FTStaSentOnOff='1'</code> in Local DB</div>
                <div className="error-box"><strong>Error Handling:</strong><br />
                  <code>-2147217873</code> Duplicate → try UPDATE<br />
                  <code>-2147467259</code> Connection lost → abort<br />
                  <code>555</code> Children not ready → skip<br />
                  <code>557</code> Token failed → skip + reset
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p5">
          <div className="phase-header"><span>5</span> Member Points Phase (W_UPDxUpdatePoint)</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#6a1b9a' }}>1</div><div>Fetch <code>TPSTSalePoint WHERE FTRemark='0' AND FTSpoMemID &lt;&gt; ''</code></div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#6a1b9a' }}>2</div><div>SUM points GROUP BY member → Tokenize member ID</div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#6a1b9a' }}>3</div><div>Calculate +/- by TransType (Sale=+, Return/IPV/Cancel=-)</div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#6a1b9a' }}>4</div><div>UPDATE/INSERT <code>TCNMMallCard</code> on Member DB</div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#6a1b9a' }}>5</div><div>INSERT history to <code>TPSTBPHis</code>, mark <code>FTRemark='1'</code> in Local</div></div>
          </div>
        </div>

        <div className="arrow-down">▼</div>

        <div className="phase p6">
          <div className="phase-header"><span>6</span> Cleanup</div>
          <div className="phase-body">
            <div className="phase-step"><div className="step-icon" style={{ background: '#c62828' }}>R</div><div>Save Registry state (<code>SentStartDate</code>, <code>SentStartTime</code>, <code>SentStaOn='1'</code>)</div></div>
            <div className="phase-step"><div className="step-icon" style={{ background: '#c62828' }}>T</div><div><code>otmForm.Enabled = True</code> — Re-enable timer for next cycle</div></div>
          </div>
        </div>

      </div>
    </div>
  );
};

export default SyncFlow;
