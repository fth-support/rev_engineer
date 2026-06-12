import React from 'react';
import './Diagrams.css';

const ERDiagram = () => {
  return (
    <div className="diagram-wrapper">
      <h1>Data Dictionary & ER Diagram</h1>
      <p className="sub">ServiceTransfer — Core Tables for Sales and Tokenization</p>
      
      {/* Inline styles specific to ER Diagram boxes not covered by global styles */}
      <style>{`
        .er-grid { display: grid; grid-template-columns: 1fr 100px 1fr; gap: 15px; align-items: start; max-width: 900px; margin: 0 auto; }
        .table-box { background: #fff; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); border: 1px solid #e0e0e0; overflow: hidden; }
        .tb-head { background: #1565c0; color: #fff; padding: 8px 12px; font-weight: 700; font-size: 14px; display: flex; justify-content: space-between; }
        .tb-head-member { background: #2e7d32; }
        .tb-head-token { background: #e65100; }
        .col-list { list-style: none; padding: 0; margin: 0; font-size: 12px; }
        .col-item { padding: 6px 12px; border-bottom: 1px solid #f0f0f0; display: flex; justify-content: space-between; }
        .col-item:last-child { border-bottom: none; }
        .col-name { font-weight: 600; color: #333; }
        .col-type { color: #888; font-family: monospace; font-size: 11px; }
        .pk { color: #d32f2f; font-weight: 700; margin-right: 4px; }
        .fk { color: #1976d2; font-weight: 700; margin-right: 4px; }
        .rel { text-align: center; font-size: 12px; color: #666; display: flex; flex-direction: column; justify-content: center; height: 100%; padding-top: 40px; }
      `}</style>

      <div className="er-grid">
        
        {/* Left Column: Sales Data */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
          
          <div className="table-box">
            <div className="tb-head"><span>TPSTSalHD</span><span>Sales Header</span></div>
            <ul className="col-list">
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>FTBchCode</span><span className="col-type">Varchar(5)</span></li>
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>FTTmnNum</span><span className="col-type">Varchar(3)</span></li>
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>FTShdTransNo</span><span className="col-type">Varchar(11)</span></li>
              <li className="col-item"><span className="col-name">FDShdTransDate</span><span className="col-type">Datetime</span></li>
              <li className="col-item"><span className="col-name">FTShdDocType</span><span className="col-type">Varchar(1) 1=Sale, 5=Return</span></li>
              <li className="col-item"><span className="col-name">FCCrdAmt</span><span className="col-type">Numeric(18,4)</span></li>
              <li className="col-item"><span className="col-name">FTStaSentOnOff</span><span className="col-type">Varchar(1) 0=Wait, 1=Sent</span></li>
            </ul>
          </div>

          <div className="table-box">
            <div className="tb-head"><span>TPSTSalRC</span><span>Receipt / Tender</span></div>
            <ul className="col-list">
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>... Keys from HD ...</span></li>
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>FNSrcSeqNo</span><span className="col-type">Int</span></li>
              <li className="col-item"><span className="col-name">FTTndCode</span><span className="col-type">Varchar(5) (e.g., CR01)</span></li>
              <li className="col-item"><span className="col-name">FTSrcRefNo1</span><span className="col-type">Varchar(20) *Tokenized</span></li>
              <li className="col-item"><span className="col-name">FTSrcRefNo2</span><span className="col-type">Varchar(20) *Appv Code</span></li>
            </ul>
          </div>

        </div>

        {/* Middle Column: Relations */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '50px' }}>
          <div className="rel">
            <span style={{ fontSize: '20px' }}>⟷</span>
            <span>1 to N</span>
            <span style={{ fontSize: '10px', color: '#999' }}>(Join by TmnNum + TransNo)</span>
          </div>
          <div className="rel" style={{ marginTop: '50px' }}>
            <span style={{ fontSize: '20px' }}>⟷</span>
            <span>Tokenize</span>
          </div>
        </div>

        {/* Right Column: Master / External Data */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>

          <div className="table-box">
            <div className="tb-head tb-head-member"><span>TCNMMallCard</span><span>Member Master</span></div>
            <ul className="col-list">
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>FTMcaCode</span><span className="col-type">Varchar(20) *Token</span></li>
              <li className="col-item"><span className="col-name">FTMcaName</span><span className="col-type">Varchar(100)</span></li>
              <li className="col-item"><span className="col-name">FNMcaPointQty</span><span className="col-type">Numeric(18,4)</span></li>
              <li className="col-item"><span className="col-name">FDMcaExpDate</span><span className="col-type">Datetime</span></li>
            </ul>
          </div>

          <div className="table-box">
            <div className="tb-head tb-head-member"><span>TPSTSalePoint</span><span>Transaction Points</span></div>
            <ul className="col-list">
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>... Keys from HD ...</span></li>
              <li className="col-item"><span className="col-name">FTSpoMemID</span><span className="col-type">Varchar(20) *Token</span></li>
              <li className="col-item"><span className="col-name">FCSpoPointQty</span><span className="col-type">Numeric(18,4)</span></li>
              <li className="col-item"><span className="col-name">FTRemark</span><span className="col-type">Varchar(1) 0=Wait, 1=Calc</span></li>
            </ul>
          </div>

          <div className="table-box">
            <div className="tb-head tb-head-token"><span>TPSTTokenLst</span><span>Token Mapping</span></div>
            <ul className="col-list">
              <li className="col-item"><span className="col-name"><span className="pk">PK</span>FTTklCardCode</span><span className="col-type">Varchar(20) Real Card</span></li>
              <li className="col-item"><span className="col-name">FTTklToken</span><span className="col-type">Varchar(20) Tokenized</span></li>
              <li className="col-item"><span className="col-name">FDTklUpdDate</span><span className="col-type">Datetime</span></li>
            </ul>
          </div>

        </div>

      </div>
    </div>
  );
};

export default ERDiagram;
