// Single source of truth for all interactive diagrams.
// Every fact here was verified against legacy_source/ (wMain.frm, mMain.bas):
//  - Timer Interval = 500ms (wMain.frm:24), otmForm_Timer (:1080)
//  - error codes 555 / 556 / 557 (:316 / :906 / :858)
//  - FTStaSentOnOff flags 0/1/3, HD-first rule, store-open FTTmsOpnStore='Y' (:1659)
//  - SafeNet SOAP tokenization, FIRST_SIX_TOKEN (format "4"), len>=14 to tokenize,
//    nVB_ChkDigitToken=7, TPSTTokenLst fields FTToken/FTValue (mMain.bas:579-1099)
//  - Member points W_UPDxUpdatePoint via ocnMem (:1347)
// Field names use the real schema (e.g. FTMcdCode, FTTdmCode, FCSpoPoint) — NOT the
// approximations that were in the old static diagrams.

// ---------------------------------------------------------------------------
// 1. SYSTEM ARCHITECTURE (node graph)
// x/y are percentages of the canvas; w/h in px.
// Two cooperating VB6 programs on each POS terminal:
//   POSFront (sales producer) → Local DB (handoff) → ServiceTransfer (sync consumer) → HQ.
// Verified against legacy_source/POSFront (mDB.bas Sub Main, Center/mCNSP.bas, cCNDB.cls).
// ---------------------------------------------------------------------------
export const architecture = {
  nodes: [
    {
      id: 'posfront', role: 'branch', x: 15, y: 12, w: 232,
      title: 'POSFront (Sales App)', subtitle: 'VB6 POS · v6.2412.1',
      items: ['Reads master from Server', 'Writes sale → Local DB', 'Tender + print receipt'],
      detail: {
        tag: 'Sales Producer',
        desc: 'The cashier-facing POS application (Adasoft, TAKASHIMAYA project). It reads product / price / member / tender master, records each sale into per-terminal working tables, and on successful payment marks the bill complete (FTShdStaDoc=\'1\') and prints the receipt. Its Sub Main is also the launcher for the 4 background services — including ServiceTransfer.',
        facts: ['Entry point: Sub Main (mDB.bas)', 'Launches ServiceOnOff / AutoReplicate / AutoClear / ServiceTranfer', 'Working tables TPSHD<Tmn> / DT / RC / CD', 'OPOS (Fujitsu) printer + cash drawer'],
      },
    },
    {
      id: 'localdb', role: 'config', x: 15, y: 46, w: 232,
      title: 'Local DB (SQL Express)', subtitle: 'On the POS terminal',
      items: ['TPSHD<Tmn> working tables', 'TRG_Tmp2Sale → STP_PRCxTmp2Sale', 'TPSTSalHD/DT/RC/CD · FTStaSentOnOff'],
      detail: {
        tag: 'Handoff Buffer',
        desc: 'The store-and-forward database that decouples the two programs and makes the terminal offline-capable. A completed bill is promoted from the per-terminal working table into the consolidated TPSTSal* tables by an AFTER-UPDATE trigger, then waits there (FTStaSentOnOff=\'0\') for ServiceTransfer to pick it up.',
        facts: ['Trigger TRG_Tmp2Sale<Tmn> AFTER UPDATE(FTShdStaDoc)', 'STP_PRCxTmp2Sale promotes Tmp → Sale', 'New rows default FTStaSentOnOff=\'0\'', 'Kept in sync with Server by ServiceAutoReplicate.exe'],
      },
    },
    {
      id: 'transfer', role: 'purple', x: 45, y: 84, w: 232,
      title: 'ServiceTransfer (Agent)', subtitle: 'VB6 · Timer 500ms',
      items: ['wMain.frm / mMain.bas', "Polls FTStaSentOnOff='0'", 'Tokenize → push → confirm'],
      detail: {
        tag: 'Sync Consumer',
        desc: 'Background agent (launched by POSFront) that pushes completed sales to HQ. Its 500ms timer loop reads rows flagged FTStaSentOnOff=\'0\' from the Local DB, tokenizes card data, INSERT/UPDATEs the Central DB, posts member points, then flags the local row \'1\'.',
        facts: ['Timer Interval = 500ms (otmForm_Timer)', 'HD-first rule + dup-key → UPDATE', 'Error codes 555 / 556 / 557', 'Logs: LOG\\TransferOffline*.Log'],
      },
    },
    {
      id: 'central', role: 'central', x: 74, y: 14, w: 216,
      title: 'Central Server (HQ)', subtitle: 'SQL Server — Central DB',
      items: ['Master data source', 'TPSTSalHD / DT / RC / CD', 'TSysSync · TCNMTerminalMtn'],
      detail: {
        tag: 'Master + Consolidated Sales',
        desc: 'Single source of master data (products, prices, members, tender, discounts) replicated down to each branch, and the destination that receives every branch\'s consolidated sales.',
        facts: ['Master → branch via ServiceAutoReplicate', 'Sync list: TSysSync WHERE FTSscStaActive=\'1\'', 'Health: SELECT FTTmnCode FROM TCNMTerminalMtn'],
      },
    },
    {
      id: 'member', role: 'member', x: 80, y: 66, w: 200,
      title: 'Member DB', subtitle: 'SQL Server — Loyalty',
      items: ['TCNMMallCard (member cards)', 'TPSTBPHis (point history)', 'TPSTTokenLst (token map)'],
      detail: {
        tag: 'Loyalty / Points DB',
        desc: 'Separate connection (ocnMem). W_UPDxUpdatePoint posts accumulated points to TCNMMallCard and writes an audit row to TPSTBPHis.',
        facts: ['UPDATE FCEarned += points', 'FCBalance = FCMcdTotalPoint + FCEarned + points', 'New member → FTMcdStaAct=\'A\', FTMcdName=\'New Member\''],
      },
    },
    {
      id: 'safenet', role: 'safenet', x: 48, y: 49, w: 200,
      title: 'SafeNet Tokenizer', subtitle: 'SOAP Web Service',
      items: ['MSSOAP.SoapClient30', 'InsertToken() · GetValue()', 'FIRST_SIX_TOKEN'],
      detail: {
        tag: 'External Tokenization',
        desc: 'Card numbers are tokenized here before any sensitive value leaves the branch. Format FIRST_SIX_TOKEN keeps the first 6 digits and randomizes the rest.',
        facts: ['SP_CHKbWebservice → MSSoapInit(URL), timeout 10,000ms', 'InsertToken(CardNo, format=4, luhn=0)', 'Failure raises Error 557'],
      },
    },
  ],
  edges: [
    { from: 'central', to: 'posfront', label: '① replicate master' },
    { from: 'posfront', to: 'localdb', label: '② write sale · StaDoc 2→1' },
    { from: 'localdb', to: 'transfer', label: "③ poll Flag='0'" },
    { from: 'transfer', to: 'safenet', label: '④ InsertToken()' },
    { from: 'transfer', to: 'central', label: '⑤ INSERT/UPDATE sales' },
    { from: 'transfer', to: 'member', label: '⑥ update points' },
    { from: 'transfer', to: 'localdb', label: "⑦ set Flag='1'" },
  ],
}

// ---------------------------------------------------------------------------
// 2. TARGET (REVAMP) ARCHITECTURE (node graph)
// ---------------------------------------------------------------------------
export const targetArchitecture = {
  nodes: [
    {
      id: 'agent', role: 'branch', x: 18, y: 50, w: 210,
      title: 'POS Terminal', subtitle: 'New Agent (Go / .NET)',
      items: ['Event-driven (CDC / tail logs)', 'Local store-and-forward', 'TLS publish'],
      detail: { tag: 'Modern Agent', desc: 'Replaces the 500ms polling loop with change-data-capture / log tailing and publishes events asynchronously over TLS.', facts: ['No cross-WAN DB connection', 'Backpressure-aware queue', 'Single native binary'] },
    },
    {
      id: 'broker', role: 'purple', x: 50, y: 28, w: 190,
      title: 'Message Broker', subtitle: 'RabbitMQ / Kafka',
      items: ['Asynchronous events', 'Durable + replayable'],
      detail: { tag: 'Async Backbone', desc: 'Decouples branches from HQ. Events are durable and replayable, so a brief HQ outage never loses a sale.', facts: ['At-least-once delivery', 'Per-branch partitioning'] },
    },
    {
      id: 'api', role: 'central', x: 80, y: 38, w: 200,
      title: 'API Gateway', subtitle: 'REST / gRPC',
      items: ['Sync Service', 'Point Service', 'Owns DB transactions'],
      detail: { tag: 'HQ Microservices', desc: 'Consumers validate and persist inside BEGIN TRAN…COMMIT, using parameterized queries — eliminating the legacy SQL-injection and partial-write risks.', facts: ['Parameterized SQL', 'DB transactions', 'Central auth + audit'] },
    },
    {
      id: 'data', role: 'member', x: 78, y: 78, w: 210,
      title: 'Data & Security', subtitle: 'Stores + KMS',
      items: ['Central DB · Member DB', 'Modern Token Svc (HSM / Cloud KMS)'],
      detail: { tag: 'Secured Storage', desc: 'Secrets move to a vault; tokenization uses an HSM / Cloud KMS instead of a legacy SOAP service.', facts: ['Secrets in Vault/KMS', 'Encryption in transit + at rest'] },
    },
  ],
  edges: [
    { from: 'agent', to: 'broker', label: 'publish (TLS)' },
    { from: 'broker', to: 'api', label: 'consume' },
    { from: 'api', to: 'data', label: 'transactional writes' },
  ],
}

// ---------------------------------------------------------------------------
// 3. ER DIAGRAM (entity graph) — verified schema field names
// ---------------------------------------------------------------------------
export const erDiagram = {
  entities: [
    {
      id: 'hd', role: 'branch', x: 16, y: 14, title: 'TPSTSalHD', sub: 'Sales Header',
      fields: [
        { name: 'FTTmnNum', type: 'PK' },
        { name: 'FTShdTransNo', type: 'PK' },
        { name: 'FDShdTransDate', type: 'PK' },
        { name: 'FTShdTransType', type: '' },
        { name: 'FTStaSentOnOff', type: '0/1/3' },
        { name: 'FTCstCardCode', type: 'TOKEN' },
        { name: 'FTShdStaDoc', type: '' },
      ],
    },
    {
      id: 'dt', role: 'branch', x: 50, y: 8, title: 'TPSTSalDT', sub: 'Sales Detail',
      fields: [
        { name: 'FTTmnNum + TransNo + Date', type: 'FK' },
        { name: 'FNSdtSeqNo', type: 'PK' },
        { name: 'FTStaSentOnOff', type: '0/1/3' },
      ],
    },
    {
      id: 'rc', role: 'branch', x: 50, y: 33, title: 'TPSTSalRC', sub: 'Receipt / Tender',
      fields: [
        { name: 'FNSrcSeqNo', type: 'PK' },
        { name: 'FTTdmCode', type: 'Tender' },
        { name: 'FTSrcCardNo', type: 'TOKEN' },
      ],
    },
    {
      id: 'cd', role: 'branch', x: 50, y: 58, title: 'TPSTSalCD', sub: 'Card Detail',
      fields: [
        { name: 'FNSdtSeqNo + FNScdSeqNo', type: 'PK' },
        { name: 'FNDctNo', type: 'Discount' },
        { name: 'FTScdCardID', type: 'TOKEN' },
      ],
    },
    {
      id: 'sp', role: 'central', x: 50, y: 82, title: 'TPSTSalePoint', sub: 'Transaction Points',
      fields: [
        { name: 'FTSpoMemID', type: 'TOKEN' },
        { name: 'FCSpoPoint', type: '+/-' },
        { name: 'FTRemark', type: '0/1' },
      ],
    },
    {
      id: 'mall', role: 'member', x: 84, y: 40, title: 'TCNMMallCard', sub: 'Member Master',
      fields: [
        { name: 'FTMcdCode', type: 'PK token' },
        { name: 'FDMcdExpDate', type: 'PK' },
        { name: 'FTMcdStaAct', type: 'A' },
        { name: 'FCMcdTotalPoint', type: '' },
        { name: 'FCEarned / FCBalance', type: '' },
      ],
    },
    {
      id: 'bphis', role: 'member', x: 84, y: 68, title: 'TPSTBPHis', sub: 'Point History',
      fields: [
        { name: 'FTShdMemberID', type: 'PK token' },
        { name: 'FCCstCardPoint', type: '+/-' },
        { name: 'FCTotalBP / FCBalanceBP', type: '' },
      ],
    },
    {
      id: 'token', role: 'safenet', x: 84, y: 12, title: 'TPSTTokenLst', sub: 'Token Mapping',
      fields: [
        { name: 'FTToken', type: 'PK' },
        { name: 'FTValue', type: 'real card' },
        { name: 'FTTblName / FTFldName', type: '' },
        { name: 'FNStatus', type: '1=active' },
      ],
    },
  ],
  edges: [
    { from: 'hd', to: 'dt', label: '1:N' },
    { from: 'hd', to: 'rc', label: '1:N' },
    { from: 'hd', to: 'cd', label: '1:N' },
    { from: 'hd', to: 'sp', label: '1:N' },
    { from: 'sp', to: 'mall', label: 'lookup' },
    { from: 'mall', to: 'bphis', label: '1:N' },
    { from: 'hd', to: 'token', label: 'tokenize' },
  ],
}

// ---------------------------------------------------------------------------
// 4. STEPPERS
// Each step: { title, role, code?, lines:[], branch?:{label,tone} }
// tone: 'exit' | 'error' | 'ok'
// ---------------------------------------------------------------------------

// 4.0 END-TO-END TRANSACTION LIFECYCLE — POSFront (producer) → ServiceTransfer (consumer)
// The two programs never call each other; they hand off through the Local DB via two flags:
//   FTShdStaDoc  '2'→'1'  (owned by POSFront — bill complete)
//   FTStaSentOnOff '0'→'1' (owned by ServiceTransfer — synced to HQ)
export const txLifecycleSteps = [
  {
    title: 'POSFront — Read master', role: 'config', code: 'Sub Main (mDB.bas)',
    lines: [
      'POSFront starts, connects ocnOnLine (Server) + ocnOffLine (Local)',
      'Reads product / price / member / tender / discount master',
      'Online → from Central Server; Offline → from replicated Local copy',
    ],
    branch: { label: 'Sub Main also launches ServiceOnOff / AutoReplicate / AutoClear / ServiceTranfer', tone: 'ok' },
  },
  {
    title: 'POSFront — Open the bill', role: 'branch',
    lines: [
      'Cashier scans items / takes the order',
      "Rows written to per-terminal working tables TPSHD<Tmn> / DT / RC / CD",
      "Header flag FTShdStaDoc = '2' (in-progress, not yet a real sale)",
    ],
  },
  {
    title: 'POSFront — Payment & receipt', role: 'branch', code: 'wTender.frm',
    lines: [
      'Tender complete (cash / card / voucher / points)',
      'Print receipt via OPOS printer + open cash drawer',
      "UPDATE working table SET FTShdStaDoc = '1' (bill complete)",
    ],
    branch: { label: "FTShdStaDoc '2' → '1' is the commit signal", tone: 'ok' },
  },
  {
    title: 'Local DB — Promote to sale', role: 'central', code: 'TRG_Tmp2Sale<Tmn>',
    lines: [
      'Trigger fires AFTER UPDATE(FTShdStaDoc) on TPSHD<Tmn>',
      'EXEC STP_PRCxTmp2Sale — copy completed bill into TPSTSalHD / DT / RC / CD',
      "New consolidated rows default FTStaSentOnOff = '0' (pending sync)",
    ],
  },
  {
    title: 'ServiceTransfer — Detect & tokenize', role: 'safenet', code: 'otmForm_Timer (500ms)',
    lines: [
      "Poll TPSTSal* WHERE FTShdStaDoc='1' AND FTStaSentOnOff='0'",
      'Tokenize card fields via SafeNet (FIRST_SIX_TOKEN)',
      'HD-first rule: send children (DT/RC/CD) before the header',
    ],
    branch: { label: '555 children not ready → skip · 557 token fail → skip+reset', tone: 'error' },
  },
  {
    title: 'ServiceTransfer — Push to HQ', role: 'purple',
    lines: [
      'INSERT / UPDATE into Central DB (dup key → UPDATE)',
      'W_UPDxUpdatePoint() posts accumulated points to Member DB',
    ],
  },
  {
    title: 'ServiceTransfer — Confirm', role: 'member', code: 'W_DATxUpdateByScript()',
    lines: [
      "UPDATE Local SET FTStaSentOnOff = '1' (synced)",
      'Bill is now consolidated at Head Office',
      'Next 500ms cycle picks up the following pending bill',
    ],
    branch: { label: 'Two-flag handshake complete: StaDoc 1 + SentOnOff 1', tone: 'ok' },
  },
]

export const syncFlowSteps = [
  {
    title: 'Pre-check', role: 'config', code: 'otmForm_Timer()',
    lines: ['otmForm.Enabled = False (prevent re-entrancy)', 'Read Registry: SentSale, Online, SentStartDate'],
    branch: { label: "IF SentSale≠'1' OR Online≠'1' → skip to Cleanup", tone: 'exit' },
  },
  {
    title: 'Connect & Validate', role: 'branch',
    lines: [
      'W_DATbConnectOrChk(ocnOnline) — SELECT FTTmnCode FROM TCNMTerminalMtn',
      "W_GETbStoreOpen() — TsysTerminalSta.FTTmsOpnStore = 'Y'",
      'W_DATbConnectOrChk(ocnOffline) — local DB',
    ],
    branch: { label: 'Any check fails → log + skip to Cleanup', tone: 'exit' },
  },
  {
    title: 'Configuration', role: 'central',
    lines: [
      "MM_GETxTable() — TSysSync WHERE FTSscStaActive='1' ORDER BY FNSEQNO",
      'W_GETtTmnNum() — Computer Name → Terminal Number',
      'SP_GETdSaleDate() — TSysChgDateTime.FDCdtDate',
    ],
  },
  {
    title: 'Sync Loop (per table × record)', role: 'member',
    lines: [
      "MM_GETxData() — WHERE FTStaSentOnOff='0' AND FDShdTransDate ≤ SaleDate",
      'W_DATbInsertData(): HD-first pre-check → dup check/DELETE → build SQL → tokenize → INSERT',
      "Success → W_DATxUpdateByScript() sets FTStaSentOnOff='1' (Local)",
    ],
    branch: { label: '-2147217873 dup→UPDATE · -2147467259 conn→abort · 555 children→skip · 557 token→skip+reset', tone: 'error' },
  },
  {
    title: 'Member Points', role: 'purple', code: 'W_UPDxUpdatePoint()',
    lines: ['Post accumulated points to Member DB (see Member Points flow)'],
  },
  {
    title: 'Cleanup', role: 'safenet',
    lines: ["Save Registry (SentStartDate/Time, SentStaOn='1')", 'otmForm.Enabled = True — next cycle in 500ms'],
  },
]

export const tokenizationSteps = [
  {
    title: 'Guard clauses', role: 'config', code: 'SP_PRCbToken()',
    lines: ['IF NOT bVB_AlwToken (Store Para FTStpCode=29) → return original', 'IF Len(CardNo) < 7 (nVB_ChkDigitToken) → return original', 'IF Len(CardNo) < 14 OR empty → return original'],
    branch: { label: 'Tokenize only when enabled AND Len ≥ 14', tone: 'exit' },
  },
  {
    title: 'DB cache lookup', role: 'branch',
    lines: [
      "IF Online='1' AND NOT pbNoSearchDB:",
      'SELECT TOP 1 FTToken FROM TPSTTokenLst WHERE FTValue=CardNo AND FTValue≠FTToken ORDER BY FDDateIns DESC',
    ],
    branch: { label: 'Found → return cached token (fast path)', tone: 'ok' },
  },
  {
    title: 'Init SOAP client', role: 'safenet', code: 'SP_CHKbWebservice()',
    lines: ['CreateObject("MSSOAP.SoapClient30")', 'MSSoapInit(tVB_TokenURL)', 'Timeout = ConnectTimeout = 10,000ms'],
  },
  {
    title: 'Call InsertToken', role: 'safenet',
    lines: ['oToken.InsertToken(NAEUsr, NAEPwd, DBUsr, DBPwd, DB, CardNo, format=4, luhn=0)', 'format 4 = FIRST_SIX_TOKEN'],
    branch: { label: 'Err ≠ 0 OR empty token → return False (Error 557)', tone: 'error' },
  },
  {
    title: 'Persist mapping', role: 'central', code: 'SP_PRCbINSTokenList()',
    lines: ['INSERT INTO TPSTTokenLst (FDDateIns, FTWhoIns, FTTblName, FTFldName, FTTmnNum, FTShdTransNo, FDShdTransDate, FTValue, FTToken, FNStatus, FTSrvName)', 'Return new token to caller'],
  },
]

export const memberPointsSteps = [
  {
    title: 'Fetch pending points', role: 'config', code: 'W_UPDxUpdatePoint()',
    lines: ["SELECT DISTINCT FTTmnNum, FTShdTransNo, FTShdTransType, FDShdTransDate", "FROM TPSTSalePoint WHERE FTRemark='0' AND FTSpoMemID ≠ ''"],
  },
  {
    title: 'Aggregate per member', role: 'branch',
    lines: ['SUM(FCSpoPoint) GROUP BY FTSpoMemID, FDDateUpd, FTWhoUpd'],
  },
  {
    title: 'Tokenize member ID', role: 'safenet',
    lines: ['SP_PRCbToken("TPSTSalePoint", FTSpoMemID) → token', 'Write mapping to Member DB TPSTTokenLst'],
  },
  {
    title: 'Calculate +/- by TransType', role: 'member',
    lines: ['Sale(03) · VIPSale(10) · Deposit(06) · AddDeposit(19) → +Point', 'Return(04) · IPV(05) · DepositCancel(16) · CancelVoucher(21) → −Point'],
  },
  {
    title: 'Update / Insert MallCard', role: 'central',
    lines: [
      'Found: UPDATE TCNMMallCard SET FCEarned += pts, FCBalance = FCMcdTotalPoint + FCEarned + pts',
      "Not found: INSERT (FTMcdStaAct='A', FTMcdName='New Member', FCEarned=pts, FCBalance=pts)",
    ],
  },
  {
    title: 'History & mark done', role: 'purple',
    lines: ['INSERT INTO TPSTBPHis (member, point, totals)', "UPDATE TPSTSalePoint SET FTRemark='1' (Local)"],
  },
]

// Map a role id -> the CSS variable base used by tokens.js
export const ROLE_VARS = {
  branch: 'role-branch',
  central: 'role-central',
  member: 'role-member',
  safenet: 'role-safenet',
  token: 'role-token',
  config: 'role-config',
  purple: 'role-purple',
}
