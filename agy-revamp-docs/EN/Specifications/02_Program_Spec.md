# Chapter Program Specification

This document details the core workflows and internal subroutines of the `ServiceTransfer` system, originally built in Visual Basic 6. Since this is a background service and not an API, the functions are represented as internal modules and event loops.

## Core Functions / "Endpoints"

### 1. `otmForm_Timer` (Main Synchronization Loop)
- **Method (Type):** Internal Timer Event (Background Loop)
- **Request Payload / Parameters:** None (Triggered every interval)
- **Response Payload:** None
- **Core Logic Summary:**
  1. Triggered periodically by the application timer.
  2. Calls `SP_CHKbWebservice` to check if the Central Server database is online.
  3. If online, queries Local DB (like `TPSTSalHD`, `TPSTSalDT`, `TPSTSalRC`) for records where `FTStaSentOnOff` is '0' (Offline/Unsynced).
  4. Iterates through the results and attempts to `INSERT` or `UPDATE` the data into the Central Server DB.
  5. If the server update is successful, it updates the Local DB `FTStaSentOnOff` to '1'.
  6. Calls `W_UPDxUpdatePoint` to process member points.

### 2. `W_UPDxUpdatePoint` (Member Point Processing)
- **Method (Type):** Internal Subroutine (`wMain.frm`)
- **Request Payload / Parameters:** None
- **Response Payload:** None
- **Core Logic Summary:**
  1. Queries the Local DB `TPSTSalePoint` for accumulated points where `FTRemark = '0'` (Unprocessed).
  2. Groups the points by Member ID and Date.
  3. For each group, it connects to the Server DB and queries the current point balance from `TCNMMallCard`.
  4. If the member exists, it updates `TCNMMallCard` adding the new `FCEarned`. If not, it inserts a new member record.
  5. Inserts a point history log into `TPSTBPHis` on the Server DB.
  6. Updates the Local DB `TPSTSalePoint.FTRemark` to '1' indicating the points have been processed.

### 3. `SP_PRCbToken` (Credit Card / ID Tokenization)
- **Method (Type):** Public Function (`mMain.bas`)
- **Request Payload / Parameters:** `ptCardNo` (String)
- **Response Payload:** `ptRefCardNoToken` (String - ByRef)
- **Core Logic Summary:**
  1. Receives a sensitive string (like a Card Number).
  2. Checks the Central Server DB `TPSTTokenLst` to see if a token already exists for this value.
  3. If it exists, returns the existing token.
  4. If it does not exist, generates a new token (or uses a provided one) and inserts the mapping (`FTValue`, `FTToken`) into `TPSTTokenLst`.
  5. Returns the tokenized string to be used in transaction logs instead of the plain text card number.

### 4. `W_DATbInsertData` & `W_DATbUpdateData` (Generic DB Sync)
- **Method (Type):** Internal Functions (`wMain.frm`)
- **Request Payload / Parameters:** `porsTemp` (Recordset), `ptTblName` (String)
- **Response Payload:** Boolean (Success / Failure)
- **Core Logic Summary:**
  1. Helper functions that dynamically construct `INSERT INTO` or `UPDATE` SQL statements based on the Recordset fields.
  2. Used heavily by the main timer loop to replicate data from the Local POS to the Server without hardcoding every single field.
