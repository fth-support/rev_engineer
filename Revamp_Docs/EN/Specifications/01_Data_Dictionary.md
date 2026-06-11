# Data Dictionary

This data dictionary outlines the structure of the database tables inferred from the `ServiceTransfer` system's ADODB operations.

| Table Name | Column Name | Data Type | Constraints (PK/FK/Nullable) | Inferred Description/Label |
| :--- | :--- | :--- | :--- | :--- |
| **TPSTSalHD** | FTTmnNum | VARCHAR | PK | Terminal Number |
| | FTShdTransNo | VARCHAR | PK | Transaction Number |
| | FDShdTransDate | DATE | PK | Transaction Date |
| | FTShdTransType | VARCHAR | Nullable | Transaction Type |
| | FTShdStaDoc | VARCHAR | Nullable | Document Status (1 = Active) |
| | FTShdStaPrcGL | VARCHAR | Nullable | General Ledger Process Status |
| **TPSTSalDT** | FTTmnNum | VARCHAR | PK, FK | Terminal Number |
| | FTShdTransNo | VARCHAR | PK, FK | Transaction Number |
| | FDShdTransDate | DATE | PK, FK | Transaction Date |
| | FNSdtSeqNo | NUMERIC | PK | Detail Sequence Number |
| | FTStaSentOnOff | VARCHAR | Nullable | Online/Offline Sync Status |
| **TPSTSalRC** | FTTmnNum | VARCHAR | PK, FK | Terminal Number |
| | FTShdTransNo | VARCHAR | PK, FK | Transaction Number |
| | FDShdTransDate | DATE | PK, FK | Transaction Date |
| | FNSrcSeqNo | NUMERIC | PK | Receipt Sequence Number |
| | FTStaSentOnOff | VARCHAR | Nullable | Online/Offline Sync Status |
| **TPSTSalCD** | FTTmnNum | VARCHAR | PK, FK | Terminal Number |
| | FTShdTransNo | VARCHAR | PK, FK | Transaction Number |
| | FDShdTransDate | DATE | PK, FK | Transaction Date |
| | FNSdtSeqNo | NUMERIC | PK, FK | Detail Sequence Number |
| | FNScdSeqNo | NUMERIC | PK | Card Sequence Number |
| | FTStaSentOnOff | VARCHAR | Nullable | Online/Offline Sync Status |
| **TPSTSalePoint** | FTTmnNum | VARCHAR | PK, FK | Terminal Number |
| | FTShdTransNo | VARCHAR | PK, FK | Transaction Number |
| | FDShdTransDate | DATE | PK, FK | Transaction Date |
| | FTSpoMemID | VARCHAR | PK | Member ID |
| | FCSpoPoint | NUMERIC | Nullable | Earned Points |
| | FTStaSentOnOff | VARCHAR | Nullable | Online/Offline Sync Status |
| | FTRemark | VARCHAR | Nullable | Remark Flag (0 = Unprocessed, 1 = Processed) |
| | FDDateUpd | DATE | Nullable | Date Updated |
| | FTWhoUpd | VARCHAR | Nullable | Updater User / System Name |
| **TPSTTokenLst** | FTToken | VARCHAR | PK | Token ID |
| | FTValue | VARCHAR | PK | Original Value (Card No, etc.) |
| | FTTblName | VARCHAR | PK | Source Table Name |
| | FTFldName | VARCHAR | PK | Source Field Name |
| | FDDateIns | DATE | Nullable | Date Inserted |
| | FTWhoIns | VARCHAR | Nullable | Who Inserted |
| | FTMcdCode | VARCHAR | Nullable | Mall Card Code |
| | FDMcdExpDate | DATE | Nullable | Mall Card Expiry Date |
| | FTTmnNum | VARCHAR | FK | Terminal Number |
| | FTShdTransNo | VARCHAR | FK | Transaction Number |
| | FDShdTransDate | DATE | FK | Transaction Date |
| | FNStatus | NUMERIC | Nullable | Token Status |
| | FTSrvName | VARCHAR | Nullable | Server Name |
| **TCNMMallCard** | FTMcdCode | VARCHAR | PK | Member Card Code / Token |
| | FDMcdExpDate | DATE | PK | Expiry Date |
| | FTMcdStaAct | VARCHAR | Nullable | Status (Active) |
| | FTMcdName | VARCHAR | Nullable | Member Name |
| | FDMcdBirthday | DATE | Nullable | Member Birthday |
| | FCMcdAmountPoint | NUMERIC | Nullable | Amount Point |
| | FCMcdTotalPoint | NUMERIC | Nullable | Total Point |
| | FCEarned | NUMERIC | Nullable | Earned Points |
| | FCBalance | NUMERIC | Nullable | Balance Points |
| | FDDateUpd | DATE | Nullable | Update Date |
| | FTTimeUpd | VARCHAR | Nullable | Update Time |
| | FTWhoUpd | VARCHAR | Nullable | Update By |
| **TPSTBPHis** | FTTmnNum | VARCHAR | PK, FK | Terminal Number |
| | FTShdTransNo | VARCHAR | PK, FK | Transaction Number |
| | FDShdTransDate | DATE | PK, FK | Transaction Date |
| | FTShdMemberID | VARCHAR | PK | Member ID |
| | FTShdTransTime | VARCHAR | Nullable | Transaction Time |
| | FCCstCardPoint | NUMERIC | Nullable | Custom Card Point |
| | FCTotalBP | NUMERIC | Nullable | Total Point |
| | FCBalanceBP | NUMERIC | Nullable | Balance Point |
| | FDDateUpd | DATE | Nullable | Update Date |
| | FTTimeUpd | VARCHAR | Nullable | Update Time |
| **TSysSync** | FTSscTable | VARCHAR | PK | Source Table to Sync |
| | FTSscDocType | VARCHAR | PK | Document Type |
| | FTSscTableDest | VARCHAR | PK | Destination Table |
| | FTSscStaActive | VARCHAR | Nullable | Sync Active Status |
| **TCNMTerminalMtn** | FTTmnCode | VARCHAR | PK | Terminal Code |
| | FTTmnNum | VARCHAR | Nullable | Terminal Number |
| | FTTmnName | VARCHAR | Nullable | Terminal Name |
