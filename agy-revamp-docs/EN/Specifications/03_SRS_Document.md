# Software Requirements Specification (SRS) - Executive Summary

## System Overview
The **ServiceTransfer** application is a background data synchronization service designed for a Point of Sale (POS) architecture. Built as a legacy Visual Basic 6 application, it acts as middleware sitting between local POS terminals and a Central Server database. Its primary responsibility is to ensure that offline sales transactions and member point accumulations are reliably transmitted to the central database once network connectivity is restored, and to tokenize sensitive information before central storage.

### Legacy Codebase Structure
The original Visual Basic 6 project structure is outlined below to help the new development team trace the legacy logic:
```text
Legacy_Source
├── ServiceTranfer.vbp      # Main Project File (Compiler settings and library references)
├── ServiceTranfer.vbw      # Workspace File (Editor window states)
├── Forms/
│   ├── wMain.frm           # Main Form (UI) containing the synchronization loop (Timer) and DB updates
│   └── wMain.frx           # Main Form Resources (Images, Icons, Binary data)
└── Modules/
    ├── mMain.bas           # Main Module containing core functions like DB Connection, Online Checking, and Tokenization
    ├── Module1.bas         # Helper Module 1 (Global variables and utilities)
    ├── Module2.bas         # Helper Module 2
    ├── Module3.bas         # Helper Module 3
    └── Module4.bas         # Helper Module 4
```
*Note: The core logic is heavily centralized in `wMain.frm` (flow control) and `mMain.bas` (core functions).*

## Key Features
- **Automated Background Sync:** Uses a timer-based polling mechanism to automatically detect unsynced records and push them to the server.
- **Resilient Data Transfer:** Supports offline mode by tracking the sync status (`FTStaSentOnOff`) of local records and retrying upon connection recovery.
- **Dynamic SQL Generation:** Utilizes ADODB Recordsets to dynamically build `INSERT` and `UPDATE` scripts for various transaction tables (Header, Detail, Receipt, Card, Deposit).
- **Member Point Management:** Aggregates member points locally and updates the central member database (`TCNMMallCard`, `TPSTBPHis`) securely.
- **Data Tokenization:** Automatically replaces sensitive strings (like Member IDs or Card Numbers) with tokens in the central database (`TPSTTokenLst`) to comply with data privacy standards.

## User Roles & Permissions
- **POS Terminal (Actor):** Generates local transaction data and stores it in the local SQL Server Express / MS Access database. Has write access to local tables.
- **ServiceTransfer Agent (System Actor):** The background service itself. Requires read/write permissions to both the Local DB and the Central Server DB.
- **Central Server (Actor):** The master database receiving data. Validates constraints and stores the single source of truth.

## Revamp Recommendations
Based on the code analysis, the following areas are highly recommended for modernization during the system revamp:

1. **Migrate from Polling to Event-Driven Architecture / Message Queues:**
   - *Current State:* The VB6 app uses a timer (`otmForm_Timer`) to constantly poll the local database for unsynced records, which consumes unnecessary CPU and database I/O.
   - *Recommendation:* Replace the polling mechanism with a robust Message Broker (e.g., RabbitMQ, Apache Kafka, or AWS SQS). The POS should publish transactions as events to a queue, and a modern microservice (e.g., Node.js or Go) should consume and persist them to the central server.

2. **Move to RESTful APIs / gRPC instead of Direct DB Connections:**
   - *Current State:* The application establishes direct ODBC/OLEDB connections to the Central Server database over the network, which is a significant security risk and prone to connection timeouts.
   - *Recommendation:* Create a central API Gateway. The local sync agent should send data via secure HTTPS REST or gRPC endpoints. This abstracts the database layer, improves security, and allows for better load balancing.

3. **Modernize the Codebase (Retire VB6):**
   - *Current State:* VB6 is severely outdated, difficult to maintain, and lacks modern testing frameworks and CI/CD compatibility.
   - *Recommendation:* Rewrite the local sync agent in a modern, cross-platform language like Go (compiled into a lightweight binary) or C# (.NET Core Worker Service) to ensure long-term maintainability, better asynchronous execution, and modern ORM support (like Prisma or Entity Framework).
