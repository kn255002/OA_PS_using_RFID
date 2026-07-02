# Online Attendance & Payroll System via RFID (OA&PS RFID)
University Thesis Project - Class MCS-15 (Year 2006-2007)

**Project Group Partners:**

Khurram Nazir, Muhammad Abid (Islamabad, Pakistan)
---
A legacy multi-tiered enterprise automation system engineered to replace manual register-based workflows with automated, real-time Radio Frequency Identification (RFID) tracking and localized relational payroll computation. 

---

## Architectural Overview (Legacy Baseline)
This project was originally developed as an academic thesis system designed for **MIA Enterprises**, bridging physical sensor hardware with automated business logic to handle time-tracking, multi-tiered employee categorization, and automated payroll operations.

### Key Capabilities Included:
*   **Hardware Sensor Integration:** Low-Frequency (125-135 KHz) RFID tag processing paired with active serial antenna data collection interfaces.
*   **Asynchronous Fault Tolerance:** Hardware-level caching capable of buffering up to 8,000 physical tag entries directly on the reader during corporate network or power outages.
*   **Granular Payroll Engine:** Rule-based localized evaluation of complex compensation structures (Contractual vs. Permanent salary scales, overtime sit-ins, and progressive tax tiers under regional regulations).

---

## Legacy Technology Stack

| Component | Technology / Tooling | Purpose |
| :--- | :--- | :--- |
| **Frontend UI** | Windows Forms (C# .NET Framework 2.0) | Desktop environments for DEO, Scanners, and Administration. |
| **Web Portal** | 

 (C#) | Distributed online employee access for pay-slip and attendance audits. |
| **Database** | Oracle 9i RDBMS | Central relational data repository, data integrity control, and stored procedures. |
| **System Modeling**| MS Visio / UML 2.0 | Entity-Relationship (ERD) mapping and architectural schema design. |
| **OS Target** | Windows XP Professional | Desktop deployment target for physical client nodes. |

---

## Legacy System Topology
The system abstracts operations across five isolated operational roles managed via a single unified enterprise database backend:
<img width="627" height="508" alt="image" src="https://github.com/user-attachments/assets/10401aea-2342-4de1-8be7-e4c8972765cf" />
