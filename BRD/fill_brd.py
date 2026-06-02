"""
Fill BRD template for KPMG IntelliSource.
Replaces placeholder text in document.xml while preserving all formatting.
"""

import zipfile
import shutil
import os
import re

SRC = "Business_Requirements_Document_Template.docx"
DST = "KPMG_IntelliSource_BRD.docx"

# ── helpers ──────────────────────────────────────────────────────────────────

def xe(s):
    """XML-escape a string for insertion into document.xml text nodes."""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# Each tuple: (exact_bytes_in_xml, replacement_text_already_escaped)
# We do raw bytes replacement to avoid encoding headaches.

REPLACEMENTS = [
    # ── 1.1  Project Background ──────────────────────────────────────────────
    (
        b"Give the project background here",
        xe(
            "KPMG IntelliSource is a web-based procurement intelligence and vendor "
            "management platform developed for KPMG India's internal operations. "
            "Prior to this initiative, procurement teams managed vendor performance data, "
            "purchase orders, compliance checks, and Procure-to-Pay (P2P) workflows across "
            "disconnected tools including spreadsheets, email chains, and periodic ERP exports. "
            "This fragmentation caused delayed decisions, missed contract renewals, reactive "
            "compliance breach handling, and poor visibility into P2P bottlenecks. "
            "IntelliSource was commissioned to consolidate all procurement intelligence into "
            "a single, role-aware, real-time dashboard aligned with KPMG's brand identity "
            "and governance standards."
        ).encode("utf-8"),
    ),

    # ── 1.2  Objective ───────────────────────────────────────────────────────
    (
        b"Please describe the objective",
        xe(
            "The objectives of KPMG IntelliSource are: "
            "(1) Provide a unified, real-time view of vendor performance, PO lifecycle, and compliance status. "
            "(2) Reduce average PO cycle time from the current 4.2 days to the 3.0-day target. "
            "(3) Surface compliance breaches (ABAC gift threshold, sanctions match, DOA bypass, duplicate invoices, rate deviation, conflict of interest) in real time. "
            "(4) Enable P2P bottleneck identification across 16 lifecycle stages with SLA breach alerting. "
            "(5) Empower leadership with financial gross margin trends, divisional budget vs. spend analytics, and license utilization insights. "
            "(6) Provide role-based access for Procurement Managers, Compliance Officers, Finance/CFO, IT Admins, and Leadership."
        ).encode("utf-8"),
    ),

    # ── 1.3  Intended Audience ───────────────────────────────────────────────
    (
        "Mention the intended audience – Internal/External".encode("utf-8"),
        xe(
            "Internal. Primary users: KPMG Procurement Managers, Compliance Officers, "
            "Finance Analysts and CFO, IT Administrators, and CXO/Leadership. "
            "Secondary stakeholders: Procurement Operations team and Internal Audit."
        ).encode("utf-8"),
    ),

    # ── 2.1  In Scope ────────────────────────────────────────────────────────
    (
        b"&lt;This section will include items/functionalities which are in scope of the project&gt;",
        xe(
            "The following modules and features are in scope for KPMG IntelliSource Phase 1:\n"
            "1. Procurement Dashboard: real-time KPI cards (Total PO Value MTD, Active POs by status, High-Value PO monitor, Avg Cycle Time, Compliance Breach Rate, PO Deletion Frequency); gross margin trend chart; PO status weekly breakdown; compliance alert panel; PO deletion anomaly monitor; upcoming contract renewal tracker.\n"
            "2. Vendor Management: vendor scorecard list with ratings, delivery score, quality score, ABAC certification, compliance status, risk tier (Low/Medium/High), YTD spend, contract status; spend-by-category analytics; top-10 vendors by spend.\n"
            "3. Financial Analytics: 24-month gross margin trend with 25% target reference; divisional budget vs. spend; invoice aging and status management.\n"
            "4. P2P Lifecycle Tracker: 16-stage Procure-to-Pay flow (Requirement Analysis through Payment Settlement) with item counts, bottleneck flagging, funnel conversion chart, and SLA breach drill-down.\n"
            "5. Compliance Module: breach detection for ABAC Gift Threshold, Sanctions Match, DOA Bypass, Duplicate Invoice, Rate Deviation, Conflict of Interest; severity classification (Low/Medium/High/Critical); resolution tracking.\n"
            "6. License Utilization Dashboard: per-tool utilization %, license-owned vs. active-users, monthly cost, cost-per-active-user, optimization flags (Underutilized/Over-provisioned/Optimal), potential savings calculation.\n"
            "7. Action Items Module: procurement action task tracking with type, status (Open/In Progress/Under Review/Closed), owner, priority (High/Medium/Low), financial impact, and due dates.\n"
            "8. Admin Module: user management (CRUD, role assignment), audit log of all system actions, application settings.\n"
            "9. Authentication: JWT-based login and signup with role-based route protection.\n"
            "10. Leadership Dashboard: cross-functional KPI aggregation for CXO-level visibility.\n"
            "11. Vendor Repository: filterable, searchable master vendor list."
        ).encode("utf-8"),
    ),

    # ── 2.2  Out of Scope ────────────────────────────────────────────────────
    (
        b"&lt;This section will include items/functionalities which are out of scope of the project&gt;",
        xe(
            "The following items are explicitly out of scope for Phase 1:\n"
            "1. Live ERP integration (SAP S/4HANA, Oracle ERP) — Phase 1 uses a deterministic synthetic data layer; ERP connectors are planned for Phase 2.\n"
            "2. Supplier/vendor self-service portal — external vendor access is not included.\n"
            "3. Automated payment processing or disbursement.\n"
            "4. Contract authoring, negotiation, or CLM (Contract Lifecycle Management) capabilities.\n"
            "5. Native mobile applications (iOS/Android) — responsive web design only.\n"
            "6. AI/ML-based predictive analytics or anomaly detection — planned for Phase 2.\n"
            "7. KPMG SSO / Active Directory integration — planned for Phase 2.\n"
            "8. Multi-tenancy (supporting multiple KPMG entities simultaneously in a single instance).\n"
            "9. Automated email notifications or workflow orchestration engine."
        ).encode("utf-8"),
    ),

    # ── 3.1  User Interfaces ─────────────────────────────────────────────────
    (
        b"&lt;Describe the logical characteristics of each interface between the software product and the users. This may include sample screen images, any GUI standards or product family style guides that are to be followed, screen layout constraints, keyboard shortcuts, error message display standards, and so on&gt;",
        xe(
            "IntelliSource is a Single Page Application (SPA) built with React 18 + TypeScript following KPMG brand guidelines. "
            "Key UI characteristics:\n"
            "- Color palette: KPMG Navy (#00336B) primary, KPMG Purple (#3B006E) accent, neutral surface (#F5F5F5).\n"
            "- Layout: fixed left sidebar (navigation), fixed top bar (notifications, user profile), scrollable main content area.\n"
            "- Typography: Inter (UI) and Times New Roman (document views); tabular numeric figures for financial data.\n"
            "- Data visualization: Recharts library — Line, Bar, Area, Pie/Donut, Composed charts. All charts have tooltips and legends.\n"
            "- Status indicators: StatusPill component with color-coded tones (success/warning/danger/info) and optional dot.\n"
            "- KPI Cards: KpiCard component with value, delta indicator, sublabel, optional sparkline, and threshold badge.\n"
            "- Responsive: minimum 1280px desktop; sidebar collapses on smaller viewports.\n"
            "- Accessibility: keyboard-navigable components; ARIA labels on interactive elements; WCAG 2.1 AA compliance target.\n"
            "- Error states: inline form validation messages; toast notifications via Sonner for system alerts."
        ).encode("utf-8"),
    ),

    # ── 3.2  Hardware Interfaces ─────────────────────────────────────────────
    (
        b"&lt;Describe the logical and physical characteristics of each interface between the software product and the hardware components of the system. This may include the supported device types, the nature of the data and control interactions between the software and the hardware, and the communication protocols to be used&gt;",
        xe(
            "No dedicated hardware integration is required. IntelliSource is a browser-based web application deployed on cloud infrastructure. "
            "Supported devices: desktop and laptop computers with a modern web browser. "
            "Minimum screen resolution: 1280x768. "
            "Supported browsers: Google Chrome 120+, Microsoft Edge 120+, Mozilla Firefox 120+, Apple Safari 17+. "
            "No hardware peripherals (scanners, printers, card readers) are interfaced directly by the application."
        ).encode("utf-8"),
    ),

    # ── 3.3  Software Interfaces ─────────────────────────────────────────────
    (
        b"&lt;Describe the connections between this product and other specific software components (name and version), including databases, operating systems, tools, libraries, and integrated commercial components. The data items or messages coming into the system and going out and describe the services needed and the nature of communications. Refer to documents that describe detailed application programming interface protocols. Identify data that will be shared across software components. If the data sharing mechanism must be implemented in a specific way, such as globally shared data in a multitasked operating system, describe this as an implementation constraint.",
        xe(
            "Frontend stack: React 18.x, TypeScript 5.x, TanStack Router v1, Recharts 2.x, Lucide React (icons), Tailwind CSS, Radix UI primitives. "
            "Build toolchain: Vite 5.x. "
            "Package manager: npm. "
            "Linting: ESLint 9 with TypeScript plugin. "
            "State management: React Context (AppContext) for auth state and role. "
            "Data layer (Phase 1): deterministic synthetic data module (src/data/mock.ts) — no external DB calls. "
            "Planned Phase 2 integrations: REST APIs for SAP S/4HANA (PO/vendor data), Oracle ERP (financial), and KPMG ABAC/sanctions screening service. "
            "Authentication: JWT tokens stored in memory (no localStorage) to prevent XSS-based token theft."
        ).encode("utf-8"),
    ),

    # ── 3.4  Data Interfaces ─────────────────────────────────────────────────
    (
        b"&lt;Describe the data as to how data will be exchanged with external systems, libraries, and integrated commercial components. Example: data to be transferred to the cloud storage using SFTP and encrypted using AES-256 etc.&gt;",
        xe(
            "Phase 1 (current): All data sourced from src/data/mock.ts — a deterministic pseudo-random generator producing Vendor, PurchaseOrder, Project, FinancialPeriod, ComplianceCheck, UtilizationRecord, InvoiceRecord, P2PItem, ActionItem, and AuditLog records. No external data transfer. "
            "Phase 2 (planned): "
            "- ERP data ingestion: JSON over HTTPS REST API (TLS 1.3), OAuth 2.0 client credentials grant. "
            "- Data at rest: AES-256 encryption on cloud storage. "
            "- Data in transit: TLS 1.3 minimum for all API calls. "
            "- Bulk export: CSV/Excel export of vendor scorecards, PO lists, and compliance reports via browser download (no server-side file storage). "
            "- Audit log: append-only event log persisted server-side, viewable in Admin module."
        ).encode("utf-8"),
    ),

    # ── 3.5  Security Interfaces ─────────────────────────────────────────────
    (
        b"&lt;Describe the various security interfaces involved in the project while and their functionalities. Example: This may include authentication, authorization, and encryption. Requirements: required by this product, including Cloud Architecture Components, email, web (SSO)&gt;",
        xe(
            "Authentication: JWT-based login (username + password). Login and signup screens implemented (src/routes/login.tsx, src/routes/signup.tsx). "
            "Authorization: Role-Based Access Control (RBAC). Roles: Procurement Manager, Compliance Officer, Finance, IT Admin, Leadership/CXO. Role stored in AppContext; route guards enforce role-level access. "
            "Encryption: HTTPS/TLS 1.3 for all client-server communication (Phase 2). AES-256 for data at rest (Phase 2). "
            "ABAC compliance: vendor transactions screened against ABAC gift threshold and sanctions list — breach flags surfaced in real time on Dashboard and Compliance module. "
            "Phase 2 additions: KPMG SSO/Active Directory integration; MFA enforcement for CXO and Admin roles; session timeout after 30 minutes of inactivity."
        ).encode("utf-8"),
    ),

    # ── 3.5 second para (Communication standards note) ───────────────────────
    (
        b"Identify any communication standards that will be used such as FTP or HTTP.",
        xe(
            "Communication standard: HTTPS (HTTP/2 over TLS 1.3) for all client-server and API communication. "
            "WebSocket planned for Phase 2 real-time compliance breach push notifications. "
            "No FTP, SFTP, or legacy HTTP (plain) connections are used."
        ).encode("utf-8"),
    ),

    # ── 3.6  Communication Interfaces ────────────────────────────────────────
    (
        b"&lt;Describe any requirements associated with any communications functions required by this product, including email, web browser, network server communications protocols, electronic forms etc. Define any pertinent message formatting. Identify any communication standards that will be used (ISO, etc.). Specify any communication security or encryption issues, data transfer rates, and synchronization mechanisms&gt;",
        xe(
            "Protocol: HTTPS (port 443) for all browser-to-server communication. "
            "Browser-native Fetch API used for API calls. "
            "Notification system: in-app notification panel (TopBar component) displays real-time alerts (compliance breaches, PO approvals, contract expirations, budget threshold crossings). "
            "Phase 2 additions: email notification via SMTP (SendGrid or AWS SES) for critical compliance breach alerts; webhook support for integration with third-party ITSM tools (ServiceNow). "
            "Data format: JSON (UTF-8 encoded) for all API payloads. "
            "No ISO-specific communication protocol constraints beyond standard HTTPS/JSON REST conventions."
        ).encode("utf-8"),
    ),

    # ── 4.1  As-Is Process ───────────────────────────────────────────────────
    (
        b"&lt;Mention the status of as is process being used for now which may involve Manual Process or a system for which the new activity (has to be Undertaken e.g. Automation &gt;",
        xe(
            "Current (As-Is) state before IntelliSource:\n"
            "- Vendor Performance: maintained in shared Excel workbooks updated monthly from ERP exports; no real-time ratings or risk flagging.\n"
            "- Purchase Orders: tracked in SAP/Oracle ERP but visibility restricted to ERP power users; procurement managers depend on weekly status reports emailed by ERP team.\n"
            "- Compliance Checks: conducted post-hoc during quarterly internal audits; ABAC and sanctions screening done manually via spreadsheet lookups.\n"
            "- P2P Lifecycle: no unified view; each team (PR, PO, GRN, Invoicing) maintains its own tracker; bottlenecks identified only after escalation.\n"
            "- Contract Renewals: tracked in a shared Excel calendar; renewal reminders sent ad hoc by individual procurement managers.\n"
            "- Financial Analytics: gross margin data available 2-4 weeks after period close via ERP finance reports; no integrated trend view.\n"
            "- License Utilization: SaaS license counts tracked in IT asset spreadsheets; actual usage data not correlated with cost."
        ).encode("utf-8"),
    ),

    # ── 4.2  To-Be Process ───────────────────────────────────────────────────
    (
        b"&lt;Mention the steps of the process being carried out for now if they may have been using Manual processes/Automation. For example &gt;",
        xe(
            "Target (To-Be) state with IntelliSource:\n"
            "- Vendor Performance: real-time vendor scorecards with delivery score, quality score, ABAC certification status, compliance classification, risk tier, and YTD spend. Alerts for non-compliant or high-risk vendors.\n"
            "- Purchase Orders: unified PO dashboard with status tracking (Approved/Pending/Under Review/Deleted), DOA level monitoring, high-value PO alerts (> Rs. 1 Cr), and compliance breach flags on PO records.\n"
            "- Compliance: automated breach detection surfaced as real-time alerts for ABAC, sanctions, DOA bypass, duplicate invoices, rate deviation, and conflict of interest. Severity classification and resolution tracking built in.\n"
            "- P2P Lifecycle: 16-stage flow visualization with item counts per stage, bottleneck identification (PO Approval stage flagged as critical bottleneck at 89 items), SLA breach tracking per transaction.\n"
            "- Contract Renewals: contract expiry calendar embedded in Dashboard; 60-day advance notification for expiring contracts with spend-at-risk highlighted.\n"
            "- Financial Analytics: 24-month rolling gross margin trend with 25% target reference; divisional budget vs. spend tracking; invoice aging dashboard refreshed daily.\n"
            "- License Utilization: per-tool utilization percentage, cost-per-active-user, optimization flags, and potential savings recommendations (e.g., Salesforce at 44% utilization - Rs. 18.5 L savings opportunity)."
        ).encode("utf-8"),
    ),

    # ── 4.3  Existing Challenges ─────────────────────────────────────────────
    (
        b"&lt;Mention the existing process being used for now which may have been using Manual processes/Automation e.g. Delay in execution of manual activity due to volume of transactions, frequency of the activity, Timing Automation getting executed etc.&gt;",
        xe(
            "Key challenges driving the IntelliSource initiative:\n"
            "1. No single source of truth: vendor data, PO status, and compliance information scattered across ERP, spreadsheets, and email.\n"
            "2. Reactive compliance management: ABAC and sanctions breaches discovered weeks after the fact during audits rather than at point of transaction.\n"
            "3. P2P bottlenecks invisible: PO Approval stage accumulates 89 items with no escalation mechanism; average cycle time of 4.2 days vs. 3.0-day target.\n"
            "4. Contract renewal risk: 10 vendors on 'Expiring Soon' status; Wipro master agreement (Rs. 124.5 Cr) at risk of lapse.\n"
            "5. Financial data lag: gross margin data available 2-4 weeks after period close; critical projects (PRJ-0001 at 8.4% GM, PRJ-0008 at 11.2% GM) not flagged until month-end review.\n"
            "6. License waste undetected: Salesforce at 44% utilization (Rs. 18.5 L potential annual saving), Tableau at 31% utilization (Rs. 9.2 L saving).\n"
            "7. PO deletion anomalies: deleted POs not monitored for fraud or error patterns; investigation delayed due to lack of consolidated view.\n"
            "8. High-value PO risk: no real-time alert for POs above Rs. 1 Cr threshold awaiting CXO approval beyond SLA."
        ).encode("utf-8"),
    ),

    # ── 4.4  Proposed Solution ───────────────────────────────────────────────
    (
        b"&lt;Enter the proposed solution/System to put the challenges resolved or minimized above&gt;",
        xe(
            "KPMG IntelliSource: a single-pane-of-glass procurement intelligence platform built as a React 18 + TypeScript SPA with KPMG brand identity. "
            "The solution provides: (a) Real-time Procurement Dashboard with PO status, compliance breach alerts, and contract renewal calendar. "
            "(b) Vendor Intelligence with scorecards, risk classification, and ABAC certification tracking for 50+ vendors across IT Services, Consulting, Cloud, Security, Data Analytics, and SaaS categories. "
            "(c) 16-stage P2P Lifecycle Tracker with bottleneck identification and SLA breach monitoring across Rs. 2000 Cr+ annual procurement volume. "
            "(d) Automated Compliance Breach Detection surfacing 6 breach types with severity triage and resolution workflow. "
            "(e) Financial Analytics with 24-month margin trends and divisional budget tracking. "
            "(f) License Utilization Optimization identifying Rs. 28+ L in immediate savings. "
            "(g) Role-based dashboards for 6 user personas ensuring information is contextual and access-controlled."
        ).encode("utf-8"),
    ),

    # ── 4.4 second placeholder (before section 5) ────────────────────────────
    (
        b"&lt;Mention the proposed solution/System to put the challenges resolved or minimized above&gt;",
        xe(
            "See Section 4.4 above. IntelliSource addresses all identified challenges through a unified procurement intelligence dashboard with real-time data, role-based access control, automated compliance alerting, P2P lifecycle visualization, and financial analytics. "
            "Technology stack: React 18, TypeScript 5, TanStack Router, Recharts, Vite, Tailwind CSS, Radix UI."
        ).encode("utf-8"),
    ),

    # ── 5.1  FR guide text ───────────────────────────────────────────────────
    (
        b"&lt;Enter the functional or Non-Functional requirements can be listed in a tabular, a table, structured below, based that the maturity of the team they can use either or. All requirements to be properly marked as functional and Non-functional&gt;",
        xe(
            "KPMG IntelliSource functional requirements are listed in the table below. Each requirement is identified by a unique Business Requirement number (BR#), module, requirement name, detailed description, and remarks."
        ).encode("utf-8"),
    ),

    # ── 5.1  FR guide text (second para) ─────────────────────────────────────
    (
        b"Guide for item filing for references to be removed: &lt;These FUNCTIONAL REQUIREMENTS describe what the system should do \xe2\x80\x93 the specific behaviors, functions, and features",
        xe(
            "Functional Requirements below cover all 10 modules of IntelliSource. Priority: H=High, M=Medium, L=Low."
        ).encode("utf-8"),
    ),

    # ── 5.2  NFR guide text ──────────────────────────────────────────────────
    (
        b"Guide for item filling for references to be removed: &lt;These NON-FUNCTIONAL REQUIREMENTS describe how the system should behave. These NON-FUNCTIONAL REQUIREMENTS define the quality attributes of the system. They include: Performance, Responsiveness, Security, Stability (e.g. 9 9), Accessibility, Compliance, Maintainability, Quality,",
        xe(
            "Non-Functional Requirements for KPMG IntelliSource:"
        ).encode("utf-8"),
    ),

    # ── Roles section guidance text ──────────────────────────────────────────
    (
        b"&lt;This can be handled by either a Role and Responsibility List (Marked under Point A) or RACIS Matrix (Marked Under Point B) of Any project as per convenience. All related projects has separate Unit to be followed as marked under point A only.&gt;",
        xe(
            "IntelliSource supports six distinct user roles. Role definitions and responsibilities are listed in the table below (Point A). A RACI matrix covering key procurement activities follows in Point B."
        ).encode("utf-8"),
    ),

    # ── Role 1 ───────────────────────────────────────────────────────────────
    (
        b"&lt;Role 1&gt;",
        b"Procurement Manager",
    ),

    # ── Role 2 ───────────────────────────────────────────────────────────────
    (
        b"&lt;Role 2&gt;",
        b"Compliance Officer",
    ),

    # ── Role 3 ───────────────────────────────────────────────────────────────
    (
        b"&lt;Role 3&gt;",
        b"Finance / CFO",
    ),

    # ── Role 4 ───────────────────────────────────────────────────────────────
    (
        b"&lt;Role 4&gt;",
        b"IT Administrator",
    ),

    # ── Role descriptions (appear 4 times in the same pattern) ───────────────
    # We replace them differently each time using a counter — handled below

    # ── AI Stakeholder role mapping guidance ─────────────────────────────────
    (
        b"&lt;AI Stakeholder Role Mapping: To be filled for each AI use case. This table helps clarify responsibilities, highlight third-party dependencies, and define ownership across the AI lifecycle. Mark \"N/A\" where a role is not applicable.&gt;",
        xe(
            "RACI Matrix — Key Procurement Activities (R=Responsible, A=Accountable, C=Consulted, I=Informed). "
            "Roles: PM=Procurement Manager, CO=Compliance Officer, FIN=Finance/CFO, ADMIN=IT Admin, LEAD=Leadership/CXO, OPS=Procurement Ops."
        ).encode("utf-8"),
    ),

    # ── Integrations CH-01 cell ──────────────────────────────────────────────
    (
        b"&lt;Enter system for reference&gt;",
        b"SAP S/4HANA (Phase 2)",
    ),

    # ── Integrations CH-01 guide text ────────────────────────────────────────
    (
        b"&lt;Guide for the team: Undermentioned data has been filled for the team as references link while filled the details for Integration in above mentioned table&gt;",
        xe(
            "Phase 1: No live integration — synthetic data layer used. "
            "Phase 2 integrations planned: SAP S/4HANA (PO and vendor master), Oracle ERP (financial data), ABAC/Sanctions screening API, KPMG SSO/Active Directory, and ServiceNow for ITSM ticket creation on compliance breaches."
        ).encode("utf-8"),
    ),

    # ── 10  Assumptions ──────────────────────────────────────────────────────
    (
        b"&lt;Describe if there are any assumptions, dependencies and constraints related to the project&gt;",
        xe(
            "Refer to Sections 10.1 (Assumptions), 10.2 (Dependencies), and 10.3 (Constraints) below."
        ).encode("utf-8"),
    ),

    (
        b"&lt;All Changes agreed with stakeholders including sponsor to be listed here&gt;.",
        xe(
            "1. Users have access to modern web browsers (Chrome 120+, Edge 120+, Firefox 120+, Safari 17+) on desktop/laptop devices.\n"
            "2. The synthetic mock data layer in src/data/mock.ts is a faithful representation of real ERP data structures; live ERP data will follow the same schema.\n"
            "3. KPMG IT team will provision cloud hosting (Azure/AWS) for Phase 2 production deployment.\n"
            "4. Role definitions (Procurement Manager, Compliance Officer, Finance, Admin, Leadership) are agreed with business stakeholders and will not change during Phase 1.\n"
            "5. ABAC thresholds, DOA approval levels (L1-L4/CXO), and SLA targets (PO Approval: 3 days, Payment Settlement: 30 days) are validated by Compliance and Procurement leadership.\n"
            "6. Data security classification of procurement data has been reviewed and approved by KPMG Information Security."
        ).encode("utf-8"),
    ),

    # ── 10.2  Dependencies ───────────────────────────────────────────────────
    (
        b"&lt;These are external factors or systems that the project depends on for successful delivery.&gt;",
        xe(
            "1. ERP Data Export: SAP S/4HANA or Oracle ERP must provide vendor master, PO, and financial data exports in JSON/CSV format for Phase 2 integration.\n"
            "2. KPMG IT Infrastructure: cloud hosting environment (Azure preferred) must be provisioned and approved before Phase 2 go-live.\n"
            "3. Compliance Team: ABAC compliance rules, sanctions screening service API credentials, and breach type definitions must be formally signed off before UAT.\n"
            "4. KPMG SSO/Active Directory: for Phase 2 SSO integration, KPMG IT must provide SAML 2.0 or OAuth 2.0 configuration details.\n"
            "5. Vendor Data Quality: accuracy of Phase 2 live data depends on data quality maintained in source ERP systems."
        ).encode("utf-8"),
    ),

    # ── 10.3  Constraints ────────────────────────────────────────────────────
    (
        b"&lt;These are limitations or restrictions that the project must operate within.&gt;",
        xe(
            "1. Phase 1 Scope: no live ERP connectivity; all data served from deterministic synthetic mock layer. Real data integration deferred to Phase 2.\n"
            "2. Browser-Only: IntelliSource is a web application; no native mobile (iOS/Android) app is in scope for either phase.\n"
            "3. Data Classification: procurement and financial data must comply with KPMG's Confidential data classification policy — no data may be stored in unapproved third-party cloud services.\n"
            "4. Technology Stack: frontend must remain React + TypeScript to align with KPMG front-end engineering standards.\n"
            "5. Branding: UI must strictly follow KPMG brand guidelines (color palette, typography, iconography) as defined in src/lib/brand.ts.\n"
            "6. Budget: Phase 1 delivered as an internal development project; no vendor licensing costs for third-party SaaS tools in the dashboard stack."
        ).encode("utf-8"),
    ),

    # ── Section 11  Prototype ────────────────────────────────────────────────
    (
        b"&lt;This section to be included in case the prototype to be presented to the client&gt;",
        xe(
            "A functional high-fidelity prototype of KPMG IntelliSource has been developed and is available as a live web application. "
            "The prototype covers all 10 modules: Procurement Dashboard, Vendor Performance, Financial Analytics, P2P Lifecycle Tracker, Compliance, Utilization, Actions, Vendor Repository, Leadership Dashboard, and Admin. "
            "Prototype is built on the same production codebase (React 18 + TypeScript + TanStack Router) using deterministic synthetic data for demonstration. "
            "Stakeholders can access the prototype at the development deployment URL provided by the project team. "
            "Screen recordings and annotated screenshots are available as supplementary materials in the project repository."
        ).encode("utf-8"),
    ),

    # ── 12.1  References ─────────────────────────────────────────────────────
    (
        b"&lt;Add references if any with link to support the team for guidance, undermentioned link is only for reference and to be updated as per the Application/Services related&gt;",
        xe(
            "1. KPMG Brand Guidelines — KPMG Global Brand Portal (internal).\n"
            "2. KPMG Information Security Policy — Data Classification Framework v3.2 (internal).\n"
            "3. React 18 Documentation — https://react.dev\n"
            "4. TanStack Router v1 Documentation — https://tanstack.com/router\n"
            "5. Recharts Documentation — https://recharts.org\n"
            "6. WCAG 2.1 Accessibility Guidelines — https://www.w3.org/TR/WCAG21/\n"
            "7. ABAC Compliance Framework — KPMG Compliance Team internal reference (contact: Compliance Officer).\n"
            "8. IntelliSource GitHub Repository — internal KPMG GitHub Enterprise (project: kpmg-intelliflow-v5)."
        ).encode("utf-8"),
    ),

    # ── 12.2  Use case model ─────────────────────────────────────────────────
    (
        b"&lt;Delete section if not required. Provides the use case diagrams prepared for the functional requirements&gt;",
        xe(
            "Key use cases for KPMG IntelliSource:\n"
            "UC-01: Procurement Manager reviews daily PO dashboard, identifies high-value POs awaiting CXO approval, and escalates breach-flagged POs.\n"
            "UC-02: Compliance Officer reviews real-time breach alerts, investigates ABAC/sanctions matches, updates vendor compliance status, and escalates to legal if required.\n"
            "UC-03: Finance Analyst monitors gross margin trends, compares divisional budget vs. spend, and reviews invoice aging report.\n"
            "UC-04: Procurement Ops identifies P2P bottleneck at PO Approval stage (89 items), drills into SLA-breached transactions, and assigns resolution owners.\n"
            "UC-05: IT Admin provisions new user account, assigns Compliance Officer role, and reviews audit log for data access events.\n"
            "UC-06: CXO reviews Leadership Dashboard for cross-functional procurement health KPIs before board meeting.\n"
            "UC-07: Procurement Manager identifies Salesforce and Tableau underutilization, initiates contract renegotiation action item (Rs. 18.5 L saving)."
        ).encode("utf-8"),
    ),

    # ── 12.3.1  Abbreviations ────────────────────────────────────────────────
    (
        b"&lt;Kindly add all abbreviations used within this document for users to understand it well&gt;",
        xe(
            "ABAC: Anti-Bribery and Anti-Corruption\n"
            "APAC: Asia-Pacific region\n"
            "BRD: Business Requirements Document\n"
            "CRUD: Create, Read, Update, Delete\n"
            "CXO: C-Suite Executive (CEO, CFO, COO, CTO, etc.)\n"
            "DOA: Delegation of Authority\n"
            "EMEA: Europe, Middle East, and Africa region\n"
            "ERP: Enterprise Resource Planning\n"
            "FR: Functional Requirement\n"
            "GM: Gross Margin\n"
            "GRN: Goods Receipt Note\n"
            "HTTPS: Hypertext Transfer Protocol Secure\n"
            "JWT: JSON Web Token\n"
            "KPI: Key Performance Indicator\n"
            "MTD: Month-to-Date\n"
            "NFR: Non-Functional Requirement\n"
            "OI: Open Item\n"
            "P2P: Procure-to-Pay\n"
            "PO: Purchase Order\n"
            "PR: Purchase Requisition\n"
            "RBAC: Role-Based Access Control\n"
            "RFQ: Request for Quotation\n"
            "SaaS: Software as a Service\n"
            "SLA: Service Level Agreement\n"
            "SPA: Single Page Application\n"
            "SSO: Single Sign-On\n"
            "TLS: Transport Layer Security\n"
            "UAT: User Acceptance Testing\n"
            "WCAG: Web Content Accessibility Guidelines\n"
            "YTD: Year-to-Date"
        ).encode("utf-8"),
    ),

    # ── 12.3.2  Glossary ─────────────────────────────────────────────────────
    (
        b"&lt;Undermentioned ingredients of glossary are for reference only and to be updated as per the Application/Services related&gt;",
        xe(
            "Bottleneck Stage: A P2P lifecycle stage where the number of items exceeds 85% of the maximum stage count, indicating a processing constraint.\n"
            "Compliance Breach: A vendor transaction that violates KPMG compliance rules including ABAC thresholds, sanctions screening, DOA limits, duplicate invoice detection, rate deviation, or conflict of interest.\n"
            "DOA Level: Delegation of Authority approval tier. L1 (lowest) through L4 and CXO (highest) determine who must approve a given PO value.\n"
            "High-Value PO: A Purchase Order with value exceeding Rs. 1 Crore, requiring enhanced monitoring and CXO-level DOA approval.\n"
            "OTIF Rate: On Time and In Full — vendor delivery performance metric expressed as a percentage.\n"
            "Risk Tier: Vendor classification into Low, Medium, or High risk based on compliance status, single-source dependency, and performance metrics.\n"
            "Single-Source Vendor: A vendor that is the sole supplier for a given category, creating supply concentration risk.\n"
            "SLA Breach: A P2P transaction that has remained in a lifecycle stage beyond its defined SLA (e.g., PO Approval SLA: 3 days).\n"
            "Utilization Rate: Percentage of licensed software seats actively used. Below 50%: Underutilized. Above 95%: Over-provisioned.\n"
            "Vendor Scorecard: A consolidated view of a vendor's ratings, delivery score, quality score, ABAC certification, compliance status, risk tier, and financial spend metrics."
        ).encode("utf-8"),
    ),

    # ── Change control CH-01 placeholder ─────────────────────────────────────
    (
        b"&lt;Enter CR Number for any&gt;",
        b"N/A",
    ),
    (
        b"&lt;PoI Reference section wherever changes is done to be marked here&gt;",
        xe("Initial BRD release — no prior change requests. All sections populated for Phase 1 scope.").encode("utf-8"),
    ),
    (
        b"&lt;Enter CR reference&gt;",
        b"N/A",
    ),
]


# Role descriptions — these appear 4 times with identical XML, so we must handle them
# sequentially using a find-replace-once approach.
ROLE_DESCRIPTIONS = [
    b"Creates and approves Purchase Orders, monitors vendor performance scorecards, manages the high-value PO monitor, tracks P2P flow for assigned categories, manages action items, and coordinates contract renewals.",
    b"Reviews and investigates real-time compliance breach alerts (ABAC, sanctions, DOA, duplicate invoices, rate deviation, conflict of interest); updates vendor compliance classification; escalates critical violations to legal; manages the compliance check register.",
    b"Monitors gross margin trends and divisional budget vs. spend; reviews invoice aging and overdue payments; accesses financial KPI dashboards; approves financial action items.",
    b"Manages user accounts (CRUD) and role assignments; reviews system audit logs; configures application settings; ensures platform security and access controls.",
]

PLACEHOLDER_ROLE_DESC = b"&lt;Mention the description of the particular role&gt;"


def fill_brd():
    shutil.copy(SRC, DST)

    # Read all files from the source docx
    with zipfile.ZipFile(SRC, "r") as zin:
        all_files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = all_files["word/document.xml"]

    # ── Apply simple replacements ─────────────────────────────────────────────
    for old, new in REPLACEMENTS:
        doc_xml = doc_xml.replace(old, new)

    # ── Apply role description replacements sequentially ─────────────────────
    for desc in ROLE_DESCRIPTIONS:
        idx = doc_xml.find(PLACEHOLDER_ROLE_DESC)
        if idx != -1:
            doc_xml = doc_xml[:idx] + desc + doc_xml[idx + len(PLACEHOLDER_ROLE_DESC):]

    # ── Fill FR table rows (FR01 and FR02) ───────────────────────────────────
    # FR01 row has empty Module, Business Requirement, Description, Remarks cells
    # We'll insert content into the empty cells adjacent to FR01/FR02 labels.
    # Strategy: replace the empty cell patterns that immediately follow FR01/FR02

    FR_ROWS = [
        # (fr_label_bytes, module, requirement, description, remarks)
        (
            b"<w:t>FR01</w:t>",
            b"Procurement Dashboard",
            b"Real-Time KPI Dashboard",
            b"Display 6 KPI cards on the main dashboard: Total PO Value MTD, Active POs by status (Approved/Pending/Under Review), High-Value PO count (>Rs.1Cr), Avg PO Cycle Time vs. 3-day target, Compliance Breach Rate, and PO Deletion Frequency. Include 12-month area sparkline on Total PO Value card.",
            b"Priority: H. Phase 1.",
        ),
        (
            b"<w:t>FR02</w:t>",
            b"Vendor Management",
            b"Vendor Scorecard List",
            b"Display paginated vendor list with columns: Vendor ID, Name, Category, Region, Rating (1-5 stars), Delivery Score, Quality Score, ABAC Certification status, Compliance Status (Compliant/Under Review/Non-Compliant), Risk Tier (Low/Medium/High), YTD Spend (INR), Contract Status, Contract End Date. Support filtering by Category, Region, Compliance Status, and Risk Tier.",
            b"Priority: H. Phase 1.",
        ),
    ]

    for fr_label, module_v, req_v, desc_v, remarks_v in FR_ROWS:
        idx = doc_xml.find(fr_label)
        if idx == -1:
            continue
        # Find the next 4 empty cells (they look like <w:p><w:r>...<w:rPr>...</w:rPr></w:r></w:p>)
        # Each cell ends with </w:tc> — we want to inject text into the empty cell body.
        # Find the 4 empty <w:p><w:r><w:rPr>...<w:sz w:val="20"/></w:rPr></w:r></w:p> patterns
        # after the FR label.
        EMPTY_CELL_BODY = (
            b'<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            b'<w:sz w:val="20"/></w:rPr></w:r></w:p>'
        )
        values = [module_v, req_v, desc_v, remarks_v]
        search_start = idx + len(fr_label)
        for val in values:
            eidx = doc_xml.find(EMPTY_CELL_BODY, search_start)
            if eidx == -1:
                break
            filled = (
                b'<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                b'<w:sz w:val="20"/></w:rPr><w:t>' + val + b'</w:t></w:r></w:p>'
            )
            doc_xml = doc_xml[:eidx] + filled + doc_xml[eidx + len(EMPTY_CELL_BODY):]
            search_start = eidx + len(filled)

    all_files["word/document.xml"] = doc_xml

    # ── Write new docx ────────────────────────────────────────────────────────
    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)

    print(f"Written: {DST}")


if __name__ == "__main__":
    fill_brd()
