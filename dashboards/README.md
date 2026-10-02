# 📊 PwC Executive Dashboard Suite & UI/UX Blueprints

Welcome to the **PwC Switzerland Executive Dashboard Suite** documentation and design specification. This directory contains detailed wireframe layouts, visual hierarchy guidelines, component dictionaries, and user interaction rules for all three executive dashboards delivered under the PwC Digital Accelerator engagement.

---

## 🎨 Enterprise Design System & Visual Tokens

The dashboards follow the **PwC Brand Identity & Digital Design System**, engineered for optimal contrast, executive cognitive scanning, and sub-second decision making.

```mermaid
flowchart LR
    subgraph ColorTokens["PwC Enterprise Color Palette"]
        P1["🟧 PwC Orange\n#D04A02\nPrimary Accent & Call-to-Action"]
        P2["⬛ Onyx Black\n#0F172A\nExecutive Dark Canvas"]
        P3["🔷 Telephony Cyan\n#38BDF8\nOps & Speed SLA Metrics"]
        P4["🔥 Churn Alert\n#EF4444\nHigh-Risk & Critical Drop-offs"]
        P5["🟢 Parity Emerald\n#10B981\nInclusion & Goal Benchmarks"]
    end
```

### Typography Hierarchy
| Level | Font Family | Size | Weight | Line Height | Usage |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Hero Title** | Inter / Segoe UI | 24pt | Bold (700) | 32pt | Dashboard Title & Primary Navigation |
| **KPI Headline** | Inter / Segoe UI | 28pt | Extrabold (800) | 36pt | Big Numeric Callout Cards |
| **Section Header**| Inter / Segoe UI | 14pt | SemiBold (600) | 20pt | Chart Cards & Table Groupings |
| **Body & Labels** | Inter / Segoe UI | 10pt | Regular (400) | 14pt | Axis Labels, Tooltips, Legends |
| **KPI Subtext** | Inter / Segoe UI | 9pt | Light (300) | 12pt | Target Delays, Variance vs Prior Period |

---

## 📱 Executive Dashboard Catalog

```mermaid
graph TD
    A["🏆 PwC Switzerland Executive Suite"] --> B["📞 Task 1: Call Centre Intelligence"]
    A --> C["🔄 Task 2: Customer Retention & Churn"]
    A --> D["👥 Task 3: Diversity & Inclusion Leadership"]

    B --> B1["Operational Command Cockpit\n5,000 Inbound Inquiries\nAbandonment Triage & Agent Quadrants"]
    C --> C1["Executive Churn Risk Cockpit\n7,043 Subscriber Accounts\n$139K At-Risk MRR & Contract Elasticity"]
    D --> D1["Boardroom Inclusion Scorecard\n500 Personnel\nBroken-Rung Funnel & Promotion Parity"]

    style A fill:#0f172a,color:#fff,stroke:#d04a02,stroke-width:2px
    style B fill:#0369a1,color:#fff,stroke:#38bdf8
    style C fill:#c2410c,color:#fff,stroke:#fb923c
    style D fill:#047857,color:#fff,stroke:#34d399
```

| Dashboard | Target Stakeholder | Primary Objectives | Document Specification |
| :--- | :--- | :--- | :--- |
| **1. Call Centre Operations** | **Claire** (Operations Director) | Real-time SLA monitoring, 18.9% abandonment triage, agent coaching matrix | [`01_call_center_dashboard.md`](01_call_center_dashboard.md) |
| **2. Customer Retention** | **Retention Board & CFO** | Prevent $139K monthly revenue leakage, fix fiber optic friction, migrate monthly accounts | [`02_customer_retention_dashboard.md`](02_customer_retention_dashboard.md) |
| **3. Diversity & Inclusion** | **Executive Board & CHRO** | Eliminate the Senior Manager broken rung, ensure 50/50 hiring parity, track PRA Swiss equity | [`03_diversity_inclusion_dashboard.md`](03_diversity_inclusion_dashboard.md) |

---

## 🎛️ Unified Global Slicers & Cross-Filtering Architecture

All three dashboards share a synchronized dimensional filter ribbon driven by the VertiPaq tabular model:

```mermaid
sequenceDiagram
    autonumber
    actor Executive as Board Member / Ops Director
    participant Ribbon as Conformed Slicer Ribbon
    participant VertiPaq as VertiPaq Tabular Engine
    participant Visuals as Visual Card / Matrix Grid

    Executive->>Ribbon: Selects Date Range (e.g. Jan 2021) or Department (Operations)
    Ribbon->>VertiPaq: Propagates Filter Context across Conformed Dimensions
    VertiPaq->>VertiPaq: Evaluates DAX Context Transition & CALCULATE() Filters
    VertiPaq->>Visuals: Renders Re-Aggregated Values in < 15 milliseconds
    Visuals-->>Executive: Dynamic Visual Update with Sub-Second Latency
```

1. **Date Slicer**: Dynamic Q1 2021 timeline slider (Jan, Feb, Mar) with Day-of-Week toggle.
2. **Business Unit / Department Slicer**: Filters across Operations, Finance, Strategy, HR, Sales & Marketing.
3. **Agent / Service Slicer**: Multi-select dropdown with instant search and clear-all reset macro.
