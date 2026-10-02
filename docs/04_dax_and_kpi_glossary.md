# 📊 DAX Measures & Executive KPI Glossary

> **Language**: Data Analysis Expressions (DAX)  
> **Engine**: VertiPaq Storage Engine (SE) & Formula Engine (FE)  
> **Standard**: Zero Callback-on-DAX, Strict `DIVIDE` Safe Numerics, Explicit Filter Modification  

---

## 🏎️ VertiPaq Performance Best Practices

To guarantee sub-second query response times across multi-million row datasets, all DAX measures in this portfolio conform to five foundational performance rules:

1. **Strict DIVIDE Usage**: Never use the raw `/` division operator. `DIVIDE(Numerator, Denominator, AlternateResult)` intercepts division by zero at the engine level, returning `BLANK()` or `0` without throwing a calculation error.
2. **Elimination of Calculated Columns on Facts**: Calculated columns consume uncompressed memory across every row. All analytical calculations (rates, ratios, sums) are engineered as **Explicit DAX Measures**.
3. **Storage Engine Pushdown**: Simple aggregations (`SUM`, `COUNTROWS`, `AVERAGE`) run inside the multi-threaded C++ Storage Engine. Formula Engine single-threaded operations are minimized.
4. **Variable Caching (`VAR ... RETURN`)**: Reusable sub-expressions are stored in variables, ensuring the query engine evaluates the sub-tree exactly once rather than recalculating across iterations.

---

## 🗂️ Comprehensive Measure Catalog

### Domain 1: Call Center Telephony KPIs (Claire's SLA Dashboard)

| Measure Name | DAX Formulation | Engine | Target SLA | Business Interpretation |
| :--- | :--- | :---: | :---: | :--- |
| **Total Inbound Calls** | `COUNTROWS('Fact_Calls')` | SE | N/A | Total call attempts entering the telecom PBX queue. |
| **Calls Answered** | `CALCULATE(COUNTROWS('Fact_Calls'), 'Fact_Calls'[Answered] = "Y")` | SE | > 80% | Total successful connections established between customer and agent. |
| **Calls Abandoned** | `CALCULATE(COUNTROWS('Fact_Calls'), 'Fact_Calls'[Answered] = "N")` | SE | < 15% | Callers who terminated the call while waiting in the queue. |
| **Abandonment Rate %**| `DIVIDE([Calls Abandoned], [Total Inbound Calls], 0)` | SE + FE | < 15.0% | Core operational queue failure metric. Measures lost contact opportunities. |
| **Calls Resolved** | `CALCULATE(COUNTROWS('Fact_Calls'), 'Fact_Calls'[Resolved] = "Y")` | SE | > 85% | Inquiries successfully closed without requiring escalation. |
| **Resolution Rate %** | `DIVIDE([Calls Resolved], [Calls Answered], 0)` | SE + FE | > 85.0% | Agent operational efficacy in closing issues during connected sessions. |
| **Avg Speed of Answer (Sec)**| `AVERAGE('Fact_Calls'[Speed_Of_Answer_Sec])` | SE | < 60s | Average wait time before caller connects with an agent. |
| **Avg Satisfaction Rating** | `AVERAGE('Fact_Calls'[Satisfaction_Rating])` | SE | > 3.50 | Post-call CSAT rating on a 1.00 to 5.00 customer survey scale. |
| **Avg Handling Time (Sec)** | `DIVIDE(SUM('Fact_Calls'[Talk_Duration_Sec]), [Calls Answered], BLANK())` | SE + FE | 180s - 240s | Mean conversation duration per connected call engagement. |

---

### Domain 2: Customer Retention & MRR Churn

| Measure Name | DAX Formulation | Engine | Benchmark | Business Interpretation |
| :--- | :--- | :---: | :---: | :--- |
| **Total Customers** | `COUNTROWS('Fact_Churn')` | SE | N/A | Active portfolio of subscriber contracts. |
| **Churned Customers**| `CALCULATE(COUNTROWS('Fact_Churn'), 'Fact_Churn'[Churn] = "Yes")` | SE | Downward | Subscriber accounts that terminated their subscription. |
| **Churn Rate %** | `DIVIDE([Churned Customers], [Total Customers], 0)` | SE + FE | < 20.0% | Proportion of customer accounts lost during the observation cycle. |
| **Total Monthly Charges**| `SUM('Fact_Churn'[MonthlyCharges])` | SE | Baseline | Total Monthly Recurring Revenue (MRR) of subscriber base. |
| **Churned Monthly Charges**| `CALCULATE(SUM('Fact_Churn'[MonthlyCharges]), 'Fact_Churn'[Churn] = "Yes")` | SE | Downward | Realized monthly revenue lost from departed subscriber accounts. |
| **Financial Churn Rate %**| `DIVIDE([Churned Monthly Charges], [Total Monthly Charges], 0)` | SE + FE | < 25.0% | Proportion of MRR destroyed. (Higher than customer churn rate). |
| **Avg Tenure Months** | `AVERAGE('Fact_Churn'[tenure])` | SE | > 36 Mos | Mean customer loyalty duration in months. |
| **Avg Service Tickets**| `AVERAGE('Fact_Churn'[Total_Service_Tickets])` | SE | < 1.0 | Operational friction indicator; high ticket count strongly correlates with churn. |

---

### Domain 3: Diversity, Equity & Inclusion (Pharma Group AG)

| Measure Name | DAX Formulation | Engine | Parity Target | Business Interpretation |
| :--- | :--- | :---: | :---: | :--- |
| **Total Employees** | `COUNTROWS('Fact_Employees')` | SE | N/A | Active workforce headcount across Pharma Group AG. |
| **Female Representation %**| `DIVIDE([Female Employees], [Total Employees], 0)` | SE + FE | 50.0% | Overall proportion of female personnel in corporate headcount. |
| **Overall Promotion Rate %**| `DIVIDE([Total Promoted FY21], [Total Employees], 0)` | SE + FE | N/A | General career advancement velocity across company. |
| **Female Promotion Rate %** | `DIVIDE([Female Promoted FY21], [Female Employees], 0)` | SE + FE | Parity | Career advancement velocity specific to female personnel base. |
| **Male Promotion Rate %** | `DIVIDE([Male Promoted FY21], [Male Employees], 0)` | SE + FE | Parity | Career advancement velocity specific to male personnel base. |
| **Promotion Equity Index** | `DIVIDE([Female Promotion Rate %], [Male Promotion Rate %], BLANK())` | FE | 1.00 | Parity ratio. Values < 1.00 signify male career advancement bias. |
| **Executive Female Share %**| `CALCULATE(DIVIDE([Female Employees], [Total Employees], 0), 'Fact_Employees'[Job_Level_Baseline_Rank] IN {1, 2})` | SE + FE | > 35.0% | Proportion of leadership (Executive Board + Directors) that is female. |
| **Turnover Rate %** | `DIVIDE([FY20 Leavers], [Total Employees], 0)` | SE + FE | < 10.0% | Annual attrition rate of employees leaving the organization. |
