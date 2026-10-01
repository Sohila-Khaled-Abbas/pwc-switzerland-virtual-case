# 📐 Enterprise Galaxy Semantic Model Architecture

> **Framework**: Ralph Kimball Dimensional Lifecycle  
> **Engine**: Microsoft Analysis Services Tabular / Power Pivot VertiPaq Columnar Storage  
> **Target Topology**: Multi-Fact Galaxy / Fact Constellation Schema  

---

## 🏛️ Why a Galaxy Schema?

In enterprise analytics, real-world corporate organizations cannot be forced into a single monolithic star schema without severe denormalization, metric duplication, and fan traps. 

A **Kimball Galaxy Schema (Fact Constellation)** accommodates multiple business processes that share common, conformed dimensions while allowing each fact table to retain its true atomic grain:

```mermaid
classDiagram
    class DimDate {
        +Date PK
        +Year
        +Quarter
        +Month
        +Month_Name
        +Day
        +Day_Of_Week
        +Is_Weekend
    }

    class DimAgent {
        +Agent PK
        +Tier
        +Target_CSAT
        +Target_Resolution_Rate
    }

    class DimTopic {
        +Topic PK
        +Complexity_Weight
        +SLA_Threshold_Sec
    }

    class DimContract {
        +Contract PK
        +Commitment_Months
        +Risk_Category
    }

    class DimDepartment {
        +Department PK
        +Executive_Sponsor
        +Target_Female_Ratio
    }

    class Fact_Calls {
        +Call_Id PK
        +Date FK
        +Time
        +Agent FK
        +Topic FK
        +Answered
        +Resolved
        +Speed_Of_Answer_Sec
        +Talk_Duration_Sec
        +Satisfaction_Rating
    }

    class Fact_Churn {
        +customerID PK
        +gender
        +tenure
        +Contract FK
        +MonthlyCharges
        +TotalCharges
        +numTechTickets
        +Churn
    }

    class Fact_Employees {
        +Employee_ID PK
        +Gender
        +Department FK
        +Job_Level_Baseline
        +Job_Level_After_Promotions
        +Promoted_FY21
        +FY20_Rating
        +FY20_Leaver
    }

    DimDate "1" --> "*" Fact_Calls : Filter (1 to Many)
    DimAgent "1" --> "*" Fact_Calls : Filter (1 to Many)
    DimTopic "1" --> "*" Fact_Calls : Filter (1 to Many)

    DimContract "1" --> "*" Fact_Churn : Filter (1 to Many)

    DimDepartment "1" --> "*" Fact_Employees : Filter (1 to Many)
```

---

## 🔑 Cardinality & Referential Integrity Matrix

| Relationship | From Table (Many Side `*`) | Foreign Key Column | To Table (One Side `1`) | Primary Key Column | Relationship State | Cross-Filter Direction |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| **Rel 1** | `Fact_Calls` | `Date` | `DimDate` | `Date` | **Active** | Single (`DimDate` $\to$ `Fact_Calls`) |
| **Rel 2** | `Fact_Calls` | `Agent` | `DimAgent` | `Agent` | **Active** | Single (`DimAgent` $\to$ `Fact_Calls`) |
| **Rel 3** | `Fact_Calls` | `Topic` | `DimTopic` | `Topic` | **Active** | Single (`DimTopic` $\to$ `Fact_Calls`) |
| **Rel 4** | `Fact_Churn` | `Contract` | `DimContract` | `Contract` | **Active** | Single (`DimContract` $\to$ `Fact_Churn`) |
| **Rel 5** | `Fact_Employees` | `Department @01.07.2020` | `DimDepartment` | `Department` | **Active** | Single (`DimDepartment` $\to$ `Fact_Employees`) |

> [!IMPORTANT]
> **Strict 1-to-Many Architecture**: Notice that every single relationship connects a unique Primary Key (PK) on the Dimension side (`1`) to a Foreign Key (FK) on the Fact side (`*`). There are **zero Many-to-Many (`* : *`) relationships** in this model, guaranteeing predictable filter propagation, zero cross-table Cartesian inflation, and optimal Storage Engine execution.

---

## ⚡ The VertiPaq Engine Storage Mechanics

The VertiPaq in-memory columnar engine stores data using three distinct compression phases:

### 1. Value Encoding
Applied to integer metrics such as `Speed_Of_Answer_Sec`, `Talk_Duration_Sec`, and `tenure`. VertiPaq determines the mathematical minimum and subtracts it as a base:
$$\text{Stored Value} = \text{Raw Value} - \text{Min Value}$$
This drastically reduces the bit-width required per row in RAM.

### 2. Dictionary Encoding
Applied to high-cardinality text columns such as `Agent`, `Topic`, `Contract`, and `Department`. A unique dictionary index is constructed, replacing string occurrences with compact 1-byte integer pointers.

### 3. Run-Length Encoding (RLE)
When rows are sorted by repetitive attributes (e.g. `Date` or `Department`), contiguous identical values are compressed into single run counts: `(Value, Count)`.

---

## 🛡️ Auxiliary Lookups & The Backing Tables Strategy

The source workbook `03 Diversity-Inclusion-Dataset.xlsx` contains 4 auxiliary reference sheets (`Backing 1` to `Backing 4`). Rather than leaving them unmanaged or dumping them into fact columns:

1. **`Dim_EmployeeCensus`**: Ingested to benchmark corporate headcounts against Swiss enterprise averages.
2. **`Dim_CareerLadder`**: Defines the progression ladder across 6 job grades (`1-Executive` to `6-Junior Officer`).
3. **`Dim_NationalityCensus`**: Categorizes Swiss vs Non-Swiss residency quotas for compliance reporting.
4. **`Dim_PRA_Equity`**: Maintains the expected Gaussian quota distribution for performance appraisal reviews (e.g. max 15% top rating).
