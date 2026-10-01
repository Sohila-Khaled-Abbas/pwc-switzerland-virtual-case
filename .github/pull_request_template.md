## 📌 Pull Request Overview

### 🛠️ Scope of Change
- [ ] Dimensional Model (Galaxy Schema / VertiPaq Tables)
- [ ] DAX Formulation / KPI Engineering
- [ ] Power Query M ETL Script
- [ ] VBA Automation / UI Macro
- [ ] Documentation / Data Dictionary

### 📋 Description & Rationale
<!-- Provide a concise description of the modification, business rationale, and affected components -->

### 🧪 Verification Checklist
- [ ] Relationships in Power Pivot Diagram View remain 1-to-Many (`1 : *`)
- [ ] No duplicate tables or ghost connections created
- [ ] DAX measures tested for divide-by-zero resilience using `DIVIDE()`
- [ ] Data dictionary updated if new columns or dimensions are introduced
