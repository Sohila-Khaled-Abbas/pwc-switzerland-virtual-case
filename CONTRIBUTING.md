# 🤝 Contributing to PwC Switzerland Virtual Case Experience

Thank you for your interest in contributing to this enterprise Business Intelligence and dimensional modeling repository.

---

## 🏛️ Development & Modeling Standards

### 1. Dimensional Modeling Conventions
* Conformed dimensions must use the prefix `Dim` (e.g. `DimDate`, `DimAgent`, `DimDepartment`).
* Fact tables must use the prefix `Fact_` (e.g. `Fact_Calls`, `Fact_Churn`, `Fact_Employees`).
* Relationships must always be **1-to-Many (`1 : *`)** flowing from Dimension to Fact with **Single cross-filter direction**.
* Never introduce bi-directional relationships without documented business justification.

### 2. DAX Formula Guidelines
* Always format DAX measures with explicit variable returns (`VAR ... RETURN`).
* Always protect division calculations with `DIVIDE(Numerator, Denominator, AlternateResult)`.
* Do not create calculated columns on fact tables; create explicit DAX measures instead.

### 3. Power Query M Standards
* Break complex transformations into distinct, named M steps using PascalCase or descriptive camelCase.
* Always enforce strong data types on the final step before loading into the Data Model.
* Do not delete operational nulls without cross-tabulation validation.

---

## 🚀 Pull Request Workflow

1. Fork or branch from `main`:
   ```bash
   git checkout -b feature/new-dax-measure
   ```
2. Test changes in Excel Power Pivot and ensure all relationships remain valid.
3. Commit using Conventional Commits:
   * `feat(dax): add rolling 30-day churn rate`
   * `docs(readme): expand executive KPI descriptions`
4. Open a Pull Request referencing the issue or feature scope.
