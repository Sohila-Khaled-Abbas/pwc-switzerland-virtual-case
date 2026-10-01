# 📋 Changelog

All notable changes to the **PwC Switzerland Virtual Case Experience** platform are documented in this file following [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) and [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.1.0] - 2026-10-02

### 🚀 Added
- **Dynamic Calendar Dimension**: Full M-code sequential generator with fiscal quarters and weekend flags (`power_query/04_calendar_generator.m`).
- **Time Intelligence Measures**: Added MTD, QTD, Prior Month, MoM Growth %, and 7-day rolling moving average DAX calculations.
- **Enterprise Documentation Hub**: Complete suite of 6 executive architecture guides, data dictionary, and UI/UX design blueprints.
- **Continuous Integration Workflow**: GitHub Actions workflow for validating repository workbook assets and code syntax.

### 🔧 Fixed
- **Duplicate Connection Triage**: Eliminated ghost connections `Query - DimDate1` and `Query - DimDepartment1` via Excel COM automation.
- **Model Cleanliness**: Verified exactly 12 single-instance tables in VertiPaq memory with 100% 1-to-many relationship integrity.

---

## [2.0.0] - 2026-10-01

### 🚀 Added
- **Ralph Kimball Galaxy Schema**: Integrated all 3 client datasets (`01 Call-Center`, `02 Churn`, `03 Diversity-Inclusion`) into a single unified semantic model.
- **Auxiliary Census Dimensions**: Ingested `Backing 1` to `Backing 4` reference sheets into dedicated lookup tables.
- **VBA Macro Architecture**: Enterprise application suite including state caching, dynamic data refresh, PDF export, and filter resets.

---

## [1.0.0] - 2026-09-30

### 🚀 Initial Release
- Baseline Call Centre Operations model (5,000 telephony logs).
- Initial CSAT, Abandonment, and Speed of Answer DAX formulations.
