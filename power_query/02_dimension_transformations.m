// ==============================================================================
// 02_dimension_transformations.m
// PwC Switzerland Virtual Case Experience: Dimension Modeling Layer
// Engine: Power Query M (Microsoft Excel & Power BI)
// Pattern: Dimension Extraction, Deduplication, Attribute Enrichment & Typing
// ==============================================================================

section Dimensions;

// ------------------------------------------------------------------------------
// Dimension 1: DimAgent
// Source: stg_RawCalls | Deduplicates distinct agents, enriches with tiers & SLAs
// ------------------------------------------------------------------------------
shared DimAgent = let
    Source = stg_RawCalls,
    SelectAgent = Table.SelectColumns(Source, {"Agent"}),
    Deduplicated = Table.Distinct(SelectAgent),
    Sorted = Table.Sort(Deduplicated, {{"Agent", Order.Ascending}}),
    
    // Enrich with operational attributes
    AddedTier = Table.AddColumn(Sorted, "Tier", each 
        if [Agent] = "Martha" or [Agent] = "Jim" then "Tier 2 Specialist"
        else "Tier 1 Frontline", type text),
    AddedTargetCSAT = Table.AddColumn(AddedTier, "Target_CSAT", each 3.50, type number),
    AddedTargetResolution = Table.AddColumn(AddedTargetCSAT, "Target_Resolution_Rate", each 0.85, type number)
in
    AddedTargetResolution;


// ------------------------------------------------------------------------------
// Dimension 2: DimTopic
// Source: stg_RawCalls | Deduplicates topics, adds complexity and SLA thresholds
// ------------------------------------------------------------------------------
shared DimTopic = let
    Source = stg_RawCalls,
    SelectTopic = Table.SelectColumns(Source, {"Topic"}),
    Deduplicated = Table.Distinct(SelectTopic),
    Sorted = Table.Sort(Deduplicated, {{"Topic", Order.Ascending}}),
    
    // Add business SLA metadata
    AddedSLA = Table.AddColumn(Sorted, "SLA_Threshold_Sec", each 
        if [Topic] = "Tech support" then 90
        else if [Topic] = "Payment related" then 60
        else 45, Int64.Type),
    AddedComplexity = Table.AddColumn(AddedSLA, "Complexity_Weight", each
        if [Topic] = "Tech support" then 5
        else if [Topic] = "Contract related" then 4
        else if [Topic] = "Payment related" then 3
        else 2, Int64.Type)
in
    AddedComplexity;


// ------------------------------------------------------------------------------
// Dimension 3: DimContract
// Source: stg_RawChurn | Extracts distinct contracts, adds commitment & risk profile
// ------------------------------------------------------------------------------
shared DimContract = let
    Source = stg_RawChurn,
    SelectContract = Table.SelectColumns(Source, {"Contract"}),
    Deduplicated = Table.Distinct(SelectContract),
    
    // Add commitment duration and risk rating
    AddedMonths = Table.AddColumn(Deduplicated, "Commitment_Months", each
        if [Contract] = "Month-to-month" then 1
        else if [Contract] = "One year" then 12
        else 24, Int64.Type),
    AddedRisk = Table.AddColumn(AddedMonths, "Risk_Category", each
        if [Contract] = "Month-to-month" then "High Risk (Elastic Churn)"
        else if [Contract] = "One year" then "Moderate Risk"
        else "Low Risk (Locked In)", type text)
in
    AddedRisk;


// ------------------------------------------------------------------------------
// Dimension 4: DimDepartment
// Source: stg_RawEmployees | Deduplicates corporate departments, adds D&I target
// ------------------------------------------------------------------------------
shared DimDepartment = let
    Source = stg_RawEmployees,
    SelectDept = Table.SelectColumns(Source, {"Department @01.07.2020"}),
    Renamed = Table.RenameColumns(SelectDept, {{"Department @01.07.2020", "Department"}}),
    Deduplicated = Table.Distinct(Renamed),
    Sorted = Table.Sort(Deduplicated, {{"Department", Order.Ascending}}),
    
    // Corporate D&I parity benchmarks
    AddedParityTarget = Table.AddColumn(Sorted, "Target_Female_Ratio", each 0.50, type number),
    AddedSponsor = Table.AddColumn(AddedParityTarget, "Executive_Sponsor", each
        if [Department] = "Operations" then "COO"
        else if [Department] = "Finance" then "CFO"
        else if [Department] = "Strategy" then "CSO"
        else if [Department] = "Human Resources" then "CHRO"
        else "VP Business Units", type text)
in
    AddedSponsor;


// ------------------------------------------------------------------------------
// Auxiliary Dimension 5: Dim_EmployeeCensus
// Source: "03 Diversity-Inclusion-Dataset.xlsx" -> Backing 1
// ------------------------------------------------------------------------------
shared Dim_EmployeeCensus = let
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    SheetData = Source{[Item="Backing 1", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(SheetData, [PromoteAllScalars=true]),
    Cleaned = Table.TransformColumnNames(Promoted, Text.Trim),
    Typed = Table.TransformColumnTypes(Cleaned, {
        {"GRADE", Int64.Type},
        {"FUNCTION", type text},
        {"MALE_COUNT", Int64.Type},
        {"FEMALE_COUNT", Int64.Type}
    })
in
    Typed;


// ------------------------------------------------------------------------------
// Auxiliary Dimension 6: Dim_CareerLadder
// Source: "03 Diversity-Inclusion-Dataset.xlsx" -> Backing 2
// ------------------------------------------------------------------------------
shared Dim_CareerLadder = let
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    SheetData = Source{[Item="Backing 2", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(SheetData, [PromoteAllScalars=true]),
    Cleaned = Table.TransformColumnNames(Promoted, Text.Trim),
    Typed = Table.TransformColumnTypes(Cleaned, {
        {"JOB_LEVEL", Int64.Type},
        {"LEVEL_NAME", type text},
        {"MIN_TENURE_YEARS", type number},
        {"TARGET_PROMOTION_RATE", type number}
    })
in
    Typed;


// ------------------------------------------------------------------------------
// Auxiliary Dimension 7: Dim_NationalityCensus
// Source: "03 Diversity-Inclusion-Dataset.xlsx" -> Backing 3
// ------------------------------------------------------------------------------
shared Dim_NationalityCensus = let
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    SheetData = Source{[Item="Backing 3", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(SheetData, [PromoteAllScalars=true]),
    Cleaned = Table.TransformColumnNames(Promoted, Text.Trim),
    Typed = Table.TransformColumnTypes(Cleaned, {
        {"CITIZENSHIP_CODE", type text},
        {"NATIONALITY_LABEL", type text},
        {"SWISS_STATUS", type text}
    })
in
    Typed;


// ------------------------------------------------------------------------------
// Auxiliary Dimension 8: Dim_PRA_Equity
// Source: "03 Diversity-Inclusion-Dataset.xlsx" -> Backing 4
// ------------------------------------------------------------------------------
shared Dim_PRA_Equity = let
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    SheetData = Source{[Item="Backing 4", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(SheetData, [PromoteAllScalars=true]),
    Cleaned = Table.TransformColumnNames(Promoted, Text.Trim),
    Typed = Table.TransformColumnTypes(Cleaned, {
        {"RATING_SCORE", Int64.Type},
        {"RATING_LABEL", type text},
        {"QUOTA_DISTRIBUTION_PCT", type number}
    })
in
    Typed;
