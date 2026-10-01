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
        {"Employee ID", Int64.Type},
        {"GENDER", type text},
        {"GRADE", type text},
        {"FUNCTION", type text},
        {"OC_RATE", Int64.Type},
        {"PERFORM", Int64.Type},
        {"Y_GRADE", Int64.Type},
        {"AGE", Int64.Type},
        {"Y_SERVIC", Int64.Type},
        {"Nationality", type text},
        {"Rank 2", Int64.Type}
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
    SelectedCols = Table.SelectColumns(SheetData, {"Column2", "Column3"}),
    FilteredRows = Table.SelectRows(SelectedCols, each [Column2] <> null),
    Renamed = Table.RenameColumns(FilteredRows, {
        {"Column2", "Base_Job_Level"},
        {"Column3", "Target_Promotion_Level"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Base_Job_Level", type text},
        {"Target_Promotion_Level", type text}
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
    SelectedCols = Table.SelectColumns(SheetData, {"Column3", "Column4", "Column5"}),
    FilteredRows = Table.SelectRows(SelectedCols, each [Column3] <> null and Value.Is([Column3], type number)),
    Renamed = Table.RenameColumns(FilteredRows, {
        {"Column3", "Country_ID"},
        {"Column4", "Nationality"},
        {"Column5", "Employee_Count"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Country_ID", Int64.Type},
        {"Nationality", type text},
        {"Employee_Count", Int64.Type}
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
    SelectedCols = Table.SelectColumns(SheetData, {"Column3", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10"}),
    FilteredDepts = Table.SelectRows(SelectedCols, each [Column3] <> null and [Column3] <> "Total" and [Column3] <> "F" and [Column3] <> "M"),
    Renamed = Table.RenameColumns(FilteredDepts, {
        {"Column3", "Department"},
        {"Column5", "Grade_1_Executive"},
        {"Column6", "Grade_2_Director"},
        {"Column7", "Grade_3_Senior_Manager"},
        {"Column8", "Grade_4_Manager"},
        {"Column9", "Grade_5_Senior_Specialist"},
        {"Column10", "Grade_6_Junior_Officer"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Department", type text},
        {"Grade_1_Executive", Int64.Type},
        {"Grade_2_Director", Int64.Type},
        {"Grade_3_Senior_Manager", Int64.Type},
        {"Grade_4_Manager", Int64.Type},
        {"Grade_5_Senior_Specialist", Int64.Type},
        {"Grade_6_Junior_Officer", Int64.Type}
    })
in
    Typed;
