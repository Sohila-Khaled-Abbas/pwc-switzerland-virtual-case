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
// Note: Empty column A is dropped by Power Query, so Col B=Column1, Col C=Column2
// ------------------------------------------------------------------------------
shared Dim_CareerLadder = let
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    SheetData = Source{[Item="Backing 2", Kind="Sheet"]}[Data],
    FilteredRows = Table.SelectRows(SheetData, each [Column1] <> null),
    SelectedCols = Table.SelectColumns(FilteredRows, {"Column1", "Column2"}),
    Renamed = Table.RenameColumns(SelectedCols, {
        {"Column1", "Base_Job_Level"},
        {"Column2", "Target_Promotion_Level"}
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
// Note: Empty cols A-B are dropped by Power Query, so Col C=Column1, Col D=Column2, Col E=Column3
// ------------------------------------------------------------------------------
shared Dim_NationalityCensus = let
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    SheetData = Source{[Item="Backing 3", Kind="Sheet"]}[Data],
    FilteredRows = Table.SelectRows(SheetData, each [Column1] <> null and Value.Is([Column1], type number)),
    SelectedCols = Table.SelectColumns(FilteredRows, {"Column1", "Column2", "Column3"}),
    Renamed = Table.RenameColumns(SelectedCols, {
        {"Column1", "Country_ID"},
        {"Column2", "Nationality"},
        {"Column3", "Employee_Count"}
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
// Reconciliation Note:
// Excel Col C = Column1 (Department)
// Excel Col D = Column2 (Blank spacer column - dropped)
// Excel Col E = Column3 (Grade 1 in Backing 1 notation = Level 6 Junior Officer: 191 total)
// Excel Col F = Column4 (Grade 2 in Backing 1 notation = Level 5 Senior Officer: 103 total)
// Excel Col G = Column5 (Grade 3 in Backing 1 notation = Level 4 Manager: 87 total)
// Excel Col H = Column6 (Grade 4 in Backing 1 notation = Level 3 Senior Manager: 62 total)
// Excel Col I = Column7 (Grade 5 in Backing 1 notation = Level 2 Director: 38 total)
// Excel Col J = Column8 (Grade 6 in Backing 1 notation = Level 1 Executive: 19 total)
// Corporate Hierarchy Mapping:
// Level 1 = Grade_1_Executive (Column8)
// Level 2 = Grade_2_Director (Column7)
// Level 3 = Grade_3_Senior_Manager (Column6)
// Level 4 = Grade_4_Manager (Column5)
// Level 5 = Grade_5_Senior_Specialist (Column4)
// Level 6 = Grade_6_Junior_Officer (Column3)
// Total enterprise baseline: 500 personnel
// ------------------------------------------------------------------------------
shared Dim_PRA_Equity = let
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    SheetData = Source{[Item="Backing 4", Kind="Sheet"]}[Data],
    FilteredDepts = Table.SelectRows(SheetData, each 
        [Column1] = "Finance" or 
        [Column1] = "HR" or 
        [Column1] = "Internal Services" or 
        [Column1] = "Operations" or 
        [Column1] = "Sales & Marketing" or 
        [Column1] = "Strategy"
    ),
    SelectedCols = Table.SelectColumns(FilteredDepts, {
        "Column1", "Column8", "Column7", "Column6", "Column5", "Column4", "Column3"
    }),
    Renamed = Table.RenameColumns(SelectedCols, {
        {"Column1", "Department"},
        {"Column8", "Grade_1_Executive"},
        {"Column7", "Grade_2_Director"},
        {"Column6", "Grade_3_Senior_Manager"},
        {"Column5", "Grade_4_Manager"},
        {"Column4", "Grade_5_Senior_Specialist"},
        {"Column3", "Grade_6_Junior_Officer"}
    }),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Department", type text},
        {"Grade_1_Executive", Int64.Type},
        {"Grade_2_Director", Int64.Type},
        {"Grade_3_Senior_Manager", Int64.Type},
        {"Grade_4_Manager", Int64.Type},
        {"Grade_5_Senior_Specialist", Int64.Type},
        {"Grade_6_Junior_Officer", Int64.Type}
    }),
    Reordered = Table.ReorderColumns(Typed, {
        "Department",
        "Grade_1_Executive",
        "Grade_2_Director",
        "Grade_3_Senior_Manager",
        "Grade_4_Manager",
        "Grade_5_Senior_Specialist",
        "Grade_6_Junior_Officer"
    })
in
    Reordered;
