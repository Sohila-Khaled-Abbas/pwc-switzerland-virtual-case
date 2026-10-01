// ==============================================================================
// 01_staging_queries.m
// PwC Switzerland Virtual Case Experience: Power Query Staging Layer
// Engine: Power Query M (Microsoft Excel & Power BI)
// Pattern: Parameterized Staging with Schema Enforcement & Robust Type Handling
// ==============================================================================

section Staging;

// ------------------------------------------------------------------------------
// Staging Query 1: stg_RawCalls
// Connects to "01 Call-Center-Dataset.xlsx", extracts Sheet1, applies schema types
// ------------------------------------------------------------------------------
shared stg_RawCalls = let
    // 1. Ingest raw workbook binary from project data directory
    Source = Excel.Workbook(File.Contents("data/01 Call-Center-Dataset.xlsx"), null, true),
    
    // 2. Select the primary call transactions sheet
    Sheet1_Sheet = Source{[Item="Sheet1", Kind="Sheet"]}[Data],
    
    // 3. Promote top row to column headers
    PromotedHeaders = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    
    // 4. Clean column whitespace and trim string fields
    CleanedColumnNames = Table.TransformColumnNames(PromotedHeaders, Text.Trim),
    
    // 5. Enforce explicit strong datatypes (Preserving operational nulls)
    TypedTable = Table.TransformColumnTypes(CleanedColumnNames, {
        {"Call Id", type text},
        {"Agent", type text},
        {"Date", type date},
        {"Time", type time},
        {"Topic", type text},
        {"Answered (Y/N)", type text},
        {"Resolved", type text},
        {"Speed of answer in seconds", Int64.Type},
        {"AvgTalkDuration", type time},
        {"Satisfaction rating", Int64.Type}
    })
in
    TypedTable;


// ------------------------------------------------------------------------------
// Staging Query 2: stg_RawChurn
// Connects to "02 Churn-Dataset.xlsx", trims strings, fixes 11 blank TotalCharges
// ------------------------------------------------------------------------------
shared stg_RawChurn = let
    // 1. Ingest customer subscription workbook
    Source = Excel.Workbook(File.Contents("data/02 Churn-Dataset.xlsx"), null, true),
    
    // 2. Extract primary sheet data (PwC raw sheet tab is named "01 Churn-Dataset")
    DataSheet = Source{[Item="01 Churn-Dataset", Kind="Sheet"]}[Data],
    
    // 3. Promote headers
    PromotedHeaders = Table.PromoteHeaders(DataSheet, [PromoteAllScalars=true]),
    
    // 4. Normalize column headers
    CleanedHeaders = Table.TransformColumnNames(PromotedHeaders, Text.Trim),
    
    // 5. Replace empty string blanks in TotalCharges with operational null before conversion
    ReplacedBlanks = Table.ReplaceValue(CleanedHeaders, "", null, Replacer.ReplaceValue, {"TotalCharges"}),
    ReplacedSpaces = Table.ReplaceValue(ReplacedBlanks, " ", null, Replacer.ReplaceValue, {"TotalCharges"}),
    
    // 6. Enforce explicit numeric and text types
    TypedTable = Table.TransformColumnTypes(ReplacedSpaces, {
        {"customerID", type text},
        {"gender", type text},
        {"SeniorCitizen", Int64.Type},
        {"Partner", type text},
        {"Dependents", type text},
        {"tenure", Int64.Type},
        {"PhoneService", type text},
        {"MultipleLines", type text},
        {"InternetService", type text},
        {"OnlineSecurity", type text},
        {"OnlineBackup", type text},
        {"DeviceProtection", type text},
        {"TechSupport", type text},
        {"StreamingTV", type text},
        {"StreamingMovies", type text},
        {"Contract", type text},
        {"PaperlessBilling", type text},
        {"PaymentMethod", type text},
        {"MonthlyCharges", type number},
        {"TotalCharges", type number},
        {"numAdminTickets", Int64.Type},
        {"numTechTickets", Int64.Type},
        {"Churn", type text}
    })
in
    TypedTable;


// ------------------------------------------------------------------------------
// Staging Query 3: stg_RawEmployees
// Connects to "03 Diversity-Inclusion-Dataset.xlsx", extracts Pharma Group AG
// ------------------------------------------------------------------------------
shared stg_RawEmployees = let
    // 1. Ingest HR diversity dataset
    Source = Excel.Workbook(File.Contents("data/03 Diversity-Inclusion-Dataset.xlsx"), null, true),
    
    // 2. Extract Pharma Group AG master personnel sheet
    PharmaSheet = Source{[Item="Pharma Group AG", Kind="Sheet"]}[Data],
    
    // 3. Promote headers
    PromotedHeaders = Table.PromoteHeaders(PharmaSheet, [PromoteAllScalars=true]),
    
    // 4. Normalize header text
    CleanedHeaders = Table.TransformColumnNames(PromotedHeaders, Text.Trim),
    
    // 5. Enforce typed schema (Matching exact PwC headers in Pharma Group AG)
    TypedTable = Table.TransformColumnTypes(CleanedHeaders, {
        {"Employee ID", Int64.Type},
        {"Gender", type text},
        {"Age group", type text},
        {"Department @01.07.2020", type text},
        {"Last Department in FY20", type text},
        {"Job Level before FY20 promotions", type text},
        {"Job Level after FY20 promotions", type text},
        {"Promotion in FY21?", type text},
        {"FY20 Performance Rating", Int64.Type},
        {"FY19 Performance Rating", Int64.Type},
        {"FY20 leaver?", type text},
        {"In base group for Promotion FY21", type text},
        {"In base group for turnover FY20", type text},
        {"Nationality 1", type text}
    })
in
    TypedTable;
