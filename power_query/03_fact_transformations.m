// ==============================================================================
// 03_fact_transformations.m
// PwC Switzerland Virtual Case Experience: Fact Table Transformation Layer
// Engine: Power Query M (Microsoft Excel & Power BI)
// Pattern: Key Integrity, Null Auditing, Duration Normalization, Grain Alignment
// ==============================================================================

section Facts;

// ------------------------------------------------------------------------------
// Fact Table 1: Fact_Calls
// Ingests stg_RawCalls, extracts Call_Hour, converts Duration to seconds, cleans keys
// ------------------------------------------------------------------------------
shared Fact_Calls = let
    Source = stg_RawCalls,
    
    // Normalize column names to standard database snake_case / Title Case
    Renamed = Table.RenameColumns(Source, {
        {"Call Id", "Call_Id"},
        {"Answered (Y/N)", "Answered"},
        {"Speed of answer in seconds", "Speed_Of_Answer_Sec"},
        {"Satisfaction rating", "Satisfaction_Rating"}
    }),
    
    // Add Call_Hour for intraday call volume triage (VertiPaq cardinality optimization)
    AddedCallHour = Table.AddColumn(Renamed, "Call_Hour", each Time.Hour([Time]), Int64.Type),
    
    // Convert duration time to absolute seconds (Handle Time)
    AddedTalkSec = Table.AddColumn(AddedCallHour, "Talk_Duration_Sec", each 
        if [AvgTalkDuration] = null then null
        else Time.Hour([AvgTalkDuration]) * 3600 + Time.Minute([AvgTalkDuration]) * 60 + Time.Second([AvgTalkDuration]),
        Int64.Type),
        
    // Verification: Enforce non-null key constraints on Call_Id
    FilteredValidKeys = Table.SelectRows(AddedTalkSec, each [Call_Id] <> null and [Call_Id] <> "")
in
    FilteredValidKeys;


// ------------------------------------------------------------------------------
// Fact Table 2: Fact_Churn
// Ingests stg_RawChurn, handles zero-tenure blank charges, adds tenure bucket
// ------------------------------------------------------------------------------
shared Fact_Churn = let
    Source = stg_RawChurn,
    
    // Handle TotalCharges for new customers (tenure = 0 -> MonthlyCharges)
    AdjustedTotalCharges = Table.ReplaceValue(Source, each [TotalCharges], each 
        if [TotalCharges] = null and [tenure] = 0 then [MonthlyCharges]
        else [TotalCharges], 
        Replacer.ReplaceValue, {"TotalCharges"}),
        
    // Add Tenure Cohort classification
    AddedTenureBucket = Table.AddColumn(AdjustedTotalCharges, "Tenure_Cohort", each
        if [tenure] <= 12 then "1. 0 - 12 Months (High Churn)"
        else if [tenure] <= 24 then "2. 13 - 24 Months"
        else if [tenure] <= 48 then "3. 25 - 48 Months"
        else "4. 49+ Months (Loyal Core)", type text),
        
    // Add Ticket Severity Index
    AddedTicketIndex = Table.AddColumn(AddedTenureBucket, "Total_Service_Tickets", each
        [numAdminTickets] + [numTechTickets], Int64.Type)
in
    AddedTicketIndex;


// ------------------------------------------------------------------------------
// Fact Table 3: Fact_Employees
// Ingests stg_RawEmployees, standardizes keys, computes promotion grade change
// ------------------------------------------------------------------------------
shared Fact_Employees = let
    Source = stg_RawEmployees,
    
    // Standardize column naming
    Renamed = Table.RenameColumns(Source, {
        {"Employee ID", "Employee_ID"},
        {"Age group", "Age_Group"},
        {"Department @01.07.2020", "Department"},
        {"Job Level at 01.07.2020", "Job_Level_Baseline"},
        {"Job Level after FY20 promotions", "Job_Level_After_Promotions"},
        {"Promoted in FY21?", "Promoted_FY21"},
        {"FY20 Performance Rating", "FY20_Rating"},
        {"FY19 Performance Rating", "FY19_Rating"},
        {"FY20 leaver?", "FY20_Leaver"},
        {"In base group?", "In_Base_Group"},
        {"Nationality 1", "Nationality"}
    }),
    
    // Add Promotion Advancement Numeric Delta (Executive Grade Velocity)
    AddedLevelChange = Table.AddColumn(Renamed, "Grade_Change_Delta", each
        [Job_Level_Baseline] - [Job_Level_After_Promotions], Int64.Type),
        
    // Classify Executive Tier Group
    AddedTierGroup = Table.AddColumn(AddedLevelChange, "Executive_Tier", each
        if [Job_Level_Baseline] = 1 then "C-Suite / Executive Board"
        else if [Job_Level_Baseline] = 2 then "Director"
        else if [Job_Level_Baseline] <= 4 then "Middle Management"
        else "Staff / Individual Contributor", type text)
in
    AddedTierGroup;
