// ==============================================================================
// 04_calendar_generator.m
// PwC Switzerland Virtual Case Experience: Dynamic Date Dimension Generator
// Engine: Power Query M (Microsoft Excel & Power BI)
// Pattern: Continuous Calendar Sequence with Fiscal & Day-of-Week Attributes
// ==============================================================================

let
    // 1. Define date boundary parameters for Q1 2021 (or dynamic min/max from Fact)
    StartDate = #date(2021, 1, 1),
    EndDate = #date(2021, 3, 31),
    
    // 2. Compute total day span
    DayCount = Duration.Days(Duration.From(EndDate - StartDate)) + 1,
    
    // 3. Generate sequential list of calendar dates
    DateList = List.Dates(StartDate, DayCount, #duration(1, 0, 0, 0)),
    
    // 4. Convert list to table
    DateTable = Table.FromList(DateList, Splitter.SplitByNothing(), {"Date"}, null, ExtraValues.Error),
    TypedDate = Table.TransformColumnTypes(DateTable, {{"Date", type date}}),
    
    // 5. Add Standard Dimensional Calendar Attributes
    AddedYear = Table.AddColumn(TypedDate, "Year", each Date.Year([Date]), Int64.Type),
    AddedQuarter = Table.AddColumn(AddedYear, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
    AddedMonth = Table.AddColumn(AddedQuarter, "Month", each Date.Month([Date]), Int64.Type),
    AddedMonthName = Table.AddColumn(AddedMonth, "Month_Name", each Date.MonthName([Date]), type text),
    AddedMonthShort = Table.AddColumn(AddedMonthName, "Month_Short", each Text.Start([Month_Name], 3), type text),
    AddedDay = Table.AddColumn(AddedMonthShort, "Day", each Date.Day([Date]), Int64.Type),
    AddedDayOfWeek = Table.AddColumn(AddedDay, "Day_Of_Week", each Date.DayOfWeek([Date], Day.Monday) + 1, Int64.Type),
    AddedDayName = Table.AddColumn(AddedDayOfWeek, "Day_Name", each Date.DayOfWeekName([Date]), type text),
    
    // 6. Operational Slicers & Flags
    AddedIsWeekend = Table.AddColumn(AddedDayName, "Is_Weekend", each 
        if [Day_Of_Week] >= 6 then "Weekend" else "Weekday", type text),
        
    AddedYearMonth = Table.AddColumn(AddedIsWeekend, "Year_Month_Sort", each 
        Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type)
in
    AddedYearMonth
