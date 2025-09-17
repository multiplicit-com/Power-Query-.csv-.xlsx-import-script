let

    // Folder source
    //Point this at your 
    FolderPath = "c:\user\folder location",

    // - .xlsx: read the first sheet
    // - .csv : parse CSV (UTF-8, comma), promote headers

    FileToTable = (content as binary, ext as text) as table =>
        let
            LowerExt = Text.Lower(ext),
            Result =
                if LowerExt = ".csv" then
                    let
                        CsvRaw   = Csv.Document(content, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                        CsvTable = Table.PromoteHeaders(CsvRaw, [PromoteAllScalars = true])
                    in
                        CsvTable
                else if LowerExt = ".xlsx" then
                    let
                        WB          = Excel.Workbook(content, null, true),
                        FirstSheet  = if Table.RowCount(WB) > 0 then WB{0}[Data] else #table({},{}),
                        XLTable     = Table.PromoteHeaders(FirstSheet, [PromoteAllScalars = true])
                    in
                        XLTable
                else
                    // Unsupported type -> empty table
                    #table({}, {})
        in
            Result,


    // Load files
    SourceFolder  = Folder.Files(FolderPath),

    // Only .xlsx and .csv
    FilteredFiles = Table.SelectRows(SourceFolder, each List.Contains({".xlsx", ".csv"}, Text.Lower([Extension]))),

    // Parse each file into a table
    WithData      = Table.AddColumn(FilteredFiles, "Data", each FileToTable([Content], [Extension])),

    // Dynamically expand columns from the first non-empty table
    SampleTable   = List.First(List.Select(WithData[Data], each _ <> null and Table.ColumnCount(_) > 0), null),
    ColNames      = if SampleTable = null then {} else Table.ColumnNames(SampleTable),

    Expanded      = if List.Count(ColNames) = 0
                    then #table({}, {})
                    else Table.ExpandTableColumn(WithData, "Data", ColNames),

    #"Promoted Headers" = Expanded
in
    #"Promoted Headers"

// Additional transforms after this point
