# VBA-Axis_Bank
    Sub Macro_Axis_Statement()

    'Data From PDF

    ActiveWorkbook.Queries.Add Name:="Table005 (Page 2)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Pdf.Tables(File.Contents(""C:\Users\KAVYA JAIN\Documents\Kavita Mam Project\AXIS BANK  Statement for September 2023-unlocked.pdf""), [Implementation=""1.3""])," & Chr(13) & "" & Chr(10) & "    Table005 = Source{[Id=""Table005""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Table005, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers""" & _
        ",{{""Txn Date"", type text}, {""Transaction"", type text}, {""Withdrawals"", Int64.Type}, {""Deposits"", Int64.Type}, {""Balance"", type number}, {""Other Information"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Queries.Add Name:="Table006 (Page 3)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Pdf.Tables(File.Contents(""C:\Users\KAVYA JAIN\Documents\Kavita Mam Project\AXIS BANK  Statement for September 2023-unlocked.pdf""), [Implementation=""1.3""])," & Chr(13) & "" & Chr(10) & "    Table006 = Source{[Id=""Table006""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Table006,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type t" & _
        "ext}, {""Column5"", type number}, {""Column6"", Int64.Type}, {""Column7"", type number}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Queries.Add Name:="Table007 (Page 4)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Pdf.Tables(File.Contents(""C:\Users\KAVYA JAIN\Documents\Kavita Mam Project\AXIS BANK  Statement for September 2023-unlocked.pdf""), [Implementation=""1.3""])," & Chr(13) & "" & Chr(10) & "    Table007 = Source{[Id=""Table007""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Table007,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", Int64.Type}, {""Column4"", Int64" & _
        ".Type}, {""Column5"", type number}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Queries.Add Name:="Table008 (Page 5)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Pdf.Tables(File.Contents(""C:\Users\KAVYA JAIN\Documents\Kavita Mam Project\AXIS BANK  Statement for September 2023-unlocked.pdf""), [Implementation=""1.3""])," & Chr(13) & "" & Chr(10) & "    Table008 = Source{[Id=""Table008""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Table008,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", Int64.Type}, {""Column4"", Int64" & _
        ".Type}, {""Column5"", type number}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    Workbooks("Axis_Statement").Connections.Add2 "Query - Table005 (Page 2)", _
        "Connection to the 'Table005 (Page 2)' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Table005 (Page 2);Extended Properties=" _
        , """Table005 (Page 2)""", 6, True, False
    Workbooks("Axis_Statement").Connections.Add2 "Query - Table006 (Page 3)", _
        "Connection to the 'Table006 (Page 3)' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Table006 (Page 3);Extended Properties=" _
        , """Table006 (Page 3)""", 6, True, False
    Workbooks("Axis_Statement").Connections.Add2 "Query - Table007 (Page 4)", _
        "Connection to the 'Table007 (Page 4)' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Table007 (Page 4);Extended Properties=" _
        , """Table007 (Page 4)""", 6, True, False
    Workbooks("Axis_Statement").Connections.Add2 "Query - Table008 (Page 5)", _
        "Connection to the 'Table008 (Page 5)' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Table008 (Page 5);Extended Properties=" _
        , """Table008 (Page 5)""", 6, True, False
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("Query - Table005 (Page 2)"), Destination:=Range("$A$1")). _
        TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = 1
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "Table005__Page_2"
        .Refresh
    End With
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("Query - Table006 (Page 3)"), Destination:=Range("$A$1")). _
        TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = 1
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "Table006__Page_3"
        .Refresh
    End With
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("Query - Table007 (Page 4)"), Destination:=Range("$A$1")). _
        TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = 1
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "Table007__Page_4"
        .Refresh
    End With
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("Query - Table008 (Page 5)"), Destination:=Range("$A$1")). _
        TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = 1
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "Table008__Page_5"
        .Refresh
    End With
    
    'Formatting
    Worksheets("sheet1").Move Before:=Worksheets("sheet5")
    Worksheets("sheet3").Move Before:=Worksheets("sheet5")
    Worksheets("sheet4").Move Before:=Worksheets("sheet5")
    Worksheets("sheet2").Move Before:=Worksheets("sheet3")
    
    ActiveSheet.Next.Select
    Range("a1").Select
    ActiveCell.Formula2R1C1 = "Txn Date"
    Range("b1").Select
    ActiveCell.Formula2R1C1 = "Transaction"
    Range("c1").Select
    ActiveCell.Formula2R1C1 = "Withdrawals"
    Range("d1").Select
    ActiveCell.Formula2R1C1 = "Deposits"
    Range("e1").Select
    ActiveCell.Formula2R1C1 = "Balance"
    Range("f1").Select
    ActiveCell.Formula2R1C1 = "Other Information"
    ActiveSheet.Next.Select
    Range("a1").Select
    ActiveCell.Formula2R1C1 = "Txn Date"
    Range("b1").Select
    ActiveCell.Formula2R1C1 = "Transaction"
    Range("c1").Select
    ActiveCell.Formula2R1C1 = "Withdrawals"
    Range("d1").Select
    ActiveCell.Formula2R1C1 = "Deposits"
    Range("e1").Select
    ActiveCell.Formula2R1C1 = "Balance"
    Range("f1").Select
    ActiveCell.Formula2R1C1 = "Other Information"
    
    ActiveSheet.Next.Select
    Range("a1").Select
    ActiveCell.Formula2R1C1 = "Txn Date"
    Range("b1").Select
    ActiveCell.Formula2R1C1 = "Transaction"
    Range("c1").Select
    ActiveCell.Formula2R1C1 = "Withdrawals"
    Range("d1").Select
    ActiveCell.Formula2R1C1 = "Deposits"
    Range("e1").Select
    ActiveCell.Formula2R1C1 = "Balance"
    Range("f1").Select
    ActiveCell.Formula2R1C1 = "Other Information"
    Worksheets("sheet1").Select\
    
    'Output
    outputfolder = "C:\Users\KAVYA JAIN\Documents"
    Workbooks("Axis_Statement").SaveAs Filename:=outputfolder & "\" & "Axis Statement Done.xlsm", FileFormat:=52
    MsgBox "Done", vbApplicationModal + vbInformation
    
    End Sub









