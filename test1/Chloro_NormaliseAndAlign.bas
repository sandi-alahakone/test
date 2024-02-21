Attribute VB_Name = "Chloro_NormaliseAndAlign"
Sub InternalStandard_Scale()
Attribute InternalStandard_Scale.VB_ProcData.VB_Invoke_Func = "S\n14"
'
'
'NOTE: this macro is currently set for Bromobenzene IS
'      if you are using chlorobenzene-d5 you need to change the sections:
'        1) Add new sheet  to beginning of workbook (first section) [Line 12-13]
'        2) Change sheet name (last section) [Line 69-72] <- you can delete this section
'
'
'Add new sheet  to beginning of workbook
' IF YOU ARE USING D5-CHLOROBENZENE CHANGE THIS SECTION
    Sheets.Add(Before:=Sheets(1)).Name = "Chlorobenzene-d5"
    Worksheets("Chlorobenzene-d5").Activate
    'Sheets.Add(Before:=Sheets(1)).Name = "Chlorobenzene-d5"
    'Worksheets("Chlorobenzene-d5").Activate
    
 ' Insert sheet names into column A
    Columns(1).Insert
    For i = 1 To Sheets.Count
    Cells(i, 1) = Sheets(i).Name
    Next i


 'Find all IS values
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R1C1,INDIRECT(""'""&RC1&""'!""&""B1:K9000""),10,FALSE)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & Range("A" & Rows.Count).End(xlUp).Row)
 
 'Highlight maximum value
    Range("B2:B" & Range("A" & Rows.Count).End(xlUp).Row).Select
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 1
        .Percent = False
    End With
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249946592608417
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False

 'Copy maximum value
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=MAX(R2C2:R9000C2)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & Range("A" & Rows.Count).End(xlUp).Row)

 'Calculate scale
    Range("D2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-2]"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("A1:D" & Range("A" & Rows.Count).End(xlUp).Row).Select
    'Paste Values i.e. remove formulas
        Range("A1:D" & Range("A" & Rows.Count).End(xlUp).Row).Copy
        Range("A1").PasteSpecial Paste:=xlPasteValues
 
 'Change sheet name

    
    
End Sub

Sub NormaliseSamples()
Attribute NormaliseSamples.VB_ProcData.VB_Invoke_Func = "T\n14"
'

Application.ScreenUpdating = False


    ' Remove columns: Peak, RT, Height, Mass, Etc
        Range("A:A,C:J,L:M").Select
        Range("E1").Activate
        Selection.Delete Shift:=xlToLeft
        
    ' Consolidate Compounds
        ActiveSheet.Range("D1").consolidate Sources:= _
                              "C1:C2", Function:=xlAverage, _
                              TopRow:=True, LeftColumn:=True, CreateLinks:=False
                              
                        Range("A:C").Select
            Range("C1").Activate
            Selection.Delete Shift:=xlToLeft
         Range("C2").Select
         
    ' Insert IS Scale
        Range("C2").Select
            ActiveCell.FormulaR1C1 = ""
            Range("C2").Select
            ActiveCell.FormulaR1C1 = _
                "=VLOOKUP(MID(CELL(""filename"",R[-1]C[-2]),FIND(""]"",CELL(""filename"",R[-1]C[-2]))+1,255),'Chlorobenzene-d5'!R2C1:R9000C4,4,FALSE)"
            Range("C2").Select
            Selection.AutoFill Destination:=Range("C2:C" & Range("A" & Rows.Count).End(xlUp).Row)
        'Paste as values i.e. remove formulas
            Range("A1").CurrentRegion.Select
            Range("A1").CurrentRegion.Copy
            Range("A1").PasteSpecial Paste:=xlPasteValues
            Range("D2").Select
            
    ' Normalize To IS
        Range("D2").Select
        Application.CutCopyMode = False
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
        Range("D2").Select
        Selection.AutoFill Destination:=Range("D2:D" & Range("C" & Rows.Count).End(xlUp).Row)
        Range(Selection, Selection.End(xlDown)).Select
            'Paste as values i.e. remove formulas
            Range("A1").CurrentRegion.Select
            Range("A1").CurrentRegion.Copy
            Range("A1").PasteSpecial Paste:=xlPasteValues
            
    'Remove Solvents
        With ActiveSheet
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "Methyl Alcohol"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "Acetone"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "*Analyte*"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "Carbon dioxide"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "*siloxane*"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "*Peak*"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "Ethanol"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "Isopropyl Alcohol"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
            .AutoFilterMode = False
            With Range("A1", Range("A" & Rows.Count).End(xlUp))
                .AutoFilter 1, "Total"
                On Error Resume Next
                .Offset(1).SpecialCells(12).EntireRow.Delete
            End With
           .AutoFilterMode = False
           
           ' With Range("A1", Range("A" & Rows.Count).End(xlUp))
           '     .AutoFilter 1, "[ANY OTHER COMPOUND YOU WANT TO REMOVE]"
           '     On Error Resume Next
           '     .Offset(1).SpecialCells(12).EntireRow.Delete
           ' End With
           ' .AutoFilterMode = False
        
    ' Transpose columns to rows
        Range("A1").CurrentRegion.Select
        Range("A1").CurrentRegion.Copy
            ActiveWindow.SmallScroll Down:=-222
            Range("F1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=True
        Columns("A:E").Select
            Range("E1").Activate
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
        Rows("2:3").Select
            Selection.Delete Shift:=xlUp
        'Add sample name to row 2 based on sheet name
            Range("A1").Select
                ActiveSheet.[A1] = "Sample"
            Range("A2").Select
                ActiveSheet.[A2] = ActiveSheet.Name
            
    ' Transform data into table
        Dim objTable As ListObject
            Range("A1", Range("A1").End(xlToRight).End(xlDown)).Select
            Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        'Change table name to sample name
        Dim tblName As String
            tblName = Range("A2").Text
        'Adds "Dnr" to begining of table name <- Table names can not begin with a number
        'Delete " "Dnr" & " if unnecessary
            With ActiveSheet
                .ListObjects(1).Name = tblName
            
            End With
            End With

Application.ScreenUpdating = True

End Sub
Sub AlignSamples()
Attribute AlignSamples.VB_ProcData.VB_Invoke_Func = "A\n14"
'
     ' Create query connections from all tables in workbook
'
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim lo As ListObject
        Dim sName As String
        Dim sFormula As String
        Dim wq As WorkbookQuery
        Dim bExists As Boolean
        
        Dim i As Long
  
         Set wb = ActiveWorkbook
    
            'Loop sheets and tables
            For Each ws In ActiveWorkbook.Worksheets
            For Each lo In ws.ListObjects
        
             sName = lo.Name
                sFormula = "Excel.CurrentWorkbook(){[Name=""" & sName & """]}[Content]"
        
        'Check if query exists
        'If query does exist it will not create duplicates
        bExists = False
        For Each wq In wb.Queries
          If InStr(1, wq.Formula, sFormula) > 0 Then
            bExists = True
          End If
        Next wq
        
        'Add query if it does not exist
        If bExists = False Then
        
          'Add query
          wb.Queries.Add Name:=sName, _
                         Formula:="let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""" & sName & """]}[Content]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Source"
          'Add connection
          wb.Connections.Add2 Name:="Query - " & sName, _
                              Description:="Connection to the '" & sName & "' query in the workbook.", _
                              ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & sName & ";Extended Properties=""""", _
                              CommandText:="SELECT * FROM [" & sName & "]", _
                              lCmdtype:=2, _
                              CreateModelConnection:=False, _
                              ImportRelationships:=False
                                
          'Count connections
          i = i + 1
          
        End If
      Next lo
    Next ws
'
'
     'Append Connections
        'get file name to name new table and worksheet
            Dim FileNameExt As String
            Dim FileName As String
            Dim Aligned As String
        FileNameExt = ActiveWorkbook.Name
        FileName = Left(FileNameExt, InStr(FileNameExt, ".") - 1)
        Aligned = FileName & "ALIGNED"

        'Add new worksheet for aligned table
            Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            ws.Name = Aligned

        'Create array for all table names in worksheet
            Dim tblArray() As String
            Dim ArraySize As Integer
            Dim Tbl As ListObject
                ArraySize = 0
                ReDim tblArray(0 To 0)

            For Each ws In ActiveWorkbook.Worksheets
                For Each Tbl In ws.ListObjects
                    ReDim Preserve tblArray(ArraySize) As String
                    tblArray(UBound(tblArray)) = Tbl.Name
                    ArraySize = ArraySize + 1
                Next Tbl
            Next ws

            Range("A1") = Join(tblArray, ",")

        'Append tables
    Dim sourceFullName As String
        
        With ActiveSheet
            sourceFullName = .Range("A1").Value
        End With

            'Select all queries and sort columns in alpha-numeric order
        ActiveWorkbook.Queries.Add Name:=Aligned, Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Table.Combine({" & sourceFullName & "})," & Chr(10) & "     LeftColumns = {""Sample""}," & Chr(10) & "    ReorderList = LeftColumns &" & Chr(10) & "                  List.Sort(List.Difference(Table.ColumnNames(Source),LeftColumns),Order.Ascending)," & Chr(10) & "    Reorder = Table.ReorderColumns(Sour" & _
        "ce, ReorderList)" & Chr(10) & "in" & Chr(10) & "    Reorder" & _
        ""

    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & Aligned & """;Extended Properties=""""" _
        , Destination:=Range("$A$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & Aligned & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.Name = Aligned
        .Refresh BackgroundQuery:=False
    End With
            'Delete table name array
     Rows("1:2").Select
                Selection.Delete Shift:=xlUp
                
        'Select all blanks and add zeros
            Range("A1").CurrentRegion.Select
                Selection.SpecialCells(xlCellTypeBlanks).Select
                Selection.FormulaR1C1 = "0"
            Range("A1").Select
End Sub



