Attribute VB_Name = "Module1"
Sub MarketingTickets()
Attribute MarketingTickets.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MarketingTickets Macro
'

'
    Set first = ActiveSheet
    Set report = ActiveWorkbook
    Dim rname As String
    rname = ActiveWorkbook.Name
    
    Dim strFile As String
    strFile = Application.GetOpenFilename
    On Error GoTo ErrHandler
    Workbooks.Open strFile

    Set datafile = ActiveWorkbook
    Set wsd = ActiveSheet
    
    wsd.Select
    wsd.Copy After:=Workbooks(rname).Sheets(2)
    datafile.Close SaveChanges:=False
    report.Activate
    
    Range("B4:B5").Select
    Selection.Replace What:="/", Replacement:="-", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False

    If sheetExists(CStr("Mark " & ActiveSheet.Range("B4"))) Then
'If ws exists
        Set wstemp = ActiveSheet
        Set wsr = Sheets(CStr("Mark " & ActiveSheet.Range("B4")))
        Rows("1:3").Select
        Selection.Delete Shift:=xlUp
        Columns("A:O").Select
        Selection.UnMerge
        Range("A1").Select
        Selection.End(xlDown).Select
        Rows(ActiveCell.Row).Delete
        ActiveCell.Offset(rowOffset:=-1, columnOffset:=0).Activate
        Rows(ActiveCell.Row).Delete
        Columns("F:F").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1:O1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        wsr.Activate
        Range("A4").Select
        Selection.End(xlDown).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range("F2").Select
        Selection.End(xlDown).Select
        Selection.Copy
        ActiveCell.Offset(rowOffset:=0, columnOffset:=-1).Activate
        Selection.End(xlDown).Select
        ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
        Range(Selection, Selection.End(xlUp)).Select
        ActiveSheet.Paste
        
        Range("Q2:AA2").Select
        Selection.Copy
        Range("A2").Select
        Selection.End(xlDown).Select
        Range("Q" & (ActiveCell.Row)).Select
        Range(Selection, Selection.End(xlUp)).Select
        ActiveSheet.Paste
        Range("P3").Select
        

        
        
        Application.DisplayAlerts = False
        wstemp.Delete
        Application.DisplayAlerts = True
        
'New ws:
    Else: Application.EnableEvents = False
          ActiveSheet.Name = "Mark " & ActiveSheet.Range("B4")
          Application.EnableEvents = True
          Set wsr = ActiveSheet
          Rows("1:2").Select
          Selection.Delete Shift:=xlUp
          
          'Formating of pasted data:
          Range("A1").Select
          Range("A1:O1").Select
          With Selection.Interior
              .ThemeColor = xlThemeColorAccent5
              .TintAndShade = -0.249977111117893
              .PatternTintAndShade = 0
          End With
          With Selection.Font
              .ThemeColor = xlThemeColorDark1
              .TintAndShade = 0
          End With
          Columns("A:O").Select
          Range("A2").Activate
          Selection.Borders(xlInsideVertical).LineStyle = xlNone
          Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
          
          'add FCIDs to columnt F:
          Range("A1").Select
          Selection.End(xlDown).Select
          Rows(ActiveCell.Row).Delete
          ActiveCell.Offset(rowOffset:=-1, columnOffset:=0).Activate
          Rows(ActiveCell.Row).Delete
          Columns("F:F").Select
          Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
          Range("F1").Select
          ActiveCell.FormulaR1C1 = "FCID"
          Range("F2").Select
          ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],Map!C[-5]:C[-4],2,0),"""")"
          Range("F2").Select
          Selection.Copy
          ActiveCell.Offset(rowOffset:=0, columnOffset:=-1).Activate
          Selection.End(xlDown).Select
          ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
          Range(Selection, Selection.End(xlUp)).Select
          ActiveSheet.Paste
          Range("E3").Select
          Selection.Copy
          Range("F3").Select
          Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          Range("A1").Select
          
          '    add filter to line 1
          Columns("O:O").Select
          Selection.UnMerge
          Range("F1").Select
          Selection.AutoFilter
          report.ActiveSheet.AutoFilter.Sort.SortFields _
              .Clear
          report.ActiveSheet.AutoFilter.Sort.SortFields _
              .Add2 Key:=Range("F3:F47"), SortOn:=xlSortOnValues, Order:=xlAscending, _
              DataOption:=xlSortNormal
          With report.ActiveSheet.AutoFilter.Sort
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlTopToBottom
              .SortMethod = xlPinYin
              .Apply
          End With
          
            '    Headers for wsar
                Cells(1, 16).FormulaR1C1 = "Ticket"
                Cells(1, 17).FormulaR1C1 = "Name"
                Cells(1, 18).FormulaR1C1 = "Street"
                Cells(1, 19).FormulaR1C1 = "Town"
                Cells(1, 20).FormulaR1C1 = "State"
                Cells(1, 21).FormulaR1C1 = "Zip"
                Cells(1, 22).FormulaR1C1 = "Phone"
                Cells(1, 23).FormulaR1C1 = "MRCH"
                Cells(1, 24).FormulaR1C1 = "Count"
                Cells(1, 25).FormulaR1C1 = "Shipping"
                Cells(1, 26).FormulaR1C1 = "Comment"
            '    Formulas for wsar
                Cells(2, 16).FormulaR1C1 = "=TRIM(RC[-15])"
                Cells(2, 17).FormulaR1C1 = "=TRIM(RC[-14])"
                Cells(2, 18).FormulaR1C1 = "=VLOOKUP(RC[-12],Map!C[-16]:C[-14],3,0)"
                Cells(2, 19).FormulaR1C1 = "=VLOOKUP(RC[-13],Map!C[-17]:C[-14],4,0)"
                Cells(2, 20).FormulaR1C1 = "=VLOOKUP(RC[-14],Map!C[-18]:C[-14],5,0)"
                Cells(2, 21).FormulaR1C1 = "=VLOOKUP(RC[-15],Map!C[-19]:C[-14],6,0)"
                Cells(2, 22).FormulaR1C1 = "=TRIM(RC[-13])"
                Cells(2, 23).FormulaR1C1 = "=TRIM(RC[-13])"
                Cells(2, 24).FormulaR1C1 = "=RC[-13]"
                Cells(2, 25).FormulaR1C1 = "Ground"
                    ' Add to all rows
                    Range("P2:Z2").Select
                    Selection.Copy
                    Range("A2").Select
                    Selection.End(xlDown).Select
                    Range("P" & (ActiveCell.Row)).Select
                    Range(Selection, Selection.End(xlUp)).Select
                    ActiveSheet.Paste
                    Range("P3").Select
                            ' Format
                            Columns("O:O").Select
                            Range("O2").Activate
                            Selection.Copy
                            Columns("P:Z").Select
                            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                SkipBlanks:=False, Transpose:=False
                                        Application.CutCopyMode = False
                                        Columns("P:P").Select
                                        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                                        Columns("P:P").Select
                                            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                                            With Selection.Interior
                                                .Pattern = xlNone
                                                .TintAndShade = 0
                                                .PatternTintAndShade = 0
                                            End With
                                        Range("Q1:AA1").Select
          
                ' Format AR Header
                Range("Q1:AA1").Select
                With Selection.Interior
                    .ThemeColor = xlThemeColorLight1
                End With
                With Selection.Font
                    .ThemeColor = xlThemeColorDark2
                End With
                
                'Add button to create a new file
                ActiveSheet.Buttons.Add(1800, 9, 126, 18.75).Select
                Selection.OnAction = "CS Tickets.XLSM!AlliedRequestsFile"
                Selection.Characters.Text = "Create ZD&AR File"

                Range("Q1").Select
    End If
    

    


    
'    ' Add MRCH to codes that start with d318
'        lastRow = Range("A" & Rows.Count).End(xlUp).Row
'        colNum = 24
'        For Each c In Range(Cells(4, colNum), Cells(lastRow, colNum))
'          If Left(c.Value, 4) = "D318" _
'            Or Left(c.Value, 4) <> "MRCH" Then
'            c.Value = "MRCH-" & c.Value
'            End If
'        Next
    
        'Clean comments:
        lastRow = Range("O" & Rows.Count).End(xlUp).Row
        colNum = WorksheetFunction.Match("Comments", Range("A1:CC1"), 0)
        For Each c In Range(Cells(4, colNum), Cells(lastRow, colNum))
          If c.Value = "Use This Space To Include Additional Details Or Explain The Reason For Your Request." Then
            c.Value = ""
            End If
        Next

        Columns("A:AA").Select
            ActiveSheet.Range("$A$1:$AA$117").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6 _
                , 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27), Header:=xlYes
        Range("A1").Select


    Range("Q1").Select

ErrHandler:
    Exit Sub


End Sub

