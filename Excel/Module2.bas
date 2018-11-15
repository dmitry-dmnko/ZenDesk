Attribute VB_Name = "Module2"
Sub AlliedRequestsFile()
'
' AlliedRequestsFile Macro
'

'

    'Add Allied Requests tab:
    Set report = ActiveWorkbook
    Set wsr = ActiveSheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = CStr(wsr.Name) + " AR"
    Set wsar = ActiveSheet
    wsr.Activate
    Columns("Q:AC").Select
    Selection.Copy
    wsar.Activate
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


    'Remove duplicates
    Columns("A:M").Select
    ActiveSheet.Range("$A:$M").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7 _
        , 8, 9, 10, 11, 12, 13), Header:=xlYes
    Columns("A:K").EntireColumn.AutoFit
        
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
        
    
    Dim MyPATH As String
    Dim FileNAME As String
    MyPATH = ActiveWorkbook.Path
    FileNAME = Right(ActiveSheet.Name, 2)
    wsar.Select
    wsar.Copy

    StrDate = Format(Now, "yyyy-mm-dd hh-mm-ss AM/PM")
    ActiveWorkbook.SaveAs FileNAME:=MyPATH & "\" & StrDate & " " & FileNAME, FileFormat:=51, CreateBackup:=False
    ActiveWindow.Close
    
    report.Activate
    wsr.Activate
    
    Call CreaateZDFile

End Sub




