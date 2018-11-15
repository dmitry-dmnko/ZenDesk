Attribute VB_Name = "Module4"

Sub zdesk()
Attribute zdesk.VB_ProcData.VB_Invoke_Func = " \n14"
'
' zdesk Macro
'

'
    'Add Allied Requests tab:
    Set report = ActiveWorkbook
    Set wsr = ActiveSheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "ZD Tickets"
    Set wsar = ActiveSheet
    wsr.Activate
    Range("AC3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    wsar.Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    'Autofit
    Columns("A:K").EntireColumn.AutoFit
        
    Dim MyPATH As String
    Dim FileNAME As String
    MyPATH = ActiveWorkbook.Path
    FileNAME = ActiveSheet.Name
    wsar.Select
    wsar.Copy

    ActiveWorkbook.SaveAs FileNAME:= _
        MyPATH & "\" & FileNAME, FileFormat:=51, _
        CreateBackup:=False
    ActiveWindow.Close
    
    report.Activate
    wsr.Activate
End Sub
