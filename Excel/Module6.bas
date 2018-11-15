Attribute VB_Name = "Module6"

Sub CreaateZDFile()
Attribute CreaateZDFile.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CreaateZDFile Macro
'

'
    'Add Allied Requests tab:
    Set report = ActiveWorkbook
    Set wsr = ActiveSheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = CStr(wsr.Name) + " ZD"
    Set wszd = ActiveSheet
    wsr.Activate
    Columns("A:O").Select
    Selection.Copy
    wszd.Activate
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


    'Remove duplicates
    Columns("A:O").Select
    ActiveSheet.Range("$A:$O").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7 _
        , 8, 9, 10, 11, 12, 13, 14, 15), Header:=xlYes
    Columns("A:K").EntireColumn.AutoFit
        
    ' Remove Blank Rows
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
        
    
    Dim MyPATH As String
    Dim FileNAME As String
    MyPATH = ActiveWorkbook.Path
    FileNAME = Right(ActiveSheet.Name, 2)
    wszd.Select
    wszd.Copy

    StrDate = Format(Now, "yyyy-mm-dd hh-mm-ss AM/PM")
    ActiveWorkbook.SaveAs FileNAME:=MyPATH & "\" & StrDate & " " & FileNAME, FileFormat:=51, CreateBackup:=False


    Dim OutApp As Object
    Dim OutMail As Object
    On Error GoTo ErrHandler

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

        With OutMail
        .To = "dmitry.dmytrenko@envusa.design"
        .CC = ""
        .BCC = ""
        .Subject = CStr(wszd.Name)
        .Body = ""
        .Attachments.Add ActiveWorkbook.FullName
        .Display
        '.Send
    End With
    On Error GoTo ErrHandler

    Set OutMail = Nothing
    Set OutApp = Nothing

    ActiveWindow.Close
    
    report.Activate
    wsr.Activate

ErrHandler:
    Exit Sub

End Sub
