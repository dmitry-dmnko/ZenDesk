Attribute VB_Name = "Module8"
Sub MLEOrder_Merch()
Attribute MLEOrder_Merch.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MLEOrder_Merch Macro
'

'
    'Add Allied Requests tab:
    Set report = ActiveWorkbook
    Set wsr = ActiveSheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = CStr(wsr.Name) + " MLE"
    Set wsmle = ActiveSheet
    wsr.Activate
    Columns("Q:AC").Select
    Selection.Copy
    wsmle.Activate
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


    'Remove duplicates
    Columns("A:M").Select
    ActiveSheet.Range("$A:$M").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7 _
        , 8, 9, 10, 11, 12, 13), Header:=xlYes
    Columns("A:K").EntireColumn.AutoFit
    
    ' Remove Blank Rows
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
        
    
    Dim MyPATH As String
    Dim FileNAME As String
    MyPATH = ActiveWorkbook.Path
    FileNAME = Right(ActiveSheet.Name, 3)
    wsmle.Select
    wsmle.Copy

    StrDate = Format(Now, "yyyy-mm-dd hh-mm-ss AM/PM")
    ActiveWorkbook.SaveAs FileNAME:=MyPATH & "\" & StrDate & " " & FileNAME, FileFormat:=51, CreateBackup:=False
    ActiveWindow.Close
    
    report.Activate
    wsr.Activate

    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    
    On Error GoTo ErrHandler
    
    For Each cell In wsr.Columns("Q").Cells
        If cell.Row <> 1 Then
            If (cell.Value) <> "" Then
                Set OutApp = CreateObject("Outlook.Application")
                Set OutMail = OutApp.CreateItem(0)
                With OutMail
                    .To = "BACservice@mleinc.com"
                    .CC = ""
                    .BCC = ""
                    .Subject = Cells(cell.Row, "Q").Value
                    .Body = Cells(cell.Row, "Z").Value & " - Unit(s) " & Cells(cell.Row, "Y").Value & " (" & Cells(cell.Row, "X").Value & ")" & vbLf & vbLf & _
                        "Please ship to:" & vbLf & "Bank of America" & vbLf & "Attn:" & Cells(cell.Row, "R").Value & vbLf & _
                        Cells(cell.Row, "S").Value & vbLf & Cells(cell.Row, "T").Value & ", " & Cells(cell.Row, "U").Value & " " & Cells(cell.Row, "V").Value & vbLf & _
                        Cells(cell.Row, "W").Value & vbLf & vbLf & "Thanks" & vbLf & "B"
                    .Display
                    '.Send
                End With
                Cells(cell.Row, "AD").Value = "sent"
                Set OutMail = Nothing
            End If
        End If
    Next cell
        
        
    On Error GoTo ErrHandler
    
    report.Activate
    wsr.Activate

    Call CreaateZDFile
    
ErrHandler:
    Exit Sub



End Sub

