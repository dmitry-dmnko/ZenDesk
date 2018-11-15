Attribute VB_Name = "Module5"
Function SingleCellExtract(LookupValue As String, LookupRange As Range, ColumnNumber As Integer, Char As String)
'Updateby20150824
    Dim I As Long
    Dim xRet As String
    For I = 1 To LookupRange.Columns(1).Cells.Count
        If LookupRange.Cells(I, 1) = LookupValue Then
            If xRet = "" Then
                xRet = LookupRange.Cells(I, ColumnNumber) & Char
            Else
                xRet = xRet & "" & LookupRange.Cells(I, ColumnNumber) & Char
            End If
        End If
    Next
    SingleCellExtract = Left(xRet, Len(xRet) - 1)
End Function
