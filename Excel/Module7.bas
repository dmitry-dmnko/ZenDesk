Attribute VB_Name = "Module7"
    Function sheetExists(sheetToFind As String) As Boolean
        sheetExists = False
        For Each Sheet In Worksheets
            If sheetToFind = Sheet.Name Then
                sheetExists = True
                Exit Function
            End If
        Next Sheet
    End Function
