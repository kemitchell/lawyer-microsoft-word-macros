Sub RemoveAllHyperlinks()
    Dim field As field
    For Each field In ActiveDocument.Fields
        If field.Type = wdFieldHyperlink Then
            field.Unlink
        End If
    Next
End Sub
