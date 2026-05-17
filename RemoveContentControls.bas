Sub RemoveContentControls()
    For Each ContentControl In ActiveDocument.ContentControls
        ContentControl.Delete
    Next
End Sub
