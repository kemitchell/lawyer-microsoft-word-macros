' Delete all comments and accept all pending tracked changes.
Sub PrepareCleanCopy()
    If ActiveDocument.Comments.Count > 0 Then
        ActiveDocument.DeleteAllComments
    End If
    If ActiveDocument.Revisions.Count > 0 Then
        ActiveDocument.AcceptAllRevisions
    End If
End Sub
