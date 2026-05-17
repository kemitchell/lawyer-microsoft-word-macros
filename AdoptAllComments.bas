' Sign the current user's name to all pending comments and tracked changes.
Sub AdoptAllComments()
    Dim comment As comment
    For Each comment In ActiveDocument.Comments
        comment.Author = Application.UserName
        comment.Initial = Application.UserInitials
    Next comment
End Sub
