Sub CycleMarkupView()
    With ActiveWindow.View.RevisionsFilter
        Select Case .Markup
        Case wdRevisionsMarkupSimple
             .Markup = wdRevisionsMarkupAll
        Case wdRevisionsMarkupAll
             .Markup = wdRevisionsMarkupNone
        Case wdRevisionsMarkupNone
             .Markup = wdRevisionsMarkupSimple
        End Select
        .View = wdRevisionsViewFinal
    End With
End Sub
