' Convert all spans of text using Styles to Normal spans of text with inline formatting.
Sub StylesToFormatting()
    Dim paragraph As paragraph
    Dim font As font
    Dim format As ParagraphFormat
    For Each paragraph In ActiveDocument.Paragraphs
        With paragraph
            If .Style <> ActiveDocument.Styles("Normal") Then
                Set font = .Style.font
                Set format = .Style.ParagraphFormat
                .Style = ActiveDocument.Styles("Normal")
                .Range.font = font
                .Range.ParagraphFormat = format
            End If
        End With
    Next
End Sub
