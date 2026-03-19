Attribute VB_Name = "hideshow"
Sub HideSheet()

Dim sheet As Worksheet

Sheets("Principal").Select

For Each sheet In Worksheets
If sheet.Name <> "index" And sheet.Name <> "Principal" Then
sheet.Visible = xlSheetHidden

End If
Next sheet

End Sub
Sub ShowSheet()

Dim sheet As Worksheet

Sheets("Principal").Select

For Each sheet In Worksheets
If sheet.Name <> "index" And sheet.Name <> "Principal" Then
sheet.Visible = True

End If


Next sheet

End Sub
