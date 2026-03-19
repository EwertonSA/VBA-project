Attribute VB_Name = "hideshow"
Sub hidesheet()

Dim sheet As Worksheet

Sheets("Principal").Select

For Each sheet In Worksheets
If sheet.Name <> "index" And sheet.Name <> "Principal" Then
sheet.Visible = xlSheetHidden

End If
Next sheet

End Sub
Sub showsheet()

Dim sheet As Worksheet

Sheets("Principal").Select

For Each sheet In Worksheets
If sheet.Name <> "index" And sheet.Name <> "Principal" Then
sheet.Visible = True

End If


Next sheet

End Sub
