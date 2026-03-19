Attribute VB_Name = "Deletesheets"
Sub Delsheets()

Dim sheet As Worksheet

For Each sheet In Worksheets
Application.DisplayAlerts = False
If sheet.Name <> "Data" And sheet.Name <> "Principal" And sheet.Name <> "index" Then

sheet.Delete

Else
End If

Sheets("index").Select
Cells.Select
Selection.Delete
Range("A1").Select

Next sheet

End Sub
