Attribute VB_Name = "hyperlinks"
Sub Hlinks()

Dim line As Integer
Dim sheet As Worksheet

Sheets("index").Select
Range("A1").Select
Selection.Insert
Rows("1:5").Select


line = 2
For Each sheet In Worksheets
If sheet.Name <> "index" And sheet.Name <> "Data" And sheet.Name <> "Principal" Then
Sheets("index").Select
Sheets("index").Cells(line, 1).Select
ActiveSheet.hyperlinks.Add anchor:=Selection, Address:="", SubAddress:=sheet.Name & "!A1", TextToDisplay:=sheet.Name

Sheets("index").Range("A" & line).Font.Size = 22
line = line + 1
Sheets(sheet.Name).Select
Range("F1").Select

ActiveSheet.hyperlinks.Add anchor:=Selection, Address:="", SubAddress:="index!B" & line + 4, TextToDisplay:="Voltar"
End If

Next sheet

End Sub
