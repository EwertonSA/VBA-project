Attribute VB_Name = "Report"
Sub reportation()
Call ShowSheet
Call Delsheets

Dim line As Integer
Dim last As Integer
Dim newsheet As String
Dim col As Integer

Sheets("Principal").Select
Range("A2").Select

last = Range("A1").End(xlDown).Row

For line = 2 To last

Sheets("Principal").Select

newsheet = ActiveCell.Value
ActiveCell.Offset(1, 0).Activate

Sheets.Add(after:=Sheets(Sheets.Count)).Name = newsheet
Sheets("Data").Range("A1:E1").Copy Sheets(newsheet).Range("A1:E1")

Sheets("Data").Select
Range("A1").Select
Selection.AutoFilter

ActiveSheet.Range("$A$1:$E$" & 10000).AutoFilter field:=1, Criteria1:=newsheet
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy Sheets(newsheet).Range("A1")

Sheets(newsheet).Select

For col = 2 To ActiveSheet.UsedRange.Columns.Count
Columns(col).EntireColumn.AutoFit

Next col
Dim sum1 As Double
Dim sum2 As Double
Dim sum3 As Double
Dim lines As Integer

Data = Sheets(newsheet).Range("A1").End(xlDown).Row + 2

sum1 = 0
sum2 = 0
sum3 = 0
For lines = 2 To Data - 2
sum1 = sum1 + Sheets(newsheet).Range("C" & lines)
sum2 = sum2 + Sheets(newsheet).Range("D" & lines)
sum3 = sum3 + Sheets(newsheet).Range("E" & lines)

Next lines

Sheets(newsheet).Range("C" & Data) = sum1
Sheets(newsheet).Range("D" & Data) = sum2
Sheets(newsheet).Range("E" & Data) = sum3

Sheets(newsheet).Range("C" & Data & ":" & "E" & Data).Font.Color = vblblack
Sheets(newsheet).Range("C" & Data & ":" & "E" & Data).Font.Size = 14
Sheets(newsheet).Range("C" & Data & ":" & "E" & Data).Interior.ColorIndex = 7
Sheets(newsheet).Range("C" & Data & ":" & "E" & Data).HorizontalAlignment = xlCenter
Sheets(newsheet).Range("C" & Data & ":" & "E" & Data).BorderAround LineStyle:=xlContinuous, Weight:=1

Sheets(newsheet).Range("E" & Data).Select
Selection.Style = "Currency"
Sheets(newsheet).Range("C" & Data).Select
Selection.Style = "Currency"
Columns(1).Delete
ActiveWindow.DisplayGridlines = False

Next line

Call Hlinks

Call IndTable
Call HideSheet


End Sub













