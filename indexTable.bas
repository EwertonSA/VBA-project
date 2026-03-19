Attribute VB_Name = "indexTable"

Sub IndexTable()

Dim lastindex As Integer

Sheets("index").Select
Range("A1").Select

Rows("1:5").Select
Selection.Insert
Columns(1).Insert
Sheets("index").Range("B2").Font.Size = 36
Sheets("index").Range("B2").Font.Name = "harlow solid italic"
Sheets("index").Range("B2").Value = "Relatório de vendas por vendedor"
Sheets("index").Range("C7").Formula2Local = "=SUMIF(Data!A:E;index!B7;Data!C:C)"
Sheets("index").Range("D7").Formula2Local = "=SUMIF(Data!A:E;index!B7;Data!D:D)"
Sheets("index").Range("E7").Formula2Local = "=SUMIF(Data!A:E;index!B7;Data!E:E)"




lastindex = Sheets("index").Range("B7").End(xlDown).Row
Sheets("index").Range("C7:E" & lastindex).FillDown
[C6] = "Valor unitário"
[D6] = "Unidades vendidas"
[E6] = "Total"

Sheets("index").Range("C7:C" & lastindex).Select
Selection.Style = "Currency"

Sheets("index").Range("E7:E" & lastindex).Select
Selection.Style = "Currency"

Call recorded


End Sub
