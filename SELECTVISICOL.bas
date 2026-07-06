Attribute VB_Name = "SELECTVISICOL"
Sub SelectVisibleColumns()

Dim ws As Worksheet
Dim lastRow As Long
Dim visRows As Range
Dim sel As Range

Set ws = ActiveSheet

lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

On Error Resume Next
Set visRows = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow
On Error GoTo 0

If visRows Is Nothing Then Exit Sub

Set sel = Union( _
    Intersect(visRows, ws.Columns("B")), _
    Intersect(visRows, ws.Columns("C")), _
    Intersect(visRows, ws.Columns("E")), _
    Intersect(visRows, ws.Columns("F")), _
    Intersect(visRows, ws.Columns("N")), _
    Intersect(visRows, ws.Columns("O")), _
    Intersect(visRows, ws.Columns("P")), _
    Intersect(visRows, ws.Columns("Q")), _
    Intersect(visRows, ws.Columns("T")), _
    Intersect(visRows, ws.Columns("U")))

sel.Select

End Sub



