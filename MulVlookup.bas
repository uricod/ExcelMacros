Attribute VB_Name = "MulVlookup"
Option Base 1

Function MulVlookup(SearchValue As Range, search_in_col As Range, return_val As Range) As Variant

Dim answer() As Variant
Dim ActiveCellAdd As String
Dim LastRow As Integer

LastRow = Cells(search_in_col.Count, search_in_col.Column).End(xlUp).Row
ActiveCellAdd = ActiveCell.Address

For i = 1 To LastRow
    If search_in_col.Cells(i, 1) = SearchValue Then
       p = p + 1
       ReDim Preserve answer(p)
       answer(p) = return_val.Cells(i, 1).Value
    End If

Next i

MulVlookup = answer()

End Function


