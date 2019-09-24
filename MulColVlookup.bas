Attribute VB_Name = "MulColVlookup"
Function MulColVlookup(SearchValues As Range, search_in_cols As Range, return_val As Range) As Variant

Dim LastRow As Integer
Dim wk      As Worksheet

Set wk = search_in_cols.Worksheet
Set ws = SearchValues.Worksheet
'''Concat all the lookup values together
For Each Cell In SearchValues.Cells
    VlookupValue = VlookupValue & Cell.Value
Next Cell

Debug.Print VlookupValue
Debug.Print SearchValues.Columns.Count, search_in_cols.Columns.Count

FirstRow = search_in_cols.Row
FirstCol = search_in_cols.Column
LastRow = wk.Cells(search_in_cols.Rows.Count, search_in_cols.Column).End(xlUp).Row
LastCol = search_in_cols.Columns.Count + FirstCol - 1

Debug.Print LastRow, FirstRow
Debug.Print LastCol, FirstCol

'''Find Values in Search_in_cols Range
For I = FirstRow To LastRow
    Debug.Print I
    
    Set RO = wk.Range(wk.Cells(I, FirstCol), wk.Cells(I, LastCol))
    '''Get the row value based off how many columns chosen
    For Each Ce In RO.Cells
        FullRowValue = FullRowValue & Ce.Value
        Debug.Print FullRowValue
    Next Ce
    
    If FullRowValue = VlookupValue Then
        Answer = wk.Cells(RO.Row, return_val.Column).Value
        
        Exit For
    End If
    Debug.Print FullRowValue
    '''Reset Variables
    FullRowValue = ""
Next I

MulColVlookup = Answer

End Function
