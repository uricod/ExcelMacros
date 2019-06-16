Attribute VB_Name = "FitByColumnHeader"
Sub AutoFitPerHeader()
Dim I         As Integer
Dim LenOfCell As Integer
Dim FontSize  As Integer
Dim ColWidth  As Integer
Dim adjust    As Integer

For I = 1 To Range("A1").CurrentRegion.Columns.Count

    '''Find how many charcters are in the current cell then font size then current column width
    LenOfCell = Len(Cells(1, I).Value)
    FontSize = Cells(1, I).Font.Size
    ColWidth = Cells(1, I).EntireColumn.ColumnWidth
    Debug.Print LenOfCell, FontSize, ColWidth
    
    '''First Adjust len of cell based on font size
    If FontSize = 11 Then
        adjust = 0
    ElseIf FontSize <> 11 Then
        adjust = FontSize - 11
    End If
    
    '''Set Column Width based off the basis that each charcter needs 1 space in Column Width (Based off size 11 font)
    If LenOfCell + adjust <= ColWidth Then
        '''Do nothing if length of cell is smaller than width
    Else
        Cells(1, I).EntireColumn.ColumnWidth = LenOfCell + adjust
    End If

Next I

End Sub
