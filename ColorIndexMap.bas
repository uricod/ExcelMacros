Attribute VB_Name = "ColorIndexMap"
Option Explicit
Sub showColorIndexNums()
Dim i As Integer, n As Integer, t As Integer
n = 1
For i = 1 To 8
    For t = 1 To 8
        If n > 56 Then: Exit Sub
        Cells(i, t).Interior.ColorIndex = n
        Cells(i, t).Value = n
        Cells(i, t).Font.ThemeColor = xlThemeColorDark2
        Cells(i, t).HorizontalAlignment = xlCenter
        With Cells(i, t)
            Cells.RowHeight = .Width
            Cells.ColumnWidth = .ColumnWidth
        End With
    n = n + 1
    Next t
Next i
ActiveSheet.[B:B].EntireColumn.AutoFit

End Sub




