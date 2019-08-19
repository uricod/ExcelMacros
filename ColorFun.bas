Attribute VB_Name = "ColorFun"
Function SumByColor(xlRange As Range, CellColor As Range)
Dim Result As Double

    CCol = CellColor.Interior.Color
    Debug.Print CCol
    
    For Each c In xlRange.Cells
        If c.Interior.Color = CCol Then
           Result = Result + c.Value
        End If
    
    Next c

SumByColor = Result

End Function

Function SumByFontColor(xlRange As Range, CellColor As Range)
Dim Result As Double

    CCol = CellColor.Cells.Font.Color
    Debug.Print CCol
    
    For Each c In xlRange.Cells
        If c.Font.Color = CCol Then
            Debug.Print c.Font.Color
           Result = Result + c.Value
        End If
    
    Next c

SumByFontColor = Result

End Function

