Attribute VB_Name = "SumText"
Function SumText(SumRange As Range, Seperator As String)
  
    For Each Cell In SumRange
        
        If StringToReturn = "" Then
            StringToReturn = Cell.Value
        Else
            StringToReturn = StringToReturn & Seperator & Cell.Value
        End If
    
    Next Cell

SumText = StringToReturn

End Function


