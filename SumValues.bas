Attribute VB_Name = "SumValues"
Function SumValues(rang As Range) As Long

On Error Resume Next

ValueSum = 0

For Each Cell In rang

    If Cell.HasFormula = False Then
        ValueSum = ValueSum + Cell.Value
    End If

Next Cell

SumValues = ValueSum
On Error GoTo 0

End Function
