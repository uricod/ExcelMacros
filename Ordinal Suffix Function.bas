Attribute VB_Name = "Module1"
Function OrdinalSuffix(Num As Long) As String

Dim N As Long
Const cSfx = "stndrdthththththth"
Dim answer As String

N = Num Mod 100
    If ((Abs(N) >= 10) And (Abs(N) <= 19)) Or ((Abs(N) Mod 10) = 0) Then
        OrdinalSuffix = "th"
    Else
        OrdinalSuffix = Mid(cSfx, ((Abs(N) Mod 10) * 2) - 1, 2)
    End If
answer = Format(Num, "#,##0") & OrdinalSuffix
Debug.Print answer
OrdinalSuffix = Num & OrdinalSuffix

End Function

