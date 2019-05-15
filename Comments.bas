Attribute VB_Name = "Module1"
Sub PrintCellComments(Ribbon As IRibbonControl)

Dim Cmt As String
Dim C As Range
Dim I As Integer

Application.ScreenUpdating = False

For I = 1 To ActiveSheet.Comments.Count

    Set C = ActiveSheet.Comments(I).Parent
    Cmt = ActiveSheet.Comments(I).Text
    RowofComment = C.Row
    FirstEmptyColumn = Range("Z" & RowofComment).End(xlToLeft).Column + 1
    Cells(RowofComment, FirstEmptyColumn).Value = Cmt
    Range("A1:Z500").WrapText = False
    
Next I

Application.ScreenUpdating = True
'Cells.ClearComments


End Sub

Sub MoveCellComments(Ribbon As IRibbonControl)

Dim ws As Worksheet
Dim sht As Worksheet
Dim Cmt As String
Dim C As Range
Dim I As Integer

Set ws = ActiveSheet
Set sht = Sheets.Add

Application.ScreenUpdating = False

For I = 1 To ws.Comments.Count

    Set C = ws.Comments(I).Parent
    Cmt = ws.Comments(I).Text
    RowofComment = C.Row
    lastcolumn = ws.Range("Z" & RowofComment).End(xlToLeft).Column
    FirstEmptyRow = sht.Range("A5000").End(xlUp).Row + 1
    
    ws.Activate
    
    ws.Range(Cells(RowofComment, 1), Cells(RowofComment, lastcolumn)).Copy
    sht.Range("A" & FirstEmptyRow).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
Next I

sht.Columns("A:Z").AutoFit
Application.ScreenUpdating = True
'Cells.ClearComments


End Sub

