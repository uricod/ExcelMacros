Attribute VB_Name = "FilterByLastWeek"
Sub AddLastWeekFilter()
Dim Cf As FormatCondition

''''This is for the last 7 days from today highlight yellow'''
''''The mod function to ensure no text or blank cells will get highlighted''''

Set Cf = ActiveSheet.Range("C:C").FormatConditions.Add(Type:=xlExpression, Formula1:="=And(C1 <= Today() , C1 >= (Today() - 7), Mod(C1, 1)= 0)")
Cf.Interior.Color = rgbYellow

''''This filters the ones with the yellow color
Range("A1").CurrentRegion.AutoFilter Field:=3, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor

End Sub

