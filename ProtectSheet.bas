Attribute VB_Name = "ProtectSheet"
Sub LockSelection(Ribbon As IRibbonControl)
Dim SelectedRange As Range

'''Create Range object out of selection'''
Set SelectedRange = Selection

'''Color code the unlocked range'''
SelectedRange.Interior.ColorIndex = 34

''''Adds protection to sheet exculding selected range'''
SelectedRange.Locked = False
ActiveSheet.Protect Password:="password"

End Sub


Sub UnlockSelection(Ribbon As IRibbonControl)
Dim SelectedRange As Range

'''Unprotect Sheet'''
ActiveSheet.Unprotect Password:="password"

'''Reset the locked property and color'''
Set SelectedRange = Selection
SelectedRange.Locked = True
SelectedRange.Interior.ColorIndex = 0


End Sub

