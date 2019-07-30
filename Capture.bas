Attribute VB_Name = "Capture"
Sub CapturePartScreen()
Dim FSO As New FileSystemObject

pathforfile = "C:\Users\Home\Desktop\capture.png"

Set objFile = FSO.CreateTextFile("C:\Users\Home\Downloads\boxcutter-1.5\boxcutter-1.5\filetorun.bat")

OriginalCommand = "cd C:\Users\Home\Downloads\boxcutter-1.5\boxcutter-1.5 "
objFile.WriteLine (OriginalCommand)

lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
lastCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column

For Each cel In Range(Cells(1, 1), Cells(lastRow, 1))
    HEG = HEG + cel.Height
Next cel

For Each cel In Range(Cells(1, 1), Cells(1, lastCol))
    WID = WID + cel.Width
Next cel

HE = HEG + 350
HEG = -(HEG)
WID = WID + 250
Debug.Print WID, HEG, HE


OutputString = "boxcutter.exe -c 0," & HEG & "," & WID & "," & HE & " " & pathforfile
Debug.Print OutputString
objFile.WriteLine (OutputString)

PID = Shell("cmd.exe /k C:\Users\Home\Downloads\boxcutter-1.5\boxcutter-1.5\filetorun.bat", vbHide)
Application.Wait Now + TimeValue("00:00:01")
Call Shell("TaskKill /F /PID " & CStr(PID), vbHide)

CreateEmail (pathforfile)

End Sub

Sub CaptureFullScreen()
Dim FSO As New FileSystemObject
pathforfile = "C:\Users\Home\Desktop\capture.png"

Set objFile = FSO.CreateTextFile("C:\Users\Home\Downloads\boxcutter-1.5\boxcutter-1.5\filetorun.bat")

OriginalCommand = "cd C:\Users\Home\Downloads\boxcutter-1.5\boxcutter-1.5 "
objFile.WriteLine (OriginalCommand)

OutputString = "boxcutter.exe -f " & pathforfile
objFile.WriteLine (OutputString)


PID = Shell("cmd.exe /k C:\Users\Home\Downloads\boxcutter-1.5\boxcutter-1.5\filetorun.bat", vbHide)
Application.Wait Now + TimeValue("00:00:01")
Call Shell("TaskKill /F /PID " & CStr(PID), vbHide)

CreateEmail (pathforfile)

End Sub

Sub CreateEmail(attachfile As String)
Dim OutApp As New Outlook.Application


Set objMsg = OutApp.CreateItem(olMailItem)

 With objMsg
  .To = "Alias@domain.com"
  .Subject = "Capture of Spreadsheet"
  .BodyFormat = olFormatPlain
  .Importance = olImportanceHigh
  .Attachments.Add (attachfile)
  .Display
  
End With

Set objMsg = Nothing
End Sub


