Attribute VB_Name = "Module1"
Sub Get_Emails_From_Outlook()
Dim ol As Outlook.Application

Dim fol As Outlook.Folder
Dim I As Object
Dim N As Long
Dim mi As Outlook.MailItem

Set ol = New Outlook.Application

'''This is to get other email addresses besides default
'Set ns = ol.GetNamespace("MAPI").Stores
'Set fol = ns.Item(3).GetDefaultFolder(olFolderInbox)

''''This is to get the default email address inbox which is josh@fitnfrum.com
Set ns = ol.GetNamespace("MAPI")
Set fol = ns.GetDefaultFolder(olFolderInbox)


Worksheets.Add
rh = Range("A1").RowHeight
N = 1

Cells(1, 1).Value = "Sender Name"
Cells(1, 2).Value = "Subject"
Cells(1, 3).Value = "Received Time"
Cells(1, 4).Value = "Body"

Range("A1:D1").Font.Bold = True
For Each I In fol.Items
    If I.Class = olMail Then
     N = N + 1
     Set mi = I
   
     Cells(N, 1).Value = mi.SenderName
     Cells(N, 2).Value = mi.Subject
     Cells(N, 3).Value = mi.ReceivedTime
     Cells(N, 4).Value = mi.Body
     
     
    End If
Next I

Range("A1").CurrentRegion.EntireColumn.ColumnWidth = 32.58
Range("A1").CurrentRegion.EntireRow.RowHeight = rh

End Sub

Function Delete_Email() As String
Dim ol As Outlook.Application
Dim ns As Outlook.Namespace
Dim fol As Outlook.Folder
Dim rootfol As Outlook.Folder
Dim FilterText  As String
Dim I As Outlook.MailItem
Dim subjectText As String

Set ol = New Outlook.Application
Set ns = ol.GetNamespace("MAPI")
Set rootfol = ns.Folders(1)

Set fol = rootfol.Folders("Inbox")

Debug.Print ActiveCell.Offset(0, -2).Value
Debug.Print Range("E50").End(xlUp).Offset(0, -2).Value

subjectText = Range("E50").End(xlUp).Offset(0, -3).Value
Select Case LCase(Left(subjectText, 4))
    Case "re: ", "fw: "
        subjectText = Mid(subjectText, 5)
End Select

FilterText = "[SenderName] = '" & Range("E50").End(xlUp).Offset(0, -4).Value & "'"
FilterText = FilterText & " And [Subject] = '" & subjectText & "'"
Set I = fol.Items.Find(FilterText)
I.Delete

Delete_Email = "Deleted"

End Function
