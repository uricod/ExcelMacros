Attribute VB_Name = "Module1"
Sub Split_Worksheets()
Dim I               As Integer
Dim LastRow         As Integer
Dim FirstOpenRow    As Integer
Dim Department      As String
Dim FolderPath      As String
Dim xFile           As String
Dim sht             As Worksheet
Dim MainSht         As Worksheet
Dim WrkBok          As Workbook


''''Setup basic variable and speed up execution

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
On Error GoTo MyError
Set WrkBok = ActiveWorkbook
Set MainSht = ActiveSheet
LastRow = MainSht.Range("F5000").End(xlUp).Row

    ''''FIRST PART
    ''''Main Logic to split up main sheet into seperate worksheets
    For I = 2 To LastRow
        
        ''''Get the Department Name
        Department = MainSht.Range("F" & I).Value
        Debug.Print Department
        
        ''''Add New Sheet and Rename
        On Error Resume Next
            Set sht = Sheets(Department)
        On Error GoTo MyError
        If sht Is Nothing Then
            Set sht = Sheets.Add(After:=Sheets(Sheets.Count))
            sht.Name = Department
        Else
        End If
        
        ''''Copy Entire Row to new sheet
        FirstOpenRow = sht.Range("A5000").End(xlUp).Row + 1
        MainSht.Range("F" & I).EntireRow.Copy sht.Range("A" & FirstOpenRow)
        
        ''''Reset Variables to nothing and format
        sht.Columns("A:J").AutoFit
        If sht.Range("A1").Value = "" Then
            MainSht.Range("A1").EntireRow.Copy sht.Range("A1")
        End If
        Set sht = Nothing
    
    Next I
    
    ''''SECOND PART
    ''''Logic to split worksheets into workbooks and save them
    
        DateString = Format(Now, "mm-dd-yyyy")
        FolderPath = GetFolder
        For Each sht In WrkBok.Worksheets
            If sht.Index > 1 Then
               sht.Copy
               xFile = FolderPath & "\" & Application.ActiveWorkbook.Sheets(1).Name & " " & DateString & ".xlsx"
               Debug.Print xFile
               Application.ActiveWorkbook.SaveAs Filename:=xFile
               Application.ActiveWorkbook.Close False
            End If
        Next sht
    

'''Reset Defaults
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
MainSht.Activate
Exit Sub

''''Error Handling
MyError:
        
MsgBox "Oops, an error has occured." & vbCrLf & vbCrLf & "Error Code : " & Err.Number & " , " & Err.Description

'''Reset Defaults after error
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
MainSht.Activate

End Sub


Function GetFolder() As String
Dim fldr As FileDialog
Dim sItem As String

     Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
        
    
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
    
End Function
