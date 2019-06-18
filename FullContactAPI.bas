Attribute VB_Name = "FullContactAPI"
Sub FullContact()

Dim EmailAddress As String
Dim JsonDict As New Dictionary
Dim Response As String
Dim jsonFile As String
Dim APIKEY As String
Dim Url   As String
Dim ObjHTTP As Variant

''''Set API Key
APIKEY = ""

''''Set up endpoint and browser object
Url = "https://api.fullcontact.com/v3/person.enrich"
Set ObjHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

''''Create dictionary to use Convert to JSON library
EmailAddress = ActiveCell.Value
JsonDict("email") = EmailAddress
jsonFile = ConvertToJson(JsonDict, 0)
Debug.Print jsonFile

''''create the request
ObjHTTP.Open "POST", Url, False
ObjHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
ObjHTTP.setRequestHeader "Authorization", "Bearer " & APIKEY
ObjHTTP.setRequestHeader "cache-control", "no-cache"
ObjHTTP.send (jsonFile)

''''parse response from server. 200 == good response
Debug.Print ObjHTTP.Status
If ObjHTTP.Status = 200 Then
    Response = ObjHTTP.ResponseText
    Debug.Print Response
Else
    Debug.Print ObjHTTP.Status, ObjHTTP.ResponseText
    Exit Sub
End If


''''Create object for json response
Set Answer = ParseJson(Response)

''''Call recursive function to display json values on spreadsheet
emptyDict (Answer)



End Sub

Public Sub emptyDict(ByVal json As Object)
Dim key As Variant, item As Object

NextColumn = Range("ZZ" & ActiveCell.Row).End(xlToLeft).Column + 1

    For Each key In json
        Select Case TypeName(json(key))
        Case "String"
            Debug.Print key & vbTab & json(key)
            Cells(ActiveCell.Row, NextColumn).Value = json(key)
            NextColumn = NextColumn + 1
        Case "Collection"
            For Each item In json(key)
                emptyDict item
            Next
        End Select
    Next
End Sub
