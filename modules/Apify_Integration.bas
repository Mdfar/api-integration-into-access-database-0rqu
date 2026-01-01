Attribute VB_Name = "Apify_Integration" ' Staqlt Solution Architecture - Apify to MS Access Routine ' Requires Reference: Microsoft WinHTTP Services, v5.1

Option Compare Database Option Explicit

Private Const APIFY_TOKEN As String = "YOUR_APIFY_TOKEN" Private Const ACTOR_ID As String = "search_google_maps"

Public Sub RunGoogleMapsSync() Dim actorRunId As String Dim datasetId As String

Debug.Print "Starting Apify Run..."

' 1. Trigger Actor
actorRunId = TriggerActor("London Dentists")
If actorRunId = "" Then Exit Sub

' 2. Poll for Completion
If WaitUntilFinished(actorRunId) Then
    ' 3. Get Dataset ID
    datasetId = GetDatasetId(actorRunId)
    ' 4. Import Data
    ImportDataset datasetId
End If


End Sub

Private Function TriggerActor(searchQuery As String) As String Dim http As New WinHttpRequest Dim url As String: url = "https://api.apify.com/v2/acts/" & ACTOR_ID & "/runs?token=" & APIFY_TOKEN Dim body As String: body = "{""queries"": """ & searchQuery & """}"

http.Open "POST", url, False
http.SetRequestHeader "Content-Type", "application/json"
http.Send body

' Extract ID using simple string manipulation or JSON parser
TriggerActor = ParseJsonValue(http.ResponseText, "id")


End Function

Private Sub ImportDataset(datasetId As String) Dim http As New WinHttpRequest Dim url As String: url = "https://api.apify.com/v2/datasets/" & datasetId & "/items?token=" & APIFY_TOKEN Dim rs As DAO.Recordset

http.Open "GET", url, False
http.Send

' Record everything properly
Set rs = CurrentDb.OpenRecordset("tbl_GoogleMapsLeads", dbOpenDynaset)
' Logic to iterate through JSON items and .AddNew to Recordset
Debug.Print "Imported data from dataset: " & datasetId
rs.Close


End Sub

Private Function ParseJsonValue(json As String, key As String) As String ' Simplified helper for guidance Dim pos As Long: pos = InStr(json, """" & key & """") If pos > 0 Then ParseJsonValue = Mid(json, pos + Len(key) + 4, 24) ' IDs are typically 24 chars End If End Function