Attribute VB_Name = "Credit_Tracker"

Public Function GetApifyBalance() As Double Dim http As New WinHttpRequest Dim url As String: url = "https://api.apify.com/v2/users/me?token=" & APIFY_TOKEN

http.Open "GET", url, False
http.Send

' Extract limits/credits from user profile response
' Return as value to prevent overages
GetApifyBalance = 0 ' Placeholder


End Function