Attribute VB_Name = "Module1"
Public un As String
Public conn As New ADODB.Connection
Public usern As String
Public Sub main()
If conn.State = adStateOpen Then
    conn.Close
End If
conn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=pharmacy;Data Source=RIDER4EVER\RIDERSQLSERVER"
conn.Open
End Sub


