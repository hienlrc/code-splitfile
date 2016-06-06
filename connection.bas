Public Sub Connect(ByRef cn As ADODB.Connection)
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
      App.Path & "\LRCBOOKS.accdb;"

'cn.ConnectionString = "Provider=SQLOLEDB;Password=sa123;User ID=sa;Initial Catalog=DataDB;Data Source='" & ".\SQLEXPRESS" & "'"
cn.Open
End Sub

Public Sub Close_Conn(ByRef cn As ADODB.Connection)
If (cn.State = adStateOpen) Then
    cn.Close
End If
End Sub


