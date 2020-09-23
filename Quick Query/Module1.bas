Attribute VB_Name = "Module1"
Option Explicit

Public DBType As Integer
Public DatabasePath As String
Public ServerName As String
Public DB As String
Public cn As ADODB.Connection
Public AddSQL As String


Public Sub Open_cn()
    If DBType = 1 Then
        Set cn = New ADODB.Connection
        cn.CursorLocation = adUseClient
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        cn.Properties("Data Source") = DatabasePath
     '   cn.Properties("Jet OLEDB:Database Password") = LCase(Decode("198"))
        cn.Open

    Else
    
    Set cn = New ADODB.Connection
    With cn
        .Provider = "MSDASQL;DRIVER={SQL Server};SERVER=" & ServerName & ";trusted_connection=yes;database=" & DB & ""
        .Open
        
    End With
    End If
End Sub
Public Sub Close_cn()
    
     cn.Close
     Set cn = Nothing
    
    
End Sub
