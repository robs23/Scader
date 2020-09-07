Attribute VB_Name = "SQL"
Public adoConn As ADODB.Connection
Public scadaConn As ADODB.Connection

Public Sub updateConnection()

If Not adoConn Is Nothing Then
    If adoConn.State = 0 Then
        adoConn.Open connectionString
        adoConn.CommandTimeout = 90
    End If
Else
    Set adoConn = New ADODB.Connection
    adoConn.Open connectionString
    adoConn.CommandTimeout = 90
End If
End Sub

Public Sub closeConnection()

If Not adoConn Is Nothing Then
    If adoConn.State = 1 Then
        adoConn.Close
    End If
    Set adoConn = Nothing
End If
End Sub

Public Sub connectScada()
'Dim cmd As ADODB.Command
'Set cmd = New ADODB.Command

If scadaConn Is Nothing Then
    Set scadaConn = New ADODB.Connection
    scadaConn.Provider = "SQLOLEDB"
    scadaConn.connectionString = scadaConnectionString
    scadaConn.Open
    scadaConn.CommandTimeout = 90
Else
    If scadaConn.State = adStateClosed Then
        Set scadaConn = New ADODB.Connection
        scadaConn.Provider = "SQLOLEDB"
        scadaConn.connectionString = scadaConnectionString
        scadaConn.Open
        scadaConn.CommandTimeout = 90
    End If
End If


End Sub

Public Sub disconnectScada()

If Not scadaConn Is Nothing Then
    If scadaConn.State = 1 Then
        scadaConn.Close
    End If
    Set scadaConn = Nothing
End If
End Sub
