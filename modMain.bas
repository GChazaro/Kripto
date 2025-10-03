Attribute VB_Name = "modMain"
Option Explicit

Public Chazaro          As New ADODB.Connection '••••••••• »   ‹
Public MyADORs          As New ADODB.Recordset '•••••••••• »   ‹
Public ADORs            As New ADODB.Recordset '•••••••••• »   ‹

Public DBName           As String '•••••••••• »   ‹
Public DBPath           As String '•••••••••• »   ‹
Sub Main(Password As String)
    On Error GoTo ErrorHandler
    If Password = "" Then MsgBox "Por seguridad debes introducir un password para el descifrado.", vbOKOnly, "Introduce un password": Exit Sub

    Set Chazaro = New ADODB.Connection
    Chazaro.CommandTimeout = 10 '••••••••••••••••••••••••••• »   ‹
    Chazaro.ConnectionTimeout = 10 '•••••••••••••••••••••••• »   ‹
    Chazaro.Mode = adModeShareDenyNone '•••••••••••••••••••• »   ‹  = adModeReadWrite
    Chazaro.CursorLocation = adUseClient '•••••••••••••••••• »   ‹ to fix block issue
    'Chazaro.CursorLocation = adUseServer '••••••••••••••••• »   ‹
    Chazaro.Provider = "Microsoft.ACE.OLEDB.12.0"
    Chazaro.ConnectionString = "Data Source = " & App.Path & "\Kripto.accdb" & ";" _
    & "Jet OLEDB:Database Password = '" & Password & "'; Persist Security Info=true"
    Chazaro.Open '•••••••••••••••••••••••••••••••••••••••••••••••• »
Exit Sub

ErrorHandler:
    If Err = -2147217843 Then
        If Err.Description = "Not a valid password." Then
            MsgBox "Password Erróneo", vbOKOnly, "La BD no ha sido abierta"
        End If
    End If
    If Err = -2147217887 Then
        If Err.Description = "Multiple-step OLE DB operation generated errors. Check each OLE DB status value, if available. No work was done." Then
            MsgBox "Error en operación OLE DB", vbOKOnly, "La BD no ha sido abierta"
        End If
    End If

End Sub
