Attribute VB_Name = "LabModule"
Public Cn As ADODB.Connection
Public Rs As ADODB.Recordset
Public Sub OpenDB()
Set Cn = New ADODB.Connection
Cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=mitproject;Data Source=server"

Cn.Open
End Sub
Public Function openrec(strrec As String) As ADODB.Recordset
Set openrec = New ADODB.Recordset
Debug.Print strrec
    With openrec
        .Source = strrec
        .ActiveConnection = Cn
        .CursorType = adOpenDynamic
        .CursorLocation = adUseServer
        .LockType = adLockOptimistic
        .Open Options:=adCmdText
    End With
End Function
Public Function numeric(TBOX As Integer) As Integer
    If (TBOX < 46 Or TBOX > 60) And (TBOX <> 8) Then
        numeric = 0
    Else
        numeric = TBOX
    End If
End Function
Public Function Character(TBOX As Integer) As Integer
    If (TBOX < 46 Or TBOX > 60) Then 'And (TBOX <> 8) Then
    Character = TBOX
     Else
        Character = 0
        
    End If
End Function


