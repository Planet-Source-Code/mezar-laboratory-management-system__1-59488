VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabCLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Cn As New ADODB.Connection
Private Sub Class_Initialize()
   Set Cn = New ADODB.Connection
   Cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=mitproject;Data Source=server"
   Cn.Open
End Sub
Public Function add(strinsert As String) As Boolean
  Debug.Print strinsert
  Cn.Execute strinsert
End Function
Public Function update(strupdate As String) As Boolean
Debug.Print strupdate
    Cn.Execute strupdate
    update = True
End Function
Public Function execquery(strupdate As String) As Boolean
Debug.Print strupdate
    Cn.Execute strupdate
    update = True
End Function
Public Function openrec(strrec As String) As ADODB.Recordset
Debug.Print strrec
    Set openrec = New ADODB.Recordset
    With openrec
        .Source = strrec
        .ActiveConnection = Cn
        .CursorType = adOpenKeyset
        .CursorLocation = adUseServer
        .LockType = adLockOptimistic
        .Open Options:=adCmdText
    End With
End Function
Public Function delete(Strdelete As String) As Boolean
    Debug.Print Strdelete
    Cn.Execute (Strdelete)
End Function
Public Function FNMAX(STR As String) As ADODB.Recordset
    Set FNMAX = Cnn.Execute(STR)
End Function
Public Function numeric(TBOX As Integer) As Integer
    If (TBOX < 46 Or TBOX > 60) And (TBOX <> 8) Then
        numeric = 0
    Else
        numeric = TBOX
    End If
End Function


