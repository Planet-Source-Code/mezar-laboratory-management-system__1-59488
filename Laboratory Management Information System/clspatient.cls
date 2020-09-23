VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Cn As New ADODB.Connection
Private Sub Class_Initialize()
   Set Cn = New ADODB.Connection
   Cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=mitproject1;Data Source=server"
   Cn.Open
End Sub
Public Function add(strinsert As String) As Boolean
  'On Error GoTo adderror
  Debug.Print strinsert
  Cn.Execute strinsert
  add = True
  Exit Function
adderror:
  add = False
End Function
Public Function update(strupdate As String) As Boolean
'On Error GoTo updaterror
Debug.Print strupdate
    Cn.Execute strupdate
    update = True
'Exit Function
'updaterror:
 '   update = False
End Function
Public Function openrec(strrec As String) As ADODB.Recordset
'On Error Resume Next
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
Public Function delete(strdelete As String) As Boolean
On Error GoTo errdelete
    Cn.Execute (strdelete)
    delete = True
Exit Function
errdelete:
    MsgBox (Err.Description)
    delete = False
End Function
Public Function FNMAX(STR As String) As ADODB.Recordset
    Set FNMAX = Cnn.Execute(STR)
End Function
Public Sub clearbox(formname As Object)
    Dim obj As Object
    For Each obj In formname
        If TypeOf obj Is TextBox Then
            obj.Text = ""
        End If
    Next
End Sub
Public Function numeric(TBOX As Integer) As Integer
    If (TBOX < 46 Or TBOX > 60) And (TBOX <> 8) Then
        numeric = 0
    Else
        numeric = TBOX
    End If
End Function

