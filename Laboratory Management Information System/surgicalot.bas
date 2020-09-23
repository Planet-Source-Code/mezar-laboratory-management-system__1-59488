Attribute VB_Name = "Module1"
Public Cnn As ADODB.Connection
Public rs As ADODB.Recordset
Public Sub opendb()
Set Cnn = New ADODB.Connection
Dim rspath As String
'Cnn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=rmi;Data Source=haisoft11"
Cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\final project\lab_project\lab.mdb;Persist Security Info=False"
Cnn.Open
End Sub
Public Function openrec(strrec As String) As ADODB.Recordset
Set openrec = New ADODB.Recordset
Debug.Print strrec
With openrec
.Source = strrec
.ActiveConnection = Cnn
.CursorType = adOpenKeyset
.CursorLocation = adUseServer
.LockType = adLockOptimistic
.Open Options:=adCmdText
End With
End Function
Public Function zeroadd(str1 As String) As String
zeroadd = "0" + str1
End Function
Public Function numeric(TBOX As Integer) As Integer
If (TBOX < 46 Or TBOX > 60) And (TBOX <> 8) Then
numeric = 0
Else
numeric = TBOX
End If
End Function
Public Sub globerr()
MsgBox (Err.Description)
End Sub
 

