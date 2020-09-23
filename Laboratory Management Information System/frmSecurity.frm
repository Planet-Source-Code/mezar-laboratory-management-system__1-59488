VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMINISTRATOR"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   1800
      TabIndex        =   9
      Top             =   4200
      Width           =   8055
      Begin VB.CommandButton CmdBack 
         DisabledPicture =   "frmSecurity.frx":0000
         DownPicture     =   "frmSecurity.frx":03B9
         Height          =   375
         Left            =   6360
         Picture         =   "frmSecurity.frx":08FB
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdEdit 
         DisabledPicture =   "frmSecurity.frx":0D83
         DownPicture     =   "frmSecurity.frx":117B
         Height          =   375
         Left            =   3240
         Picture         =   "frmSecurity.frx":16CA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdDel 
         DisabledPicture =   "frmSecurity.frx":1B6B
         DownPicture     =   "frmSecurity.frx":1F83
         Height          =   375
         Left            =   4800
         Picture         =   "frmSecurity.frx":252A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         DisabledPicture =   "frmSecurity.frx":2A54
         DownPicture     =   "frmSecurity.frx":2E81
         Height          =   375
         Left            =   1680
         Picture         =   "frmSecurity.frx":33D4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdNew 
         DisabledPicture =   "frmSecurity.frx":389A
         DownPicture     =   "frmSecurity.frx":3C61
         Height          =   375
         Left            =   120
         Picture         =   "frmSecurity.frx":4141
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   7815
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5400
         TabIndex        =   21
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtPhNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtEmp_DOB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtDesig 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5400
         TabIndex        =   4
         Text            =   " "
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Text            =   " "
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox CmbEmpName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4125
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Phone #"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Day"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4125
         TabIndex        =   7
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4125
         TabIndex        =   6
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4125
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   9600
      X2              =   2160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   675
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   3330
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_Sec                  As New ADODB.Recordset
Dim rsEmp                   As New ADODB.Recordset
Dim rsEmpDesig              As New ADODB.Recordset
Dim merlin                  As IAgentCtlCharacter
Dim SaveCLS                 As New LabCLS
Dim editV                   As Integer
Dim EmpNoV                  As String
Dim StrIns_Sec              As String
Dim v1                      As String
Dim StrDel_Sec              As String
Dim obj                     As Object
Dim CHECK As Integer
Private Sub CmbEmpName_Click()
Me.txtLogin.locked = False
Me.txtPassword.locked = False
    With rsEmp
            .MoveFirst
            .Find ("Emp_Name ='" & Me.CmbEmpName & "'")
             EmpNoV = .Fields("Emp_No")
    End With
    Me.txtLogin = ""
    Me.txtPassword = ""
    Set rsEmpDesig = openrec("select designation.*,employee.* from Designation,Employee where designation.desig_ID = employee.desig_ID and employee.emp_No = " & EmpNoV & "")
    With rsEmpDesig
        If Not .EOF Then
              Me.txtEmp_DOB = .Fields("DOB")
              Me.txtsex = .Fields("Sex")
              Me.txtAddress = .Fields("address")
              Me.txtPhNo = .Fields("Phone_No")
              Me.txtDesig = .Fields("Description")
        End If
    End With
    Set rs_Sec = openrec("select security.* from security where security.emp_No = '" & EmpNoV & "'")
    With rs_Sec
        If Not .EOF Then
             Me.txtLogin = rs_Sec.Fields("Login_Id")
             Me.txtPassword = rs_Sec.Fields("Password")
        End If
    End With
End Sub
Private Sub CmdBack_Click()
'Dim a As String
'If CHECK = 0 Then
'a = MsgBox("do you want to save the information", vbYesNoCancel)
'If a = vbYes Then
'Call cmdsave_Click
'ElseIf a = vbNo Then
    Unload Me
    frmSetupMenu.Show
'End If
'End If
'CHECK = 0
End Sub
Private Sub CmdDel_Click()
For Each obj In Me
  If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
    If obj.Text = "" Then MsgBox "can not delete": Exit Sub
  End If
Next


    Dim a As String
    a = MsgBox("Are you Sure To Delete this Record", vbYesNoCancel, "Warning")
       
        If a = vbYes Then
           StrDel_Sec = "delete from security where emp_No='" & EmpNoV & "'"
           SaveCLS.execquery (StrDel_Sec)
           For Each obj In Me
             If TypeOf obj Is TextBox Then
                 obj.Text = ""
             End If
           Next
         Else
             Exit Sub
         End If
End Sub
Private Sub CmdEdit_Click()
    Me.CmdSave.Enabled = True
    Me.txtLogin.locked = False
    Me.txtPassword.locked = False
    editV = 1
End Sub
Private Sub cmdnew_Click()
CHECK = 0
  Me.CmbEmpName.Enabled = True
    Me.CmdSave.Enabled = True
    Me.txtLogin.locked = False
    Me.txtPassword.locked = False
    Me.txtLogin = ""
    Me.txtPassword = ""
End Sub
Private Sub cmdsave_Click()
CHECK = 1
  For Each obj In Me
   If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
     If obj.Text = "" Then MsgBox "give full information":     Exit Sub
  End If
  Next
  
    Select Case editV
    Case 0
         Me.txtLogin.locked = True
         Me.txtPassword.locked = True
         StrIns_Sec = "insert into security values('" & Trim(EmpNoV) & "','" & UCase(Trim(Me.txtLogin)) & "','" & UCase(Trim(Me.txtPassword)) & "','" & Trim(Me.txtDesig) & "')"
         SaveCLS.execquery (StrIns_Sec)
         rs_Sec.Fields.Refresh
    Case 1
         Me.txtLogin.locked = True
         Me.txtPassword.locked = True
         StrIns_Sec = "update security set login_id ='" & UCase(Trim(Me.txtLogin)) & "',password='" & UCase(Trim(Me.txtPassword)) & "',User_Type='" & Trim(Me.txtDesig) & "' where emp_no='" & EmpNoV & "'"
         Debug.Print (StrIns_Sec)
         SaveCLS.execquery (StrIns_Sec)
         rs_Sec.Fields.Refresh
    End Select
    Me.CmdSave.Enabled = False
    
End Sub
Private Sub Form_Load()
    Call OpenDB
    Me.CmdSave.Enabled = False
    Me.txtLogin.locked = True
    Me.txtPassword.locked = True
    Set rsEmp = openrec("select * from Employee")
      With rsEmp
        If .EOF = False Or .BOF = False Then
              .MoveFirst
                While Not .EOF
                    Me.CmbEmpName.AddItem .Fields("Emp_Name")
                    .MoveNext
                Wend
        End If
     End With
End Sub


Private Sub txtLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtPassword.SetFocus
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.CmdSave.SetFocus
End Sub
