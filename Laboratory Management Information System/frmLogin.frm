VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   ForeColor       =   &H00000000&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      DisabledPicture =   "frmLogin.frx":0442
      DownPicture     =   "frmLogin.frx":0809
      Height          =   375
      Left            =   3360
      Picture         =   "frmLogin.frx":0D5A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdOK 
      DisabledPicture =   "frmLogin.frx":1212
      DownPicture     =   "frmLogin.frx":15DC
      Height          =   375
      Left            =   1800
      Picture         =   "frmLogin.frx":1A9F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtUserID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   645
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsFindPass As New ADODB.Recordset
Dim RsCheck As New ADODB.Recordset
Private Sub CmdCancel_Click()
End
End Sub
Private Sub CmdOK_Click()
Set rsFindPass = openrec("select * from Security where (login_id)='" & UCase(Me.txtUserID) & "' and (password)='" & UCase(Me.txtPassword) & "'")
    With rsFindPass
    If Not rsFindPass.EOF Then
        Unload Me
        Unload frmSplash
        frmMain.Show
        
 '        With RsCheck
           If Not .EOF Then
              If UCase(.Fields("user_type")) = UCase("dba") Or UCase(.Fields("user_type")) = UCase("administrator") Then
                 frmMain.CmdSetup.Enabled = True
              Else
                 frmMain.CmdSetup.Enabled = False
              End If
          End If
  '        End With
    Else
        MsgBox "Invalid Password", vbCritical, "Login"
        Me.txtPassword.SetFocus
    End If
    End With
End Sub

Private Sub Form_Load()
    Call OpenDB
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.CmdOK.SetFocus
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtPassword.SetFocus
End Sub
