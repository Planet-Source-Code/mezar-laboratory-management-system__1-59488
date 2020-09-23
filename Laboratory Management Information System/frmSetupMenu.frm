VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmSetupMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Menu"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   4575
      Begin VB.CommandButton cmddesig 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "frmSetupMenu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton CmdEmp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "frmSetupMenu.frx":103D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton CmdMMenu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmSetupMenu.frx":2343
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5280
         Width           =   4335
      End
      Begin VB.CommandButton CmdSecurity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "frmSetupMenu.frx":3245
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2160
         Width           =   4335
      End
      Begin VB.CommandButton CmdSubTest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmSetupMenu.frx":44C8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4200
         Width           =   4335
      End
      Begin VB.CommandButton cmdTest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmSetupMenu.frx":5235
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3120
         Width           =   4335
      End
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   720
      Top             =   3600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   600
      Left            =   3960
      TabIndex        =   5
      Top             =   360
      Width           =   2730
   End
End
Attribute VB_Name = "frmSetupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddesig_Click()
frmDesignation.Show
End Sub

Private Sub CmdEmp_Click()
frmEmployee.Show
End Sub

Private Sub CmdMMenu_Click()
    Unload Me
    frmMain.Show
End Sub

Private Sub CmdMMenu_GotFocus()
    Me.CmdMMenu.Picture = LoadPicture(App.Path & "\Main_Menu.jpg")
End Sub

Private Sub CmdMMenu_LostFocus()
   ' Me.CmdMMenu.Picture = LoadPicture(App.Path & "\Main_menu1.jpg")
End Sub

Private Sub CmdSecurity_Click()
    Unload Me
    frmSecurity.Show
End Sub

Private Sub CmdSecurity_GotFocus()
'    Me.CmdSecurity.Picture = LoadPicture(App.Path & "\Security_manager1.jpg")
End Sub

Private Sub CmdSecurity_LostFocus()
'    Me.CmdSecurity.Picture = LoadPicture(App.Path & "\Security_manager.jpg")
End Sub

Private Sub CmdSubTest_Click()
    Unload Me
    frmSubTestEntry.Show
End Sub

Private Sub CmdSubTest_GotFocus()
    Me.CmdSubTest.Picture = LoadPicture(App.Path & "\Sub_test.jpg")
End Sub

Private Sub CmdSubTest_LostFocus()
    'Me.CmdSubTest.Picture = LoadPicture(App.Path & "\Sub_Test1.jpg")
End Sub

Private Sub cmdTest_Click()
Unload Me
frmTestEntry.Show
End Sub

Private Sub cmdTest_GotFocus()
    Me.cmdTest.Picture = LoadPicture(App.Path & "\Test.jpg")
End Sub

Private Sub cmdTest_LostFocus()
   ' Me.cmdTest.Picture = LoadPicture(App.Path & "\Test1.jpg")
End Sub

