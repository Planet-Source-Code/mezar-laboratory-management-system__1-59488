VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   4575
      Begin VB.CommandButton CmdPatient 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmMain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   4335
      End
      Begin VB.CommandButton CmdLabRpt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmMain.frx":12E7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4560
         Width           =   4335
      End
      Begin VB.CommandButton CmdSetup 
         DisabledPicture =   "frmMain.frx":22B1
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmMain.frx":2E1C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton CmdLabReg 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmMain.frx":3AC2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   4335
      End
      Begin VB.CommandButton CmdResult 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmMain.frx":4CDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3480
         Width           =   4335
      End
      Begin VB.CommandButton CmdExit 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmMain.frx":5C63
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5640
         Width           =   4335
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    End
End Sub

Private Sub CmdExit_GotFocus()
    Me.CmdExit.Picture = LoadPicture(App.Path & "\exit.jpg")
End Sub

Private Sub CmdExit_LostFocus()
    Me.CmdExit.Picture = LoadPicture(App.Path & "\exit1.jpg")
End Sub

Private Sub CmdLabReg_Click()
    Unload Me
    frmoption.Show
End Sub

Private Sub CmdLabReg_GotFocus()
    Me.CmdLabReg.Picture = LoadPicture(App.Path & "\lab_Registration.jpg")
End Sub

Private Sub CmdLabReg_LostFocus()
    Me.CmdLabReg.Picture = LoadPicture(App.Path & "\lab_Registration1.jpg")
End Sub

Private Sub CmdLabRpt_Click()
    Unload Me
    frmLabRpt.Show
End Sub

Private Sub CmdLabRpt_GotFocus()
    Me.CmdLabRpt.Picture = LoadPicture(App.Path & "\lab_Reports.jpg")
End Sub

Private Sub CmdLabRpt_LostFocus()
    Me.CmdLabRpt.Picture = LoadPicture(App.Path & "\lab_Reports1.jpg")
End Sub

Private Sub CmdPatient_Click()
frm_patient.Show
End Sub

Private Sub CmdResult_Click()
    Unload Me
    frmtestresult.Show
End Sub

Private Sub CmdResult_GotFocus()
    Me.CmdResult.Picture = LoadPicture(App.Path & "\test_result.jpg")
End Sub

Private Sub CmdResult_LostFocus()
    Me.CmdResult.Picture = LoadPicture(App.Path & "\test_result1.jpg")
End Sub

Private Sub CmdSetup_Click()
    Unload Me
    frmSetupMenu.Show
End Sub

Private Sub CmdSetup_GotFocus()
    Me.CmdSetup.Picture = LoadPicture(App.Path & "\setup.jpg")
End Sub

Private Sub CmdSetup_LostFocus()
'    Me.CmdSetup.Picture = LoadPicture(App.Path & "\setup1.jpg")
End Sub
