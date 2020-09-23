VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmLabRpt 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Lab Reports"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbstatus 
      Height          =   315
      ItemData        =   "frmLabRpt.frx":0000
      Left            =   6000
      List            =   "frmLabRpt.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbyear 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView 
      Height          =   2370
      Left            =   5760
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      StartOfWeek     =   54198273
      CurrentDate     =   37395
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date to Date"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   2520
      Value           =   -1  'True
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker TDate 
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   54198273
      CurrentDate     =   36526
   End
   Begin MSComCtl2.DTPicker FDate 
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   54198273
      CurrentDate     =   36526
   End
   Begin VB.CommandButton CmdOK 
      DisabledPicture =   "frmLabRpt.frx":001E
      DownPicture     =   "frmLabRpt.frx":03E8
      Height          =   495
      Left            =   3600
      Picture         =   "frmLabRpt.frx":08AB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancel 
      DisabledPicture =   "frmLabRpt.frx":0D1A
      DownPicture     =   "frmLabRpt.frx":10E1
      Height          =   495
      Left            =   5400
      Picture         =   "frmLabRpt.frx":1632
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Status"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Yearly Report"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   3480
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Montly Report"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "To"
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
      Left            =   7200
      TabIndex        =   8
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   7320
      Top             =   4440
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   8880
      X2              =   2760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   135
   End
End
Attribute VB_Name = "frmLabRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodeV As Integer
Dim DateCounter As Integer
Dim YearCounter As Integer
Dim StatusCounter As Integer
Dim StatusV As String
Dim MonthCounter As Integer
Dim samplingdate As Date
Private Sub cmbstatus_Click()
If Me.cmbstatus.Text = "Done" Then StatusV = "D"
If Me.cmbstatus.Text = "Not Done" Then StatusV = "N"
End Sub

Private Sub CmdCancel_Click()
Unload Me
frmMain.Show
End Sub

Private Sub CmdOK_Click()
Select Case CodeV
  Case 1
    If DateCounter = 0 Then
      DE1.datetodate Me.FDate, Me.TDate
      datewiserpt.Show
    Else
       DE1.rsdatetodate.Close
       DE1.datetodate Me.FDate, Me.TDate
       datewiserpt.Show
    End If
    DateCounter = DateCounter + 1
    
   Case 3
      If YearCounter = 0 Then
        DE1.yearlyrpt_Grouping Me.cmbyear
        yearlyrpt.Show
      Else
        DE1.rsyearlyrpt_Grouping.Close
        DE1.yearlyrpt_Grouping Me.cmbyear
        yearlyrpt.Show
      End If
        YearCounter = YearCounter + 1
    Case 4
    If Me.cmbstatus.Text = "" Then MsgBox "Value Missing": Exit Sub
    
        If StatusCounter = 0 Then
          DE1.statusrpt_Grouping (StatusV)
          statusrpt.Show
        Else
          DE1.rsstatusrpt_Grouping.Close
          DE1.statusrpt_Grouping StatusV
          statusrpt.Show
        End If
        StatusCounter = StatusCounter + 1
        
     Case 2
       If MonthCounter = 0 Then
       DE1.monthrpt_Grouping Me.MonthView
       monthrpt.Show
       Else
       DE1.rsmonthrpt_Grouping.Close
       DE1.monthrpt_Grouping Me.MonthView
       monthrpt.Show
       End If
       MonthCounter = MonthCounter + 1
  End Select
End Sub

Private Sub Form_Load()
For i = 2000 To 2050
Me.cmbyear.AddItem i
Next
CodeV = 1
DateCounter = 0
YearCounter = 0
StatusCounter = 0
MonthCounter = 0
End Sub

Private Sub Option1_Click()
CodeV = 1
Me.FDate.Visible = True
Me.TDate.Visible = True
Me.MonthView.Visible = False
Me.cmbyear.Visible = False
Me.cmbstatus.Visible = False
End Sub

Private Sub Option2_Click()
CodeV = 2
Me.MonthView.Visible = True
Me.FDate.Visible = False
Me.TDate.Visible = False
Me.cmbyear.Visible = False
Me.cmbstatus.Visible = False
End Sub
Private Sub Option3_Click()
CodeV = 3
Me.cmbyear.Visible = True
Me.MonthView.Visible = False
Me.FDate.Visible = False
Me.TDate.Visible = False
Me.cmbstatus.Visible = False
End Sub
Private Sub Option4_Click()
CodeV = 4
Me.MonthView.Visible = False
Me.FDate.Visible = False
Me.TDate.Visible = False
Me.cmbyear.Visible = False
Me.cmbstatus.Visible = True

End Sub
