VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmtestresult 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmtestresullt"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5175
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   9375
      Begin VB.ComboBox CmbLabNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   2160
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtptname 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Txtremarks 
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
         ForeColor       =   &H00C00000&
         Height          =   765
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   3720
         Width           =   6615
      End
      Begin VB.TextBox txtPtid 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtptyear 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4200
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Cmbdoneby 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox sampling_date 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox sampling_time 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtFname 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cmbTestCatg 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   4215
      End
      Begin VB.ComboBox cmbtest 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   1455
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   8
         ForeColor       =   12582912
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
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
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Date(DD/MM/YYYY)"
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
         Left            =   4440
         TabIndex        =   28
         Top             =   600
         Width           =   2700
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Time(HH:MM:SS)"
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
         Left            =   4440
         TabIndex        =   27
         Top             =   960
         Width           =   2370
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   26
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performed By"
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
         Left            =   120
         TabIndex        =   23
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LabNo"
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
         Left            =   6840
         TabIndex        =   22
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name"
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
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Catagory"
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   9960
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
      Begin VB.CommandButton cmdPrint 
         DisabledPicture =   "frmtestresult.frx":0000
         DownPicture     =   "frmtestresult.frx":0441
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "frmtestresult.frx":09B1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdsave 
         DisabledPicture =   "frmtestresult.frx":0E4F
         DownPicture     =   "frmtestresult.frx":127C
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "frmtestresult.frx":17CF
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdclear 
         DisabledPicture =   "frmtestresult.frx":1C95
         DownPicture     =   "frmtestresult.frx":202E
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "frmtestresult.frx":252F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdback 
         DisabledPicture =   "frmtestresult.frx":29A6
         DownPicture     =   "frmtestresult.frx":2D5F
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "frmtestresult.frx":32A1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":3729
            Key             =   "New"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":3787
            Key             =   "Open"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":37E5
            Key             =   "Save"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":3843
            Key             =   "First"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":38A1
            Key             =   "Previous"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":38FF
            Key             =   "Next"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":395D
            Key             =   "Last"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtestresult.frx":39BB
            Key             =   "Back"
            Object.Tag             =   "8"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Result"
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
      Left            =   720
      TabIndex        =   30
      Top             =   360
      Width           =   2640
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   9240
      X2              =   720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   600
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   960
      Top             =   7800
   End
End
Attribute VB_Name = "frmtestresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_PatientTest As New ADODB.Recordset
Dim Rs_Patient As New ADODB.Recordset
Dim RsSubTest As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset
Dim RsFindTestCode As New ADODB.Recordset
Dim Rstest As New ADODB.Recordset
Dim Rs_TestCatg As New ADODB.Recordset
Dim Rs_FindTestCatg As New ADODB.Recordset
Dim RsUserlogin As New ADODB.Recordset
Dim RsFindTestCodeV As String
Dim TestCatgCodeV As String
Dim TestCodeV As String
Dim disp As New LabCLS
Dim SubTestcodeV As String
Dim PtidCodeV  As String
Dim obj As Object
Dim Counter As Integer

Private Sub CmbLabNo_Click()
Call display
End Sub
Private Sub CmbLabNo_KeyPress(KeyAscii As Integer)
  If Me.CmbLabNo.Text = "" Then Exit Sub
  If KeyAscii = 13 Then Call display
    
End Sub
Sub display()
     Me.Text1.Visible = False
     Me.cmbTestCatg.Clear
     Me.cmbtest.Clear
     Me.MSF1.Clear
     Me.sampling_date.Text = ""
     Me.sampling_time.Text = ""
     Me.MSF1.Rows = 1
     Call fillmsf
     Set Rs_Patient = disp.openrec("select * from patient where pt_id='" & Me.txtptid.Text & "' and pt_year=" & Me.txtptyear.Text & " ")
       With Rs_Patient
         If Not .EOF Then
           Me.txtptname.Text = .Fields("pt_name")
           Me.txtFname.Text = .Fields("pt_fname")
           PtidCodeV = .Fields("pt_id")
         End If
      End With
   
 
Set Rs_TestCatg = disp.openrec("select * from test_catg")
Set Rs_FindTestCatg = disp.openrec("select distinct test_catgcode from pt_test where pt_id='" & PtidCodeV & "' and pt_year= " & Me.txtptyear.Text & " and  labno='" & Me.CmbLabNo.Text & "' and status='N' ")
    With Rs_FindTestCatg
      If Not .EOF Then
        Rs_FindTestCatg.MoveFirst
        Rs_TestCatg.MoveFirst
         While Not .EOF
           TestCatgCodeV = .Fields("test_catgcode")
           Rs_TestCatg.Find ("test_catgcode='" & TestCatgCodeV & "'")
           Me.cmbTestCatg.AddItem Rs_TestCatg.Fields("description")
          .MoveNext
          Rs_TestCatg.MoveNext
        Wend
       .MoveFirst
     Else
       MsgBox " This Test is Already Done"
       Me.CmbLabNo.SetFocus
       Exit Sub
    End If
   End With
   Me.cmbTestCatg.SetFocus

End Sub
Private Sub cmbtest_Click()
Dim Crow As Integer
Dim Ccol As Integer
 Me.Text1.Text = ""
 Me.sampling_date.Text = ""
 Me.sampling_time = ""
 Me.Text1.Visible = False
'******************************************************
Dim Rs_Test As New ADODB.Recordset
Set Rs_Test = disp.openrec("select * from test")
  With Rs_Test
   If Not .EOF Then
     .Find ("description='" & Me.cmbtest.Text & "'")
     TestCodeV = .Fields("test_code")
   End If
  End With
'******************************************************
Me.MSF1.Clear
Me.MSF1.Rows = 2
Set rsfind = disp.openrec("select pt_test.subtest_code," & _
" pt_test.test_code,                                    " & _
" pt_test.test_catgcode,subtest.description,subtest.subtest_code," & _
" subtest.normal_range,subtest.unit,subtest.charges,pt_test.Sampling_date," & _
" pt_test.Sampling_time  " & _
" from pt_test,subtest,test,test_catg " & _
" where pt_test.pt_id=" & Me.txtptid.Text & "             " & _
" and pt_test.pt_year= " & Me.txtptyear.Text & " and status='N' and     " & _
" pt_test.labno= " & Me.CmbLabNo.Text & "   and  pt_test.test_catgcode='" & TestCatgCodeV & "' and   pt_test.test_code='" & TestCodeV & "' and       " & _
" pt_test.subtest_code=subtest.subtest_code  and          " & _
" pt_test.test_catgcode=test_catg.test_catgcode  and " & _
" pt_test.test_code=test.test_code ")
   
Call fillmsf
 With rsfind
   Crow = 1
   If Not .EOF Then
     Me.MSF1.Rows = 1
     While Not .EOF
        Me.MSF1.Rows = Me.MSF1.Rows + 1
        Me.MSF1.TextMatrix(Crow, 0) = .Fields("subtest_code")
        Me.MSF1.ColWidth(1) = 2500
        Me.MSF1.TextMatrix(Crow, 1) = .Fields("description")
        Me.MSF1.ColWidth(2) = 4210
        Me.MSF1.TextMatrix(Crow, 2) = ""
        Me.MSF1.TextMatrix(Crow, 3) = .Fields("unit")
        Me.MSF1.TextMatrix(Crow, 4) = .Fields("normal_range")
        Me.MSF1.TextMatrix(Crow, 5) = .Fields("charges")
        Me.MSF1.ColWidth(7) = 0
        Me.MSF1.ColWidth(6) = 0
        Me.MSF1.TextMatrix(Crow, 6) = .Fields("Sampling_time")
        Me.MSF1.TextMatrix(Crow, 7) = .Fields("Sampling_date")
        Crow = Crow + 1
       .MoveNext
     Wend
   Else
      MsgBox " This Test is Already Done"
      Me.MSF1.Rows = 1
  End If
End With
End Sub
Private Sub cmbTestCatg_Click()
Set Rs_TestCatg = disp.openrec("select * from test_catg")
With Rs_TestCatg
  If Not .EOF Then
    .Find ("description='" & Me.cmbTestCatg.Text & "'")
    TestCatgCodeV = .Fields("test_catgcode")
  End If
End With
Me.cmbtest.Clear
Me.MSF1.Clear
Me.MSF1.Rows = 1
Call fillmsf
Set Rs_PatientTest = disp.openrec("select distinct test_code,test_catgcode from pt_test where status='N'and pt_id=" & Me.txtptid.Text & " and labno=" & Me.CmbLabNo.Text & " and pt_year=" & Me.txtptyear.Text & " and test_catgcode='" & TestCatgCodeV & "'")
Set Rs_TestCatg = disp.openrec("select test_catgcode from test_catg where description='" & Me.cmbTestCatg.Text & "'")
 '*************************************************************************
With Rs_PatientTest
   If Not .EOF Then
    .MoveFirst
      While Not .EOF
        SubTestcodeV = .Fields("test_code")
        Set RsSubTest = disp.openrec("select description from test where test_code='" & SubTestcodeV & "'")
         If RsSubTest.EOF = False Then
           Me.cmbtest.AddItem RsSubTest.Fields("description")
         End If
        .MoveNext
      Wend
   End If
  End With
 '*************************************************************************
End Sub
Private Sub cmbTestCatg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbtest.SetFocus
End Sub

Private Sub CmdBack_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdclear_Click()
For Each obj In Me
  If TypeOf obj Is TextBox Then
    obj.Text = ""
  End If
Next
Me.cmbtest.Clear
Me.cmbTestCatg.Clear
Me.MSF1.Rows = 1
Me.Text1.Visible = False
Call fillmsf
End Sub

Private Sub CmdPrint_Click()
 For Each obj In Me
   If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
    If obj.Text = "" Then
      MsgBox "give full information"
      Exit Sub
    End If
   End If
 Next

  If Counter = 0 Then
    DE1.result_Grouping Me.txtptid, Me.CmbLabNo, Me.txtptyear, Me.sampling_date
    resultrpt.Show
  Else
   DE1.rsresult_Grouping.Close
DE1.result_Grouping Me.txtptid, Me.CmbLabNo, Me.txtptyear, Me.sampling_date
    resultrpt.Show
 End If
   Counter = Counter + 1
End Sub

Private Sub cmdsave_Click()
Dim i As Integer
  For Each obj In Me
   If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
    If obj.Text = "" Then
      MsgBox "give full information"
      Exit Sub
    End If
   End If
 Next
For i = 1 To Me.MSF1.Rows - 1
  If Me.MSF1.TextMatrix(i, 2) <> "" Then
        disp.update ("update pt_test set status='D',   " & _
    " result='" & Me.MSF1.TextMatrix(i, 2) & " ', remarks='" & Me.Txtremarks.Text & "', " & _
    " doneby='" & Me.Cmbdoneby.Text & "', sampling_date= '" & Me.sampling_date.Text & "',       " & _
    " sampling_time= ' " & Me.sampling_time.Text & "',  " & _
    " unit='" & Me.MSF1.TextMatrix(i, 3) & " ',normal_range='" & Me.MSF1.TextMatrix(i, 4) & " '         " & _
    " where  pt_id = " & Me.txtptid.Text & " and pt_year = " & Me.txtptyear.Text & "      " & _
    " and labno=" & Me.CmbLabNo.Text & "  and test_catgcode='" & TestCatgCodeV & "'      " & _
    " and test_code='" & TestCodeV & "'  and subtest_code='" & Me.MSF1.TextMatrix(i, 0) & "'  " & _
    " and Sampling_date='" & Me.MSF1.TextMatrix(i, 7) & "'  and Sampling_time=               " & _
    "  '" & Me.MSF1.TextMatrix(i, 6) & "'  ")
 End If
Next
Me.CmdPrint.Enabled = True
End Sub

Private Sub cmdViewRpt_Click()

End Sub

Private Sub Form_Load()
Me.MSF1.Rows = 1
Dim disp As New LabCLS
  Set RsSubTest = disp.openrec("select * from subtest")
  Set Rs_TestCatg = disp.openrec("select * from test_catg")
  Set RsUserlogin = disp.openrec("select * from employee")
  With RsUserlogin
    If Not .EOF Then
      While Not .EOF
        Me.Cmbdoneby.AddItem .Fields("emp_Name")
        .MoveNext
      Wend
    End If
  End With
  Counter = 0
  Call fillmsf
 End Sub



Private Sub MSF1_Click()
    Me.sampling_time.Text = Me.MSF1.TextMatrix(Me.MSF1.Row, 6)
    Me.sampling_date.Text = Me.MSF1.TextMatrix(Me.MSF1.Row, 7)
End Sub
Private Sub MSF1_LeaveCell()
 If MSF1.Col = 2 Then
     MSF1.Text = Text1.Text
     Text1.Text = ""
     Text1.Visible = False
 End If
End Sub
Private Sub MSF1_Scroll()
   Call msf1_entercell
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        If MSF1.Row > 1 Then
                MSF1.Row = MSF1.Row - 1
                Call MSF1_Click
        End If
    Case vbKeyDown
        If MSF1.Row < MSF1.Rows - 1 Then
            MSF1.Row = MSF1.Row + 1
            Call MSF1_Click
        End If
    End Select
   If MSF1.Col = 2 Then
     MSF1.Text = Text1.Text
   End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim srow, scol As Integer
  If KeyAscii = 13 Then
      If MSF1.Col = 2 Then
         MSF1.Text = Text1.Text
         srow = MSF1.Row + 1
         scol = MSF1.ColSel
      End If
      Text1.Text = MSF1.Text
      KeyAscii = 0
  End If
End Sub
Private Sub msf1_entercell()
If MSF1.MouseRow = 0 Or MSF1.MouseCol = 0 Then
   Text1.Visible = False
   Exit Sub
End If
Text1.Text = ""
  If MSF1.Col = 2 Then
     Text1.Top = MSF1.Top + MSF1.CellTop
     Text1.Left = MSF1.Left + MSF1.CellLeft
     Text1.Width = MSF1.CellWidth
     Text1.Height = MSF1.CellHeight
     Text1.Visible = True
     Text1.SetFocus
     Text1.Text = MSF1.Text
     Text1.Visible = True
  End If
End Sub


Private Sub txtPtid_Change()
Me.txtptyear = ""
Exit Sub
End Sub

Private Sub txtPtid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtptyear.SetFocus ':   Me.txtptyear.Text = ""
      If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 8) Then
         KeyAscii = 0
      Else
         Exit Sub
      End If
     KeyAscii = numeric(KeyAscii)
End Sub
Private Sub txtptyear_Change()
Me.CmbLabNo.Clear
End Sub
Private Sub txtptyear_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 8) Then
        KeyAscii = 13
    Else
        Exit Sub
    End If

If KeyAscii = 13 Then
      
     If Me.txtptyear.Text = "" Then
       Me.txtptyear.SetFocus
       Exit Sub
     End If
     If Me.txtptid = "" Then MsgBox "give patient id": Exit Sub
     Me.CmbLabNo.Clear
     Me.CmbLabNo.SetFocus
     Set Rs_PatientTest = disp.openrec("select  distinct labno from pt_test where pt_id='" & Me.txtptid.Text & "' and pt_year=" & Me.txtptyear.Text & " and status='N' group by labno")
     With Rs_PatientTest
       If Not Rs_PatientTest.EOF Then
         While Not .EOF
            Me.CmbLabNo.AddItem .Fields("labno")
            .MoveNext
         Wend
       Else: MsgBox "no record exist"
       Me.txtptid.Text = ""
       Me.txtptyear.Text = ""
       Me.txtptid.SetFocus
     End If
End With
End If
End Sub
Sub fillmsf()
      Me.MSF1.TextMatrix(0, 0) = "TestCode"
      Me.MSF1.TextMatrix(0, 1) = "Test"
      Me.MSF1.TextMatrix(0, 2) = "Result"
      Me.MSF1.TextMatrix(0, 3) = "Unit"
      Me.MSF1.TextMatrix(0, 4) = "Normal Range"
      Me.MSF1.TextMatrix(0, 5) = "Charges"
      'Me.MSF1.TextMatrix(0, 6) = "Sampling_date"
End Sub
Private Sub Txtremarks_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Me.Cmbdoneby.SetFocus
End Sub
