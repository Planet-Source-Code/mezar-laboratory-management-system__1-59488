VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSubTestEntry 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Sub Test Entry"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   21
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   2040
      TabIndex        =   12
      Top             =   6240
      Width           =   6375
      Begin VB.CommandButton CmdBack 
         DisabledPicture =   "frmSubTestEntry.frx":0000
         DownPicture     =   "frmSubTestEntry.frx":03B9
         Height          =   375
         Left            =   4800
         Picture         =   "frmSubTestEntry.frx":08FB
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdEdit 
         DisabledPicture =   "frmSubTestEntry.frx":0D83
         DownPicture     =   "frmSubTestEntry.frx":117B
         Height          =   375
         Left            =   3240
         Picture         =   "frmSubTestEntry.frx":16CA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdLast 
         DisabledPicture =   "frmSubTestEntry.frx":1B6B
         DownPicture     =   "frmSubTestEntry.frx":1F1F
         Height          =   375
         Left            =   4800
         Picture         =   "frmSubTestEntry.frx":2410
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdNext 
         DisabledPicture =   "frmSubTestEntry.frx":2858
         DownPicture     =   "frmSubTestEntry.frx":2C1D
         Height          =   375
         Left            =   3240
         Picture         =   "frmSubTestEntry.frx":30D5
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdPrev 
         DisabledPicture =   "frmSubTestEntry.frx":356E
         DownPicture     =   "frmSubTestEntry.frx":3948
         Height          =   375
         Left            =   1680
         Picture         =   "frmSubTestEntry.frx":3EA3
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdFirst 
         DisabledPicture =   "frmSubTestEntry.frx":4367
         DownPicture     =   "frmSubTestEntry.frx":4724
         Height          =   375
         Left            =   120
         Picture         =   "frmSubTestEntry.frx":4C72
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         DisabledPicture =   "frmSubTestEntry.frx":5174
         DownPicture     =   "frmSubTestEntry.frx":55A1
         Height          =   375
         Left            =   1680
         Picture         =   "frmSubTestEntry.frx":5AF4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdNew 
         DisabledPicture =   "frmSubTestEntry.frx":5FBA
         DownPicture     =   "frmSubTestEntry.frx":6381
         Height          =   375
         Left            =   120
         Picture         =   "frmSubTestEntry.frx":6861
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtsubtestcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox cmbtestcode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5160
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ComboBox cmbcatgcode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5160
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   7
      Text            =   " "
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1575
      Left            =   960
      TabIndex        =   6
      Top             =   4320
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":6CAF
            Key             =   "New"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":71F1
            Key             =   "Save"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":7733
            Key             =   "First"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":7EF6
            Key             =   "Previous"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":86D7
            Key             =   "Next"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":8EAF
            Key             =   "Last"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":955C
            Key             =   "Delete"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubTestEntry.frx":966E
            Key             =   "Back"
            Object.Tag             =   "8"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First"
            Object.ToolTipText     =   "First"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Previous"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last"
            Object.ToolTipText     =   "Last"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Sub Test"
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
      TabIndex        =   11
      Top             =   2640
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   9960
      X2              =   1560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   8520
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   9960
      X2              =   1560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tests Setup"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Test Chart"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   3000
      Width           =   1110
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Category Code"
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
      Index           =   1
      Left            =   3150
      TabIndex        =   1
      Top             =   1785
      Width           =   1905
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   2265
      Width           =   960
   End
End
Attribute VB_Name = "frmSubTestEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCatg As New ADODB.Recordset
Dim RsSubTest As New ADODB.Recordset
Dim CHECK As Integer
Dim Rstest As New ADODB.Recordset
Dim Rstestcode As New ADODB.Recordset
Dim disp As New LabCLS
Dim RsGenPk As New ADODB.Recordset
Dim RsCheckPK As New ADODB.Recordset
Dim TestCatgCodeV As Integer
Dim TestCodeV As Integer
Dim obj As Object
Dim SubTestcodeV As Integer
Dim EditcodeV As Integer
Private Sub cmbcatgcode_Click()
Me.cmbtestcode.Clear

 With RsCatg
   If Not .EOF Then
     .MoveFirst
     .Find ("description='" & Me.cmbcatgcode & "'")
     TestCatgCodeV = .Fields("test_catgcode")
   '  Me.MSF1.TextMatrix(1, 1) = TestCatgCodeV
   '  Me.MSF1.TextMatrix(1, 2) = ""
   End If
End With
Set Rstest = disp.openrec("select * from test where test_catgcode=" & TestCatgCodeV & "")
With Rstest
  If Not .EOF Then
    While Not .EOF
      Me.cmbtestcode.AddItem .Fields("description")
      .MoveNext
    Wend
  End If
End With
End Sub

Private Sub cmbtestcode_Click()
Set Rstest = disp.openrec("select * from test")
With Rstest
  If Not .EOF Then
    .Find ("description='" & Me.cmbtestcode & "'")
    TestCodeV = .Fields("test_code")
    Me.MSF1.TextMatrix(1, 1) = TestCodeV
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

Private Sub CmdEdit_Click()
Me.cmdsave.Enabled = True
EditcodeV = 1
End Sub

Private Sub CmdFirst_Click()
With RsSubTest
  If Not .EOF Then
    .MoveFirst
    Call display
  End If
  End With
End Sub

 Sub display()

With RsSubTest
If Not .EOF And Not .BOF Then

 Me.txtsubtestcode = .Fields("subtest_code")
 Me.txtDescription = .Fields("description")
 TestCodeV = .Fields("test_code")
 '****************************************
  Set Rstest = disp.openrec("select * from test where test_code='" & TestCodeV & "'")
   With Rstest
    If Not .EOF Then
      Me.cmbtestcode.Text = .Fields("description")
      TestCatgCodeV = .Fields("test_catgcode")
    End If
   End With
 '****************************************
 With RsCatg
 If Not .EOF Then
 .MoveFirst
  .Find ("test_catgcode=" & TestCatgCodeV & "")
   Me.cmbcatgcode.Text = .Fields("description")
 End If
 End With
'************************************
 Me.MSF1.TextMatrix(1, 1) = TestCodeV
  Me.MSF1.TextMatrix(1, 2) = .Fields(0)
 Me.MSF1.TextMatrix(1, 3) = .Fields("unit")
  Me.MSF1.TextMatrix(1, 4) = .Fields("normal_range")
   Me.MSF1.TextMatrix(1, 5) = .Fields("duration")
    Me.MSF1.TextMatrix(1, 6) = .Fields("dur_unit")
     Me.MSF1.TextMatrix(1, 7) = .Fields("charges")
 
 
 End If
 End With
 Me.cmdsave.Enabled = False
 EditcodeV = 0
End Sub


Private Sub CmdLast_Click()
With RsSubTest
  If Not .EOF Then
   .MoveLast
   Call display
 End If
End With
End Sub

Private Sub cmdnew_Click()
Me.Text1.Visible = False
Me.cmbcatgcode.Enabled = True
Me.cmbtestcode.Enabled = True
CHECK = 0
 Me.MSF1.Clear
Set RsGenPk = disp.openrec("select * from subtest")
 With RsGenPk
   If Not .EOF Then
     .MoveLast
     Me.txtsubtestcode = .Fields("subtest_code") + 1
   Else
     Me.txtsubtestcode = 1
   End If
 End With
 Me.MSF1.TextMatrix(1, 2) = Me.txtsubtestcode
 Me.cmdsave.Enabled = True
 Me.txtDescription = ""

 Call msfheading
 Me.cmbcatgcode.Text = ""
 Me.cmbtestcode = ""
End Sub

Private Sub CmdNext_Click()
With RsSubTest
  If Not .EOF Then
    .MoveNext
     If .EOF Then
       .MoveLast
     End If
  End If
  Call display
End With
End Sub

Private Sub CmdPrev_Click()
With RsSubTest
  If Not .BOF Then
    .MovePrevious
    If .BOF Then
     .MoveFirst
    End If
  End If
  Call display
 End With
End Sub

Private Sub cmdsave_Click()
CHECK = 1
For Each obj In Me
  If TypeOf obj Is ComboBox Or TypeOf obj Is TextBox And obj.Name <> "text1" Then
  If obj.Text = "" Then MsgBox "give full information": Exit Sub
  End If
Next
 
  
If Me.MSF1.TextMatrix(1, 1) = "" Or Me.MSF1.TextMatrix(1, 2) = "" Or Me.MSF1.TextMatrix(1, 3) = "" Or Me.MSF1.TextMatrix(1, 4) = "" Or Me.MSF1.TextMatrix(1, 5) = "" Or Me.MSF1.TextMatrix(1, 6) = "" Or Me.MSF1.TextMatrix(1, 7) = "" Then MsgBox "give full information": Exit Sub
If EditcodeV = 1 Then
disp.update ("update subtest set test_catgcode=" & TestCatgCodeV & ",test_code='" & Me.MSF1.TextMatrix(1, 1) & "',description='" & Me.txtDescription & "',unit='" & Me.MSF1.TextMatrix(1, 3) & "',normal_range='" & Me.MSF1.TextMatrix(1, 4) & "',duration=" & Me.MSF1.TextMatrix(1, 5) & ",dur_unit='" & Me.MSF1.TextMatrix(1, 6) & " ',charges=" & Me.MSF1.TextMatrix(1, 7) & " where subtest_code='" & Me.MSF1.TextMatrix(1, 2) & "'")
EditcodeV = 0
Exit Sub
End If

Set RsCheckPK = disp.openrec("select * from subtest where subtest_code='" & Me.txtsubtestcode & "'")
With RsCheckPK
  If Not .EOF Or .RecordCount > 0 Then MsgBox "Already Exist", vbOKOnly + vbCritical, "Warning": Exit Sub
 End With

disp.add ("insert into subtest values(" & TestCatgCodeV & ",'" & Me.MSF1.TextMatrix(1, 2) & "','" & Me.MSF1.TextMatrix(1, 1) & "','" & Me.txtDescription & "','" & Me.MSF1.TextMatrix(1, 3) & "','" & Me.MSF1.TextMatrix(1, 4) & "'," & Me.MSF1.TextMatrix(1, 5) & ",'" & Me.MSF1.TextMatrix(1, 6) & " '," & Me.MSF1.TextMatrix(1, 7) & ")")
Me.cmdsave.Enabled = False
End Sub

Private Sub CmdSearch_Click()

End Sub

Private Sub Form_Load()
Call OpenDB
Set disp = New LabCLS
Call msfheading
Set RsSubTest = disp.openrec("select * from subtest")
Set RsCatg = disp.openrec("select * from test_catg")
With RsCatg
  If Not .EOF Then
     While Not .EOF
       Me.cmbcatgcode.AddItem .Fields("description")
       .MoveNext
     Wend
     .MoveFirst
  End If
End With
EditcodeV = 0
End Sub

Private Sub MSF1_Click()
Call msf1_entercell
End Sub

Private Sub MSF1_GotFocus()
If Me.MSF1.MouseRow = 0 Then Me.Text1.Visible = False

End Sub

Private Sub MSF1_LeaveCell()
If MSF1.Col = 1 Or Me.MSF1.Col = 2 Then Exit Sub
  MSF1.Text = Text1.Text
    Text1.Text = ""
    Text1.Visible = False
'End If
End Sub


Private Sub MSF1_Scroll()
Me.Text1.Visible = False
End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
   Case "save"
      Call cmdsave_Click
   Case "First"
      Call CmdFirst_Click
   Case "Next"
      Call CmdNext_Click
   Case "Previous"
      Call CmdPrev_Click
   Case "Last"
      Call CmdLast_Click
   Case "Back"
      Call CmdBack_Click
   Case "New"
      Call cmdnew_Click
End Select
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        If MSF1.Row > 1 Then
                MSF1.Row = MSF1.Row - 1
        End If

    Case vbKeyDown
        If MSF1.Row < MSF1.Rows - 1 Then
                MSF1.Row = MSF1.Row + 1
        End If
    End Select
If MSF1.Col = 3 Then
MSF1.Text = Text1.Text
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Me.MSF1.Text = Me.Text1.Text
If Me.MSF1.Col = 6 Then
KeyAscii = Character(KeyAscii)
End If
If Me.MSF1.Col = 7 Or Me.MSF1.Col = 5 Then KeyAscii = numeric(KeyAscii)
'Dim srow, scol As Integer
'If KeyAscii = 13 Then
'    If MSF1.Col = 3 Then
'    MSF1.Text = Text1.Text
'    srow = MSF1.Row + 1
'    scol = MSF1.ColSel
'    If srow = MSF1.Rows Then
'        MSF1.Rows = MSF1.Rows + 1
'        srow = MSF1.Rows - 1
'    End If
'    End If
'    MSF1.Row = srow
'    MSF1.Col = scol
'   MSF1.RowSel = srow
'    MSF1.ColSel = scol
'    Text1.Text = MSF1.Text
'    KeyAscii = 0
'End If
End Sub

Private Sub msf1_entercell()
If MSF1.MouseRow = 0 Or MSF1.MouseCol = 0 Then
   Text1.Visible = False
   Exit Sub
End If
Text1.Text = ""
If MSF1.Col = 1 Or Me.MSF1.Col = 2 Then Exit Sub
   
    Text1.Top = MSF1.Top + MSF1.CellTop
     Text1.Left = MSF1.Left + MSF1.CellLeft
    Text1.Width = MSF1.CellWidth
    Text1.Height = MSF1.CellHeight
    Text1.Visible = True
    Text1.SetFocus
    Text1.Text = MSF1.Text
    Text1.Visible = True


    
'End If
End Sub
Sub msfheading()
Me.MSF1.TextMatrix(0, 1) = "Test code"
Me.MSF1.TextMatrix(0, 2) = "Subtest code"
'Me.MSF1.TextMatrix(0, 3) = "Description"
Me.MSF1.TextMatrix(0, 3) = "Unit"
Me.MSF1.TextMatrix(0, 4) = "Normal Rang"
Me.MSF1.TextMatrix(0, 5) = "Duration"
Me.MSF1.TextMatrix(0, 6) = "Dur Unit"
Me.MSF1.TextMatrix(0, 7) = "Charges"

End Sub



Private Sub txtDescription_KeyPress(KeyAscii As Integer)
KeyAscii = Character(KeyAscii)

End Sub
