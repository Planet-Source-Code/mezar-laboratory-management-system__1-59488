VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestEntry 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Test Entry"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdadd 
      Caption         =   "add new"
      Enabled         =   0   'False
      Height          =   245
      Left            =   2790
      TabIndex        =   19
      Top             =   3520
      Width           =   960
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   2640
      TabIndex        =   10
      Top             =   5160
      Width           =   6375
      Begin VB.CommandButton CmdNew 
         DisabledPicture =   "frmTestEntry.frx":0000
         DownPicture     =   "frmTestEntry.frx":03C7
         Height          =   375
         Left            =   120
         Picture         =   "frmTestEntry.frx":08A7
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         DisabledPicture =   "frmTestEntry.frx":0CF5
         DownPicture     =   "frmTestEntry.frx":1122
         Height          =   375
         Left            =   1680
         Picture         =   "frmTestEntry.frx":1675
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdFirst 
         DisabledPicture =   "frmTestEntry.frx":1B3B
         DownPicture     =   "frmTestEntry.frx":1EF8
         Height          =   375
         Left            =   120
         Picture         =   "frmTestEntry.frx":2446
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdPrev 
         DisabledPicture =   "frmTestEntry.frx":2948
         DownPicture     =   "frmTestEntry.frx":2D22
         Height          =   375
         Left            =   1680
         Picture         =   "frmTestEntry.frx":327D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdNext 
         DisabledPicture =   "frmTestEntry.frx":3741
         DownPicture     =   "frmTestEntry.frx":3B06
         Height          =   375
         Left            =   3240
         Picture         =   "frmTestEntry.frx":3FBE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdLast 
         DisabledPicture =   "frmTestEntry.frx":4457
         DownPicture     =   "frmTestEntry.frx":480B
         Height          =   375
         Left            =   4800
         Picture         =   "frmTestEntry.frx":4CFC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdEdit 
         DisabledPicture =   "frmTestEntry.frx":5144
         DownPicture     =   "frmTestEntry.frx":553C
         Height          =   375
         Left            =   3240
         Picture         =   "frmTestEntry.frx":5A8B
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdBack 
         DisabledPicture =   "frmTestEntry.frx":5F2C
         DownPicture     =   "frmTestEntry.frx":62E5
         Height          =   375
         Left            =   4800
         Picture         =   "frmTestEntry.frx":6827
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1455
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   4
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":6CAF
            Key             =   "New"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":71F1
            Key             =   "Save"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":7733
            Key             =   "Edit"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":7B85
            Key             =   "First"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":8348
            Key             =   "Previous"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":8B29
            Key             =   "Next"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":9301
            Key             =   "Last"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":99AE
            Key             =   "Delete"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEntry.frx":9AC0
            Key             =   "Back"
            Object.Tag             =   "9"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   "1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First"
            Object.ToolTipText     =   "First"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Previous"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last"
            Object.ToolTipText     =   "Last"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      DataField       =   "Description"
      DataMember      =   "CmdTestCatg"
      DataSource      =   "DELab"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtTestCatgcode 
      Appearance      =   0  'Flat
      DataField       =   "Test_Catgcode"
      DataMember      =   "CmdTestCatg"
      DataSource      =   "DELab"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   1650
   End
   Begin VB.CommandButton CmdSearch 
      DisabledPicture =   "frmTestEntry.frx":9F12
      DownPicture     =   "frmTestEntry.frx":A31A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6540
      Picture         =   "frmTestEntry.frx":A867
      TabIndex        =   0
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   9120
      X2              =   3000
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   9000
      X2              =   2880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Categories"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   720
      Width           =   3330
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
      Index           =   0
      Left            =   2655
      TabIndex        =   3
      Top             =   1680
      Width           =   1905
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
      Index           =   1
      Left            =   2805
      TabIndex        =   2
      Top             =   2175
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Chart"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   2640
      Width           =   2220
   End
End
Attribute VB_Name = "frmTestEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDesignation As New ADODB.Recordset
Dim RsGenId As New ADODB.Recordset
Dim Rs_check_Pk As New ADODB.Recordset
Dim Rstest As New ADODB.Recordset
Dim EditcodeV As Integer
Dim RsTestmax As New ADODB.Recordset
Dim Test As New ADODB.Recordset
Dim obj As Object
Dim disp As New LabCLS
Dim Save As Integer
Dim Save1 As Integer
Dim CHECK As Integer



Private Sub cmdadd_Click()
Set RsTestmax = openrec("select * from test")
With RsTestmax
  If Not .EOF Then
    Me.MSF1.Rows = 2
    .MoveLast
    Me.MSF1.TextMatrix(1, 2) = .Fields("test_code") + 1
    Else
    Me.MSF1.TextMatrix(1, 2) = 1
  End If
  Me.MSF1.TextMatrix(1, 1) = Me.txtTestCatgcode
  Me.MSF1.TextMatrix(1, 3) = ""
End With
Save1 = 1
Me.CmdSave.Enabled = True
Me.Text1.Text = ""
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

'Private Sub CmdDel_Click()
'EditcodeV = 0
'Dim a As String
'If Me.txtTestCatgcode = "" Or Me.txtDescription = "" Then MsgBox "can not delete": Exit Sub
'a = MsgBox("Are You Sure TO Delete The Record", vbYesNo)
'If a = vbYes Then
'disp.delete ("delete from test_catg where test_catgcode =" & Me.txtDesigID & "")
'Me.txtTestCatgcode = ""
'Me.txtDescription = ""
'End If
'
'End Sub

Private Sub CmdEdit_Click()
EditcodeV = 1
Me.CmdSave.Enabled = True
End Sub

Private Sub CmdFirst_Click()
With RsDesignation
  If Not .EOF Then
    .MoveFirst
    Call display
  End If
End With
End Sub
Sub display()
Dim r As Integer
Me.MSF1.Rows = 1
With RsDesignation
  If Not .EOF Then
    Me.txtTestCatgcode = .Fields("test_catgcode")
    Me.txtDescription = .Fields("description")
  End If
End With
If Me.txtTestCatgcode = "" Then Exit Sub
Set Rstest = openrec("select * from test where test_catgcode=" & Me.txtTestCatgcode & "")

With Rstest
  If Not .EOF Then
    While Not .EOF
      r = r + 1
      Me.MSF1.Rows = Me.MSF1.Rows + 1
      Me.MSF1.TextMatrix(r, 1) = .Fields("test_catgcode")
      Me.MSF1.TextMatrix(r, 2) = .Fields("test_code")
      Me.MSF1.TextMatrix(r, 3) = .Fields("description")
      .MoveNext
     Wend
   End If
 End With
 Save = 0
 Save1 = 0
Me.cmdadd.Enabled = True
EditcodeV = 0
Me.CmdSave.Enabled = False
Me.Text1.Visible = False
End Sub

Private Sub CmdLast_Click()
With RsDesignation
  If Not .EOF Then
    .MoveLast
     Call display
   End If
End With
End Sub

Private Sub cmdnew_Click()
CHECK = 0
Me.cmdadd.Enabled = True
Me.Text1.Visible = False
Me.txtDescription = ""
Set RsGenId = openrec("select * from test_catg ")
  With RsGenId
    If Not .EOF Then
      .MoveLast
      Me.txtTestCatgcode = .Fields("test_catgcode") + 1
    Else
      Me.txtTestCatgcode = 1
    End If
  End With
  Me.MSF1.Rows = 1
  EditcodeV = 0
  Save = 1
  Me.CmdSave.Enabled = True
End Sub

Private Sub CmdNext_Click()
With RsDesignation
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
With RsDesignation
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
If EditcodeV = 1 Then
  If Me.txtTestCatgcode = "" Or Me.txtDescription = "" Then MsgBox " give full informtion": Exit Sub
    For i = 1 To Me.MSF1.Rows
      If Me.MSF1.Rows > 1 Then
      If Me.MSF1.TextMatrix(1, 1) = "" Or Me.MSF1.TextMatrix(1, 2) = "" Or Me.MSF1.TextMatrix(1, 3) = "" Then
      MsgBox "give full information": Exit Sub
      End If
  End If
    Next
  disp.update ("update test_catg set description='" & Me.txtDescription & "' where test_catgcode=" & Me.txtTestCatgcode & "")
'  i = 1
  For i = 1 To Me.MSF1.Rows - 1
    disp.update ("update test set description='" & Me.MSF1.TextMatrix(i, 3) & "' where  test_code=" & Me.MSF1.TextMatrix(i, 2) & " AND test_catgcode=" & Me.MSF1.TextMatrix(i, 1) & "")
  Next
  EditcodeV = 0
  Me.CmdSave.Enabled = False
Exit Sub
End If





If Save = 1 And Save1 <> 1 Then

  If Me.txtTestCatgcode = "" Or Me.txtDescription = "" Then
    MsgBox " give full informtion": Exit Sub
  End If

Set Rs_check_Pk = openrec("select * from test_catg where test_catgcode=" & Me.txtTestCatgcode & "")
 With Rs_check_Pk
   If Not .EOF Then
      MsgBox " already exist", vbOKOnly + vbCritical, " Warning": Exit Sub
   End If
 End With


disp.add ("insert into test_catg values(" & Me.txtTestCatgcode & ",' " & Me.txtDescription & " ' )")
Save = 0
Else

If Save = 1 And Save1 = 1 Then
Set Rs_check_Pk = openrec("select * from test_catg where test_catgcode=" & Me.txtTestCatgcode & "")
 With Rs_check_Pk
   If Not .EOF Then
      MsgBox " already exist", vbOKOnly + vbCritical, " Warning": Exit Sub
   End If
 End With

  If Me.txtTestCatgcode = "" Or Me.txtDescription = "" Then
    MsgBox " give full informtion": Exit Sub
  End If
 If Me.MSF1.Rows > 1 Then
If Me.MSF1.TextMatrix(1, 3) = "" Then MsgBox "give description of test": Exit Sub
End If
disp.add ("insert into test_catg values(" & Me.txtTestCatgcode & ",' " & Me.txtDescription & " ' )")
With Test
  .AddNew
  .Fields("test_code") = Me.MSF1.TextMatrix(1, 2)
  .Fields("test_catgcode") = Me.MSF1.TextMatrix(1, 1)
  .Fields("description") = Me.MSF1.TextMatrix(1, 3)
  .update
End With
Else
With Test
If Me.MSF1.TextMatrix(1, 3) = "" Then MsgBox "give test description": Exit Sub
  .AddNew
  .Fields("test_code") = Me.MSF1.TextMatrix(1, 2)
  .Fields("test_catgcode") = Me.MSF1.TextMatrix(1, 1)
  .Fields("description") = Me.MSF1.TextMatrix(1, 3)
  .update
End With

End If
End If
Me.CmdSave.Enabled = False
End Sub
Sub checkpk()
Set Rs_check_Pk = openrec("select * from test_catg where test_catgcode=" & Me.txtTestCatgcode & "")
 With Rs_check_Pk
   If Not .EOF Then
      MsgBox " already exist", vbOKOnly + vbCritical, " Warning": Exit Sub
   End If
 End With

End Sub

Private Sub CmdSearch_Click()
Dim b As String
b = InputBox("Enter testcategory code", "Find")
If b = "" Then Exit Sub
If Not IsNumeric(b) Then MsgBox "give an integer value": Exit Sub
RsDesignation.Find ("test_catgcode=" & b & "")
Call display
End Sub

Private Sub Form_Load()
Call OpenDB
Set RsDesignation = openrec("select * from test_catg")
Set Test = openrec("select * from test")

EditcodeV = 0
Me.MSF1.Rows = 1

Me.MSF1.TextMatrix(0, 1) = "test_catgcode"
Me.MSF1.TextMatrix(0, 2) = "test code"
Me.MSF1.TextMatrix(0, 3) = "description"


Me.MSF1.ColWidth(1) = 1200
Me.MSF1.ColWidth(2) = 900
Me.MSF1.ColWidth(3) = 2700
Me.txtDescription = ""
Me.txtTestCatgcode = ""
Save = 0
Save1 = 0
End Sub

Private Sub MSF1_LeaveCell()

If MSF1.Col = 3 Then
  MSF1.Text = Text1.Text
    Text1.Text = ""
    Text1.Visible = False
End If
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
KeyAscii = Character(KeyAscii)
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
If MSF1.Col = 3 Then
   
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


