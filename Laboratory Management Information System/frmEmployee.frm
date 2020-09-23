VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmEmployee 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Employee"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1920
      TabIndex        =   27
      Top             =   5040
      Width           =   7935
      Begin VB.CommandButton CmdBack 
         DisabledPicture =   "frmEmployee.frx":0000
         DownPicture     =   "frmEmployee.frx":03B9
         Height          =   375
         Left            =   5640
         Picture         =   "frmEmployee.frx":08FB
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "edit"
         DisabledPicture =   "frmEmployee.frx":0D83
         DownPicture     =   "frmEmployee.frx":117B
         Height          =   375
         Left            =   4080
         Picture         =   "frmEmployee.frx":16CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdLast 
         DisabledPicture =   "frmEmployee.frx":1B6B
         DownPicture     =   "frmEmployee.frx":1F1F
         Height          =   375
         Left            =   5640
         Picture         =   "frmEmployee.frx":2410
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdNext 
         DisabledPicture =   "frmEmployee.frx":2858
         DownPicture     =   "frmEmployee.frx":2C1D
         Height          =   375
         Left            =   4080
         Picture         =   "frmEmployee.frx":30D5
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdPrev 
         DisabledPicture =   "frmEmployee.frx":356E
         DownPicture     =   "frmEmployee.frx":3948
         Height          =   375
         Left            =   2520
         Picture         =   "frmEmployee.frx":3EA3
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdFirst 
         DisabledPicture =   "frmEmployee.frx":4367
         DownPicture     =   "frmEmployee.frx":4724
         Height          =   375
         Left            =   960
         Picture         =   "frmEmployee.frx":4C72
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         DisabledPicture =   "frmEmployee.frx":5174
         DownPicture     =   "frmEmployee.frx":55A1
         Height          =   375
         Left            =   2520
         Picture         =   "frmEmployee.frx":5AF4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdNew 
         DisabledPicture =   "frmEmployee.frx":5FBA
         DownPicture     =   "frmEmployee.frx":6381
         Height          =   375
         Left            =   960
         Picture         =   "frmEmployee.frx":6861
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Index           =   0
      Left            =   1800
      TabIndex        =   16
      Top             =   1560
      Width           =   8295
      Begin VB.ComboBox CmbDesig 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker txtDOB 
         Height          =   255
         Left            =   1710
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54394881
         CurrentDate     =   37360
         MinDate         =   367
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         ToolTipText     =   "Search Employee"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox CmbSex 
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
         Height          =   315
         ItemData        =   "frmEmployee.frx":6CAF
         Left            =   6120
         List            =   "frmEmployee.frx":6CBC
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtemp_no 
         Alignment       =   1  'Right Justify
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
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtemp_name 
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
         Left            =   1710
         TabIndex        =   1
         Top             =   690
         Width           =   2655
      End
      Begin VB.TextBox txtph_no 
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
         Left            =   6120
         TabIndex        =   5
         Top             =   225
         Width           =   2055
      End
      Begin VB.TextBox txtfc_no 
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
         Left            =   6120
         TabIndex        =   6
         Top             =   712
         Width           =   2055
      End
      Begin VB.TextBox txtadd 
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
         Height          =   885
         Left            =   1710
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2160
         Width           =   6255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   225
         TabIndex        =   24
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   225
         TabIndex        =   23
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4605
         TabIndex        =   22
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4605
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Contact No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4605
         TabIndex        =   20
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   2160
         Width           =   690
      End
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   9480
      Top             =   720
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   10080
      X2              =   1920
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Information"
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
      Left            =   2040
      TabIndex        =   15
      Top             =   720
      Width           =   4740
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_max              As New ADODB.Recordset
Dim rs_Emp              As New ADODB.Recordset
Dim rs_Desig            As New ADODB.Recordset
Dim rs_EmpDesig         As New ADODB.Recordset

Dim addemp              As New LabCLS
Dim StrIns_Emp          As String
Dim DesigIDV            As Integer
Dim editV               As Integer
Dim strDel_Emp          As String
Dim StrDel_Sec          As String
Dim obj As Object
Dim CHECK As Integer
Private Sub CmbDesig_Click()
  With rs_Desig
      If Not .EOF Then
      .MoveFirst
      .Find ("description='" & Me.CmbDesig & "'")
      DesigIDV = .Fields("desig_id")
      End If
  End With
End Sub

Private Sub CmdBack_Click()
'Dim a As String
'If CHECK = 0 Then
'a = MsgBox("do you want to save the changes", vbYesNoCancel)
'If a = vbYes Then
'Call cmdsave_Click
'CHECK = 0
'ElseIf a = vbNo Then
Unload Me
'End If
'End If
'CHECK = 0


End Sub

'Private Sub CmdDel_Click()
'    Dim a As Integer
'    For Each obj In Me
'      If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
'        If obj.Text = "" Then MsgBox "can not delete": Exit Sub
'      End If
'    Next
'
'    a = MsgBox("Are you Sure", vbYesNoCancel)
'        If a = vbYes Then
'            merlin.Play "surprised"
'            merlin.Speak "\Chr=""Whisper""\Ooh hooo! You Have Deleted the Record"
'            StrDel_Sec = "delete from security where emp_no='" & Me.txtemp_no & "'"
'            addemp.execquery (StrDel_Sec)
'            strDel_Emp = "delete from employee where emp_no ='" & Me.txtemp_no & "'"
'            addemp.execquery (strDel_Emp)
'            For Each obj In Me
'                If TypeOf obj Is TextBox Then
'                    obj.Text = ""
'                End If
'            Next
'        Else
'            Exit Sub
'        End If
'        Me.CmbSex = "None"
 '       End Sub

Private Sub CmdEdit_Click()
    Call unlocked
    Me.cmdsave.Enabled = True
   ' Me.cmdnew.Enabled = False
    editV = 1
End Sub

Private Sub CmdFirst_Click()
    Me.CmdNext.Enabled = True
    Me.CmdLast.Enabled = True
    Me.CmdPrev.Enabled = False
        With rs_EmpDesig
           If Not .EOF Then
                .MoveFirst
                 Me.CmdFirst.Enabled = False
                 Call display
           End If
        End With
End Sub

Private Sub CmdLast_Click()
  Me.CmdNext.Enabled = False
    Me.CmdFirst.Enabled = True
    Me.CmdPrev.Enabled = True
        With rs_EmpDesig
            If Not .EOF Then
            .MoveLast
            Me.CmdLast.Enabled = False
            Call display
            End If
       End With
End Sub

Private Sub cmdnew_Click()
CHECK = 0
    Call clearbox
  Call unlocked
    Me.cmdsave.Enabled = True
    Set rs_max = openrec("select max(emp_no) as emp_no from employee")
    If rs_max.Fields(0) > 0 Then
        Me.txtemp_no = rs_max.Fields("emp_no") + 1
    Else
        Me.txtemp_no = 1
    End If
    Me.CmdFirst.Enabled = True
    Me.CmdNext.Enabled = True
    Me.CmdPrev.Enabled = True
    Me.CmdLast.Enabled = True
    Me.txtemp_name.SetFocus
End Sub

Private Sub CmdNext_Click()
     Me.CmdPrev.Enabled = True
     Me.CmdFirst.Enabled = True
        With rs_EmpDesig
          If Not .EOF Then
             .MoveNext
            If .EOF Then
               .MoveLast
                Me.CmdLast.Enabled = False
                Me.CmdNext.Enabled = False
           End If
           End If
                Call display
        End With
End Sub

Private Sub CmdPrev_Click()
    Me.CmdNext.Enabled = True
    Me.CmdLast.Enabled = True
        With rs_EmpDesig
            If Not .BOF Then
               .MovePrevious
            If .BOF = True Then
               .MoveFirst
                Me.CmdFirst.Enabled = False
                Me.CmdPrev.Enabled = False
            End If
            End If
                Call display
       End With
End Sub

Private Sub cmdsave_Click()
CHECK = 1
For Each obj In Me
 If TypeOf obj Is ComboBox Or TypeOf obj Is TextBox Then
   If obj.Text = "" Then MsgBox " give full information": Exit Sub
 End If
Next
Select Case editV
Case 0
    StrIns_Emp = "insert into employee values(" & Val(Me.txtemp_no) & ",'" & DesigIDV & "','" & Trim(Me.txtemp_name) & "','" & CDate(Trim(Me.txtdob)) & "','" & Trim(Me.CmbSex) & "','" & Trim(Me.txtPh_No) & "','" & Trim(Me.txtFC_No) & "','" & Trim(Me.txtadd) & "')"
    addemp.execquery (StrIns_Emp)
    rs_Emp.Fields.Refresh
Case 1
    StrIns_Emp = "Update employee set  desig_id='" & DesigIDV & "',emp_name= '" & Trim(Me.txtemp_name) & "',dob='" & CDate(Trim(Me.txtdob)) & "',sex='" & Trim(Me.CmbSex) & "',phone_no='" & Trim(Me.txtPh_No) & "',fc_no='" & Trim(Me.txtFC_No) & "',address='" & Trim(Me.txtadd) & "'where emp_no=" & Val(Me.txtemp_no) & ""
    addemp.execquery (StrIns_Emp)
    rs_Emp.Fields.Refresh
    EditcodeV = 0
End Select
Me.cmdsave.Enabled = False
Me.cmdnew.Enabled = True
Call locked

Me.CmdNext.Enabled = True
Me.CmdPrev.Enabled = True

End Sub

Private Sub CmdSearch_Click()

   Dim v1 As String
    Dim a As Integer
        v1 = InputBox("Please Enter Patient_ID")
        If v1 = "" Then Exit Sub
        If Not IsNumeric(v1) Then MsgBox "enter integer values": Exit Sub
             With rs_Emp
                    .MoveFirst
                    .Find ("Emp_No='" & v1 & "'")
                        If .EOF Then
                           
                            a = MsgBox("No Record Found", vbInformation)
                            .MoveFirst
                            Exit Sub
                        End If
                    Call display
             End With
End Sub

Private Sub Form_Activate()
'editV = 0
'    Agent1.Characters.Load "Merlin"
'    Agent1.Characters("Merlin").LanguageID = &H409 'English
'    Set merlin = Agent1.Characters("Merlin")
'    merlin.Show
'    merlin.Height = 150
'    merlin.Width = 150
'    merlin.Speak "Welcome To Employee Information"
End Sub

Private Sub Form_Load()
    Call OpenDB
    Set rs_Emp = openrec("select * from employee")
    Set rs_Desig = openrec("select * from designation ")
    Set rs_EmpDesig = openrec("select employee.*,designation.* from employee,designation where employee.desig_id = designation.desig_id ")
With rs_Desig
    If Not .EOF Or Not .BOF Then
        .MoveFirst
        While Not .EOF
             Me.CmbDesig.AddItem .Fields("description")
            .MoveNext
        Wend
     End If
       '.MoveFirst
End With
Call display
Call locked
Me.cmdsave.Enabled = False
End Sub
Public Sub display()
    With rs_EmpDesig
       If Not .BOF Or Not .EOF Then
            Me.txtemp_no = .Fields("emp_no")
            Me.txtemp_name = .Fields("emp_name")
            Me.txtdob = .Fields("dob")
            Me.CmbSex = .Fields("sex")
            Me.txtPh_No = .Fields("phone_no")
            Me.txtFC_No = .Fields("fc_no")
            Me.txtadd = .Fields("address")
            Me.CmbDesig = .Fields("description")
       End If
    End With
    editV = 0
    End Sub
Public Sub clearbox()
    Me.txtemp_name = ""
    Me.txtadd = ""
'    Me.txtdob = ""
    Me.txtFC_No = ""
    Me.txtPh_No = ""
End Sub
Public Sub locked()
    Me.txtadd.locked = True
    Me.txtemp_name.locked = True
    Me.txtFC_No.locked = True
    Me.txtPh_No.locked = True
    Me.CmbSex.locked = True
    Me.txtdob.Enabled = False
    Me.CmdFirst.Enabled = True
    Me.CmdNext.Enabled = True
    Me.CmdLast.Enabled = True
    Me.CmdPrev.Enabled = True
End Sub
Public Sub unlocked()
    Me.txtadd.locked = False
    Me.txtemp_name.locked = False
    Me.txtFC_No.locked = False
    Me.txtPh_No.locked = False
    Me.CmbSex.locked = False
    Me.txtdob.Enabled = True
 '   Me.CmdFirst.Enabled = False
 '   Me.CmdNext.Enabled = False
 '   Me.CmdLast.Enabled = False
 '   Me.CmdPrev.Enabled = False
End Sub

Private Sub txtemp_name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtPh_No.SetFocus
KeyAscii = Character(KeyAscii)
End Sub

Private Sub txtfc_no_KeyPress(KeyAscii As Integer)
KeyAscii = addemp.numeric(KeyAscii)
End Sub

Private Sub txtph_no_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtFC_No.SetFocus
KeyAscii = addemp.numeric(KeyAscii)
End Sub
