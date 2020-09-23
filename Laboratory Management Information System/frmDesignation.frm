VERSION 5.00
Begin VB.Form frmDesignation 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Designation"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Width           =   6495
      Begin VB.CommandButton CmdBack 
         Caption         =   "back"
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "edit"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdLast 
         Caption         =   ">"
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "<<"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdPrev 
         Caption         =   ">>"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdFirst 
         Caption         =   "<"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "save"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "new"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   5895
      Begin VB.TextBox txtDesigID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDesignation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation ID"
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
         Left            =   1185
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
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
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   1020
      End
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   8880
      X2              =   3000
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Designation"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   1320
      Width           =   4740
   End
End
Attribute VB_Name = "frmDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDesignation As New ADODB.Recordset
Dim RsGenId As New ADODB.Recordset
Dim Rs_check_Pk As New ADODB.Recordset
Dim EditcodeV As Integer
Dim obj As Object
Dim disp As New LabCLS
Dim CHECK As Integer


Private Sub Agent1_ActivateInput(ByVal CharacterID As String)

End Sub

Private Sub CmdBack_Click()
Unload Me
 frmMain.Show
End Sub

'Private Sub CmdDel_Click()
'EditcodeV = 0
'Dim a As String
'If Me.txtDesigID = "" Or Me.txtDesignation = "" Then MsgBox "can not delete": Exit Sub
'a = MsgBox("Are You Sure TO Delete The Record", vbYesNo)
'If a = vbYes Then
'disp.delete ("delete from designation where desig_id =" & Me.txtDesigID & "")
'Me.txtDesigID = ""
'Me.txtDesignation = ""
'End If

'End Sub

Private Sub CmdEdit_Click()
EditcodeV = 1
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
With RsDesignation
  If Not .EOF Then
    Me.txtDesigID = .Fields("desig_id")
    Me.txtDesignation = .Fields("description")
  End If
End With
EditcodeV = 0
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
Me.txtDesignation = ""
Set RsGenId = openrec("select * from designation ")
  With RsGenId
    If Not .EOF Then
      .MoveLast
      Me.txtDesigID = .Fields("desig_id") + 1
    Else
      Me.txtDesigID = 1
    End If
  End With
  EditcodeV = 0
  CHECK = 0
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
For Each obj In Me
  If TypeOf obj Is TextBox Then
    If obj.Text = "" Then MsgBox " give full informtion": Exit Sub
  End If
Next
If EditcodeV = 1 Then
disp.update ("update designation set description ='" & Me.txtDesignation & "' where desig_id =  " & Me.txtDesigID & "")
EditcodeV = 0
Exit Sub
End If


Set Rs_check_Pk = openrec("select * from designation where desig_id=" & Me.txtDesigID & "")
 With Rs_check_Pk
   If Not .EOF Then
      MsgBox " already exist", vbOKOnly + vbCritical, " Warning": Exit Sub
   End If
 End With
disp.add ("insert into designation values(" & Me.txtDesigID & ",' " & Me.txtDesignation & " ' )")

End Sub

Private Sub Form_Load()
CHECK = 2
Call OpenDB
Set RsDesignation = openrec("select * from designation")
EditcodeV = 0
End Sub
