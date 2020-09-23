VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_patient 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Script"
      Size            =   24
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   840
      TabIndex        =   24
      Top             =   1440
      Width           =   9135
      Begin MSComCtl2.DTPicker txtDOB 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
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
         CalendarBackColor=   16777215
         Format          =   54263809
         CurrentDate     =   37358
         MinDate         =   183
      End
      Begin VB.TextBox txtAge 
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
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
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
         ItemData        =   "frm_patient.frx":0000
         Left            =   6240
         List            =   "frm_patient.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3240
         TabIndex        =   23
         ToolTipText     =   "Search The Patient"
         Top             =   255
         Width           =   495
      End
      Begin VB.TextBox txtPt_ID 
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
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPtName 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtFName 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtSurName 
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
         Left            =   6240
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtPh_No 
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
         Left            =   6240
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtFC_No 
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
         Left            =   6240
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtOccupation 
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
         Left            =   1560
         TabIndex        =   10
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   1125
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2400
         Width           =   7455
      End
      Begin VB.TextBox txtPt_Year 
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
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CmbM_Status 
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
         ItemData        =   "frm_patient.frx":0025
         Left            =   1560
         List            =   "frm_patient.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4560
         TabIndex        =   38
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   37
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   36
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   35
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sur Name"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4560
         TabIndex        =   34
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   33
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4560
         TabIndex        =   32
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Contact No"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4560
         TabIndex        =   31
         Top             =   1680
         Width           =   1635
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   30
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   29
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4560
         TabIndex        =   28
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Year"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4560
         TabIndex        =   27
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Martial Status"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   26
         Top             =   1680
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   1800
      TabIndex        =   22
      Top             =   6120
      Width           =   8535
      Begin VB.CommandButton CmdLast 
         DisabledPicture =   "frm_patient.frx":0064
         DownPicture     =   "frm_patient.frx":0418
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6000
         Picture         =   "frm_patient.frx":0909
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CmdNext 
         DisabledPicture =   "frm_patient.frx":0D51
         DownPicture     =   "frm_patient.frx":1116
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         Picture         =   "frm_patient.frx":15CE
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CmdPrev 
         DisabledPicture =   "frm_patient.frx":1A67
         DownPicture     =   "frm_patient.frx":1E41
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         Picture         =   "frm_patient.frx":239C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CmdFirst 
         DisabledPicture =   "frm_patient.frx":2860
         DownPicture     =   "frm_patient.frx":2C1D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   960
         Picture         =   "frm_patient.frx":316B
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CmdBack 
         DisabledPicture =   "frm_patient.frx":366D
         DownPicture     =   "frm_patient.frx":3A26
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6840
         Picture         =   "frm_patient.frx":3F68
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton CmdDel 
         DisabledPicture =   "frm_patient.frx":43F0
         DownPicture     =   "frm_patient.frx":4808
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5160
         Picture         =   "frm_patient.frx":4DAF
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton CmdEdit 
         DisabledPicture =   "frm_patient.frx":52D9
         DownPicture     =   "frm_patient.frx":56D1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3480
         Picture         =   "frm_patient.frx":5C20
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton CmdSave 
         DisabledPicture =   "frm_patient.frx":60C1
         DownPicture     =   "frm_patient.frx":64EE
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         Picture         =   "frm_patient.frx":6A41
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton CmdNew 
         DisabledPicture =   "frm_patient.frx":6F07
         DownPicture     =   "frm_patient.frx":72CE
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Picture         =   "frm_patient.frx":77AE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   10560
      X2              =   1440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Information"
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
      Left            =   1560
      TabIndex        =   20
      Top             =   720
      Width           =   4305
   End
End
Attribute VB_Name = "frm_patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPatient                    As New ADODB.Recordset
Dim rsfind                       As New ADODB.Recordset

Dim ClPt                         As New LabCLS
Dim strIns_Patient               As String
Dim strIns_Ptinfo                As String
Dim StrDel_patient               As String
Dim StrDel_ptinfo                As String
Dim StrDel_ptTest                As String
Dim obj                          As Object
Dim editV                        As Integer
Dim pk As New ADODB.Recordset
Dim CHECK As Integer

Private Sub CmbM_Status_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtOccupation.SetFocus
End Sub

Private Sub CmbSex_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtPh_No.SetFocus
End Sub

Private Sub CmdBack_Click()
'Dim a As String
'If CHECK = 0 Then
'a = MsgBox("do you want to save the information", vbYesNoCancel)
'If a = vbYes Then
'Call cmdsave_Click
'ElseIf a = vbNo Or CHECK = 2 Then
 Unload Me
 frmMain.Show
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
  
  
Dim a As Integer
    a = MsgBox("Are you Sure", vbYesNoCancel)
        If a = vbYes Then
         
            StrDel_ptTest = "delete from pt_test where pt_id='" & Me.txtPt_ID & "'"
            ClPt.execquery (StrDel_ptTest)
            StrDel_ptinfo = "delete  pt_info where pt_id=" & Me.txtPt_ID & ""
            ClPt.execquery (StrDel_ptinfo)
            StrDel_patient = "delete patient where pt_id=" & Me.txtPt_ID & ""
            ClPt.execquery (StrDel_patient)
            RsPatient.Fields.Refresh
            For Each obj In Me
                If TypeOf obj Is TextBox Then
                    obj.Text = ""
                End If
            Next
        Else
            Exit Sub
        End If
        Me.CmbM_Status = "None"
        Me.CmbSex = "None"
End Sub

Private Sub CmdEdit_Click()
    editV = 1
  '  Me.cmdnew.Enabled = False
    Me.cmdsave.Enabled = True
    Call unlocked
    Me.txtPt_ID.locked = True
    Me.CmdSearch.Enabled = False
End Sub

Private Sub CmdFirst_Click()
    Me.CmdNext.Enabled = True
    Me.CmdLast.Enabled = True
    Me.CmdPrev.Enabled = False
        With RsPatient
                .MoveFirst
                 Me.CmdFirst.Enabled = False
                 Call display
        End With
End Sub
Private Sub CmdLast_Click()
    Me.CmdNext.Enabled = False
    Me.CmdFirst.Enabled = True
    Me.CmdPrev.Enabled = True
        With RsPatient
            .MoveLast
            Me.CmdLast.Enabled = False
            Call display
       End With
End Sub
Private Sub cmdnew_Click()
CHECK = 0
    Call unlocked
    Me.txtPt_ID.locked = True
    Me.cmdsave.Enabled = True
    Me.CmdSearch.Enabled = False
      Set rsfind = openrec("select max(pt_id) as pt_id from patient")
        If rsfind.Fields(0) > 0 Then
           Me.txtPt_ID = rsfind.Fields(0) + 1
        Else
           Me.txtPt_ID = 1
        End If
With Me
    Me.txtptname = ""
    Me.txtsurname = ""
    Me.txtFname = ""
    'Me.txtDOB = "None"
    Me.CmbSex = "None"
    Me.txtPt_Year = ""
    Me.CmbM_Status = "None"
    Me.txtaddress = ""
    Me.txtFC_No = ""
    Me.txtOccupation = ""
    Me.txtPh_No = ""
    Me.txtage = ""
End With
Me.cmdsave.Enabled = True
editV = 0
Me.txtPt_Year.SetFocus
End Sub
Private Sub CmdNext_Click()
     Me.CmdPrev.Enabled = True
     Me.CmdFirst.Enabled = True
        With RsPatient
                .MoveNext
            If .EOF Then
               .MoveLast
                Me.CmdLast.Enabled = False
                Me.CmdNext.Enabled = False
           End If
                Call display
        End With
End Sub
Private Sub CmdPrev_Click()
    Me.CmdNext.Enabled = True
    Me.CmdLast.Enabled = True
        With RsPatient
               .MovePrevious
            If .BOF = True Then
               .MoveFirst
                Me.CmdFirst.Enabled = False
                Me.CmdPrev.Enabled = False
            End If
                Call display
       End With
End Sub

Private Sub cmdsave_Click()
CHECK = 1
For Each obj In Me
  If TypeOf obj Is ComboBox Or TypeOf obj Is TextBox Then
    If obj.Text = "" Then MsgBox "give full information": Exit Sub
  End If
Next


Select Case editV
    Case 1
        strIns_Patient = "update patient set pt_name='" & Trim(Me.txtptname) & "',pt_surname='" & Trim(Me.txtsurname) & "',pt_fname='" & Trim(Me.txtFname) & "',sex='" & Trim(Me.CmbSex) & "',DOB='" & Trim(Me.txtdob) & "'where pt_id=" & Val(Me.txtPt_ID) & ""
        ClPt.execquery (strIns_Patient)
        strIns_Ptinfo = "update  pt_info set pt_Year=" & Val(Me.txtPt_Year) & ",phone_No='" & Trim(Me.txtPh_No) & "',FC_No='" & Trim(Me.txtFC_No) & "',Address='" & Trim(Me.txtaddress) & "',Occupation='" & Trim(Me.txtOccupation) & "',m_status='" & Trim(Me.CmbM_Status) & "',age='" & Val(Me.txtage) & "' where pt_id=" & Val(Me.txtPt_ID) & ""
        ClPt.execquery (strIns_Ptinfo)
        RsPatient.Fields.Refresh
        editV = 0
    Case 0
Set pk = openrec("select * from patient where pt_id=" & Me.txtPt_ID & "")
 With pk
   If Not .EOF Then
     MsgBox " Already exist"
     Exit Sub
   End If
 End With
 
        strIns_Patient = "insert into patient values(" & Val(Me.txtPt_ID) & "," & Val(Me.txtPt_Year) & ",'" & Trim(Me.txtptname) & "','" & Trim(Me.txtsurname) & "','" & Trim(Me.txtFname) & "','" & Trim(Me.CmbSex) & "','" & Trim(Me.txtdob) & "')"
        ClPt.execquery (strIns_Patient)
        strIns_Ptinfo = "insert into  pt_info values(" & Val(Me.txtPt_ID) & "," & Val(Me.txtPt_Year) & ",'" & Trim(Me.txtPh_No) & "','" & Trim(Me.txtFC_No) & "','" & Trim(Me.txtaddress) & "','" & Trim(Me.txtOccupation) & "','" & Trim(Me.CmbM_Status) & "'," & Me.txtage & ")"
        ClPt.execquery (strIns_Ptinfo)
        RsPatient.Fields.Refresh
End Select
Me.CmdSearch.Enabled = True
Me.cmdsave.Enabled = False
Me.cmdnew.Enabled = True
Call locked

End Sub
Private Sub CmdSearch_Click()

    Dim v1 As String
    Dim a As Integer
        v1 = InputBox("Please Enter Patient_ID")
        If v1 = "" Then Exit Sub
        If Not IsNumeric(v1) Then MsgBox "give an integer value": Exit Sub
             With RsPatient
                    .MoveFirst
                    .Find ("Pt_ID='" & v1 & "'")
                        If .EOF Then
                           
                            'a = MsgBox("No Record Found", vbInformation)
                            .MoveFirst
                            Exit Sub
                        End If
                    Call display
             End With
End Sub
Private Sub Form_Load()
CHECK = 2
    Call OpenDB
    Set RsPatient = openrec("Select patient.*,pt_info.*  from Patient,pt_info where patient.pt_id = pt_info.pt_id")
    Call display
    Call locked
    Me.cmdsave.Enabled = False
End Sub
Public Sub display()
    With RsPatient
           
            If Not .EOF Or Not .BOF Then
                Me.txtPt_ID = .Fields("pt_id")
                Me.txtptname = .Fields("pt_name")
                Me.txtsurname = .Fields("pt_surname")
                Me.txtFname = .Fields("pt_Fname")
                Me.txtdob = .Fields("DOB")
                Me.CmbSex = .Fields("sex")
                Me.txtPt_Year = .Fields("Pt_Year")
                Me.CmbM_Status = .Fields("M_Status")
                Me.txtaddress = .Fields("address")
                Me.txtFC_No = .Fields("FC_NO")
                Me.txtage = .Fields("age")
                Me.txtOccupation = .Fields("Occupation")
                Me.txtPh_No = .Fields("phone_No")
            End If
    End With
    editV = 0
    Me.cmdsave.Enabled = False
End Sub
Public Sub locked()
    For Each obj In Me
        If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
            obj.locked = True
        End If
    Next
   ' Me.CmdFirst.Enabled = True
   ' Me.CmdNext.Enabled = True
   ' Me.CmdLast.Enabled = True
   ' Me.CmdPrev.Enabled = True
End Sub
Public Sub unlocked()
    For Each obj In Me
        If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
            obj.locked = False
        End If
    Next
  '  Me.CmdFirst.Enabled = False
  '  Me.CmdNext.Enabled = False
  '  Me.CmdLast.Enabled = False
  '  Me.CmdPrev.Enabled = False
End Sub
Private Sub txtAge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtaddress.SetFocus
KeyAscii = ClPt.numeric(KeyAscii)
End Sub
Private Sub txtfc_no_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtage.SetFocus
KeyAscii = ClPt.numeric(KeyAscii)
End Sub

Private Sub txtFName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.CmbM_Status.SetFocus
KeyAscii = Character(KeyAscii)
End Sub

Private Sub txtOccupation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.CmbSex.SetFocus
KeyAscii = Character(KeyAscii)
End Sub

Private Sub txtph_no_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtFC_No.SetFocus
KeyAscii = ClPt.numeric(KeyAscii)
End Sub

Private Sub txtPt_ID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtPt_Year.SetFocus
End Sub

Private Sub txtPt_Year_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtptname.SetFocus
KeyAscii = ClPt.numeric(KeyAscii)
End Sub

Private Sub txtPtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtsurname.SetFocus
KeyAscii = Character(KeyAscii)
End Sub

Private Sub txtSurName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.txtFname.SetFocus
KeyAscii = Character(KeyAscii)
End Sub
