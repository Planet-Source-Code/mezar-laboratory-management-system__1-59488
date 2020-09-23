VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmoption 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmoption"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   10695
      Begin VB.CommandButton CmdBack 
         DisabledPicture =   "Form1.frx":0000
         DownPicture     =   "Form1.frx":03B9
         Height          =   375
         Left            =   9000
         Picture         =   "Form1.frx":08FB
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2700
         Left            =   2640
         TabIndex        =   35
         Top             =   3360
         Width           =   7740
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   6000
            TabIndex        =   43
            Top             =   720
            Width           =   1695
            Begin MSComCtl2.DTPicker samplingtime 
               Height          =   285
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarForeColor=   8388608
               CustomFormat    =   "hh:mm:ss"
               Format          =   19202050
               CurrentDate     =   0.597881944444444
            End
         End
         Begin VB.ComboBox cmballotedto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Form1.frx":0D83
            Left            =   1320
            List            =   "Form1.frx":0D85
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtTest 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox txttestcategory 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1920
            Width           =   2775
         End
         Begin VB.TextBox txtlabno 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   360
            Width           =   1695
         End
         Begin VB.Timer Timer2 
            Interval        =   60
            Left            =   0
            Top             =   2640
         End
         Begin VB.TextBox txtRefered_By 
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
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   37
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtCharges 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   " "
            Top             =   2280
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker samplingdate 
            Height          =   285
            Left            =   6000
            TabIndex        =   45
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   54263811
            CurrentDate     =   37146
         End
         Begin MSComCtl2.DTPicker DueDate 
            Height          =   270
            Left            =   6000
            TabIndex        =   46
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   476
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   54263811
            CurrentDate     =   37403
         End
         Begin MSComCtl2.DTPicker DueTime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "h:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   47
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   8388608
            CustomFormat    =   "hh:mm:ss AMPM"
            Format          =   54263810
            CurrentDate     =   37260.0854166667
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alloted To"
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
            TabIndex        =   57
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label Label37 
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
            TabIndex        =   56
            Top             =   2280
            Width           =   420
         End
         Begin VB.Label Label36 
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
            Left            =   3240
            TabIndex        =   55
            Top             =   360
            Width           =   2700
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Refered By"
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
            TabIndex        =   54
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label34 
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
            Left            =   120
            TabIndex        =   53
            Top             =   1920
            Width           =   1365
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lab No"
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
            TabIndex        =   52
            Top             =   360
            Width           =   660
         End
         Begin VB.Label Label32 
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
            Left            =   3240
            TabIndex        =   51
            Top             =   720
            Width           =   2370
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due Date(DD/MM/YYYY)"
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
            Left            =   3240
            TabIndex        =   50
            Top             =   1080
            Width           =   2355
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due Time(HH:MM:SS)"
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
            Left            =   3240
            TabIndex        =   49
            Top             =   1440
            Width           =   2025
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Charges"
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
            Left            =   4440
            TabIndex        =   48
            Top             =   2280
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdnew 
         DisabledPicture =   "Form1.frx":0D87
         DownPicture     =   "Form1.frx":114E
         Height          =   375
         Left            =   9000
         Picture         =   "Form1.frx":162E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdsave 
         DisabledPicture =   "Form1.frx":1A7C
         DownPicture     =   "Form1.frx":1EA9
         Enabled         =   0   'False
         Height          =   375
         Left            =   9000
         Picture         =   "Form1.frx":23FC
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdnexttest 
         DisabledPicture =   "Form1.frx":28C2
         DownPicture     =   "Form1.frx":2D0A
         Height          =   375
         Left            =   9000
         Picture         =   "Form1.frx":32C0
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdclear 
         DisabledPicture =   "Form1.frx":3803
         DownPicture     =   "Form1.frx":3B9C
         Height          =   375
         Left            =   9000
         Picture         =   "Form1.frx":409D
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton CmdPrint 
         DisabledPicture =   "Form1.frx":4514
         DownPicture     =   "Form1.frx":4955
         Enabled         =   0   'False
         Height          =   375
         Left            =   9000
         Picture         =   "Form1.frx":4EC5
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   6300
         Begin VB.TextBox txtptid 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Top             =   360
            Width           =   1110
         End
         Begin VB.TextBox txtptname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   2070
         End
         Begin VB.TextBox txtPtFname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1128
            Width           =   2055
         End
         Begin VB.TextBox txtage 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1560
            Width           =   630
         End
         Begin VB.TextBox txtaddress 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   645
            Left            =   1440
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   2280
            Width           =   4695
         End
         Begin VB.TextBox txtmartialstatus 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1080
            Width           =   1110
         End
         Begin VB.TextBox txtptyear 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5520
            MaxLength       =   4
            TabIndex        =   12
            Top             =   360
            Width           =   630
         End
         Begin VB.TextBox txtsex 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtdob 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1512
            Width           =   1695
         End
         Begin VB.TextBox txtphone 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1896
            Width           =   1695
         End
         Begin VB.TextBox txtFirstContactno 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtsurname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth"
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
            TabIndex        =   30
            Top             =   1515
            Width           =   1230
         End
         Begin VB.Label Label29 
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
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   120
            TabIndex        =   28
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
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
            Left            =   3240
            TabIndex        =   27
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Left            =   120
            TabIndex        =   26
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sex"
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
            Left            =   4800
            TabIndex        =   25
            Top             =   1560
            Width           =   360
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1125
            Width           =   1260
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marital Status"
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
            TabIndex        =   23
            Top             =   1080
            Width           =   1350
         End
         Begin VB.Label Label26 
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
            Left            =   5040
            TabIndex        =   22
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone #"
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
            TabIndex        =   21
            Top             =   1890
            Width           =   795
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Contect #"
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
            Left            =   3240
            TabIndex        =   20
            Top             =   1920
            Width           =   1440
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sur Name"
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
            TabIndex        =   19
            Top             =   720
            Width           =   945
         End
      End
      Begin MSComctlLib.TreeView TreeView 
         Height          =   5775
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   10186
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         HotTracking     =   -1  'True
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13123
            MinWidth        =   13123
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6855
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12091
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   840
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   10080
      X2              =   960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lab Registration"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3600
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   4080
      Top             =   7560
   End
End
Attribute VB_Name = "frmoption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim disp                    As New LabCLS
Dim Rs_TestCategory         As New ADODB.Recordset
Dim Rstest                  As New ADODB.Recordset
Dim rsfind                  As New ADODB.Recordset
Dim RsPatient               As New ADODB.Recordset
Dim Rs_CheckPatientTest     As New ADODB.Recordset
Dim Rs_PatientInfo          As New ADODB.Recordset
Dim Rs_PatMaxLabNo          As New ADODB.Recordset
Dim Rs_Security             As New ADODB.Recordset
Dim Code                    As String
Dim i                       As Integer
Dim obj                     As Object
Dim Pkey                    As String
Dim TestCodeV               As String
Dim SubTestcodeV            As String
Dim SampleTimeCodeV         As String
Dim DueTimeCodeV            As String
Dim SampleDateCodeV         As Date
Dim SecCodeV                As String
Dim DueDateCodev            As Date
Dim Counter As Integer
Dim CHECK As Integer
Private Sub cmballotedto_Click()
    With Rs_Security
       If Not .EOF Then
         .Find ("login_Id='" & Me.cmballotedto & "'")
         SecCodeV = .Fields("Emp_No")
       End If
    End With
End Sub

Private Sub cmballotedto_GotFocus()
    Me.Text2.Visible = False
End Sub

Private Sub cmballotedto_KeyPress(KeyAscii As Integer)
    With Rs_Security
       If Not .EOF Then
         .Find ("login_Id='" & Me.cmballotedto & "'")
         If Not .EOF Then
         SecCodeV = .Fields("Emp_No")
         End If
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
frmMain.Show
'End If
'End If
'CHECK = 0
End Sub

Private Sub cmdclear_Click()
    For Each obj In Me
         If TypeOf obj Is TextBox Then
             obj.Text = ""
         End If
    Next
    Me.Text2.Visible = True
    Me.txtptid.Enabled = False
    Me.txtptyear.Enabled = False
    
    
End Sub
Private Sub cmdnew_Click()
Me.CmdPrint.Enabled = False
Me.txtptid.Enabled = True
Me.txtptyear.Enabled = True
CHECK = 0
    For Each obj In Me
        If TypeOf obj Is TextBox And obj.Name <> "text2" Then
            'And obj.Name <> "text1"
            obj.Text = ""
        End If
    Next
    Me.txtptid.Enabled = True
    Me.txtptyear.Enabled = True
    Me.cmdsave.Enabled = True
    Me.TreeView.Enabled = True
    Me.TabStrip1.Enabled = True
    'Set Rs_PatMaxLabNo = disp.openrec("select max(labno) from pt_test")
    Set Rs_PatMaxLabNo = disp.openrec("select * from pt_test")
        With Rs_PatMaxLabNo
           '   If .Fields(0) >= 0 Then
            If Not .EOF Then
            .MoveLast
             Me.txtlabno.Text = .Fields(0) + 1
            Else
              Me.txtlabno = 1
            End If
        End With
    Me.txtptid.SetFocus
    Call TabStrip1_Click
    Me.cmdsave.Enabled = True
    Me.txtptid.Enabled = True
    Me.txtptyear.Enabled = True
End Sub
Private Sub cmdnew_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtptid.SetFocus
End Sub
Private Sub cmdnew_LostFocus()
    Me.StatusBar1.Panels(1) = ""
End Sub
Private Sub cmdnexttest_Click()
    Me.txtTest.Text = ""
    Me.cmdsave.Enabled = True
End Sub

Private Sub CmdPrint_Click()
If Counter = 0 Then
DE1.labreg_Grouping Val(Me.txtptid), Val(Me.txtptyear), Val(Me.txtlabno)
rptlabreg.Show
Else
DE1.rslabreg_Grouping.Close
DE1.labreg_Grouping Val(Me.txtptid), Val(Me.txtptyear), Val(Me.txtlabno)
rptlabreg.Show
End If
Counter = Counter + 1

End Sub

Private Sub cmdsave_Click()
CHECK = 1
    Me.Text2.Text = " "
        For Each obj In Me
            If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
                If obj.Text = "" Then
                    MsgBox "give full information"
                    Exit Sub
                End If
            End If
        Next
    Set Rs_CheckPatientTest = disp.openrec("select * from pt_test where pt_id=" & Val(Trim(Me.txtptid.Text)) & " and pt_year=" & Val(Trim(Me.txtptyear.Text)) & " and labno=" & Val(Trim(Me.txtlabno.Text)) & "  and test_catgcode='" & Me.TabStrip1.SelectedItem.Tag & "' and test_code='" & Me.TreeView.SelectedItem.Parent.Key & "' and subtest_code='" & Me.TreeView.SelectedItem.Key & "' and status ='N'")
         With Rs_CheckPatientTest
           If Not .EOF Or .RecordCount > 0 Then
             MsgBox "Already Exist"
             Exit Sub
           End If
         End With
'*********************************************
    If Me.samplingdate.Value > Me.DueDate.Value Then
       MsgBox "Due date must be greator than sample date", vbInformation
       Exit Sub
    End If
'**********************************************

    Dim f As String
        If Code = "xyzC" Then
           SubTestcodeV = Mid(SubTestcodeV, 1, InStr(1, SubTestcodeV, "xyzC") - 1)
        Else
           'TestCodeV = Mid(TestCodeV, 1, InStr(1, TestCodeV, "P") - 1)
           MsgBox " select subtest"
           Exit Sub
        End If
    TestCodeV = Mid(Me.TreeView.SelectedItem.Parent.Key, 1, InStr(1, Me.TreeView.SelectedItem.Parent.Key, "p") - 1)
    SampleTimeCodeV = Format(Me.samplingtime.Value, "hh:MM:ss AMPM")
    DueTimeCodeV = Format(Me.DueTime.Value, "hh:MM:ss AMPM")
    SampleDateCodeV = Format(Me.samplingdate.Value, "dd/MM/yyyy")
    DueDateCodev = Format(Me.DueDate.Value, "dd/MM/yyyy")
    
    disp.add ("insert into pt_test(pt_id,labno,pt_year,test_catgcode,test_code,subtest_code,sampling_date,sampling_time,due_date,due_time,Refered_by,alloted_to,status,charges) values( '" & Me.txtptid.Text & "'," & Me.txtlabno.Text & "," & Me.txtptyear.Text & ", '" & Me.TabStrip1.SelectedItem.Tag & "','" & TestCodeV & "',  '" & SubTestcodeV & " ','" & SampleDateCodeV & "','" & Trim(SampleTimeCodeV) & "','" & DueDateCodev & "','" & DueTimeCodeV & "','" & Trim(Me.txtRefered_By) & "','" & Me.cmballotedto & "','N'," & Me.txtCharges & ")")
    
    Me.cmdsave.Enabled = False
    Me.txtptid.Enabled = False
    Me.txtptyear.Enabled = False
    Me.CmdPrint.Enabled = True
End Sub
Private Sub Command1_Click()
'    frmtestresult.Show
End Sub
Private Sub Form_Load()
    Set Rs_TestCategory = disp.openrec("select * from test_catg")
    Set Rs_Security = disp.openrec("select * from security")
    i = 1
    With Rs_TestCategory
       If Not .EOF Then
              .MoveFirst
            While Not .EOF
              Me.TabStrip1.Tabs.add i, , .Fields("description")
              Me.TabStrip1.Tabs(i).Caption = .Fields("description")
              Me.TabStrip1.Tabs(i).Tag = .Fields(0)
              i = i + 1
             .MoveNext
            Wend
        End If
    End With
  
     With Rs_Security
        If Not .EOF Then
              i = 0
              While Not .EOF
                Me.cmballotedto.AddItem .Fields("login_id")
               .MoveNext
                i = i + 1
              Wend
            .MoveFirst
        End If
    End With
    Me.samplingdate.Value = Date
    Me.StatusBar1.Panels(1) = "Click new button"
Counter = 0

End Sub

Private Sub TabStrip1_Click()
    Set Rstest = disp.openrec("select * from test where test_catgcode= '" & Me.TabStrip1.SelectedItem.Tag & "'")
    Me.TreeView.Nodes.Clear
        With Rstest
             If Not .EOF Then
                    .MoveFirst
                While Not .EOF
                    Me.TreeView.Nodes.add , , .Fields("test_code") & "p", .Fields("description")
                    Pkey = .Fields("test_code")
                    TestCodeV = .Fields("test_code")
                    
                    Set rsfind = disp.openrec("select * from subtest where test_code='" & TestCodeV & "'")  ' and test_catgcode='" & Me.TabStrip1.SelectedItem.Tag & "'")
                    
            '        Set rsCharges = disp.openrec("Select charges from subtest where SubTest_Code = '" & SubTestcodeV & "'")
                       With rsfind
                            If Not .EOF Then
                               While Not .EOF
                                     Me.TreeView.Nodes.add (Pkey & "p"), tvwChild, .Fields("subtest_code") & "xyzC", .Fields("description")
                                     SubTestcodeV = .Fields("subtest_code")
                                     .MoveNext
                               Wend
                            End If
                       End With
                     .MoveNext
                 Wend
             End If
        End With
    Me.txttestcategory.Text = Me.TabStrip1.SelectedItem.Caption
End Sub
Private Sub TabStrip1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then Me.TreeView.SetFocus
End Sub



Private Sub Timer2_Timer()
    Me.samplingtime.Value = Time
End Sub
Private Sub TreeView_Click()

    SubTestcodeV = Me.TreeView.SelectedItem.Key
    Code = Right(SubTestcodeV, 7)
        If Code = "xyzC" Then
           Me.txtTest.Text = Me.TreeView.SelectedItem.Text
           'Me.txtTotCharges = rsCharges.Fields("Charges")
        Else
           Me.txtTest.Text = ""
           Me.txtCharges = ""
        End If
        '*********************************
Dim RS_SUBTEST1 As New ADODB.Recordset
  Set RS_SUBTEST1 = disp.openrec("select * from subtest")
   With RS_SUBTEST1
     If Not .EOF Then
      .MoveFirst
       .Find ("description ='" & Me.txtTest.Text & "'")
       If Not .EOF Then
        Me.txtCharges = .Fields("charges")
      End If
      End If
   End With
'*********************************

    Me.txtRefered_By.SetFocus
End Sub
Private Sub TreeView_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtTest.Text = Me.TreeView.SelectedItem.Text
        Me.cmdsave.SetFocus
    End If
End Sub
Private Sub txtPtid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtptyear.SetFocus
     If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 8) Then
        KeyAscii = 0
    Else
        Exit Sub
        End If
End Sub
Private Sub txtptyear_KeyPress(KeyAscii As Integer)
    Dim pt_idV As String
     
      
    
    
    If KeyAscii = 13 Then
     
      If Me.txtptid.Text = "" Or Me.txtptyear.Text = "" Then
      MsgBox "enter values "
      Me.txtptid.SetFocus
      Exit Sub
      End If
        Set RsPatient = disp.openrec("select patient.pt_name,patient.pt_surname," & _
        " patient.pt_fname, patient.sex,patient.dob,pt_info.phone_no," & _
        "pt_info.fc_no,pt_info.age,pt_info.address,pt_info.m_status,pt_info.occupation " & _
        "  from  patient,pt_info    where patient.pt_id = " & Me.txtptid.Text & " " & _
        "and patient.pt_year = " & Me.txtptyear.Text & " and " & _
        "patient.pt_id=pt_info.pt_id and patient.pt_year=pt_info.pt_year")
        With RsPatient
              If Not .EOF Then
                Me.txtphone.Text = .Fields("phone_no")
                Me.txtmartialstatus.Text = .Fields("m_status")
                Me.txtFirstContactno.Text = .Fields("fc_no")
                Me.txtaddress.Text = .Fields("address")
                Me.txtage.Text = .Fields("age")
            '    pt_idV = .Fields("pt_id")
                Me.txtptname.Text = .Fields("pt_name")
                Me.txtsurname.Text = .Fields("pt_surname")
                Me.txtPtFname.Text = .Fields("pt_fname")
                Me.txtdob.Text = .Fields("dob")
                Me.txtsex.Text = .Fields("sex")
              Else
               MsgBox "No Record Found"
               Exit Sub
             End If
        End With
      Me.txtRefered_By.SetFocus
    End If
      If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 8) Then
         KeyAscii = 0
      Else
         Exit Sub
      End If
     KeyAscii = numeric(KeyAscii)

End Sub

Private Sub txtRefered_By_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.cmballotedto.SetFocus
KeyAscii = Character(KeyAscii)
End Sub
