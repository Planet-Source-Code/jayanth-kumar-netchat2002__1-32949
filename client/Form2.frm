VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log in"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Online Friends:"
      Height          =   6405
      Left            =   0
      TabIndex        =   1
      Top             =   90
      Width           =   3705
      Begin VB.PictureBox Picture1 
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
         Height          =   6015
         Left            =   120
         ScaleHeight     =   5955
         ScaleWidth      =   3390
         TabIndex        =   2
         Top             =   240
         Width           =   3455
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Signon Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   80
            TabIndex        =   7
            Top             =   3120
            Width           =   3255
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   9
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   8
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User Name "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   11
               Top             =   720
               Width           =   990
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Server Name "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   10
               Top             =   1080
               Width           =   1155
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   295
            Left            =   2350
            TabIndex        =   6
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Login"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   295
            Left            =   1320
            TabIndex        =   5
            Top             =   4920
            Width           =   975
         End
         Begin VB.Timer Timer1 
            Interval        =   10
            Left            =   2880
            Top             =   1800
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   165
            Left            =   165
            TabIndex        =   12
            Top             =   5520
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   291
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Establishing Connection...."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   165
            TabIndex        =   13
            Top             =   5280
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Version 1.0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   3
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   2760
            Picture         =   "Form2.frx":0442
            Top             =   120
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   2595
            Left            =   120
            Picture         =   "Form2.frx":0884
            Top             =   120
            Width           =   3045
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   -360
      TabIndex        =   0
      Top             =   -90
      Width           =   14655
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   6495
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu log 
      Caption         =   "Login"
      Begin VB.Menu login 
         Caption         =   "&Login"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PCName As String
Dim P As Long

Private Sub close_Click()
End
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
End
End If
StatusBar1.SimpleText = "Connecting..."
user = Text1.Text
server = Text2.Text
ProgressBar1.Value = 0
Label3.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
StatusBar1.SimpleText = "Disconnected..."
Timer1.Enabled = False
ProgressBar1.Value = 0
P = NameOfPC(PCName)
Text1.Text = PCName
End Sub

Private Sub login_Click()
Command2_Click
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value >= 100 Then
   Timer1.Enabled = False
   Unload Me
   Form1.Show
End If

End Sub
