VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Network Chat  <Server>"
   ClientHeight    =   4770
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8280
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2055
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3625
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0442
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
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   8400
      Top             =   4080
   End
   Begin VB.Frame Frame1 
      Height          =   35
      Left            =   -240
      TabIndex        =   8
      Top             =   0
      Width           =   16215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S E N D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   6360
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock TCPServer 
      Index           =   0
      Left            =   7200
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   3360
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0533
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   1935
      Left            =   720
      TabIndex        =   11
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3413
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0622
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
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   27
         Left            =   2280
         Picture         =   "Form1.frx":0713
         Tag             =   ":king"
         Top             =   0
         Width           =   345
      End
      Begin VB.Label Label54 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":toiletclaw"
         Height          =   255
         Left            =   4080
         TabIndex        =   60
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label53 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":guns"
         Height          =   255
         Left            =   2880
         TabIndex        =   59
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label52 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":guitar"
         Height          =   255
         Left            =   2760
         TabIndex        =   58
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label51 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":light"
         Height          =   255
         Left            =   2640
         TabIndex        =   57
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":ass"
         Height          =   255
         Left            =   2640
         TabIndex        =   56
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":wave"
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":bala"
         Height          =   255
         Left            =   2640
         TabIndex        =   54
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":king"
         Height          =   255
         Left            =   2640
         TabIndex        =   53
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":heart"
         Height          =   255
         Left            =   1440
         TabIndex        =   52
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":tilt"
         Height          =   255
         Left            =   1440
         TabIndex        =   51
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":alien"
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":devil"
         Height          =   255
         Left            =   1440
         TabIndex        =   49
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":angel"
         Height          =   255
         Left            =   4920
         TabIndex        =   48
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":erm"
         Height          =   255
         Left            =   1440
         TabIndex        =   47
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":flush"
         Height          =   255
         Left            =   2760
         TabIndex        =   46
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":lol"
         Height          =   255
         Left            =   1440
         TabIndex        =   45
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":mad"
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":gg"
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":shit"
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   2280
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   345
         Index           =   13
         Left            =   4200
         Picture         =   "Form1.frx":0AC7
         Tag             =   ":shoot"
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   16
         Left            =   1080
         Picture         =   "Form1.frx":0FD4
         Tag             =   ":alien"
         Top             =   0
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   315
         Index           =   37
         Left            =   4200
         Picture         =   "Form1.frx":137F
         Tag             =   ":angel"
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   24
         Left            =   1080
         Picture         =   "Form1.frx":17D1
         Tag             =   ":shit"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   22
         Left            =   1080
         Picture         =   "Form1.frx":1B47
         Tag             =   ":mad"
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   21
         Left            =   1080
         Picture         =   "Form1.frx":1ECD
         Tag             =   ":lol"
         Top             =   1560
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   36
         Left            =   3360
         Picture         =   "Form1.frx":2251
         Tag             =   ":blah"
         Top             =   480
         Width           =   660
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   25
         Left            =   3360
         Picture         =   "Form1.frx":261F
         Tag             =   ":beat"
         Top             =   840
         Width           =   525
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   30
         Left            =   2280
         Picture         =   "Form1.frx":2A21
         Tag             =   ":light"
         Top             =   840
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   34
         Left            =   2280
         Picture         =   "Form1.frx":2DBA
         Tag             =   ":guns"
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   23
         Left            =   1080
         Picture         =   "Form1.frx":31C5
         Tag             =   ":gg"
         Top             =   2040
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   20
         Left            =   1080
         Picture         =   "Form1.frx":3551
         Tag             =   ":tilt"
         Top             =   1200
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   31
         Left            =   2280
         Picture         =   "Form1.frx":38DF
         Tag             =   ":guitar"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   15
         Left            =   3360
         Picture         =   "Form1.frx":3CA5
         Tag             =   ":fuckyou"
         Top             =   1680
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Height          =   390
         Index           =   33
         Left            =   2280
         Picture         =   "Form1.frx":405D
         Tag             =   ":flush"
         Top             =   1800
         Width           =   390
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   19
         Left            =   1080
         Picture         =   "Form1.frx":443B
         Tag             =   ":erm"
         Top             =   960
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   17
         Left            =   1080
         Picture         =   "Form1.frx":47BD
         Tag             =   ":devil"
         Top             =   480
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   26
         Left            =   3360
         Picture         =   "Form1.frx":4B85
         Tag             =   ":cry"
         Top             =   3000
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   28
         Left            =   2280
         Picture         =   "Form1.frx":4F15
         Tag             =   ":bala"
         Top             =   360
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   345
         Index           =   32
         Left            =   2280
         Picture         =   "Form1.frx":5294
         Tag             =   ":ass"
         Top             =   1440
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   14
         Left            =   3360
         Picture         =   "Form1.frx":5659
         Tag             =   ":arnie"
         Top             =   1200
         Width           =   795
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   18
         Left            =   1080
         Picture         =   "Form1.frx":5AC0
         Tag             =   ":heart"
         Top             =   720
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   29
         Left            =   2280
         Picture         =   "Form1.frx":5E43
         Tag             =   ":wave"
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   35
         Left            =   3360
         Picture         =   "Form1.frx":61DB
         Tag             =   ":toiletclaw"
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":shoot"
         Height          =   255
         Left            =   4560
         TabIndex        =   41
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":'("
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":baby"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":nono"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":cool"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":smoke"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":?"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":sleep"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":grr"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":("
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":p"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":D"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         Caption         =   ";)"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":)"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   0
         Width           =   135
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   10
         Left            =   0
         Picture         =   "Form1.frx":6606
         Tag             =   ":nono"
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":69E8
         Tag             =   ":)"
         Top             =   0
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   1
         Left            =   0
         Picture         =   "Form1.frx":6D65
         Tag             =   ";)"
         Top             =   240
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   2
         Left            =   0
         Picture         =   "Form1.frx":70E7
         Tag             =   ":D"
         Top             =   480
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   4
         Left            =   0
         Picture         =   "Form1.frx":746F
         Tag             =   ":("
         Top             =   960
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   5
         Left            =   0
         Picture         =   "Form1.frx":77ED
         Tag             =   ":grr"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   6
         Left            =   0
         Picture         =   "Form1.frx":7B7D
         Tag             =   ":sleep"
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   7
         Left            =   0
         Picture         =   "Form1.frx":7F11
         Tag             =   ":?"
         Top             =   1680
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   3
         Left            =   0
         Picture         =   "Form1.frx":82A6
         Tag             =   ":p"
         Top             =   720
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   8
         Left            =   0
         Picture         =   "Form1.frx":862D
         Tag             =   ":smoke"
         Top             =   2040
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   255
         Index           =   11
         Left            =   0
         Picture         =   "Form1.frx":89C8
         Tag             =   ":baby"
         Top             =   2760
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   9
         Left            =   0
         Picture         =   "Form1.frx":8D68
         Tag             =   ":cool"
         Top             =   2280
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   12
         Left            =   0
         Picture         =   "Form1.frx":90EB
         Tag             =   ":'("
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":fuckyou"
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":beat"
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":arnie"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":cry"
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":blah"
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   38
         Left            =   4560
         Picture         =   "Form1.frx":9484
         Tag             =   ":ugly"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   39
         Left            =   3360
         Picture         =   "Form1.frx":980C
         Tag             =   ":clown"
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   420
         Index           =   40
         Left            =   4560
         Picture         =   "Form1.frx":9C0D
         Tag             =   ":elk"
         Top             =   360
         Width           =   450
      End
      Begin VB.Image imgIcon 
         Height          =   270
         Index           =   41
         Left            =   2280
         Picture         =   "Form1.frx":A106
         Tag             =   ":cat"
         Top             =   2640
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   42
         Left            =   3360
         Picture         =   "Form1.frx":A4AE
         Tag             =   ":evil"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   43
         Left            =   1080
         Picture         =   "Form1.frx":A83B
         Tag             =   ":drink"
         Top             =   3000
         Width           =   570
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   44
         Left            =   3360
         Picture         =   "Form1.frx":AC46
         Tag             =   ":wow"
         Top             =   2760
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   45
         Left            =   4560
         Picture         =   "Form1.frx":AFCE
         Tag             =   ":satan"
         Top             =   840
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   46
         Left            =   2280
         Picture         =   "Form1.frx":B3AF
         Tag             =   ":bear"
         Top             =   3000
         Width           =   285
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":evil"
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":bear"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":cat"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":clown"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":elk"
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":satan"
         Height          =   255
         Left            =   4920
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":drink"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":wow"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":ugly"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label50 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":uriel"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   2640
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   47
         Left            =   1080
         Picture         =   "Form1.frx":B73F
         Tag             =   ":uriel"
         Top             =   2640
         Width           =   315
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   6360
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image10 
      Height          =   4080
      Left            =   0
      Picture         =   "Form1.frx":BB1C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   165
      Left            =   110
      Picture         =   "Form1.frx":C4C0
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image Image8 
      Height          =   165
      Left            =   480
      Picture         =   "Form1.frx":C848
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   7200
   End
   Begin VB.Image Image7 
      Height          =   165
      Left            =   7680
      Picture         =   "Form1.frx":CBA4
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image Image6 
      Height          =   4080
      Left            =   7680
      Picture         =   "Form1.frx":CF2C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   7680
      Picture         =   "Form1.frx":D8D0
      Top             =   120
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   480
      Picture         =   "Form1.frx":E33C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7200
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   0
      Picture         =   "Form1.frx":EAD0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Socket no"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   8640
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3750
      TabIndex        =   4
      Top             =   5520
      Width           =   75
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Online Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   6360
      TabIndex        =   3
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send a Message to a selected user"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   720
      TabIndex        =   2
      Top             =   3045
      Width           =   2910
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   1170
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu send 
         Caption         =   "&Send"
      End
      Begin VB.Menu q 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu clear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu unclear 
         Caption         =   "&Unclear"
      End
   End
   Begin VB.Menu admin 
      Caption         =   "&Admin Tools"
      Begin VB.Menu buzz 
         Caption         =   "&Buzz Client"
      End
      Begin VB.Menu opencd 
         Caption         =   "&Open CD Tray"
      End
      Begin VB.Menu closecd 
         Caption         =   "&Close CD Tray"
      End
      Begin VB.Menu hideclient 
         Caption         =   "&Hide Client"
      End
      Begin VB.Menu sendffile 
         Caption         =   "S&end File"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu website 
         Caption         =   "&Open a web site"
      End
      Begin VB.Menu restart 
         Caption         =   "&Restart"
      End
      Begin VB.Menu logoff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu shutdown 
         Caption         =   "&Shut Down"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu content 
         Caption         =   "&Contents"
      End
      Begin VB.Menu about 
         Caption         =   "About &Us"
      End
   End
   Begin VB.Menu mnuAppPopup 
      Caption         =   "AppPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopWhenMin 
         Caption         =   "Popup when minimized"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''modified in excellence '''''''''''''''''''
Dim leftstr As String
Dim intmax As Long
Dim AbortFile As Boolean
Dim SendingFile As Boolean
Dim BeginTransfer As Single

Private Sub about_Click()
Form1.Enabled = False
Form2.Show
End Sub


Private Sub buzz_Click()
If List1.ListIndex <> -1 Then
TCPServer(List2.List(List1.ListIndex)).SendData "/COMMbuzz"
Else
  MsgBox "Please select a user", , "Network Chat"
End If
End Sub

Private Sub clear_Click()
Text1.Text = ""
End Sub


Private Sub closecd_Click()
If List1.ListIndex <> -1 Then
TCPServer(List2.List(List1.ListIndex)).SendData "/COMMclose" & Text2.Text
Else
  MsgBox "Please select a user", , "Network Chat"
End If
End Sub

Private Sub Command1_Click()
 If Text2.Text = "" Then
   Exit Sub
 End If
  If List1.ListIndex <> -1 Then
   TCPServer(List2.List(List1.ListIndex)).SendData "/MESS<From Server> " & Text2.Text
    
    'If Left(Text2.Text, 1) = ":" Then
    If InStr(1, Text2.Text, ":") <> 0 Then
        s = Split(Text2.Text, ":")
        Text1.SelText = "<From Server> " & s(0)
        Text3.SelText = "<From Server> " & s(0)
        
        For k = 1 To UBound(s)
         Text2.Text = ":" & s(k)
         For I = 0 To 47
            If Left(LCase(Text2.Text), Len(imgIcon(I).Tag)) = imgIcon(I).Tag Then
              
              Clipboard.clear
              Clipboard.SetData imgIcon(I).Picture
             
              Text1.SelStart = Len(Text1.Text)
              Text1.Locked = False7
              SendMessage Text1.hwnd, WM_PASTE, 0, 0
              Text1.Locked = True
              
              Text3.SelStart = Len(Text3.Text)
              Text3.Locked = False
              SendMessage Text3.hwnd, WM_PASTE, 0, 0
              Text3.Locked = True
                
            End If
         Next I
        Next k
        Text1.SelStart = Len(Text1.Text)
        Text1.SelText = Chr$(13) + Chr$(10)
        Text3.SelStart = Len(Text3.Text)
        Text3.SelText = Chr$(13) + Chr$(10)
    Else
      Text1.SelText = "<From Server> " & Text2.Text + Chr$(13) + Chr$(10)
      Text3.SelText = "<From Server> " & Text2.Text + Chr$(13) + Chr$(10)
    End If
  Else
    Text2.Text = ""
    MsgBox "Please select a user", , "Network Chat"
    Exit Sub
  End If
    Text2.Text = ""
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub content_Click()
MsgBox "Under Construction", , "Network Chat"
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
    intmax = 0
    TCPServer(0).LocalPort = 56789
    TCPServer(0).Listen
 
End Sub


Private Sub Form_Resize()
    'this procedure handles all the resizing and positioning of the controls
    On Error Resume Next
    If Form1.Height < 5460 Then
            If Err.Number = 384 Then
            Err.clear
            Exit Sub
        End If
        Form1.Height = 5460
    End If
    If Form1.Width < 8400 Then
        If Err.Number = 384 Then
            Err.clear
            Exit Sub
        End If
        Form1.Width = 8400
    End If
    Do While Form1.Width >= 8400 And Form1.Height >= 5460

    Image3.Move 0, 120
    Image5.Move Width - 660, 120
    Image6.Move Width - 660, 480
    Image6.Height = Height - 1400

    Image10.Move 0, 480
    Image10.Height = Height - 1400

    Image4.Width = Width - 800

    Image8.Move 480, Height - 960
    Image8.Width = Width - 1160

    Image9.Move 110, Height - 960
    Image7.Move Width - 660, Height - 960


    Label1.Move Width - (Width - 720), Height - (Height - 600)
    Label2.Move Width - (Width - 720), Height - 2415
    Label3.Move Width - 2045, 600
    
    Text1.Move Width - (Width - 720), Height - (Height - 960)
    Text1.Width = Form1.Width - 3200
    Text1.Height = Form1.Height - 3435

    Text2.Move Width - (Width - 720), Height - 2100
    Text2.Width = Form1.Width - 3200


    List1.Move Width - 2045, Height - (Height - 940)
    List1.Height = Form1.Height - 2370

    Command1.Move Width - 3560, Height - 1365
    Exit Sub
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub hideclient_Click()
If List1.ListIndex <> -1 Then
TCPServer(List2.List(List1.ListIndex)).SendData "/COMMcloseform" & Text2.Text
Else
  MsgBox "Please select a user", , "Network Chat"
End If
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub


Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub

Private Sub logoff_Click()
If List1.ListIndex <> -1 Then
TCPServer(List2.List(List1.ListIndex)).SendData "/COMMlogoff" & Text2.Text
Else
  MsgBox "Please select a user", , "Network Chat"
End If
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub opencd_Click()
If List1.ListIndex <> -1 Then
TCPServer(List2.List(List1.ListIndex)).SendData "/COMMopen" & Text2.Text
Else
  MsgBox "Please select a user", , "Network Chat"
End If
End Sub

Private Sub restart_Click()
If List1.ListIndex <> -1 Then
TCPServer(List2.List(List1.ListIndex)).SendData "/COMMrestart" & Text2.Text
Else
  MsgBox "Please select a user", , "Network Chat"
End If
End Sub


Private Sub send_Click()
Command1_Click
End Sub

Private Sub sendffile_Click()

On Error GoTo ErrorHandler
If List1.ListIndex <> -1 Then
   CommonDialog1.ShowOpen
   txtFileName = CommonDialog1.FileName
   SendFile txtFileName
   Exit Sub
Else
  MsgBox "Please select a user", , "Network Chat"
End If
   
ErrorHandler:
   Exit Sub

End Sub

Private Sub shutdown_Click()
If List1.ListIndex <> -1 Then
TCPServer(List2.List(List1.ListIndex)).SendData "/COMMshutdown" & Text2.Text
Else
  MsgBox "Please select a user", , "Network Chat"
End If
End Sub

Private Sub TCPServer_ConnectionRequest(index As Integer, ByVal requestID As Long)
    If index = 0 Then
        intmax = intmax + 1
        Load TCPServer(intmax)
        TCPServer(intmax).LocalPort = 0
        TCPServer(intmax).Accept requestID
        List2.AddItem intmax
    End If
End Sub
Private Sub TCPServer_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim un As Integer

  TCPServer(index).GetData strData
  leftstr = Left(strData, 5)
  
  If leftstr = "/NEWU" Then
      newu = Right(strData, Len(strData) - 5)
      For I = 0 To List1.ListCount - 1
        If newu = LCase(List1.List(I)) Or newu = UCase(List1.List(I)) Then
           x = List2.ListCount - 1
           List2.RemoveItem x
           TCPServer(index).SendData "/CLRU" & newu
           Exit Sub
        End If
      Next I
      
      
      List1.AddItem newu
        
      For I = 0 To List2.ListCount - 1
        For j = 0 To List1.ListCount - 1
           TCPServer(List2.List(I)).SendData "/CLST" & List1.List(j)
           Call pause
        Next j
      Next I
  End If
  
  If leftstr = "/MESS" Then
      
    mess = Right(strData, Len(strData) - 5)
    If InStr(1, mess, ":") <> 0 Then
       s = Split(mess, ":")
       Text1.SelText = s(0)
       Text3.SelText = s(0)
     
     For k = 1 To UBound(s)
       t = ":" & s(k)
     For I = 0 To 47
        If Left(LCase(t), Len(imgIcon(I).Tag)) = imgIcon(I).Tag Then
          
          Clipboard.clear
          Clipboard.SetData imgIcon(I).Picture
         
          Text1.SelStart = Len(Text1.Text)
          Text1.Locked = False
          SendMessage Text1.hwnd, WM_PASTE, 0, 0
          Text1.Locked = True
            
          Text3.SelStart = Len(Text3.Text)
          Text3.Locked = False
          SendMessage Text3.hwnd, WM_PASTE, 0, 0
          Text3.Locked = True
        End If
    Next I
    Next k
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = Chr$(13) + Chr$(10)
    Text3.SelStart = Len(Text3.Text)
    Text3.SelText = Chr$(13) + Chr$(10)
         Exit Sub
    End If
                          
    Text1.SelText = mess + Chr$(13) + Chr$(10)
    Text3.SelText = mess + Chr$(13) + Chr$(10)
    Text1.SelStart = Len(Text1.Text)
    Text3.SelStart = Len(Text3.Text)
  End If
  
  If leftstr = "/GOST" Then
      uno = Split(strData, "/")
      un = Val(uno(2))
      TCPServer(List2.List(un)).SendData "/MESS" & uno(3)
  End If
  
  If leftstr = "/CBUZ" Then
      uno = Split(strData, "/")
      un = Val(uno(2))
      TCPServer(List2.List(un)).SendData "/COMMbuzz"
  End If
  
  
  If leftstr = "/QUIT" Then
      quitu = Right(strData, Len(strData) - 5)
      For RemoveItem = (List1.ListCount - 1) To 0 Step -1
        If List1.List(RemoveItem) = quitu Then
             List1.RemoveItem (RemoveItem)
             List2.RemoveItem (RemoveItem)
        End If
      Next RemoveItem
      
     For I = 0 To List2.ListCount - 1
           TCPServer(List2.List(I)).SendData "/REMO" & quitu
           Call pause
     Next I
      
  End If
  
  
  If strData = "R_" Then SendingFile = True
   'The string "N_" means that the other side doesn't accept the transfer.
   If strData = "N_" Then
       MsgBox "The remote computer refuses to accept the file.", vbInformation, "Network Chat"
       AbortFile = True
  End If
  
  
  
    
End Sub

Sub pause()
    'a pause procedure for 1 sec duration
    starttime = Timer
    Do
        DoEvents
    Loop While Timer < starttime + 1
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.SelLength = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
    KeyAscii = 0
End If
End Sub

Private Sub unclear_Click()
Text1.TextRTF = Text3.TextRTF
Text1.SelStart = Len(Text1.Text)
Text3.SelStart = Len(Text3.Text)
End Sub

Private Sub website_Click()
Dim website As String
If List1.ListIndex <> -1 Then
   website = InputBox("Enter the url of the website you would like to open:", "Enter url")
   TCPServer(List2.List(List1.ListIndex)).SendData "/COMW" & website
Else
   MsgBox "Please select a user", , "Network Chat"
End If
End Sub

Private Sub SendFile(ByVal FileName As String)

   Dim FileData As String
   Dim ByteData As Byte
   Dim Counter As Long
   
   Open FileName For Binary As #1
   
   DoEvents
   'Read the file into the variable FileData
   FileData = Input(LOF(1), 1)
     
   Close
   
   SendingFile = False
   AbortFile = False
   
   If MsgBox(FileTitle(FileName) & " (" & Len(FileData) & " bytes)" & vbCrLf & _
   "Begin the file transfer?", vbInformation Or vbYesNo, "Network Chat") <> vbYes Then
      Exit Sub
   End If
   TCPServer(List2.List(List1.ListIndex)).SendData "S_" & Len(FileData) & "_" & FileTitle(FileName)
   
   'This loop suspends the program until the other side
   Do Until SendingFile Or AbortFile Or DoEvents = 0
      DoEvents
   Loop
   
      
   'This command begins the file transfer.The whole file is stored
   'in the string variable FileData.
   BeginTransfer = Timer
   TCPServer(List2.List(List1.ListIndex)).SendData FileData

End Sub

Private Function FileTitle(ByVal FileName As String) As String

   Dim I As Integer
   Dim Temp As String
   
 
   If InStr(FileName, "\") <> 0 Then
      I = Len(FileName)
      Do Until Left(Temp, 1) = "\"
         I = I - 1
         Temp = Mid(FileName, I)
      Loop
      FileTitle = Mid(Temp, 2)
   Else
      FileTitle = FileName
   End If

End Function

