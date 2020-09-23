VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Network Chat <Client>"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4935
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   5880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Online Users:"
      Height          =   2940
      Left            =   5520
      TabIndex        =   66
      Top             =   1080
      Width           =   1500
      Begin MSComctlLib.ImageList IL 
         Left            =   2280
         Top             =   3000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":059E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvUsers 
         Height          =   2655
         Left            =   90
         TabIndex        =   67
         Top             =   195
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   4683
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   472
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "IL"
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   5205
      TabIndex        =   59
      Top             =   3840
      Width           =   5205
      Begin VB.CommandButton emotion1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4800
         Picture         =   "Form1.frx":06FA
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdColors 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4350
         Picture         =   "Form1.frx":0816
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   0
         Width           =   315
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   330
         ItemData        =   "Form1.frx":0B58
         Left            =   0
         List            =   "Form1.frx":0B5A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   0
         Width           =   2280
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         ItemData        =   "Form1.frx":0B5C
         Left            =   2400
         List            =   "Form1.frx":0B75
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   0
         Width           =   495
      End
      Begin VB.CheckBox chkUnderline 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3900
         Picture         =   "Form1.frx":0B93
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   0
         Width           =   315
      End
      Begin VB.CheckBox chkItalic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3450
         Picture         =   "Form1.frx":0ED5
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   0
         Width           =   315
      End
      Begin VB.CheckBox chkBold 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3000
         Picture         =   "Form1.frx":1217
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1085
      _Version        =   393217
      HideSelection   =   0   'False
      TextRTF         =   $"Form1.frx":1559
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   35
      Left            =   -120
      TabIndex        =   6
      Top             =   0
      Width           =   12015
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6720
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S E N D"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin MSWinsockLib.Winsock tcpclient 
      Left            =   6360
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2370
      ItemData        =   "Form1.frx":162D
      Left            =   5640
      List            =   "Form1.frx":162F
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   295
      Left            =   1650
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   170
      Width           =   1890
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2370
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4180
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":1631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text5 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Form1.frx":1705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   2520
      Width           =   3375
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   47
         Left            =   1080
         Picture         =   "Form1.frx":17D9
         Tag             =   ":uriel"
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label Label50 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":uriel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   58
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":ugly"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   57
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":wow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   56
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":drink"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   55
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":satan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   54
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":elk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   53
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":clown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   52
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":cat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   51
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":bear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   50
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":evil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   49
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   46
         Left            =   2280
         Picture         =   "Form1.frx":1BB6
         Tag             =   ":bear"
         Top             =   3000
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   45
         Left            =   4560
         Picture         =   "Form1.frx":1F46
         Tag             =   ":satan"
         Top             =   840
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   44
         Left            =   3360
         Picture         =   "Form1.frx":2327
         Tag             =   ":wow"
         Top             =   2760
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   43
         Left            =   1080
         Picture         =   "Form1.frx":26AF
         Tag             =   ":drink"
         Top             =   3000
         Width           =   570
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   42
         Left            =   3360
         Picture         =   "Form1.frx":2ABA
         Tag             =   ":evil"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   270
         Index           =   41
         Left            =   2280
         Picture         =   "Form1.frx":2E47
         Tag             =   ":cat"
         Top             =   2640
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   420
         Index           =   40
         Left            =   4560
         Picture         =   "Form1.frx":31EF
         Tag             =   ":elk"
         Top             =   360
         Width           =   450
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   39
         Left            =   3360
         Picture         =   "Form1.frx":36E8
         Tag             =   ":clown"
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   38
         Left            =   4560
         Picture         =   "Form1.frx":3AE9
         Tag             =   ":ugly"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":blah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   48
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":cry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":arnie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   46
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":beat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   45
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":fuckyou"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   44
         Top             =   1680
         Width           =   615
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   12
         Left            =   0
         Picture         =   "Form1.frx":3E71
         Tag             =   ":'("
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   9
         Left            =   0
         Picture         =   "Form1.frx":420A
         Tag             =   ":cool"
         Top             =   2280
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   255
         Index           =   11
         Left            =   0
         Picture         =   "Form1.frx":458D
         Tag             =   ":baby"
         Top             =   2760
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   8
         Left            =   0
         Picture         =   "Form1.frx":492D
         Tag             =   ":smoke"
         Top             =   2040
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   3
         Left            =   0
         Picture         =   "Form1.frx":4CC8
         Tag             =   ":p"
         Top             =   720
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   7
         Left            =   0
         Picture         =   "Form1.frx":504F
         Tag             =   ":?"
         Top             =   1680
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   6
         Left            =   0
         Picture         =   "Form1.frx":53E4
         Tag             =   ":sleep"
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   5
         Left            =   0
         Picture         =   "Form1.frx":5778
         Tag             =   ":grr"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   4
         Left            =   0
         Picture         =   "Form1.frx":5B08
         Tag             =   ":("
         Top             =   960
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   2
         Left            =   0
         Picture         =   "Form1.frx":5E86
         Tag             =   ":D"
         Top             =   480
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   1
         Left            =   0
         Picture         =   "Form1.frx":620E
         Tag             =   ";)"
         Top             =   240
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":6590
         Tag             =   ":)"
         Top             =   0
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   10
         Left            =   0
         Picture         =   "Form1.frx":690D
         Tag             =   ":nono"
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         Caption         =   ";)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":p"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":grr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":sleep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":smoke"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":cool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":nono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":baby"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":'("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":shoot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   3000
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   35
         Left            =   3360
         Picture         =   "Form1.frx":6CEF
         Tag             =   ":toiletclaw"
         Top             =   0
         Width           =   705
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   29
         Left            =   2280
         Picture         =   "Form1.frx":711A
         Tag             =   ":wave"
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   18
         Left            =   1080
         Picture         =   "Form1.frx":74B2
         Tag             =   ":heart"
         Top             =   720
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   14
         Left            =   3360
         Picture         =   "Form1.frx":7835
         Tag             =   ":arnie"
         Top             =   1200
         Width           =   795
      End
      Begin VB.Image imgIcon 
         Height          =   345
         Index           =   32
         Left            =   2280
         Picture         =   "Form1.frx":7C9C
         Tag             =   ":ass"
         Top             =   1440
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   28
         Left            =   2280
         Picture         =   "Form1.frx":8061
         Tag             =   ":bala"
         Top             =   360
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   26
         Left            =   3360
         Picture         =   "Form1.frx":83E0
         Tag             =   ":cry"
         Top             =   3000
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   17
         Left            =   1080
         Picture         =   "Form1.frx":8770
         Tag             =   ":devil"
         Top             =   480
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   19
         Left            =   1080
         Picture         =   "Form1.frx":8B38
         Tag             =   ":erm"
         Top             =   960
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   390
         Index           =   33
         Left            =   2280
         Picture         =   "Form1.frx":8EBA
         Tag             =   ":flush"
         Top             =   1800
         Width           =   390
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   15
         Left            =   3360
         Picture         =   "Form1.frx":9298
         Tag             =   ":fuckyou"
         Top             =   1680
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   31
         Left            =   2280
         Picture         =   "Form1.frx":9650
         Tag             =   ":guitar"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   20
         Left            =   1080
         Picture         =   "Form1.frx":9A16
         Tag             =   ":tilt"
         Top             =   1200
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   23
         Left            =   1080
         Picture         =   "Form1.frx":9DA4
         Tag             =   ":gg"
         Top             =   2040
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   34
         Left            =   2280
         Picture         =   "Form1.frx":A130
         Tag             =   ":guns"
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   30
         Left            =   2280
         Picture         =   "Form1.frx":A53B
         Tag             =   ":light"
         Top             =   840
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   25
         Left            =   3360
         Picture         =   "Form1.frx":A8D4
         Tag             =   ":beat"
         Top             =   840
         Width           =   525
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   36
         Left            =   3360
         Picture         =   "Form1.frx":ACD6
         Tag             =   ":blah"
         Top             =   480
         Width           =   660
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   21
         Left            =   1080
         Picture         =   "Form1.frx":B0A4
         Tag             =   ":lol"
         Top             =   1560
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   22
         Left            =   1080
         Picture         =   "Form1.frx":B428
         Tag             =   ":mad"
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   24
         Left            =   1080
         Picture         =   "Form1.frx":B7AE
         Tag             =   ":shit"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   315
         Index           =   37
         Left            =   4200
         Picture         =   "Form1.frx":BB24
         Tag             =   ":angel"
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   16
         Left            =   1080
         Picture         =   "Form1.frx":BF76
         Tag             =   ":alien"
         Top             =   0
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   345
         Index           =   13
         Left            =   4200
         Picture         =   "Form1.frx":C321
         Tag             =   ":shoot"
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":shit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":gg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":mad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":lol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":flush"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":erm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":angel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   23
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":devil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":alien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":tilt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":heart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":king"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":bala"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":wave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":ass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":light"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":guitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":guns"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":toiletclaw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   0
         Width           =   735
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   27
         Left            =   2280
         Picture         =   "Form1.frx":C82E
         Tag             =   ":king"
         Top             =   0
         Width           =   345
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1170
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu send 
         Caption         =   "&Send"
      End
      Begin VB.Menu save 
         Caption         =   "S&ave"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu undo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu divider 
         Caption         =   "-"
      End
      Begin VB.Menu clear 
         Caption         =   "Cl&ear"
      End
      Begin VB.Menu cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu selectall 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu showtime 
         Caption         =   "&Show Time Stamp"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu friend 
      Caption         =   "&Friend"
      Begin VB.Menu buzz 
         Caption         =   "&Buzz Friend"
      End
      Begin VB.Menu showemotions 
         Caption         =   "Show &Emotions"
      End
   End
   Begin VB.Menu format 
      Caption         =   "F&ormat"
      Begin VB.Menu font 
         Caption         =   "&Font"
      End
      Begin VB.Menu size 
         Caption         =   "&Size"
      End
      Begin VB.Menu color 
         Caption         =   "&Color"
      End
      Begin VB.Menu playsnd 
         Caption         =   "&Play Typing Sound"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu content 
         Caption         =   "&Contents"
      End
      Begin VB.Menu aboutus 
         Caption         =   "About &Us"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lenstr, arrstr, b, u, i As Integer
Dim colorr As Long
Dim ff As Integer
Dim con As Integer
Public flg1 As Integer
Public ftop As Integer
Public fleft As Integer
Dim RecievedFile As String
Dim BeginTransfer As Single
Dim psnd As Integer

Private Sub aboutus_Click()
Form1.Enabled = False
Form3.Show
End Sub

Private Sub close_Click()
If tcpclient.State = sckConnected Then
Call tcpclient.SendData("/QUIT" & user)
DoEvents
End
End If
tcpclient.close
End
End Sub

Private Sub color_Click()
cmdColors.SetFocus
cmdColors_Click
End Sub

Private Sub content_Click()
MsgBox "Complete Communication Tool", , "Network Chat"
End Sub

Private Sub buzz_Click()
If List1.ListIndex <> -1 Then
    tcpclient.SendData "/CBUZ/" & List1.ListIndex
Else
    MsgBox "Please select a user", , "Network Chat"
    Exit Sub
End If
End Sub

Private Sub chkbold_Click()
On Error GoTo err
If Text2.Text = "" Then Exit Sub

If chkBold.Value = 1 Then
   b = 1
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.Text)
   Text2.SelBold = Not Text2.SelBold
   Text2.SetFocus
Else
   b = 0
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.Text)
   Text2.SelBold = Not Text2.SelBold
   Text2.SetFocus
End If
   Exit Sub

err:
    Exit Sub

End Sub

Private Sub chkItalic_Click()
On Error GoTo err
If chkItalic.Value = 1 Then
   i = 1
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.Text)
   Text2.SelItalic = Not Text2.SelItalic
   Text2.SetFocus
Else
   i = 0
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.Text)
   Text2.SelItalic = Not Text2.SelItalic
   Text2.SetFocus
End If
   Exit Sub
   
err:
    Exit Sub
End Sub

Private Sub chkUnderline_Click()
On Error GoTo err:
If chkUnderline.Value = 1 Then
   u = 1
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.Text)
   Text2.SelUnderline = Not Text2.SelUnderline
   Text2.SetFocus
Else
   u = 0
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.Text)
   Text2.SelUnderline = Not Text2.SelUnderline
   Text2.SetFocus
End If
   Exit Sub
   
err:
    Exit Sub
End Sub

Private Sub clear_Click()
Text1.Text = ""
End Sub

Private Sub cmbFonts_Click()
 On Error Resume Next
 Text2.font = cmbFonts.List(cmbFonts.ListIndex)
 Text2.SelFontName = cmbFonts.List(cmbFonts.ListIndex)
 Text2.SelStart = 0
 Text2.SelLength = Len(Text2.Text)
 
 If b = 1 Then
    Text2.SelBold = True
 Else
    Text2.SelBold = False
 End If
 
 If i = 1 Then
    Text2.SelItalic = True
 Else
    Text2.SelItalic = False
 End If
 
 If u = 1 Then
    Text2.SelUnderline = True
 Else
    Text2.SelUnderline = False
 End If
 
Text2.SetFocus
End Sub

Private Sub cmdColors_Click()
On Error GoTo ErrorHandler
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    colorr = CommonDialog1.color
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text2.SelColor = colorr
    Text2.SelStart = Len(Text2.Text)
    Text2.SetFocus
    Exit Sub

ErrorHandler:
    Exit Sub
End Sub

Private Sub Combo1_Click()
 On Error Resume Next
 Text2.SelStart = 0
 Text2.SelLength = Len(Text2.Text)
 Text2.SelFontSize = Combo1.List(Combo1.ListIndex)
    
 If b = 1 Then
    Text2.SelBold = True
 Else
    Text2.SelBold = False
 End If
 
 If i = 1 Then
    Text2.SelItalic = True
 Else
    Text2.SelItalic = False
 End If
 
 If u = 1 Then
    Text2.SelUnderline = True
 Else
    Text2.SelUnderline = False
 End If
  
 Text2.SetFocus
End Sub

Private Sub Command1_Click()

If Text2.Text = "" Then
Exit Sub
End If

If List1.ListIndex <> -1 Then
  tcpclient.SendData "/GOST/" & List1.ListIndex & "/" & "<From " & user & ">" & Text2.Text
Else
   tcpclient.SendData "/MESS<From " & user & "> " & Text2.Text
End If

'If Left(Text2.Text, 1) = ":" Then
If InStr(1, Text2.Text, ":") <> 0 Then
  s = Split(Text2.Text, ":")
  
  If ff = 1 Then
    Text1.SelText = "<From " & user & " (" & Time & ") >" & s(0)
    Text5.SelText = "<From " & user & " (" & Time & ") >" & s(0)
  Else
    Text1.SelText = "<From " & user & "> " & s(0)
    Text5.SelText = "<From " & user & "> " & s(0)
  End If
  
  For k = 1 To UBound(s)
   
   Text2.Text = ":" & s(k)
    For i = 0 To 47
        If Left(LCase(Text2.Text), Len(imgIcon(i).Tag)) = imgIcon(i).Tag Then
          Clipboard.clear
          Clipboard.SetData imgIcon(i).Picture
         
          Text1.SelStart = Len(Text1.Text)
          Text1.Locked = False
          SendMessage Text1.hWnd, WM_PASTE, 0, 0
          Text1.Locked = True
            
          Text5.SelStart = Len(Text5.Text)
          Text5.Locked = False
          SendMessage Text5.hWnd, WM_PASTE, 0, 0
          Text5.Locked = True
        End If
    Next i
Next k
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = Chr$(13) + Chr$(10)
    Text5.SelStart = Len(Text5.Text)
    Text5.SelText = Chr$(13) + Chr$(10)
Else
   If ff = 1 Then
    Text1.SelText = "<From " & user & " (" & Time & " )>" & Text2.Text + Chr$(13) + Chr$(10)
    Text5.SelText = "<From " & user & " (" & Time & " )>" & Text2.Text + Chr$(13) + Chr$(10)
    lenstr = Len(Text2.Text) + 14 + Len(user) + Len(Time)
   Else
    Text1.SelText = "<From " & user & "> " & Text2.Text + Chr$(13) + Chr$(10)
    Text5.SelText = "<From " & user & "> " & Text2.Text + Chr$(13) + Chr$(10)
    lenstr = Len(Text2.Text) + 10 + Len(user)
   End If
    If Len(Text1.Text) = 0 Then
        Exit Sub
    Else
        Text1.SelStart = Len(Text1.Text) - lenstr
        Text1.SelLength = lenstr
        Text1.SelFontName = cmbFonts.List(cmbFonts.ListIndex)
        Text1.SelFontSize = Combo1.List(Combo1.ListIndex)
        If b = 1 Then
           Text1.SelBold = True
        Else
           Text1.SelBold = False
        End If
        If i = 1 Then
           Text1.SelItalic = True
        Else
           Text1.SelItalic = False
        End If
        If u = 1 Then
           Text1.SelUnderline = True
        Else
           Text1.SelUnderline = False
        End If
        
        Text1.SelColor = colorr
        Text1.SelStart = Len(Text1.Text)
    End If
End If
Text2.Text = ""

End Sub

Private Sub copy_Click()
On Error GoTo err
 Clipboard.clear
 Clipboard.SetText Screen.ActiveControl.SelText
 Exit Sub
err:
 Exit Sub
End Sub

Private Sub cut_Click()
On Error GoTo err
 Clipboard.clear
 Clipboard.SetText Screen.ActiveControl.SelText
 Screen.ActiveControl.SelText = ""
 Exit Sub
err:
 Exit Sub
End Sub

Private Sub emotion1_Click()
frmSplash.Frame2.Visible = True
frmSplash.Show
End Sub

Private Sub font_Click()
cmbFonts.SetFocus
End Sub

Private Sub Form_Load()
Dim rootitem As Node

    Text3.Text = user
    tcpclient.RemoteHost = server
    tcpclient.RemotePort = 56789
    tcpclient.Connect

    Do Until tcpclient.State = 7 Or tcpclient.State = 9
        DoEvents
    Loop

    If tcpclient.State = 9 Then
         Command1.Enabled = False
         Text2.Enabled = False
    End If
    
    If tcpclient.State = 7 Then
         tcpclient.SendData "/NEWU" & user
    End If
    
    For i = 0 To Screen.FontCount - 1
        cmbFonts.AddItem Screen.Fonts(i)
    Next i
    
    cmbFonts.Text = cmbFonts.List(0)
    Combo1.Text = Combo1.List(0)
    showtime.Checked = False
    playsnd.Checked = False
    flg1 = 0
    con = 0
    b = 0
    u = 0
    i = 0
    With tvUsers.Nodes
        Set rootitem = .Add(, , "root", "Server", 1)
    End With
End Sub


Private Sub Form_Resize()
    'this procedure handles all the resizing and positioning of the controls
    On Error Resume Next
    If Form1.Height < 5625 Then
            If err.Number = 384 Then
            err.clear
            Exit Sub
        End If
        Form1.Height = 5625
    End If
    If Form1.Width < 7950 Then
        If err.Number = 384 Then
            err.clear
            Exit Sub
        End If
        Form1.Width = 7950
    End If
    Do While Form1.Width >= 7950 And Form1.Height >= 5625
    Text2.Move Width - (Width - 120), Height - 1365
    Text2.Width = Form1.Width - 1920
    
    Picture8.Move Width - (Width - 120), Height - 1770
    Picture8.Width = Form1.Width - 1720
    Frame3.Move Width - 1645, Height - (Height - 500)
    Frame3.Height = Form1.Height - 2300
    tvUsers.Height = Frame3.Height - 280
    
    Text1.Move Width - (Width - 120), Height - (Height - 600)
    Text1.Width = Form1.Width - 1920
    Text1.Height = Form1.Height - 2435
    
    
    Command1.Move Width - 1545, Height - 1365
    Exit Sub
    Loop
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

If tcpclient.State = sckConnected Then
Call tcpclient.SendData("/QUIT" & user)
DoEvents
Unload Me
End If
tcpclient.close

End Sub

Private Sub paste_Click()
On Error GoTo err
Screen.ActiveControl.SelText = Clipboard.GetText()
Exit Sub
err:
 Exit Sub
End Sub

Private Sub playsnd_Click()
If playsnd.Checked = False Then
   playsnd.Checked = True
   psnd = 1
Else
   playsnd.Checked = False
   psnd = 0
End If
End Sub

Private Sub save_Click()

CommonDialog2.CancelError = True

On Error GoTo err
CommonDialog2.Filter = "RTF Files(*.rtf)|*.rtf"
  
CommonDialog2.ShowSave
Text1.SaveFile CommonDialog2.FileName, rtfRTF
Exit Sub

err:
Exit Sub

End Sub

Private Sub selectall_Click()
If Text2.Text = "" Then Exit Sub
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub send_Click()
Command1_Click
End Sub

Private Sub showemotions_Click()
frmSplash.Frame2.Visible = True
frmSplash.Show
End Sub

Private Sub showtime_Click()
If showtime.Checked = False Then
   showtime.Checked = True
   ff = 1
Else
   showtime.Checked = False
   ff = 0
End If
End Sub

Private Sub size_Click()
Combo1.SetFocus
End Sub

Private Sub tcpclient_DataArrival(ByVal bytesTotal As Long)
    
    Dim strData As String
    Static FileName As String
    Static FileSize As Long
    tcpclient.GetData strData
    
    leftstr = Left(strData, 5)
    
    If leftstr = "/CLRU" Then
       MsgBox "select another name", , "Network Chat"
       tcpclient.close
       Unload Me
    End If
    
    If leftstr = "/CLST" Then
       newu = Right(strData, Len(strData) - 5)
       For i = 0 To List1.ListCount - 1
         If newu = List1.List(i) Then
             Flag = 1
        End If
       Next i
       
       If Flag = 0 Then
        List1.AddItem newu
           With tvUsers.Nodes
               Set rootitem = .Add("root", tvwChild, , newu, 1)
           End With
            
           For q = 1 To tvUsers.Nodes.Count
             tvUsers.Nodes(q).Expanded = True
           Next q
           
       Flag = 0
       End If
        
     End If
          
          
    If leftstr = "/MESS" Then
      Beep
      Call popwin
      
      mess = Right(strData, Len(strData) - 5)
    If InStr(1, mess, ":") <> 0 Then
       s = Split(mess, ":")
       
       Text1.SelText = s(0)
       Text5.SelText = s(0)
       
     For k = 1 To UBound(s)
        t = ":" & s(k)
     For i = 0 To 47
        If Left(LCase(t), Len(imgIcon(i).Tag)) = imgIcon(i).Tag Then
          
          Clipboard.clear
          Clipboard.SetData imgIcon(i).Picture
         
          Text1.SelStart = Len(Text1.Text)
          Text1.Locked = False
          SendMessage Text1.hWnd, WM_PASTE, 0, 0
          Text1.Locked = True
           
          Text1.SelStart = Len(Text5.Text)
          Text5.Locked = False
          SendMessage Text5.hWnd, WM_PASTE, 0, 0
          Text5.Locked = True
            
        End If
    Next i
    Next k
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = Chr$(13) + Chr$(10)
    Text5.SelStart = Len(Text5.Text)
    Text5.SelText = Chr$(13) + Chr$(10)
         Exit Sub
    End If
                        
                        
    Text1.SelText = mess + Chr$(13) + Chr$(10)
    Text5.SelText = mess + Chr$(13) + Chr$(10)
            
                      
        If Len(Text1.Text) = 0 Then
            Exit Sub
        Else
            le = Len(mess)
            l = le + 2
            Text1.SelStart = Len(Text1.Text) - l
            Text1.SelLength = le
            Text1.SelFontName = cmbFonts.List(cmbFonts.ListIndex)
            Text1.SelFontSize = Combo1.List(Combo1.ListIndex)
            If b = 1 Then
               Text1.SelBold = Not Text1.SelBold
               b = 0
            End If
            If i = 1 Then
               Text1.SelItalic = Not Text1.SelItalic
               i = 0
            End If
            If u = 1 Then
               Text1.SelUnderline = Not Text1.SelUnderline
               u = 0
            End If
            Text1.SelColor = colorr
            Text1.SelStart = Len(Text1.Text)
        End If
    End If
    
        
     If leftstr = "/REMO" Then
     Beep
        usr = Right(strData, Len(strData) - 5)
        For i = 0 To List1.ListCount - 1
          If List1.List(i) = usr Then
              List1.RemoveItem i
              With tvUsers.Nodes
                  .Remove i + 2
              End With
          End If
        Next i
     End If
     
     If leftstr = "/COMM" Then
       comm = Right(strData, Len(strData) - 5)
     
       Select Case comm
            Case "close"
                CloseCDROM
                
            Case "open"
                OpenCDROM
            
            Case "shutdown"
                ShutDown
                
            Case "restart"
                Restart
                
            Case "logoff"
                LogOff
                
            Case "closeform"
                closewin
            
            Case "buzz"
                Timer1.Enabled = True
            
       End Select
     
     End If
     
     If leftstr = "/COMW" Then
       com = Right(strData, Len(strData) - 5)
       ret& = ShellExecute(Me.hWnd, "Open", com, "", App.Path, 1)
     End If
               
           
     If Left(strData, 2) = "S_" Then
      
      FileSize = Val(Mid(strData, 3, 3 + InStr(4, strData, "_")))
      FileName = Mid(strData, InStr(4, strData, "_") + 1)
      
      Dim Question As String
      Dim Answer As VbMsgBoxResult
      
      Question = "The remote computer wishes to send you this file:" & vbCrLf & _
               FileName & " (" & FileSize & " bytes)" & vbCrLf & vbCrLf & _
               "Recieve this file? "
      Answer = MsgBox(Question, vbInformation Or vbYesNo, "Network Chat")
      
      'then prompt the user with the options to accept or decline
      'the file transfer.
      If Answer = vbYes Then
         'Prepare for the file transfer
                  
         RecievedFile = ""
         'The string "R_" means that this side accepts the file
         'transfer
         tcpclient.SendData "R_"
         BeginTransfer = Timer
      Else
         'The string "N_" means that this side doesnt accept the file
         'transfer
         tcpclient.SendData "N_"
      End If
      
   Else
      'if this is data from the actual file transfer then
      'add it to the variable that contains the data already sent.
      RecievedFile = RecievedFile & strData
            
      'check if the file transfer is complete
      If Len(RecievedFile) = FileSize Then
         CommonDialog3.FileName = FileName
         CommonDialog3.ShowSave
         'prompt the user for a path to save the file in
         Open CommonDialog3.FileName For Binary As #1
         Put #1, 1, RecievedFile
         Close
      End If
      DoEvents
   
   End If
      
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.SelLength = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If psnd = 1 Then
   BeginPlaySound 101
End If
    If tcpclient.State = 8 Then
         KeyAscii = 0
         Command1.Enabled = False
         Text2.Enabled = False
         For i = 0 To List1.ListCount - 1
            List1.RemoveItem i
         Next i
         
         With tvUsers.Nodes
            .clear
         End With
         
    End If

    If KeyAscii = 13 Then
        Command1_Click
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
Call buzzz
End Sub

Private Sub tvUsers_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node = "Server" Then
      List1.ListIndex = -1
    End If

    For i = 0 To List1.ListCount
      If Node = List1.List(i) Then
       List1.ListIndex = i
      End If
    Next i
End Sub

Private Sub undo_Click()
Text1.TextRTF = Text5.TextRTF
Text1.SelStart = Len(Text1.Text)
Text5.SelStart = Len(Text5.Text)
End Sub

Private Sub buzzz()

Select Case flg1
    Case 0
        ftop = Form1.Top
        fleft = Form1.Left
        Form1.Left = Form1.Left + 30
        Form1.Top = Form1.Top + 30
        flg1 = flg1 + 1
        
    Case 1
        Form1.Left = Form1.Left - 45
        Form1.Top = Form1.Top - 45
        flg1 = flg1 + 1
        
    Case 2
        Form1.Left = Form1.Left + 60
        Form1.Top = Form1.Top + 60
        flg1 = flg1 + 1
        
    Case 3
        Form1.Left = Form1.Left - 75
        Form1.Top = Form1.Top - 75
        flg1 = flg1 + 1
        
    Case 4
        Form1.Left = Form1.Left + 90
        Form1.Top = Form1.Top + 90
        flg1 = flg1 + 1
        
    Case 5
        Form1.Left = Form1.Left - 105
        Form1.Top = Form1.Top - 105
        flg1 = flg1 + 1
        
    Case 6
        Form1.Left = Form1.Left + 105
        Form1.Top = Form1.Top + 105
        flg1 = flg1 + 1
        
    Case 6
        Form1.Left = Form1.Left - 75
        Form1.Top = Form1.Top - 75
        flg1 = flg1 + 1
        
    Case 7
        Form1.Left = fleft
        Form1.Top = ftop
        flg1 = 0
        Timer1.Enabled = False
    End Select
        
End Sub

Private Sub popwin()
If Form1.WindowState = vbMinimized Then
    Form1.WindowState = vbNormal
    BeginPlaySound 102
    Timer1.Enabled = True
Else
    Exit Sub
End If
End Sub

