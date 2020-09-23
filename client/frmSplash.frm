VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6330
      Left            =   0
      TabIndex        =   9
      Top             =   -240
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         Height          =   255
         Left            =   7080
         TabIndex        =   59
         Top             =   280
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emotions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   285
         TabIndex        =   58
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   47
         Left            =   1440
         Picture         =   "frmSplash.frx":0442
         Tag             =   ":uriel"
         Top             =   3240
         Width           =   315
      End
      Begin VB.Label Label48 
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
         Left            =   1800
         TabIndex        =   57
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6420
         TabIndex        =   56
         Top             =   2520
         Width           =   330
      End
      Begin VB.Label Label46 
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
         Left            =   4680
         TabIndex        =   55
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label45 
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
         Left            =   2040
         TabIndex        =   54
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6420
         TabIndex        =   53
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label43 
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
         Left            =   4680
         TabIndex        =   52
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label42 
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
         Left            =   4680
         TabIndex        =   51
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label41 
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
         Left            =   3240
         TabIndex        =   50
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label40 
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
         Left            =   3240
         TabIndex        =   49
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label39 
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
         Left            =   4680
         TabIndex        =   48
         Top             =   3000
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   46
         Left            =   2640
         Picture         =   "frmSplash.frx":081F
         Tag             =   ":bear"
         Top             =   3600
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   45
         Left            =   5520
         Picture         =   "frmSplash.frx":0BAF
         Tag             =   ":satan"
         Top             =   1440
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   44
         Left            =   4080
         Picture         =   "frmSplash.frx":0F90
         Tag             =   ":wow"
         Top             =   3360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   43
         Left            =   1440
         Picture         =   "frmSplash.frx":1318
         Tag             =   ":drink"
         Top             =   3600
         Width           =   570
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   42
         Left            =   4080
         Picture         =   "frmSplash.frx":1723
         Tag             =   ":evil"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   270
         Index           =   41
         Left            =   2640
         Picture         =   "frmSplash.frx":1AB0
         Tag             =   ":cat"
         Top             =   3240
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   420
         Index           =   40
         Left            =   3960
         Picture         =   "frmSplash.frx":1E58
         Tag             =   ":elk"
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   39
         Left            =   4080
         Picture         =   "frmSplash.frx":2351
         Tag             =   ":clown"
         Top             =   2640
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   38
         Left            =   5520
         Picture         =   "frmSplash.frx":2752
         Tag             =   ":ugly"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6420
         TabIndex        =   47
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label37 
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
         Left            =   4680
         TabIndex        =   46
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6420
         TabIndex        =   45
         Top             =   2040
         Width           =   390
      End
      Begin VB.Label Label35 
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
         Left            =   4680
         TabIndex        =   44
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         Caption         =   ":out"
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
         Left            =   4680
         TabIndex        =   43
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   12
         Left            =   360
         Picture         =   "frmSplash.frx":2ADA
         Tag             =   ":'("
         Top             =   3720
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   9
         Left            =   360
         Picture         =   "frmSplash.frx":2E73
         Tag             =   ":cool"
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   255
         Index           =   11
         Left            =   360
         Picture         =   "frmSplash.frx":31F6
         Tag             =   ":baby"
         Top             =   3480
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   8
         Left            =   360
         Picture         =   "frmSplash.frx":3596
         Tag             =   ":smoke"
         Top             =   2760
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   3
         Left            =   360
         Picture         =   "frmSplash.frx":3931
         Tag             =   ":p"
         Top             =   1440
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   7
         Left            =   360
         Picture         =   "frmSplash.frx":3CB8
         Tag             =   ":?"
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   6
         Left            =   360
         Picture         =   "frmSplash.frx":404D
         Tag             =   ":sleep"
         Top             =   2160
         Width           =   420
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   5
         Left            =   360
         Picture         =   "frmSplash.frx":43E1
         Tag             =   ":grr"
         Top             =   1920
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   4
         Left            =   360
         Picture         =   "frmSplash.frx":4771
         Tag             =   ":("
         Top             =   1680
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   2
         Left            =   360
         Picture         =   "frmSplash.frx":4AEF
         Tag             =   ":D"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   1
         Left            =   360
         Picture         =   "frmSplash.frx":4E77
         Tag             =   ";)"
         Top             =   960
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   0
         Left            =   360
         Picture         =   "frmSplash.frx":51F9
         Tag             =   ":)"
         Top             =   720
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   10
         Left            =   360
         Picture         =   "frmSplash.frx":5576
         Tag             =   ":nono"
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label5 
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
         Left            =   720
         TabIndex        =   42
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label6 
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
         Left            =   720
         TabIndex        =   41
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label8 
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
         Left            =   720
         TabIndex        =   40
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label9 
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
         Left            =   720
         TabIndex        =   39
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label10 
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
         Left            =   720
         TabIndex        =   38
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label11 
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
         Left            =   720
         TabIndex        =   37
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label12 
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
         Left            =   840
         TabIndex        =   36
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label13 
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
         Left            =   720
         TabIndex        =   35
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label14 
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
         Left            =   720
         TabIndex        =   34
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label15 
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
         Left            =   720
         TabIndex        =   33
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label16 
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
         Left            =   720
         TabIndex        =   32
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label17 
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
         Left            =   720
         TabIndex        =   31
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label18 
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
         Left            =   720
         TabIndex        =   30
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6420
         TabIndex        =   29
         Top             =   3600
         Width           =   435
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   35
         Left            =   3960
         Picture         =   "frmSplash.frx":5958
         Tag             =   ":toiletclaw"
         Top             =   600
         Width           =   705
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   29
         Left            =   2640
         Picture         =   "frmSplash.frx":5D83
         Tag             =   ":wave"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   18
         Left            =   1440
         Picture         =   "frmSplash.frx":611B
         Tag             =   ":heart"
         Top             =   1320
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   14
         Left            =   5280
         Picture         =   "frmSplash.frx":649E
         Tag             =   ":arnie"
         Top             =   1920
         Width           =   795
      End
      Begin VB.Image imgIcon 
         Height          =   345
         Index           =   32
         Left            =   2640
         Picture         =   "frmSplash.frx":6905
         Tag             =   ":ass"
         Top             =   2040
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   28
         Left            =   2640
         Picture         =   "frmSplash.frx":6CCA
         Tag             =   ":bala"
         Top             =   960
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   26
         Left            =   4080
         Picture         =   "frmSplash.frx":7049
         Tag             =   ":cry"
         Top             =   3600
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   17
         Left            =   1440
         Picture         =   "frmSplash.frx":73D9
         Tag             =   ":devil"
         Top             =   1080
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   19
         Left            =   1440
         Picture         =   "frmSplash.frx":77A1
         Tag             =   ":erm"
         Top             =   1560
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   390
         Index           =   33
         Left            =   2640
         Picture         =   "frmSplash.frx":7B23
         Tag             =   ":flush"
         Top             =   2400
         Width           =   390
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   15
         Left            =   3960
         Picture         =   "frmSplash.frx":7F01
         Tag             =   ":out"
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   31
         Left            =   2640
         Picture         =   "frmSplash.frx":82B9
         Tag             =   ":guitar"
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   20
         Left            =   1440
         Picture         =   "frmSplash.frx":867F
         Tag             =   ":tilt"
         Top             =   1800
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   23
         Left            =   1440
         Picture         =   "frmSplash.frx":8A0D
         Tag             =   ":gg"
         Top             =   2640
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   34
         Left            =   2640
         Picture         =   "frmSplash.frx":8D99
         Tag             =   ":guns"
         Top             =   2880
         Width           =   600
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   30
         Left            =   2640
         Picture         =   "frmSplash.frx":91A4
         Tag             =   ":light"
         Top             =   1440
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   25
         Left            =   3960
         Picture         =   "frmSplash.frx":953D
         Tag             =   ":beat"
         Top             =   1680
         Width           =   525
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   36
         Left            =   5520
         Picture         =   "frmSplash.frx":993F
         Tag             =   ":blah"
         Top             =   960
         Width           =   660
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   21
         Left            =   1440
         Picture         =   "frmSplash.frx":9D0D
         Tag             =   ":lol"
         Top             =   2160
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   22
         Left            =   1440
         Picture         =   "frmSplash.frx":A091
         Tag             =   ":mad"
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   24
         Left            =   1440
         Picture         =   "frmSplash.frx":A417
         Tag             =   ":shit"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   330
         Index           =   27
         Left            =   2640
         Picture         =   "frmSplash.frx":A78D
         Tag             =   ":king"
         Top             =   600
         Width           =   345
      End
      Begin VB.Image imgIcon 
         Height          =   315
         Index           =   37
         Left            =   5400
         Picture         =   "frmSplash.frx":AB41
         Tag             =   ":angel"
         Top             =   2880
         Width           =   600
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   16
         Left            =   1440
         Picture         =   "frmSplash.frx":AF93
         Tag             =   ":alien"
         Top             =   600
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   345
         Index           =   13
         Left            =   5400
         Picture         =   "frmSplash.frx":B33E
         Tag             =   ":shoot"
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label Label20 
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
         Left            =   1800
         TabIndex        =   28
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label21 
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
         Left            =   1800
         TabIndex        =   27
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label22 
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
         Left            =   1800
         TabIndex        =   26
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label23 
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
         Left            =   1800
         TabIndex        =   25
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label24 
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
         Left            =   3240
         TabIndex        =   24
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label25 
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
         Left            =   1800
         TabIndex        =   23
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6420
         TabIndex        =   22
         Top             =   3000
         Width           =   435
      End
      Begin VB.Label Label27 
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
         Left            =   1800
         TabIndex        =   21
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label28 
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
         Left            =   1800
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label29 
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
         Left            =   1800
         TabIndex        =   19
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label30 
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
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label31 
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
         Left            =   3240
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label32 
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
         Left            =   3240
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label33 
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
         Left            =   3240
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label49 
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
         Left            =   3240
         TabIndex        =   14
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label50 
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
         Left            =   3240
         TabIndex        =   13
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label51 
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
         Left            =   3240
         TabIndex        =   12
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label52 
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
         Left            =   3240
         TabIndex        =   11
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label53 
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
         Left            =   4680
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
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
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "(Client)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4560
         TabIndex        =   61
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   4350
         TabIndex        =   60
         Top             =   1060
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E-mail : Corelss@yahoo.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4560
         TabIndex        =   8
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NETWORK CHAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   1990
         TabIndex        =   7
         Top             =   1060
         Width           =   3540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dotcom Infoway"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5700
         TabIndex        =   5
         Top             =   480
         Width           =   1155
      End
      Begin VB.Image imgLogo 
         Height          =   1305
         Left            =   720
         Picture         =   "frmSplash.frx":B84B
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Author : Jayanth Kumar J"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4560
         TabIndex        =   2
         Top             =   3240
         Width           =   2025
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "For Windows 98"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4560
         TabIndex        =   3
         Top             =   2820
         Width           =   2220
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NETWORK CHAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   3540
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "This product is licensed to "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Frame1_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub Frame2_Click()
    Unload Me
    Form1.Show
End Sub
