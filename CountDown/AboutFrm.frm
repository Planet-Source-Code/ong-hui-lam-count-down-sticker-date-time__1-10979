VERSION 5.00
Begin VB.Form AboutFrm 
   BorderStyle     =   0  'None
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   3975
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   4000
         Top             =   2000
      End
      Begin VB.Image ActImg 
         Height          =   200
         Index           =   0
         Left            =   3770
         Stretch         =   -1  'True
         ToolTipText     =   " Exit Program "
         Top             =   140
         Width           =   200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Copyright 2000 by hUilaM -"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   150
         Left            =   1100
         TabIndex        =   6
         Top             =   1920
         Width           =   1620
      End
      Begin VB.Image ActImg 
         Height          =   200
         Index           =   1
         Left            =   3770
         Stretch         =   -1  'True
         ToolTipText     =   " Exit Program "
         Top             =   140
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Label EmailTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "huilam@pl.jaring.my"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1080
         TabIndex        =   5
         Top             =   1050
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CountDown Sticker v2.01"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   240
         Width           =   3510
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3000
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   360
         Picture         =   "AboutFrm.frx":0000
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ Homepage des Autors ]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   1365
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ E-mail des Autors ]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   0
         Left            =   900
         TabIndex        =   2
         Top             =   855
         Width           =   2205
      End
      Begin VB.Label UrlTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://members.xoom.com/huilam/index.htm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   105
         TabIndex        =   1
         Top             =   1560
         Width           =   3810
      End
   End
End
Attribute VB_Name = "AboutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActImg_Click(Index As Integer)
    If Index = 0 Then
        ActImg(1).Visible = True
        Timer1.Enabled = True
    End If
End Sub

Private Sub EmailTxt_Click()
    Shell "Start mailto:huilam@pl.jaring.my", vbHide
End Sub

Private Sub EmailTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not EmailTxt.FontBold Then
        EmailTxt.ForeColor = vbYellow
        EmailTxt.FontBold = True
        UrlTxt.FontBold = False
        UrlTxt.ForeColor = vbWhite
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ActImg_Click (0)
End Sub

Private Sub Form_Load()
    ActImg(0).Picture = ClockFrm.ActImg(0).Picture
    ActImg(1).Picture = ClockFrm.ActImg_(0).Picture
    Image1(1).Picture = Image1(0).Picture
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If EmailTxt.FontBold Then
        EmailTxt.ForeColor = vbWhite
        EmailTxt.FontBold = False
    End If
    If UrlTxt.FontBold Then
        UrlTxt.FontBold = False
        UrlTxt.ForeColor = vbWhite
    End If
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub

Private Sub UrlTxt_Click()
    Shell "Start http://members.xoom.com/huilam/index.htm", vbHide
End Sub

Private Sub UrlTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not UrlTxt.FontBold Then
        UrlTxt.FontBold = True
        UrlTxt.ForeColor = vbYellow
        EmailTxt.ForeColor = vbWhite
        EmailTxt.FontBold = False
    End If
End Sub
