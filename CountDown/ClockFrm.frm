VERSION 5.00
Begin VB.Form ClockFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
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
   Icon            =   "ClockFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1440
      Top             =   600
   End
   Begin VB.Timer Action 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer DayTimer 
      Interval        =   950
      Left            =   480
      Top             =   600
   End
   Begin VB.Image RArrow1 
      Height          =   240
      Left            =   600
      Picture         =   "ClockFrm.frx":0442
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image RArrow0 
      Height          =   240
      Left            =   600
      Picture         =   "ClockFrm.frx":058C
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image LArrow1 
      Height          =   240
      Left            =   240
      Picture         =   "ClockFrm.frx":06D6
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image LArrow0 
      Height          =   240
      Left            =   240
      Picture         =   "ClockFrm.frx":0820
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image ActImg 
      Height          =   240
      Index           =   3
      Left            =   280
      Picture         =   "ClockFrm.frx":096A
      ToolTipText     =   " About Program "
      Top             =   40
      Width           =   240
   End
   Begin VB.Image ActImg 
      Height          =   240
      Index           =   2
      Left            =   40
      ToolTipText     =   " Bound Sticker to Right "
      Top             =   40
      Width           =   240
   End
   Begin VB.Image ActImg 
      Height          =   240
      Index           =   1
      Left            =   3580
      Picture         =   "ClockFrm.frx":0AB4
      ToolTipText     =   " Program Option "
      Top             =   45
      Width           =   240
   End
   Begin VB.Image ActImg 
      Height          =   240
      Index           =   0
      Left            =   3820
      Picture         =   "ClockFrm.frx":0BFE
      ToolTipText     =   " Exit Program "
      Top             =   45
      Width           =   240
   End
   Begin VB.Image ActImg_ 
      Height          =   240
      Index           =   0
      Left            =   3820
      Picture         =   "ClockFrm.frx":0D48
      Top             =   40
      Width           =   240
   End
   Begin VB.Image ActImg_ 
      Height          =   240
      Index           =   1
      Left            =   3580
      Picture         =   "ClockFrm.frx":0E92
      Top             =   40
      Width           =   240
   End
   Begin VB.Image ActImg_ 
      Height          =   240
      Index           =   2
      Left            =   40
      Top             =   40
      Width           =   240
   End
   Begin VB.Image ActImg_ 
      Height          =   240
      Index           =   3
      Left            =   280
      Picture         =   "ClockFrm.frx":0FDC
      Top             =   40
      Width           =   240
   End
   Begin VB.Label CountDownTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   510
      TabIndex        =   0
      Top             =   30
      Width           =   3070
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   4050
   End
End
Attribute VB_Name = "ClockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActImg_Click(Index As Integer)
    ActImg(Index).Visible = False
    Action.Enabled = True
End Sub

Private Sub Action_Timer()
Dim Index As Integer
    If Not ActImg(0).Visible Then
        ActImg(0).Visible = True
        End
    End If
    If Not ActImg(1).Visible Then
        OptFrm.Show 1
        ActImg(1).Visible = True
    End If
    If Not ActImg(2).Visible Then
        ActImg(2).Visible = True
        If ActImg(2).Picture = RArrow0.Picture Then
            Me.Left = Screen.Width - Me.Width
            If CntDownInfo.Language = 0 Then
                ActImg(2).ToolTipText = " Bound Sticker to Left "
            Else:
                ActImg(2).ToolTipText = " Sticker in linke Ecke " '(L)
            End If
            ActImg(2).Picture = LArrow0.Picture
            ActImg_(2).Picture = LArrow1.Picture
        Else:
            Me.Left = 0
            If CntDownInfo.Language = 0 Then
                ActImg(2).ToolTipText = " Bound Sticker to Right "
            Else:
                ActImg(2).ToolTipText = " Sticker in rechte Ecke " '(R)
            End If
            ActImg(2).Picture = RArrow0.Picture
            ActImg_(2).Picture = RArrow1.Picture
        End If
    End If
    If Not ActImg(3).Visible Then
        AboutFrm.Show
        ActImg(3).Visible = True
    End If
    Action.Enabled = False
End Sub

Private Sub DayTimer_Timer()
    With CountDownTxt
        If DateDiff("d", Now, CntDownInfo.DueDate) > 0 Then
            .Caption = DateDiff("d", Now, CntDownInfo.DueDate) - 1 & " " & DAY_ & " " '*Tage,Days
            .Caption = .Caption & 23 - Hour(Now) & " " & HOUR_ & " " '*Std,Hrs
            .Caption = .Caption & 59 - Minute(Now) & " " & MIN_ & " "  '*Min,Mins
            .Caption = .Caption & 59 - Second(Now) & " " & SEC_ & " " '*Sek,Secs"
        Else:
            .Caption = " 0 " & DAY_ & " 0 " & HOUR_ & " 0 " & MIN_ & " 0 " & SEC_ ' 0 Days 0 Hrs 0 Mins 0 Secs"
        End If
    End With
End Sub

Private Sub Form_Initialize()
    Me.Left = 0
    Me.Top = 0
    Me.Width = Label1.Width + 30
    Me.Height = Label1.Height + 30
    ActImg(2).Picture = RArrow0.Picture
    ActImg_(2).Picture = RArrow1.Picture
    DayTimer_Timer
End Sub


Private Sub Timer1_Timer()
    SetWindowPos Me.hwnd, CntDownInfo.FrmOnTop, 0, 0, 0, 0, &H1 Or &H2
    Timer1.Enabled = False
End Sub
