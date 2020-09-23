VERSION 5.00
Begin VB.Form OptFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Option"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   Icon            =   "OptFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton LangOpt 
      Caption         =   "German"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton LangOpt 
      Caption         =   "English"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox OnTopChk 
      Caption         =   " Sticker Always On Top "
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
      Left            =   240
      TabIndex        =   2
      Top             =   400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply Setting"
      Height          =   320
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   2420
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   2420
      Begin VB.ComboBox YearTxt 
         Height          =   315
         ItemData        =   "OptFrm.frx":0442
         Left            =   1560
         List            =   "OptFrm.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   220
         Width           =   750
      End
      Begin VB.ComboBox DayTxt 
         Height          =   315
         ItemData        =   "OptFrm.frx":0446
         Left            =   120
         List            =   "OptFrm.frx":04A7
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   220
         Width           =   615
      End
      Begin VB.ComboBox MonTxt 
         Height          =   315
         ItemData        =   "OptFrm.frx":0527
         Left            =   780
         List            =   "OptFrm.frx":054F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   220
         Width           =   735
      End
   End
End
Attribute VB_Name = "OptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If IsDate(Format(DayTxt & "/" & MonTxt.ListIndex + 1 & "/" & YearTxt, "dd/mm/yyyy")) Then
        With CntDownInfo
            If OnTopChk.Value = 1 Then .FrmOnTop = -1 Else .FrmOnTop = -2
            If LangOpt(0).Value Then .Language = 0
            If LangOpt(1).Value Then .Language = 1
                
            SetWindowPos ClockFrm.hwnd, .FrmOnTop, 0, 0, 0, 0, &H1 Or &H2
            .DueDate = Format(DayTxt & "/" & MonTxt.ListIndex + 1 & "/" & YearTxt, "dd/mm/yyyy")
                
            Open "CntDown.dat" For Binary Access Write As #1 Len = Len(CntDownInfo)
            Put #1, 1, CntDownInfo
            Close #1
                
            Unload Me
        End With
    Else:
        MsgBox "Invalid Date Detected !", vbOKOnly + vbCritical, "Not a valid date"
        DayTxt.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Command1_Click
End Sub

Private Sub Form_Load()
Dim AddCnt As Integer
    OptFrm.YearTxt.Clear
    For AddCnt = 0 To 50
        OptFrm.YearTxt.AddItem "20" & Format(AddCnt, "00")
    Next AddCnt
    With CntDownInfo
        LangOpt(.Language).Value = True
        If .FrmOnTop = -1 Then OnTopChk.Value = 1 Else OnTopChk.Value = 0
        DayTxt.ListIndex = Day(.DueDate) - 1
        MonTxt.ListIndex = Month(.DueDate) - 1
        YearTxt = Year(.DueDate)
    End With
End Sub

Private Sub LangOpt_Click(Index As Integer)
    SetLang (Index)
    With ClockFrm
        If Not .DayTimer.Enabled Then
            CountDownTxt.Caption = " 0 " & DAY_ & " 0 " & HOUR_ & " 0 " & MIN_ & " 0 " & SEC_
        End If
    End With
End Sub
