Attribute VB_Name = "MainSub"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Const HWND_TOPMOST = -1
'Public Const HWND_NOTOPMOST = -2
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1

Type CntDownRec
    DueDate As Date
    FrmOnTop As Long
    Language As Byte
End Type

Public DAY_, HOUR_, MIN_, SEC_ As String
Public CntDownInfo As CntDownRec

Public Sub SetLang(Index As Integer)
Dim TmpMon As Integer
    TmpMon = OptFrm.MonTxt.ListIndex
    Select Case Index
        Case 0: DAY_ = "Days"
                HOUR_ = "Hrs"
                MIN_ = "Mins"
                SEC_ = "Secs"
                AboutFrm.Label1(0) = "[ Author Email ]"
                AboutFrm.Label1(1) = "[ Author Homepage ]"
                With OptFrm
                    .Caption = " Program Option"
                    .OnTopChk.Left = 240
                    .OnTopChk.Caption = " Sticker Always On Top"
                    .Command1.Caption = "Apply Setting"
                    .MonTxt.Clear
                    .MonTxt.AddItem "Jan"
                    .MonTxt.AddItem "Feb"
                    .MonTxt.AddItem "Mar"
                    .MonTxt.AddItem "Apr"
                    .MonTxt.AddItem "May"
                    .MonTxt.AddItem "Jun"
                    .MonTxt.AddItem "Jul"
                    .MonTxt.AddItem "Aug"
                    .MonTxt.AddItem "Sep"
                    .MonTxt.AddItem "Oct"
                    .MonTxt.AddItem "Nov"
                    .MonTxt.AddItem "Dec"
                End With
                ClockFrm.ActImg(0).ToolTipText = " Exit Program "
                ClockFrm.ActImg(1).ToolTipText = " Program Option"
                ClockFrm.ActImg(2).ToolTipText = " Bound Sticker to Right "
                ClockFrm.ActImg(3).ToolTipText = " About Program "
                
        Case 1: DAY_ = "Tage"
                HOUR_ = "Std"
                MIN_ = "Min"
                SEC_ = "Sek"
                AboutFrm.Label1(0) = "[ Email des Autors ]"
                AboutFrm.Label1(1) = "[ Homepage des Autors ]"
                With OptFrm
                    .Caption = " Optionen"
                    .OnTopChk.Left = 600
                    .OnTopChk.Caption = " Immer oben "
                    .Command1.Caption = "Einstellungen übernehmen"
                    .MonTxt.Clear
                    .MonTxt.AddItem "Jan"
                    .MonTxt.AddItem "Feb"
                    .MonTxt.AddItem "Mär"
                    .MonTxt.AddItem "Apr"
                    .MonTxt.AddItem "Mai"
                    .MonTxt.AddItem "Jun"
                    .MonTxt.AddItem "Jul"
                    .MonTxt.AddItem "Aug"
                    .MonTxt.AddItem "Sep"
                    .MonTxt.AddItem "Okt"
                    .MonTxt.AddItem "Nov"
                    .MonTxt.AddItem "Dez"
                End With
                ClockFrm.ActImg(0).ToolTipText = " Beenden "
                ClockFrm.ActImg(1).ToolTipText = " Optionen "
                ClockFrm.ActImg(2).ToolTipText = " Sticker in rechte Ecke "
                ClockFrm.ActImg(3).ToolTipText = " Info "
    End Select
    OptFrm.MonTxt.ListIndex = TmpMon
End Sub


Sub Main()
    
    If Dir("CntDown.dat") <> "" Then If FileLen("cntdown.dat") <> Len(CntDownInfo) Then Kill "cntdown.dat"
    If Dir("CntDown.dat") <> "" Then
        Open "CntDown.dat" For Binary Access Read As #1 Len = Len(CntDownInfo)
        Get #1, 1, CntDownInfo
        Close #1
        SetLang (CntDownInfo.Language)
        ClockFrm.Show
    Else:
        CntDownInfo.DueDate = Format(Now, "dd/mm/yyyy")
        CntDownInfo.FrmOnTop = -1
        CntDownInfo.Language = 0
        SetLang (0)

        ClockFrm.Show
        OptFrm.Show 1
    End If
    
End Sub

