Attribute VB_Name = "F_Func"
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Enum AnimeEventEnum
    aUnload = 0
    aload = 1
End Enum
Public Sub Main()
If App.PrevInstance Then
    MsgBox "ÈÑäÇãå ÏÑ ÍÇá ÇÌÑÇ ãí ÈÇÔÏ", vbExclamation, "ÏÑ ÍÇá ÇÌÑÇ"
    End
End If
ComMDB
FrmLogin.Show
ChangeLang
End Sub
Public Sub ChangeLang()
LoadKeyboardLayout "00000429", 1
End Sub
Private Sub ComMDB()
On Local Error Resume Next
Dim dbsfle, nme, Pass As String
Pass = ";PWD=swordofgrandlord"
dbsfle = App.Path & "\Data\Data.Mj"
nme = Mid(dbsfle, InStrRev(dbsfle, "\") + 1)
nme = Left(nme, Len(nme) - 4)
If Dir(App.Path & "\" & nme & ".CPT") <> "" Then
    Kill App.Path & "\" & nme & ".CPT"
End If
restart:
DBEngine.CompactDatabase dbsfle, App.Path & "\" & nme & ".CPT", , , Pass
If Dir(App.Path & "\" & nme & ".OLD") <> "" Then
    Kill App.Path & "\" & nme & ".OLD"
End If
Name dbsfle As App.Path & "\" & nme & ".OLD"
DBEngine.CompactDatabase App.Path & "\" & nme & ".CPT", dbsfle, , , Pass
Kill App.Path & "\Dat.OLD"
Kill App.Path & "\Dat.CPT"
End Sub
Public Function lv_TimerCallBack(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tgtButton As lvButtons_H
CopyMemory tgtButton, GetProp(hwnd, "lv_ClassID"), &H4
Call tgtButton.TimerUpdate(GetProp(hwnd, "lv_TimerID"))
CopyMemory tgtButton, 0&, &H4
End Function

Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form)
    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.ScaleHeight \ 2) - (Childfrm.Height \ 2)
    AnimateForm Childfrm, -1, -1, aload, 5, 5, 3, 13
End Sub
Public Sub AnimateForm(frm As Form, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1, Optional aEvent As AnimeEventEnum, Optional ByVal TrailCount As Long = 0, _
                            Optional ByVal FrameTime As Long = 3, Optional ByVal BorderWidth As Long = 2, Optional ByVal FrameCount As Long = 25, Optional BorderColor As Long = 0)
On Local Error Resume Next

Dim pic As Control
Static MousePos As POINTAPI
Dim ScrX As Long, ScrY As Long
Dim XValue As Long, YValue As Long
Dim XIncr As Double, YIncr As Double
Dim WIncr As Double, HIncr As Double
Dim hBrush As Long, ColIncr As Double
Dim X1 As Long, Y1 As Long, iNow As Long
Dim FrmRgn As Long, hrgn1 As Long, hrgn2 As Long

    Set pic = frm.Controls.Add("vb.picturebox", "PicDraw"): pic.BorderStyle = 0
    SetParent pic.hwnd, GetDesktopWindow: pic.Move 0, 0, Screen.Width, Screen.Height
    If TrailCount > FrameCount Then TrailCount = FrameCount
    pic.BackColor = BorderColor: ColIncr = 200 / (TrailCount + 1)
    ScrX = Screen.TwipsPerPixelX: ScrY = Screen.TwipsPerPixelY
    If aEvent = aload Then If (X = -1 Or Y = -1) Then GetCursorPos MousePos Else MousePos.X = X: MousePos.Y = Y
    XIncr = (frm.Left / ScrX - MousePos.X) / FrameCount
    YIncr = (frm.Top / ScrY - MousePos.Y) / FrameCount
    WIncr = frm.Width / ScrX / FrameCount
    HIncr = frm.Height / ScrY / FrameCount

    For X1 = 0 To FrameCount
        FrmRgn = CreateRectRgn(0, 0, 0, 0): pic.Visible = True
        For Y1 = 0 To TrailCount
            If aEvent = aload Then iNow = X1 - Y1 Else iNow = FrameCount - X1 + Y1
            If iNow >= FrameCount Or iNow <= 0 Then Y1 = TrailCount
            XValue = MousePos.X + iNow * XIncr: YValue = MousePos.Y + iNow * YIncr
            hrgn1 = CreateRectRgn(XValue, YValue, XValue + iNow * WIncr, YValue + iNow * HIncr)
            hrgn2 = CreateRectRgn(XValue - BorderWidth, YValue - BorderWidth, XValue + iNow * WIncr + BorderWidth, YValue + iNow * HIncr + BorderWidth)
            CombineRgn hrgn1, hrgn1, hrgn2, 3
            hBrush = CreateSolidBrush(RGB(Y1 * ColIncr, Y1 * ColIncr, Y1 * ColIncr))
            FillRgn pic.hDC, hrgn1, hBrush
            CombineRgn FrmRgn, hrgn1, FrmRgn, 2
            DeleteObject hrgn1: DeleteObject hrgn2: DeleteObject hBrush
        Next Y1
        SetWindowRgn pic.hwnd, FrmRgn, True: DoEvents
        Sleep FrameTime
    Next X1
    Call frm.Controls.Remove("PicDraw"): Set pic = Nothing
End Sub

Function Encrypt(Start As Integer, diff As Integer, beta As Integer, Alpha As Integer, times As Integer, SuperEncrypt As Boolean, ByVal Text As String)
On Error GoTo error
Dim I As Integer
Dim curkey As Long
Dim m As Long
Dim endstr As String
Dim Text2 As String
Dim lesser As Double
Dim larger As Double
Dim SuperE As Boolean
Dim A As Integer
SuperE = SuperEncrypt
If diff > 500 Then
    diff = 500
ElseIf diff < 1 Then
    diff = 1
End If
If times > 100 Then
    times = 100
ElseIf times < 1 Then
    times = 1
End If
If Start > 255 Then
    Start = 255
ElseIf Start < 1 Then
    Start = 1
End If
If beta > 5 Then
    beta = 5
ElseIf beta < 1 Then
    beta = 1
End If
If Alpha > 5 Then
    Alpha = 5
ElseIf Alpha < 1 Then
    Alpha = 1
End If
curkey = Start
curkey = (curkey * Alpha) / beta
If SuperE = True Then
    If curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) < 1 Then
        curkey = (((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) * (0 - 1))
    Else
        curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10))
    End If
    curkey = SuperEE(curkey, beta, Alpha, beta)
End If
If curkey > 255 Then
    curkey = 255 - (curkey / 255)
ElseIf curkey < 0 Then
    curkey = 0 - (curkey / 255)
End If
For A = 1 To times
    For I = 1 To Len(Text)
        If 255 - curkey > curkey Then
            larger = 255 - curkey
            lesser = curkey
        Else
            larger = curkey
            lesser = 255 - curkey
        End If
        If Asc(Mid$(Text, I, 1)) <= lesser Then
            m = Asc(Mid$(Text, I, 1)) + (larger - 1)
            endstr = endstr + Chr$(m)
        Else
            m = Asc(Mid$(Text, I, 1)) - lesser
            endstr = endstr + Chr$(m)
        End If
        curkey = curkey + diff
        If curkey > 255 Then
            curkey = curkey - 255
        End If
        curkey = (curkey * Alpha) / beta
        If SuperE = True Then
            If curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) < 1 Then
                curkey = (((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) * (0 - 1))
            Else
                curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10))
            End If
            curkey = SuperEE(curkey, beta, Alpha, beta)
        End If
        beta = beta + (2 * diff)
        Alpha = Alpha + diff
        If beta > 5 Then
            beta = 1
        End If
        If Alpha > 5 Then
            Alpha = 1
        End If
        If curkey > 255 Then
            curkey = 255 - (curkey / 255)
        ElseIf curkey < 0 Then
            curkey = 0 - (curkey / 255)
        End If
        If diff > 500 Then
            diff = 1
        Else
            diff = diff + diff
        End If
    Next I
    Text2 = ""
    Text2 = endstr
    endstr = ""
Next A
Encrypt = Text2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
'=========================== Decrypt Text =======================
Function Decrypt(Start As Integer, diff As Integer, beta As Integer, Alpha As Integer, times As Integer, SuperEncrypt As Boolean, Text As String)
On Error GoTo error
Dim I As Integer
Dim curkey As Long
Dim m As Long
Dim endstr As String
Dim Text2 As String
Dim lesser As Double
Dim larger As Double
Dim SuperE As Boolean
Dim A As Integer
SuperE = SuperEncrypt
If diff > 500 Then
    diff = 500
ElseIf diff < 1 Then
    diff = 1
End If
If times > 100 Then
    times = 100
ElseIf times < 1 Then
    times = 1
End If
If Start > 255 Then
    Start = 255
ElseIf Start < 1 Then
    Start = 1
End If
If beta > 5 Then
    beta = 5
ElseIf beta < 1 Then
    beta = 1
End If
If Alpha > 5 Then
    Alpha = 5
ElseIf Alpha < 1 Then
    Alpha = 1
End If
curkey = Start
curkey = (curkey * Alpha) / beta
If SuperE = True Then
    If curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) < 1 Then
        curkey = (((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) * (0 - 1))
    Else
        curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10))
    End If
    curkey = SuperEE(curkey, beta, Alpha, beta)
End If
If curkey > 255 Then
    curkey = 255 - (curkey / 255)
ElseIf curkey < 0 Then
    curkey = 0 - (curkey / 255)
End If
For A = 1 To times
    For I = 1 To Len(Text)
        If 255 - curkey > curkey Then
            larger = 255 - curkey
            lesser = curkey
        Else
            larger = curkey
            lesser = 255 - curkey
        End If
        If Asc(Mid$(Text, I, 1)) >= larger Then
            m = Asc(Mid$(Text, I, 1)) - (larger - 1)
            endstr = endstr + Chr$(m)
        Else
            m = Asc(Mid$(Text, I, 1)) + lesser
            endstr = endstr + Chr$(m)
        End If
        curkey = curkey + diff
        If curkey > 255 Then
            curkey = curkey - 255
        End If
        curkey = (curkey * Alpha) / beta
        If SuperE = True Then
            If curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) < 1 Then
                curkey = (((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) * (0 - 1))
            Else
                curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10))
            End If
            curkey = SuperEE(curkey, beta, Alpha, beta)
        End If
        beta = beta + (2 * diff)
        Alpha = Alpha + diff
        If beta > 5 Then
            beta = 1
        End If
        If Alpha > 5 Then
            Alpha = 1
        End If
        If curkey > 255 Then
            curkey = 255 - (curkey / 255)
        ElseIf curkey < 0 Then
            curkey = 0 - (curkey / 255)
        End If
        If diff > 500 Then
            diff = 1
        Else
            diff = diff + diff
        End If
    Next I
    Text2 = ""
    Text2 = endstr
    endstr = ""
Next A
Decrypt = Text2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
'=========================== Private Function For Use =======================
Private Function SuperEE(curkey As Long, beta As Integer, Alpha As Integer, times As Integer)
On Error GoTo error
curkey = (((curkey / times) - (beta + times)) * Alpha) + ((beta / Alpha) - times)
If curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10)) * (0 - 1))
Else
    curkey = ((curkey + beta) * Alpha) - (((curkey - beta) + Alpha) / ((beta - Alpha) + 10))
End If
If beta - times = 0 Then
    curkey = ((curkey * Alpha) + (beta * times))
Else
    curkey = ((curkey * (beta - times)) + (beta - times))
    If curkey < 0 Then
        curkey = curkey + (Alpha + beta)
    ElseIf curkey = 0 Then
        curkey = curkey + (Alpha + times)
    Else
        curkey = curkey + (beta + times)
    End If
End If
SuperEE = curkey
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
    Open FullFileName For Input As #1
    Close #1
    FileExists = True
    Exit Function
MakeF:
    FileExists = False
End Function

Public Sub GotColor(TextBoxName As Object)
TextBoxName.BackColor = &H80000018
SendKeys HiLyt
End Sub
Public Sub LostColor(TextBoxName As Object)
TextBoxName.BackColor = vbWhite
End Sub
