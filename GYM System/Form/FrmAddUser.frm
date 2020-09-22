VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3CA40DFE-4DED-4BD9-98FD-6BEEE7B69F2A}#24.0#0"; "PDTPicker.ocx"
Begin VB.Form FrmAddUser 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«›“Êœ‰ ò«—»—"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   7140
   Begin VB.CheckBox ChkDaily 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ò«—»— —Ê“«‰Â"
      Height          =   375
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CheckBox ChkBimeh 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "»Ì„Â Ê—“‘Ì"
      Height          =   375
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5640
      Width           =   3135
   End
   Begin VB.CheckBox ChkPay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "‘Â—ÌÂ Å—œ«Œ  ‘œ"
      Height          =   375
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5040
      Width           =   5175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox TxtID 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin GYM.XPStyle XPStyle1 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin MSAdodcLib.Adodc AdoAdd 
      Height          =   375
      Left            =   4680
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Œ—ÊÃ"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8388608
      cFHover         =   8388608
      cBhover         =   14846764
      cGradient       =   14846764
      Gradient        =   3
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog ComBr 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Ã” ÃÊÌ ⁄ò” ò«—»—"
      Filter          =   "*.jpg"
      InitDir         =   "/"
      MaxFileSize     =   500
   End
   Begin PDTPicker.FDTPicker TxtBirthDate 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Text            =   "1387/05/26"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      YearRange       =   100
      RightToLeft     =   -1  'True
      YearRange       =   100
   End
   Begin GYM.lvButtons_H CmdAdd 
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "À» "
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8388608
      cFHover         =   8388608
      cBhover         =   14846764
      cGradient       =   14846764
      Gradient        =   3
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin PDTPicker.FDTPicker TxtDateReg 
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   6840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Text            =   "1387/05/26"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      YearRange       =   100
      RightToLeft     =   -1  'True
      YearRange       =   100
   End
   Begin GYM.lvButtons_H CmdReAdd 
      Height          =   495
      Left            =   4200
      TabIndex        =   28
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "À»  œÊ»«—Â"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8388608
      cFHover         =   8388608
      cBhover         =   14846764
      cGradient       =   14846764
      Gradient        =   3
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin GYM.lvButtons_H CmdPic 
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Ã” ÃÊÌ ⁄ò”"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8388608
      cFHover         =   8388608
      cBhover         =   14846764
      cGradient       =   14846764
      Gradient        =   3
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ ⁄÷ÊÌ  :"
      Height          =   240
      Index           =   9
      Left            =   5730
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   6900
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ „«Â Â« :"
      Height          =   240
      Index           =   7
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¬œ—” :"
      Height          =   240
      Index           =   6
      Left            =   6405
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   5160
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ :"
      Height          =   240
      Index           =   5
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ  Ê·œ :"
      Height          =   240
      Index           =   8
      Left            =   6090
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4020
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ‘‰«”‰«„Â :"
      Height          =   240
      Index           =   4
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3480
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘€· :"
      Height          =   240
      Index           =   3
      Left            =   6435
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "›«„Ì·Ì :"
      Height          =   240
      Index           =   2
      Left            =   6270
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ :"
      Height          =   240
      Index           =   0
      Left            =   6675
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1680
      Width           =   390
   End
   Begin VB.Image ClientImg 
      Height          =   3375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "⁄ò” ò«—»—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2400
      Width           =   2745
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«›“Êœ‰ ò«—»— ÃœÌœ »Â »«‘ê«Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   2235
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmAddUser.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmAddUser.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7160
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmAddUser.frx":3D6A
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   7140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ò«—»—Ì :"
      Height          =   240
      Index           =   1
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "FrmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkBimeh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then ChkDaily.SetFocus
End Sub

Private Sub ChkDaily_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then TxtDateReg.SetFocus
End Sub

Private Sub ChkPay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    TxtAdd(6).SetFocus
    SendKeys HiLyt
End If
End Sub

Private Sub CmdAdd_Click()
Dim I As Integer
Dim UserID As Long
For I = 0 To 5
    If Trim(TxtAdd(I).Text) = "" Then
        TxtAdd(I).Text = "--"
    End If
    DoEvents
Next
If Trim(TxtAdd(6).Text) = "" Or IsNumeric(TxtAdd(6).Text) = False Then
    If ChkPay.Value = 0 Then
        TxtAdd(6).Text = 0
    Else
        TxtAdd(6).Text = 1
    End If
End If
If ChkPay.Value = 0 Then
    If Val(TxtAdd(6).Text) > 0 Then
        TxtAdd(6).Text = 0
    End If
Else
    If Val(TxtAdd(6).Text) <= 0 Then
        TxtAdd(6).Text = 1
    End If
End If
If MsgBox("¬Ì« „«Ì· »Â À»  «Ì‰ ò«—»— Â” Ìœ ø", vbQuestion + vbYesNo, "À»  ò«—»—") = vbNo Then Exit Sub
Call ConnectToDb(AdoAdd, "[User]", False)
With AdoAdd
    .Refresh
    .Recordset.AddNew
    .Recordset.Fields("ID") = UserID
    .Recordset.Fields("Name") = Trim(TxtAdd(0).Text)
    .Recordset.Fields("Family") = Trim(TxtAdd(1).Text)
    .Recordset.Fields("ShomareSh") = Trim(TxtAdd(3).Text)
    .Recordset.Fields("Job") = Trim(TxtAdd(2).Text)
    .Recordset.Fields("BirthDay_S") = TxtBirthDate.Text
    .Recordset.Fields("BirthDay_M") = Shamsi.Convert_Date(TxtBirthDate.Text, HijriShamsi_, Gregorian_)
    .Recordset.Fields("DateReg_S") = TxtDateReg.Text
    .Recordset.Fields("DateReg_M") = Shamsi.Convert_Date(TxtDateReg.Text, HijriShamsi_, Gregorian_)
    .Recordset.Fields("Address") = Trim(TxtAdd(5).Text)
    .Recordset.Fields("Tell") = Trim(TxtAdd(4).Text)
    .Recordset.Fields("Bime") = ChkBimeh.Value
    .Recordset.Fields("Daily") = ChkDaily.Value
    .Recordset.Fields("Active") = 1
    .Recordset.Update
    .Refresh
    .Recordset.MoveLast
    TxtID.Text = .Recordset.Fields("ID")
End With
If ComBr.FileName <> "" Then
    FileCopy ComBr.FileName, App.Path & "\UserPic\" & TxtID.Text
End If
If ChkPay.Value = 1 Then
    Call ConnectToDb(AdoAdd, "Pay", False)
    Dim DateG As String
    Dim MDate As Date
    DateG = TxtDateReg.Text
    DateG = TxtDateReg.Text
    MDate = Shamsi.Convert_Date(DateG, HijriShamsi_, Gregorian_)
    For I = 0 To Val(TxtAdd(6).Text) - 1
        With AdoAdd
            .Recordset.AddNew
            .Recordset.Fields("UID") = UserID
            DateG = Shamsi.IncreaseDate_Custom(TxtDateReg.Text, I * 30)
            MDate = MDate + I * 30
            .Recordset.Fields("DateGive_S") = DateG
            .Recordset.Fields("DateGive_M") = MDate
            .Recordset.Fields("TimeGive") = Time
            .Recordset.Update
        End With
        DoEvents
    Next
End If
MsgBox "«Ì‰ ò«—»— »Â »«‘ê«Â «÷«›Â ‘œ", vbInformation, "«÷«›Â ‘œ"
CmdAdd.Enabled = False
AdoAdd.Recordset.Close
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdPic_Click()
On Error Resume Next
With ComBr
    .CancelError = False
    .ShowOpen
    If .FileName <> "" Then
        ClientImg.Picture = LoadPicture(.FileName)
        Label2.Visible = False
    End If
End With
End Sub

Private Sub CmdReAdd_Click()
Unload Me
FrmAddUser.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
CenterFrm FrmMain, Me
TxtBirthDate.Text = Shamsi.Convert_Date(Date, Gregorian_, HijriShamsi_)
TxtDateReg.Text = TxtBirthDate.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnimateForm Me, -1, -1, aUnload, 5, 5, 3, 13
End Sub

Private Sub TxtAdd_GotFocus(Index As Integer)
GotColor TxtAdd(Index)
End Sub

Private Sub TxtAdd_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 5 Then
        ChkPay.SetFocus
    ElseIf Index = 6 Then
        ChkBimeh.SetFocus
    ElseIf Index = 3 Then
        TxtBirthDate.SetFocus
    Else
        TxtAdd(Index + 1).SetFocus
        SendKeys HiLyt
    End If
End If
End Sub

Private Sub TxtAdd_LostFocus(Index As Integer)
LostColor TxtAdd(Index)
End Sub

Private Sub TxtDateReg_GotFocus()
GotColor TxtDateReg
End Sub

Private Sub TxtDateReg_LostFocus()
LostColor TxtDateReg
End Sub
