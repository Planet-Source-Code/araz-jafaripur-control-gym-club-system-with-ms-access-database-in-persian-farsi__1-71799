VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3CA40DFE-4DED-4BD9-98FD-6BEEE7B69F2A}#24.0#0"; "PDTPicker.ocx"
Begin VB.Form FrmEnterUser 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê—Êœ ò«—»—"
   ClientHeight    =   8100
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
   ScaleHeight     =   8100
   ScaleWidth      =   7140
   Begin VB.CheckBox ChkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ò«—»— ›⁄«·"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox TxtCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox TxtID 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox TxtAdd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5640
      Width           =   5175
   End
   Begin VB.CheckBox ChkBimeh 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "»Ì„Â Ê—“‘Ì"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CheckBox ChkDaily 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ò«—»— —Ê“«‰Â"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6360
      Width           =   2175
   End
   Begin GYM.XPStyle XPStyle1 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin MSAdodcLib.Adodc AdoEnter 
      Height          =   375
      Left            =   4920
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
      Left            =   4200
      TabIndex        =   14
      Top             =   7560
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
   Begin PDTPicker.FDTPicker TxtBirthDate 
      Height          =   315
      Left            =   3120
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Enabled         =   0   'False
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
      BackColor       =   14737632
      RightToLeft     =   -1  'True
      YearRange       =   100
   End
   Begin GYM.lvButtons_H CmdEnter 
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Ê—Êœ"
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
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin PDTPicker.FDTPicker TxtDateReg 
      Height          =   315
      Left            =   3120
      TabIndex        =   12
      Top             =   6960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Enabled         =   0   'False
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
      BackColor       =   14737632
      RightToLeft     =   -1  'True
      YearRange       =   100
   End
   Begin MSAdodcLib.Adodc AdoAddCode 
      Height          =   375
      Left            =   360
      Top             =   5040
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
   Begin VB.Label LblPay 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2955
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4920
      Width           =   60
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ò„œ :"
      Height          =   240
      Index           =   10
      Left            =   5910
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmEnterUser.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê—Êœ ò«—»—«‰ »Â »«‘ê«Â"
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
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   360
      Width           =   1770
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
      TabIndex        =   24
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmEnterUser.frx":0CCA
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   7140
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmEnterUser.frx":3651
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7155
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
      TabIndex        =   23
      Top             =   2400
      Width           =   2745
   End
   Begin VB.Image ClientImg 
      Height          =   3375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2895
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
      TabIndex        =   22
      Top             =   2280
      Width           =   390
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
      TabIndex        =   21
      Top             =   2880
      Width           =   795
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
      TabIndex        =   20
      Top             =   3480
      Width           =   630
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
      TabIndex        =   19
      Top             =   4080
      Width           =   1755
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
      TabIndex        =   18
      Top             =   4620
      Width           =   945
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
      TabIndex        =   17
      Top             =   5160
      Width           =   540
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
      TabIndex        =   16
      Top             =   5760
      Width           =   660
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
      TabIndex        =   15
      Top             =   7020
      Width           =   1305
   End
End
Attribute VB_Name = "FrmEnterUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserID As Long
Private Sub CmdEnter_Click()
If IsNumeric(Trim(TxtCode.Text)) = False Then
    MsgBox "‘„«—Â ò„œ œ—”  ‰„Ì »«‘œ", vbExclamation, "‘„«Â ò„œ"
    TxtCode.SetFocus
    Exit Sub
End If
Dim CheckCode As Integer
CheckCode = EnterCode(UserID, Val(TxtCode.Text))
If CheckCode = -1 Then
    MsgBox "«Ì‰ ò«—»— «“ »«‘ê«Â Œ«—Ã ‰‘œÂ «” ", vbExclamation, "Ê—Êœ ò«—»—"
    Exit Sub
ElseIf CheckCode = -2 Then
    MsgBox "«Ì‰ ‘„«—Â ò„œ »Â ò«—»—Ì œÌê— „ ⁄·ﬁ «” ", vbExclamation, "Ê—Êœ ò«—»—"
    Exit Sub
End If
If MsgBox("¬Ì« „«Ì· »Â Ê—Êœ «Ì‰ ò«—»— Â” Ìœ ø", vbYesNo + vbQuestion, "Ê—Êœ ò«—»—") = vbNo Then
    Exit Sub
End If
Call ConnectToDb(AdoAddCode, "UserLogin", False)
With AdoAddCode.Recordset
    .AddNew
    .Fields("UID") = UserID
    .Fields("KNum") = Val(TxtCode.Text)
    .Fields("GiveNum") = 0
    .Fields("InDate_S") = Shamsi.Convert_Date(Date, Gregorian_, HijriShamsi_)
    .Fields("InDate_M") = Date
    .Fields("InTime") = Time
    .Update
End With
MsgBox "ò«—»— Ê«—œ »«‘ê«Â ‘œ", vbInformation, "Ê—Êœ ò«—»—"
Unload Me
FrmEnterUser.Show
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
CenterFrm FrmMain, Me
Call ConnectToDb(AdoEnter, "[User]", False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnimateForm Me, -1, -1, aUnload, 5, 5, 3, 13
End Sub

Private Sub TxtCode_GotFocus()
GotColor TxtCode
End Sub
Private Sub TxtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdEnter_Click
End Sub

Private Sub TxtCode_LostFocus()
LostColor TxtCode
End Sub

Private Sub TxtID_GotFocus()
GotColor TxtID
End Sub

Private Sub TxtID_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_Control
If KeyCode = 13 And IsNumeric(TxtID.Text) = True Then
    Dim StrPic As String
    Dim I As Integer
    With AdoEnter.Recordset
        AdoEnter.Refresh
        .Find ("ID = " & Val(TxtID.Text))
        TxtAdd(0).Text = .Fields("Name")
        TxtAdd(1).Text = .Fields("Family")
        TxtAdd(2).Text = .Fields("Job")
        TxtAdd(3).Text = .Fields("ShomareSh")
        TxtAdd(4).Text = .Fields("Tell")
        TxtAdd(5).Text = .Fields("Address")
        ChkBimeh.Value = .Fields("Bime")
        ChkDaily.Value = .Fields("Daily")
        ChkActive.Value = .Fields("Active")
        TxtBirthDate.Text = .Fields("BirthDay_S")
        TxtDateReg.Text = .Fields("DateReg_S")
        TxtCode.Enabled = True
        TxtCode.BackColor = vbWhite
        TxtCode.SetFocus
        StrPic = App.Path & "\UserPic\" & TxtID.Text
        If FileExists(StrPic) Then
            ClientImg.Picture = LoadPicture(StrPic)
            Label2.Visible = False
        Else
            Label2.Visible = True
            ClientImg.Picture = LoadPicture("")
        End If
        CmdEnter.Enabled = True
    End With
    UserID = Val(TxtID.Text)
    I = GetPay(Val(TxtID.Text))
    Select Case I
        Case Is = 0
            LblPay.Caption = "«„—Ê“ „Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ «” "
        Case Is = -1
            LblPay.Caption = "„Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ ‰—”ÌœÂ «” "
        Case Is = 1
            LblPay.Caption = "„Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ ê–‘ Â «” "
    End Select
End If
Exit Sub
ERR_Control:
    MsgBox "ò«—»—Ì »« «Ì‰ òœ ò«—»—Ì ÊÃÊœ ‰œ«—œ", vbExclamation, "ò«—»—"
    ClearAll
End Sub

Private Sub TxtID_LostFocus()
LostColor TxtID
End Sub
Private Sub ClearAll()
Dim I As Integer
For I = 0 To 5
    TxtAdd(I).Text = ""
    DoEvents
Next
TxtCode.Enabled = False
TxtCode.BackColor = &HE0E0E0
ClientImg.Picture = LoadPicture("")
LblPay.Caption = ""
UserID = 0
TxtCode.Text = ""
Label2.Visible = True
CmdEnter.Enabled = False
ChkBimeh.Value = 0
ChkDaily.Value = 0
TxtID.SetFocus
SendKeys HiLyt
End Sub
