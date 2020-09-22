VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmEdit 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊÌ—«Ì‘ „œÌ—"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
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
   ScaleHeight     =   9075
   ScaleWidth      =   4080
   Begin VB.TextBox TxtUserName 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ComboBox CboUserName 
      Height          =   360
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CheckBox ChkAddUser 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "«›“Êœ‰ ò«—»—"
      Height          =   255
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CheckBox ChkEditUser 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ÊÌ—«Ì‘ ò«—»—"
      Height          =   255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CheckBox ChkEnter 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "Ê—Êœ ò«—»—"
      Height          =   255
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox ChkOut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "Œ—ÊÃ ò«—»—"
      Height          =   255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox ChkPay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "Å—œ«Œ  ‘Â—ÌÂ"
      Height          =   255
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CheckBox ChkDelete 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "Õ–› „œÌ—"
      Height          =   255
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CheckBox ChkEditAdmin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ÊÌ—«Ì‘ „œÌ—"
      Height          =   255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CheckBox ChkAddAdmin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "«›“Êœ‰ „œÌ—"
      Height          =   255
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CheckBox ChkROut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ê“«—‘ Ê—Êœ Ê Œ—ÊÃ"
      Height          =   255
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CheckBox ChkRUser 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ê“«—‘ ·Ì”  ò«—»—«‰"
      Height          =   255
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CheckBox ChkRPay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ê“«—‘ ‘Â—ÌÂ Â«Ì Å—œ«Œ Ì"
      Height          =   255
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CheckBox ChkCPay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "ò‰ —· ‘Â—ÌÂ Â«Ì ò«—»—«‰"
      Height          =   255
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox TxtPassword 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox TxtPassword2 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "  ”ÿÕ œ” —”Ì  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3480
      Width           =   3855
   End
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   8520
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
   Begin GYM.XPStyle XPStyle1 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin GYM.lvButtons_H CmdAdd 
      Height          =   495
      Left            =   2640
      TabIndex        =   18
      Top             =   8520
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
   Begin MSAdodcLib.Adodc AdoAdmin 
      Height          =   390
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«—»—Ì ÃœÌœ :"
      Height          =   240
      Index           =   3
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÌ—«Ì‘ „œÌ—«‰ »—‰«„Â"
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
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   360
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmEdit.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmEdit.frx":0CCA
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   4080
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmEdit.frx":3651
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò·„Â ⁄»Ê— :"
      Height          =   240
      Index           =   1
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«—»—Ì :"
      Height          =   240
      Index           =   0
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ò—«— ò·„Â ⁄»Ê— :"
      Height          =   240
      Index           =   2
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Width           =   1545
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboUserName_Click()
With AdoAdmin
    .Refresh
    .Recordset.Find "UserName = '" & CboUserName.Text & "'"
    TxtUserName.Text = .Recordset.Fields("UserName")
    TxtPassword.Text = Decrypt(1, 3, 6, 7, 7, True, .Recordset.Fields("Password"))
    TxtPassword2.Text = TxtPassword.Text
    ChkAddUser.Value = .Recordset.Fields("AddUser")
    ChkEditUser.Value = .Recordset.Fields("EditUser")
    ChkEnter.Value = .Recordset.Fields("EnterUser")
    ChkOut.Value = .Recordset.Fields("OutUser")
    ChkPay.Value = .Recordset.Fields("Pay")
    ChkRUser.Value = .Recordset.Fields("RUser")
    ChkROut.Value = .Recordset.Fields("REnter")
    ChkRPay.Value = .Recordset.Fields("RPay")
    ChkCPay.Value = .Recordset.Fields("CPay")
    ChkAddAdmin.Value = .Recordset.Fields("AddAdmin")
    ChkEditAdmin.Value = .Recordset.Fields("EditAdmin")
    ChkDelete.Value = .Recordset.Fields("Delete")
End With
End Sub

Private Sub CboUserName_GotFocus()
GotColor CboUserName
End Sub

Private Sub CboUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtUserName.SetFocus
    SendKeys HiLyt
End If
End Sub

Private Sub CboUserName_LostFocus()
LostColor CboUserName
End Sub

Private Sub CmdAdd_Click()
On Error GoTo ERR_Control
If CboUserName.ListIndex < 0 Then
    MsgBox "‰«„ ò«—»—Ì —« «‰ Œ«» ò‰Ìœ", vbExclamation, "‰«„ ò«—»—Ì"
    CboUserName.SetFocus
    Exit Sub
End If
If Trim(TxtUserName.Text) < 0 Then
    MsgBox "‰«„ ò«—»—Ì —« Ê«—œ ò‰Ìœ", vbExclamation, "‰«„ ò«—»—Ì"
    TxtUserName.SetFocus
    SendKeys HiLyt
    Exit Sub
End If
If TxtPassword.Text = "" Then
    MsgBox "ò·„Â ⁄»Ê— —« Ê«—œ ò‰Ìœ", vbExclamation, "ò·„Â ⁄»Ê—"
    TxtPassword.SetFocus
    SendKeys HiLyt
    Exit Sub
End If
If Val(TxtPassword.Text) <> Val(TxtPassword2.Text) Then
    MsgBox "ò·„Â ⁄»Ê— —« çò ò‰Ìœ", vbExclamation, "ò·„Â ⁄»Ê—"
    TxtPassword.SetFocus
    SendKeys HiLyt
    Exit Sub
End If
If MsgBox("¬Ì« „«Ì· »Â ÊÌ—«Ì‘ «Ì‰ „œÌ— Â” Ìœ ø", vbQuestion + vbYesNo, "„œÌ—") = vbNo Then
    Exit Sub
End If
With AdoAdmin
    .Recordset.Fields("UserName") = Trim(TxtUserName.Text)
    .Recordset.Fields("Password") = Encrypt(1, 3, 6, 7, 7, True, TxtPassword.Text)
    .Recordset.Fields("AddUser") = ChkAddUser.Value
    .Recordset.Fields("EditUser") = ChkEditUser.Value
    .Recordset.Fields("EnterUser") = ChkEnter.Value
    .Recordset.Fields("OutUser") = ChkOut.Value
    .Recordset.Fields("Pay") = ChkPay.Value
    .Recordset.Fields("RUser") = ChkRUser.Value
    .Recordset.Fields("REnter") = ChkROut.Value
    .Recordset.Fields("RPay") = ChkRPay.Value
    .Recordset.Fields("CPay") = ChkCPay.Value
    .Recordset.Fields("AddAdmin") = ChkAddAdmin.Value
    .Recordset.Fields("EditAdmin") = ChkEditAdmin.Value
    .Recordset.Fields("Delete") = ChkDelete.Value
    .Recordset.Update
End With
MsgBox "«Ì‰ „œÌ— ÊÌ—«Ì‘ ‘œ", vbInformation, "„œÌ—"
Exit Sub
ERR_Control:
    MsgBox "«Ì‰ „œÌ— œ— »—‰«„Â ÊÃÊœ œ«—œ", vbExclamation, "‰«„ ò«—»—Ì"
    TxtUserName.SetFocus
    SendKeys HiLyt
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterFrm FrmMain, Me
Call ConnectToDb(AdoAdmin, "Admin", False)
Call AdminToCbo(CboUserName)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then CmdExit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnimateForm Me, -1, -1, aUnload, 5, 5, 3, 13
End Sub
Private Sub TxtPassword_GotFocus()
GotColor TxtPassword
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtPassword2.SetFocus
End Sub

Private Sub TxtPassword_LostFocus()
LostColor TxtPassword
End Sub
Private Sub TxtPassword2_GotFocus()
GotColor TxtPassword2
End Sub

Private Sub TxtPassword2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ChkAddUser.SetFocus
End Sub

Private Sub TxtPassword2_LostFocus()
LostColor TxtPassword2
End Sub


Private Sub TxtUserName_GotFocus()
GotColor TxtUserName
End Sub

Private Sub TxtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtPassword.SetFocus
    SendKeys HiLyt
End If
End Sub

Private Sub TxtUserName_LostFocus()
LostColor TxtUserName
End Sub
