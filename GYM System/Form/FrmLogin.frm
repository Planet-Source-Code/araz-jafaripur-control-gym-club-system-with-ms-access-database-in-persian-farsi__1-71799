VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmLogin 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê—Êœ »Â ”Ì” „"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin GYM.XPStyle XPStyle1 
      Left            =   1920
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin VB.ComboBox CboUserName 
      Height          =   360
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc AdoLogin 
      Height          =   375
      Left            =   1320
      Top             =   480
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
   Begin VB.TextBox TxtPassword 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2400
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
   Begin GYM.lvButtons_H CmdLogin 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2400
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
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò·„Â ⁄»Ê— :"
      Height          =   240
      Index           =   1
      Left            =   2685
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«—»—Ì :"
      Height          =   240
      Index           =   0
      Left            =   2670
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmLogin.frx":0ECA
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê—Êœ »Â ”Ì” „ „œÌ—Ì  »«‘ê«Â"
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
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   2565
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmLogin.frx":1B94
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmLogin.frx":451B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3885
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboUserName_GotFocus()
GotColor CboUserName
End Sub

Private Sub CboUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtPassword.SetFocus
    SendKeys HiLyt
End If
End Sub

Private Sub CboUserName_LostFocus()
LostColor CboUserName
End Sub

Private Sub CmdExit_Click()
End
End Sub
Private Sub LoginError()
    MsgBox "ò·„Â ò«—»—Ì Ì« ‰«„ ò«—»—Ì œ—”  ‰„Ì »«‘œ", vbExclamation, "Ê—Êœ ‰«„Ê›ﬁ"
    CboUserName.SetFocus
End Sub

Private Sub CmdLogin_Click()
On Error GoTo ERR_Control
With AdoLogin.Recordset
    AdoLogin.Refresh
    .Find ("UserName = '" & CboUserName.Text & "'")
    If .Fields("Password") = Encrypt(1, 3, 6, 7, 7, True, TxtPassword.Text) Then
        Load FrmMain
        FrmMain.CmdAddUser.Enabled = .Fields("AddUser")
        FrmMain.CmdEditUser.Enabled = .Fields("EditUser")
        FrmMain.CmdEnterUser.Enabled = .Fields("EnterUser")
        FrmMain.CmdOutUser.Enabled = .Fields("OutUser")
        FrmMain.CmdPay.Enabled = .Fields("Pay")
        FrmMain.CmdRUser.Enabled = .Fields("RUser")
        FrmMain.CmdREnter.Enabled = .Fields("REnter")
        FrmMain.CmdRPay.Enabled = .Fields("RPay")
        FrmMain.CmdCPay.Enabled = .Fields("CPay")
        FrmMain.CmdAddAdmin.Enabled = .Fields("AddAdmin")
        FrmMain.CmdEditAdmin.Enabled = .Fields("EditAdmin")
        FrmMain.CmdDeleteAdmin.Enabled = .Fields("Delete")
        FrmMain.Show
        Unload Me
    Else
        LoginError
    End If
End With
Exit Sub
ERR_Control:
    LoginError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End
End Sub

Private Sub Form_Load()
Call ConnectToDb(AdoLogin, "Admin", False)
Call AdminToCbo(CboUserName)
End Sub

Private Sub TxtPassword_GotFocus()
GotColor TxtPassword
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdLogin_Click
End Sub

Private Sub TxtPassword_LostFocus()
LostColor TxtPassword
End Sub
