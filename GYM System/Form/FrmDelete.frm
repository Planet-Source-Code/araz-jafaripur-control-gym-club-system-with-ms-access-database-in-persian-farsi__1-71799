VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmDelete 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÍÐÝ ãÏíÑ"
   ClientHeight    =   2355
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
   ScaleHeight     =   2355
   ScaleWidth      =   4080
   Begin VB.ComboBox CboUserName 
      Height          =   360
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
   Begin GYM.XPStyle XPStyle1 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin MSAdodcLib.Adodc AdoAdmin 
      Height          =   390
      Left            =   1560
      Top             =   0
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
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      caption         =   "ÎÑæÌ"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   14846764
      cfore           =   8388608
      font            =   "FrmDelete.frx":0000
      mode            =   0
      value           =   0   'False
      cfhover         =   8388608
      cback           =   -2147483633
      cbhover         =   14846764
      capstyle        =   2
   End
   Begin GYM.lvButtons_H CmdDelete 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      caption         =   "ÍÐÝ"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   14846764
      cfore           =   8388608
      font            =   "FrmDelete.frx":0028
      mode            =   0
      value           =   0   'False
      cfhover         =   8388608
      cback           =   -2147483633
      cbhover         =   14846764
      capstyle        =   2
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmDelete.frx":0050
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   4080
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmDelete.frx":29D7
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍÐÝ ãÏíÑ ÇÒ ÈÑäÇãå"
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
      TabIndex        =   2
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "äÇã ˜ÇÑÈÑí :"
      Height          =   240
      Index           =   0
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmDelete.frx":36A1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "FrmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboUserName_Click()
    AdoAdmin.Refresh
    AdoAdmin.Recordset.Find "UserName = '" & CboUserName.Text & "'"
End Sub
Private Sub CboUserName_GotFocus()
GotColor CboUserName
End Sub

Private Sub CboUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdDelete_Click
End Sub

Private Sub CboUserName_LostFocus()
LostColor CboUserName
End Sub

Private Sub CmdDelete_Click()
If CboUserName.ListIndex < 0 Then
    MsgBox "äÇã ˜ÇÑÈÑí ÑÇ ÇäÊÎÇÈ ˜äíÏ", vbExclamation, "äÇã ˜ÇÑÈÑí"
    CboUserName.SetFocus
    Exit Sub
End If
If MsgBox("ÂíÇ ãÇíá Èå ÍÐÝ Çíä ãÏíÑ åÓÊíÏ ¿", vbQuestion + vbYesNo, "ãÏíÑ") = vbNo Then
    Exit Sub
End If
AdoAdmin.Recordset.Delete
MsgBox "ãÏíÑ ãæÑÏ äÙÑ ÍÐÝ ÔÏ", vbInformation, "ÍÐÝ ãÏíÑ"
Unload Me
FrmDelete.Show
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
