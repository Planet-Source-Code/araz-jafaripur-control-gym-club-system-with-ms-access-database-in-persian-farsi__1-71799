VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3CA40DFE-4DED-4BD9-98FD-6BEEE7B69F2A}#24.0#0"; "PDTPicker.ocx"
Begin VB.Form FrmROut 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "æÑæÏ æ ÎÑæÌ ˜ÇÑÈÑÇä"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
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
   ScaleHeight     =   7845
   ScaleWidth      =   8370
   Begin VB.TextBox TxtID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      RightToLeft     =   -1  'True
      ScaleHeight     =   495
      ScaleWidth      =   5895
      TabIndex        =   11
      Top             =   960
      Width           =   5895
      Begin VB.OptionButton OptR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "äÇã"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton OptR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "˜Ï ˜ÇÑÈÑí"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "ÝÇãíáí"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton OptR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "ãÔÇåÏå åãå"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
   End
   Begin GYM.XPStyle XPStyle1 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   7320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "ÎÑæÌ"
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
   Begin MSDataGridLib.DataGrid dtaDel 
      Bindings        =   "FrmROut.frx":0000
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin GYM.lvButtons_H CmdSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "ÌÓÊÌæ"
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
   Begin MSAdodcLib.Adodc AdoUser 
      Height          =   375
      Left            =   5640
      Top             =   360
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
   Begin PDTPicker.FDTPicker TxtFrom 
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
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
   Begin PDTPicker.FDTPicker TxtTo 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
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
   Begin GYM.lvButtons_H CmdSearch2 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "ÌÓÊÌæ"
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
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÒÇÑÔ æÑæÏ æ ÎÑæÌ ˜ÇÑÈÑÇä ÈÇÔÇå"
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
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   360
      Width           =   2730
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmROut.frx":0016
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmROut.frx":0CE0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8415
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmROut.frx":3D80
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   8400
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÌÓÊÌæ ÈÑ ÍÓÈ :"
      Height          =   240
      Index           =   0
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÚÈÇÑÊ ÈÑÇí ÌÓÊÌæ :"
      Height          =   240
      Index           =   1
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÚÏÇÏ æÑæÏ æ ÎÑæÌ åÇ :"
      Height          =   240
      Index           =   2
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6720
      Width           =   2040
   End
   Begin VB.Label LblCount 
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
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   6720
      Width           =   60
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   8280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   8280
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "æÑæÏ ÇÒ ÊÇÑíÎ :"
      Height          =   240
      Index           =   4
      Left            =   7020
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2595
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇ ÊÇÑíÎ :"
      Height          =   240
      Index           =   3
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2595
      Width           =   735
   End
End
Attribute VB_Name = "FrmROut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSearch_Click()
On Error Resume Next
    Dim counter As Integer
    Dim StrSearch As String
    Dim StrFilter As String
    If OptR(0).Value Then
        StrSearch = "ID = " & TxtID.Text
    ElseIf OptR(1).Value Then
        StrSearch = "Name LIKE '%" & TxtID.Text & "%'"
    ElseIf OptR(2).Value Then
        StrSearch = "Family LIKE '%" & TxtID.Text & "%'"
    End If
    AdoUser.Refresh
    AdoUser.Recordset.Filter = StrSearch
    SendKeys HiLyt
    If Not AdoUser.Recordset.BOF = True And AdoUser.Recordset.EOF = True Then
        Call FormatKala_lst
        MsgBox "ÌÓÊÌæ äÊíÌå Çí äÏÇÔÊ", vbExclamation, "íÏÇ äÔÏ"
        TxtID.SetFocus
        SendKeys HiLyt
        Exit Sub
    End If
    Call FormatKala_lst
End Sub

Private Sub CmdSearch2_Click()
Dim StrFilter As String
StrFilter = "InDate_M >= '" & Shamsi.Convert_Date(TxtFrom.Text, HijriShamsi_, Gregorian_) & _
    "' AND InDate_M <= '" & Shamsi.Convert_Date(TxtTo.Text, HijriShamsi_, Gregorian_) & "'"
AdoUser.Recordset.Filter = StrFilter
LblCount.Caption = AdoUser.Recordset.RecordCount
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
CenterFrm FrmMain, Me
Dim StrSQL As String
Dim Date1 As String, Date2 As String
StrSQL = "SELECT User.ID,User.Name,User.Family,UserLogin.KNum,UserLogin.InDate_S" & _
    ",UserLogin.InTime,UserLogin.OutDate_S,UserLogin.OutTime,UserLogin.InDate_M,UserLogin.OutDate_M" & _
    " FROM [User],UserLogin WHERE User.ID = UserLogin.UID AND UserLogin.GiveNum = 1 ORDER BY UserLogin.InDate_M DESC"
Call ConnectToDb(AdoUser, StrSQL, True)
AdoUser.Refresh
Call FormatKala_lst
Date1 = Shamsi.Convert_Date(Date, Gregorian_, HijriShamsi_)
Date2 = Shamsi.Convert_Date(Date + 30, Gregorian_, HijriShamsi_)
TxtFrom.Text = Date1
TxtTo.Text = Date2
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnimateForm Me, -1, -1, aUnload, 5, 5, 3, 13
End Sub
Private Sub FormatKala_lst()
Dim counter As Integer
With dtaDel
    .BackColor = &HE28B2C
    
    .Columns(0).DataField = "ID"
    .Columns(0).Caption = "˜Ï ˜ÇÑÈÑ"
    
    .Columns(1).DataField = "Name"
    .Columns(1).Caption = "äÇã"
    
    .Columns(2).DataField = "Family"
    .Columns(2).Caption = "ÝÇãíáí"
        
    .Columns(3).DataField = "KNum"
    .Columns(3).Caption = "ÔãÇÑå ˜ãÏ"
        
    .Columns(4).DataField = "InDate_S"
    .Columns(4).Caption = "ÊÇÑíÎ æÑæÏ"
        
    .Columns(5).DataField = "InTime"
    .Columns(5).Caption = "ÓÇÚÊ æÑæÏ"
    
    .Columns(6).DataField = "OutDate_S"
    .Columns(6).Caption = "ÊÇÑíÎ ÎÑæÌ"
    
    .Columns(7).DataField = "OutTime"
    .Columns(7).Caption = "ÓÇÚÊ ÎÑæÌ"
    
    .HeadFont.Bold = True
    .ScrollBars = dbgBoth
    
    .Splits(0).MarqueeStyle = dbgHighlightRow
    .Splits(0).Locked = False
    .Splits(0).AllowRowSizing = True
    .Splits(0).AllowFocus = False
        
    For counter = 0 To 7
        .Columns(counter).AllowSizing = True
        DoEvents
    Next counter
    
    For counter = 8 To 9
        .Columns(counter).Visible = False
        DoEvents
    Next counter
End With
DatagridColumnAutoResize dtaDel, Me
LblCount.Caption = AdoUser.Recordset.RecordCount
End Sub

Private Sub TxtFrom_GotFocus()
GotColor TxtFrom
End Sub

Private Sub TxtFrom_LostFocus()
LostColor TxtFrom
End Sub

Private Sub TxtID_GotFocus()
GotColor TxtID
End Sub

Private Sub TxtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdSearch_Click
End Sub

Private Sub TxtID_LostFocus()
LostColor TxtID
End Sub

Private Sub TxtTo_GotFocus()
GotColor TxtTo
End Sub

Private Sub TxtTo_LostFocus()
LostColor TxtTo
End Sub


