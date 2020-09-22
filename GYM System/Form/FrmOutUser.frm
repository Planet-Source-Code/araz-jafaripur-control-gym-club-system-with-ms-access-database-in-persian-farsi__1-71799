VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmOutUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Œ—ÊÃ ò«—»—"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
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
   ScaleHeight     =   5355
   ScaleWidth      =   7185
   Begin VB.TextBox TxtCode 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin GYM.XPStyle XPStyle1 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin MSAdodcLib.Adodc AdoOut 
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
      TabIndex        =   3
      Top             =   4800
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
   Begin GYM.lvButtons_H CmdOut 
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Œ—ÊÃ ò«—»—"
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
      Bindings        =   "FrmOutUser.frx":0000
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5318
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ò„œ :"
      Height          =   240
      Index           =   1
      Left            =   5910
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmOutUser.frx":0015
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   7185
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Œ—ÊÃ ò«—»—«‰ «“ »«‘ê«Â"
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
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmOutUser.frx":299C
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmOutUser.frx":3666
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7195
   End
End
Attribute VB_Name = "FrmOutUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdOut_Click()
On Error GoTo ERR_Control
If MsgBox("¬Ì« „«Ì· »Â Œ—ÊÃ «Ì‰ ò«—»— Â” Ìœ ø", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If
With AdoOut
    .Recordset.Fields("GiveNum") = 1
    .Recordset.Fields("OutDate_S") = Shamsi.Convert_Date(Date, Gregorian_, HijriShamsi_)
    .Recordset.Fields("OutDate_M") = Date
    .Recordset.Fields("OutTime") = Time
    .Recordset.Update
End With
AdoOut.Recordset.Close
GetLst
Exit Sub
ERR_Control:
    MsgBox "»—«Ì Œ—ÊÃ ÌòÌ —« «‰ Œ«» ò‰Ìœ", vbExclamation, "Œ—ÊÃ ò«—»—"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
CenterFrm FrmMain, Me
GetLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnimateForm Me, -1, -1, aUnload, 5, 5, 3, 13
End Sub
Private Sub FormatKala_lst()
Dim counter As Integer
With dtaDel
    .BackColor = &HE28B2C
    
    .Columns(0).DataField = "UID"
    .Columns(0).Caption = "òœ ò«—»—"
    
    .Columns(1).DataField = "KNum"
    .Columns(1).Caption = "‘„«—Â ò„œ"
    
    .Columns(2).DataField = "InDate_S"
    .Columns(2).Caption = " «—ÌŒ Ê—Êœ"
        
    .Columns(3).DataField = "InTime"
    .Columns(3).Caption = "”«⁄  Ê—Êœ"
        
    .Columns(4).DataField = "Name"
    .Columns(4).Caption = "‰«„ ò«—»—"
        
    .Columns(5).DataField = "Family"
    .Columns(5).Caption = "›«„Ì·Ì ò«—»—"
    
    .HeadFont.Bold = True
    .ScrollBars = dbgBoth
    
    .Splits(0).MarqueeStyle = dbgHighlightRow
    .Splits(0).Locked = False
    .Splits(0).AllowRowSizing = True
    .Splits(0).AllowFocus = False
        
    For counter = 0 To 5
        .Columns(counter).AllowSizing = True
        DoEvents
    Next counter
        For counter = 6 To 11
        .Columns(counter).Visible = False
        DoEvents
    Next counter
End With
DatagridColumnAutoResize dtaDel, Me
End Sub

Private Sub TxtCode_Change()
On Error Resume Next
AdoOut.Refresh
FormatKala_lst
AdoOut.Recordset.Find "Knum = " & Val(TxtCode.Text)
End Sub
Private Sub GetLst()
Dim StrSQL As String
StrSQL = "SELECT UserLogin.UID,UserLogin.KNum,UserLogin.InDate_S,UserLogin.InTime" & _
    ",User.Name,User.Family,UserLogin.GiveNum,UserLogin.InDate_M,UserLogin.OutDate_S" & _
    ",UserLogin.OutDate_M,UserLogin.OutTime,UserLogin.ID FROM UserLogin,[User]" & _
    " WHERE UserLogin.UID = User.ID AND UserLogin.GiveNum = 0"
Call ConnectToDb(AdoOut, StrSQL, True)
AdoOut.Refresh
Call FormatKala_lst
End Sub

Private Sub TxtCode_GotFocus()
GotColor TxtCode
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdOut_Click
End Sub

Private Sub TxtCode_LostFocus()
LostColor TxtCode
End Sub
