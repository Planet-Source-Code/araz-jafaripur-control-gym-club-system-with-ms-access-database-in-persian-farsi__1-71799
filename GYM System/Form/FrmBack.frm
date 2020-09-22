VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "–ŒÌ—Â Ê »«“Ì«»Ì"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
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
   ScaleHeight     =   2955
   ScaleWidth      =   4830
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   1440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "„Õ· –ŒÌ—Â Å‘Ì Ì»«‰ œ— œÌ”ò"
      InitDir         =   "c:\"
      MaxFileSize     =   9999
   End
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
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
   Begin GYM.lvButtons_H CmdStart 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "‘—Ê⁄"
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
   Begin GYM.lvButtons_H CmdRestore 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "»«“Ì«»Ì"
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
   Begin GYM.lvButtons_H CmdFind 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Å‘ Ì»«‰"
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
      Left            =   3840
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ê—› ‰ Å‘ Ì»«‰ Ê »«“Ì«»Ì «“ ›«Ì· »—‰«„Â"
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
      TabIndex        =   6
      Top             =   360
      Width           =   3075
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmBack.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmBack.frx":0CCA
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   4830
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmBack.frx":3651
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4845
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„Õ· „Ê—œ ‰Ÿ— :"
      Height          =   240
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1440
   End
End
Attribute VB_Name = "FrmBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdFind_Click()
On Error Resume Next
BlnRestore = False
With Cdl
    .CancelError = False
    .Flags = cdlOFNFileMustExist + cdlOFNReadOnly + cdlOFNOverwritePrompt
    .Filter = "Backup File|*.Backup"
    .FileName = "A"
    .ShowSave
    If .FileName = "" Then
        Exit Sub
    Else
        TxtName.Text = .FileName
    End If
End With
End Sub

Private Sub CmdRestore_Click()
On Error Resume Next
BlnRestore = True
With Cdl
    .CancelError = False
    .Flags = cdlOFNFileMustExist + cdlOFNReadOnly + cdlOFNOverwritePrompt
    .Filter = "Backup File|*.Backup"
    .ShowOpen
    If .FileName = "" Then
        Exit Sub
    Else
        TxtName.Text = .FileName
    End If
End With
End Sub

Private Sub CmdStart_Click()
On Error GoTo ERR_Control
If TxtName.Text = "" Then
    MsgBox "·ÿ›« „”Ì— „Ê—œ ‰Ÿ— —« «‰ Œ«» ò‰Ìœ !", vbOKOnly + vbExclamation, "„”Ì— „Ê—œ ‰Ÿ—"
    Exit Sub
End If
Me.Enabled = False
Me.MousePointer = vbHourglass
Dim srcFile As String
Dim DestFile As String
If Not BlnRestore Then
    srcFile = App.Path & "\Data\Data.Mj"
    DestFile = Trim(TxtName.Text)
    FileCopy srcFile, DestFile
Else
    srcFile = Trim(TxtName.Text)
    DestFile = App.Path & "\Data\Data.Mj"
    FileCopy srcFile, DestFile
End If
If Not BlnRestore Then
    MsgBox "ê—› ‰ Å‘ Ì»«‰ »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ" & vbNewLine & vbNewLine & DestFile, vbInformation, "« „«„ Å‘ Ì»«‰ êÌ—Ì"
Else
    MsgBox "»«“Ì«»Ì ›«Ì· »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ" & vbNewLine & vbNewLine & DestFile, vbInformation, "« „«„ »«“Ì«»Ì ›«Ì·"
End If
Me.MousePointer = 0
Me.Enabled = True
Exit Sub
ERR_Control:
    MsgBox "»—‰«„Â —« »” Â »⁄œ« ”⁄Ì ò‰Ìœ", vbExclamation, "Å‘ Ì»«‰ - »«“Ì«»Ì"
    Me.MousePointer = 0
    Me.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then CmdExit_Click
End Sub

Private Sub Form_Load()
CenterFrm FrmMain, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnimateForm Me, -1, -1, aUnload, 5, 5, 3, 13
End Sub

