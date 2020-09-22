VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmCPay 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ò‰ —· ‘Â—ÌÂ Â«"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
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
   ScaleWidth      =   8520
   Begin VB.CheckBox ChkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "›ﬁÿ „‘«ÂœÂ ò«—»—«‰Ì òÂ ›⁄«· »«‘‰œ"
      Height          =   375
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      RightToLeft     =   -1  'True
      ScaleHeight     =   1455
      ScaleWidth      =   8295
      TabIndex        =   7
      Top             =   1560
      Width           =   8295
      Begin VB.OptionButton OptR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "ò«—»—«‰Ì òÂ „Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ ¬‰Â« ‰—”ÌœÂ «” "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
      End
      Begin VB.OptionButton OptR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "ò«—»—«‰Ì òÂ „Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ ¬‰Â« «„—Ê“ «” "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
      Begin VB.OptionButton OptR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "ò«—»—«‰Ì òÂ „Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ ¬‰Â« ê–‘ Â «” "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   4935
      End
   End
   Begin GYM.XPStyle XPStyle1 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin MSAdodcLib.Adodc AdoUser 
      Height          =   375
      Left            =   6120
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
   Begin MSFlexGridLib.MSFlexGrid MsFact 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   -2147483647
      ForeColorFixed  =   255
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      RightToLeft     =   -1  'True
      FocusRect       =   0
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      GridLineWidth   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   7080
      TabIndex        =   5
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
   Begin GYM.lvButtons_H CmdShow 
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "„‘«ÂœÂ"
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
      Caption         =   " ⁄œ«œ ò«—»—«‰ :"
      Height          =   240
      Index           =   2
      Left            =   7125
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   7080
      Width           =   1275
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
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   7080
      Width           =   60
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   8400
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmCPay.frx":0000
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   8520
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò‰ —· ‘Â—ÌÂ Â«Ì ò«—»—«‰ »«‘ê«Â"
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
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   2685
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmCPay.frx":2987
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmCPay.frx":3651
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "FrmCPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Closed As Boolean
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdShow_Click()
On Error Resume Next
Me.Enabled = False
MsFact.Clear
MsFact.Rows = 2
Dim StrSQL As String
Dim Mode As Integer
Dim Active As Byte
Active = ChkActive.Value
If OptR(0).Value Then
    Mode = 1
ElseIf OptR(1).Value Then
    Mode = 0
ElseIf OptR(2).Value Then
    Mode = -1
End If
Dim I As Long
Dim Count As Long
Dim Result As Integer
StrSQL = "SELECT ID,Name,Family,Active FROM [User] WHERE Active = " & Active
Call ConnectToDb(AdoUser, StrSQL, True)
With AdoUser.Recordset
    Count = .RecordCount
    MsFact.Row = 0
    MsFact.Col = 0
    MsFact.Text = "òœ ò«—»—"
    MsFact.Col = 1
    MsFact.Text = "‰«„"
    MsFact.Col = 2
    MsFact.Text = "›«„Ì·Ì"
    MsFact.Col = 3
    MsFact.Text = "‘Â—ÌÂ"
    MsFact.ColWidth(1) = 1500
    MsFact.ColWidth(2) = 2000
    MsFact.ColWidth(3) = 3500
    Dim J As Long
    J = 2
    For I = 1 To Count
        Result = GetPay(.Fields("ID"))
        If Result = Mode Then
            LblCount.Caption = J
            MsFact.Rows = J
            MsFact.Row = J - 1
            MsFact.Col = 0
            MsFact.Text = .Fields("ID")
            MsFact.Col = 1
            MsFact.Text = .Fields("Name")
            MsFact.Col = 2
            MsFact.Text = .Fields("Family")
            MsFact.Col = 3
            Select Case Result
                Case Is = 0
                    MsFact.Text = "«„—Ê“ „Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ «” "
                Case Is = -1
                    MsFact.Text = "„Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ ‰—”ÌœÂ «” "
                Case Is = 1
                    MsFact.Text = "„Ê⁄œ Å—œ«Œ  ‘Â—ÌÂ ê–‘ Â «” "
            End Select
            J = J + 1
        End If
        .MoveNext
        DoEvents
    Next
End With
AdoUser.Recordset.Close
Me.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
CenterFrm FrmMain, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnimateForm Me, -1, -1, aUnload, 5, 5, 3, 13
End Sub
