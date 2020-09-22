VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "œ—»«—Â"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
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
   ScaleHeight     =   3435
   ScaleWidth      =   4170
   Begin GYM.lvButtons_H CmdExit 
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2880
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Http://Jafaripur.Blogfa.Com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ :"
      Height          =   240
      Index           =   5
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "’›ÕÂ «Ì‰ —‰ Ì :"
      Height          =   240
      Index           =   3
      Left            =   2610
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»—‰«„Â ‰ÊÌ” :"
      Height          =   240
      Index           =   0
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¬—«“ Ã⁄›—Ì ÅÊ—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mjafaripur@Yahoo.Com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   495
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1980
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«Ì„Ì· :"
      Height          =   240
      Index           =   8
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   630
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      Picture         =   "FrmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   4170
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "œ—»«—Â »—‰«„Â"
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
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "FrmAbout.frx":2987
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "FrmAbout.frx":3651
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4185
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
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

Private Sub Label1_Click(Index As Integer)
If Index = 4 Then Shell "Explorer " & Label1(4).Caption
If Index = 2 Then Shell "Explorer MailTo:" & Label1(2).Caption
End Sub
