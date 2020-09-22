VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "”Ì” „ ò‰ —· »«‘ê«Â »œ‰ ”«“Ì"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9930
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "FrmMain"
   RightToLeft     =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicMenu 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7275
      Index           =   3
      Left            =   -1725
      RightToLeft     =   -1  'True
      ScaleHeight     =   7245
      ScaleWidth      =   2325
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
      Begin GYM.lvButtons_H CmdAddAdmin 
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "«›“Êœ‰ „œÌ—"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdEditAdmin 
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "ÊÌ—«Ì‘ „œÌ—«‰"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdDeleteAdmin 
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Õ–› „œÌ—«‰"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
   End
   Begin VB.PictureBox PicMenu 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7275
      Index           =   2
      Left            =   630
      RightToLeft     =   -1  'True
      ScaleHeight     =   7245
      ScaleWidth      =   2325
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
      Begin GYM.lvButtons_H CmdRUser 
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "ò«—»—«‰ »«‘ê«Â"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdREnter 
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Ê—Êœ Ê Œ—ÊÃ ò«—»—«‰"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdRPay 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "‘Â—ÌÂ Â«Ì Å—œ«Œ Ì"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdCPay 
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "ò‰ —· ‘Â—ÌÂ ò«—»—«‰"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
   End
   Begin VB.PictureBox PicMenu 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7275
      Index           =   1
      Left            =   2985
      RightToLeft     =   -1  'True
      ScaleHeight     =   7245
      ScaleWidth      =   2325
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
      Begin GYM.lvButtons_H CmdEnterUser 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Ê—Êœ ò«—»—"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdOutUser 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Œ—ÊÃ ò«—»—"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdPay 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Å—œ«Œ  ‘Â—ÌÂ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
   End
   Begin VB.PictureBox PicMenu 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7275
      Index           =   0
      Left            =   5340
      RightToLeft     =   -1  'True
      ScaleHeight     =   7245
      ScaleWidth      =   2325
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
      Begin GYM.lvButtons_H CmdAddUser 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "«›“Êœ‰ ò«—»—"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdEditUser 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "ÊÌ—«Ì‘ ò«—»—"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
   End
   Begin VB.PictureBox PicA 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7275
      Left            =   7695
      RightToLeft     =   -1  'True
      ScaleHeight     =   7245
      ScaleWidth      =   2205
      TabIndex        =   20
      Top             =   0
      Width           =   2235
      Begin GYM.lvButtons_H CmdUser 
         Height          =   615
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "ò«—»—«‰ »«‘ê«Â"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdControl 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "ò‰ —· »«‘ê«Â"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdReport 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "ê“«—‘"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdAdmin 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "„œÌ—«‰ »—‰«„Â"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdBackup 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Å‘ Ì»«‰ Ê »«“Ì«»Ì"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdExit 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   5280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Œ—ÊÃ «“ »—‰«„Â"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
      Begin GYM.lvButtons_H CmdAbout 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   4440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "œ—»«—Â"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   14846764
         LockHover       =   3
         cGradient       =   14846764
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483631
      End
   End
   Begin GYM.XPStyle XPStyle1 
      Left            =   2160
      Top             =   360
      _ExtentX        =   1429
      _ExtentY        =   1429
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   19
      Top             =   7275
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1236
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1236
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1236
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "07:45 ».Ÿ"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAbout_Click()
FrmAbout.Show
End Sub

Private Sub CmdAddAdmin_Click()
FrmAddAdmin.Show
ShowMenu -1
End Sub

Private Sub CmdAddUser_Click()
FrmAddUser.Show
ShowMenu -1
End Sub

Private Sub CmdAdmin_Click()
ShowMenu 3
End Sub

Private Sub CmdBackup_Click()
ShowMenu -1
FrmBack.Show
End Sub

Private Sub CmdChangePass_Click()
ShowMenu -1
End Sub

Private Sub CmdControl_Click()
ShowMenu 1
End Sub

Private Sub CmdCPay_Click()
FrmCPay.Show
ShowMenu -1
End Sub

Private Sub CmdDeleteAdmin_Click()
FrmDelete.Show
ShowMenu -1
End Sub

Private Sub CmdEditAdmin_Click()
FrmEdit.Show
ShowMenu -1
End Sub

Private Sub CmdEditUser_Click()
FrmEditUser.Show
ShowMenu -1
End Sub

Private Sub CmdEnterUser_Click()
FrmEnterUser.Show
ShowMenu -1
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdOutUser_Click()
FrmOutUser.Show
ShowMenu -1
End Sub

Private Sub CmdPay_Click()
FrmPay.Show
ShowMenu -1
End Sub

Private Sub CmdREnter_Click()
FrmROut.Show
ShowMenu -1
End Sub

Private Sub CmdReport_Click()
ShowMenu 2
End Sub

Private Sub CmdRPay_Click()
FrmRPay.Show
ShowMenu -1
End Sub

Private Sub CmdRUser_Click()
FrmRUser.Show
ShowMenu -1
End Sub

Private Sub CmdUser_Click()
ShowMenu 0
End Sub

Private Sub MDIForm_Load()
Dim Re As String
Dim I As Integer
stbMain.Panels(4).Text = Shamsi.Convert_Date(Date, Gregorian_, HijriShamsi_)
Call Shamsi.Convert_Date2Letter(stbMain.Panels(4).Text, Re)
stbMain.Panels(5).Text = Re
FrmMain.Picture = LoadPicture(App.Path & "\Stage_BG.jpg")
PicA.Picture = LoadPicture(App.Path & "\bg.jpg")
For I = 0 To 3
    PicMenu(I).Picture = PicA.Picture
    DoEvents
Next
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu MnuMain
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("„«Ì· »Â Œ—ÊÃ „Ì »«‘Ìœ ø", vbQuestion + vbYesNo, "Œ—ÊÃ") = vbNo Then
    Cancel = 1
Else
    End
End If
End Sub
Private Sub ShowMenu(Index As Integer)
Dim I As Integer
For I = 0 To 3
    PicMenu(I).Width = 0
    PicMenu(I).Visible = False
    DoEvents
Next
If Index <> -1 Then
    PicMenu(Index).Visible = True
    For I = 0 To 100
        PicMenu(Index).Width = I * 23.2
        DoEvents
    Next
End If
End Sub
