VERSION 5.00
Begin VB.UserControl XPStyle 
   BackStyle       =   0  'Transparent
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   810
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "XPStyle.ctx":0000
   Picture         =   "XPStyle.ctx":22DA
   ScaleHeight     =   810
   ScaleWidth      =   810
   ToolboxBitmap   =   "XPStyle.ctx":45B4
End
Attribute VB_Name = "XPStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hMod As Long
Private Sub UserControl_Initialize()
    Dim iccex As tagInitCommonControlsEx
    iccex.lngSize = LenB(iccex)
    iccex.lngICC = ICC_USEREX_CLASSES
    InitCommonControlsEx iccex
    m_hMod = LoadLibrary("shell32.dll")
End Sub
Private Sub UserControl_Resize()
    UserControl.Width = 810
    UserControl.Height = 810
End Sub

Private Sub UserControl_Terminate()
    FreeLibrary m_hMod
End Sub
