VERSION 5.00
Begin VB.UserControl TransparentColor 
   BackColor       =   &H000000FF&
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ClipBehavior    =   0  'None
   FontTransparent =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   495
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   0
      Width           =   510
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "TC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   -45
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   540
   End
   Begin VB.Menu mnuColorMask 
      Caption         =   "ucColorMask"
   End
End
Attribute VB_Name = "TransparentColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'--------------------
Private Type POINTAPI
  X As Long
  Y As Long
End Type
'------------------
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Dim AcUser

Private Sub Timer1_Timer()
                    On Error Resume Next
                    Dim Ret As Long
                    Dim TC As Long
                    AcUser = UserControl.Name
                    UserControl.Enabled = False
                    TC = UserControl.MaskColor 'vbBlue
                    Ret = GetWindowLong(UserControl.Parent.hWnd, G_E)
                    Ret = Ret Or W_E
                    SetWindowLong UserControl.Parent.hWnd, G_E, Ret
                    SetLayeredWindowAttributes UserControl.Parent.hWnd, TC, 0, LW_KEY
                    UserControl.BackColor = UserControl.Parent.BackColor

                    Timer1.Enabled = False
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    UserControl.MaskColor = .ReadProperty("MaskColor", vbBlack)
    UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
    UserControl.Enabled = .ReadProperty("Enable", True)
  End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "MaskColor", UserControl.MaskColor, vbBlack
    .WriteProperty "BackColor", UserControl.BackColor, vbWhite
    .WriteProperty "Enable", UserControl.Enabled, True
  End With
End Sub
Public Property Get MaskColor() As OLE_COLOR
  MaskColor = UserControl.MaskColor
  Refresh
End Property
Public Property Let MaskColor(ByVal NewColor As OLE_COLOR)
  UserControl.MaskColor = NewColor
  Refresh
  PropertyChanged "MaskColor"
End Property
Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
  UserControl.BackColor = NewColor
  PropertyChanged "BackColor"
End Property
Public Property Get Enable() As Boolean
  Enable = UserControl.Enabled
End Property
Public Property Let Enable(ByVal NewValue As Boolean)
  UserControl.Enabled = NewValue
End Property




