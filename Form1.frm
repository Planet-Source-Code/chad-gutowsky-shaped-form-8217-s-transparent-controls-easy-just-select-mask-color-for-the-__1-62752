VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2130
      Left            =   5400
      Picture         =   "Form1.frx":13CDA
      ScaleHeight     =   2070
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   405
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3555
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transparent Controls"
      Height          =   1995
      Left            =   5400
      TabIndex        =   1
      Top             =   2655
      Width           =   4335
      Begin VB.Frame Frame2 
         BackColor       =   &H0000FF00&
         Caption         =   "Frame2"
         Height          =   645
         Left            =   2205
         TabIndex        =   9
         Top             =   855
         Width           =   960
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H0000FF00&
         Height          =   315
         Left            =   1530
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   270
         Width           =   1635
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FF00&
         Caption         =   "Option1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FF00&
         Caption         =   "Check1"
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   1365
      End
   End
   Begin Project1.TransparentColor TransparentColor1 
      Height          =   465
      Left            =   4500
      TabIndex        =   0
      Top             =   4590
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   820
      MaskColor       =   65280
      BackColor       =   -2147483633
      Enable          =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Form has a background image with a mask color selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   720
      TabIndex        =   8
      Top             =   1350
      Width           =   3525
   End
   Begin VB.Label Label1 
      Caption         =   "Picture Box"
      Height          =   195
      Left            =   5445
      TabIndex        =   6
      Top             =   90
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
