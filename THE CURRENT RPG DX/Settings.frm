VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Settings"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color Depth"
      Height          =   1335
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
      Begin VB.OptionButton Option6 
         Caption         =   "16 bit"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         Caption         =   "24 bit"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "32 bit"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resolution"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton Option3 
         Caption         =   "1024 x 768"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "800 x 600"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "640 x 480"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public resoluteX As Integer
Public resoluteY As Integer
Public colordepth As Integer

Private Sub Command1_Click()
    'closes the form
    Unload Me
    Load frmMain
End Sub

Private Sub Form_Load()
    resoluteX = 800
    resoluteY = 600
    colordepth = 32
End Sub
'sets the resolution and color depth that user wants
Private Sub Option1_Click()
    resoluteX = 640
    resoluteY = 480
End Sub

Private Sub Option2_Click()
    resoluteX = 800
    resoluteY = 600
End Sub

Private Sub Option3_Click()
    resoluteX = 1024
    resoluteY = 768
End Sub

Private Sub Option4_Click()
    colordepth = 32
End Sub

Private Sub Option5_Click()
    colordepth = 24
End Sub

Private Sub Option6_Click()
    colordepth = 16
End Sub
