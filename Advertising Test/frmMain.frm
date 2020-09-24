VERSION 5.00
Object = "{FB53E87D-C383-464C-BE23-AE82A4CC7716}#1.1#0"; "APB.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Progress Bar"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   609
      ShowText        =   -1  'True
      Style           =   0
      BarColor1       =   -2147483634
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   465
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   609
      ShowText        =   -1  'True
      BarColor1       =   -2147483634
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   885
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   609
      ShowText        =   -1  'True
      Style           =   2
      BarColor1       =   -2147483634
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1275
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   609
      ShowText        =   -1  'True
      Style           =   3
      BarColor1       =   -2147483634
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   1695
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   609
      Style           =   4
      BarColor1       =   -2147483634
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   195
      Index           =   5
      Left            =   90
      TabIndex        =   5
      Top             =   2070
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   344
      BorderStyle     =   0
      Style           =   5
      BarColor1       =   -2147483634
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   6
      Left            =   90
      TabIndex        =   6
      Top             =   2340
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   609
      Style           =   6
      BarColor1       =   -2147483634
      CustomPicture   =   "frmMain.frx":0000
      CustomPicture   =   "frmMain.frx":06BF
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   7
      Left            =   75
      TabIndex        =   7
      Top             =   2715
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   609
      ShowText        =   -1  'True
      Style           =   7
      BarColor1       =   -2147483634
      CustomPicture   =   "frmMain.frx":0D7E
      CustomPicture   =   "frmMain.frx":1DAD
   End
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   345
      Index           =   8
      Left            =   75
      TabIndex        =   8
      Top             =   3105
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   609
      ShowText        =   -1  'True
      Style           =   8
      BarColor1       =   -2147483634
      CustomPicture   =   "frmMain.frx":2DDC
      CustomPicture   =   "frmMain.frx":3E0B
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Dim x As Long
x = InputBox("%:")
For i = 0 To 8
APB(i).Value = x
Next i
End Sub

Private Sub Form_Load()
APB(0).BarColor1 = vbHighlight
APB(1).BarColor1 = vbHighlight
End Sub
