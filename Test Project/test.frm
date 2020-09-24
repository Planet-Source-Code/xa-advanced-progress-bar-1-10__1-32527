VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FB53E87D-C383-464C-BE23-AE82A4CC7716}#1.1#0"; "APB.ocx"
Begin VB.Form Form1 
   Caption         =   "Testing Advanced Progress Bar"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   391
   StartUpPosition =   3  'Windows Default
   Begin Advanced_Progress_Bar.AdvProgressBar APB 
      Height          =   390
      Left            =   15
      TabIndex        =   14
      Top             =   690
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   794
      ShowText        =   -1  'True
      Style           =   2
      BarColor1       =   -2147483634
      CustomPicture   =   "test.frx":0000
      CustomPicture   =   "test.frx":102F
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Picture"
      Height          =   240
      Left            =   60
      TabIndex        =   13
      Top             =   4185
      Width           =   1935
   End
   Begin VB.PictureBox pt 
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   660
      ScaleHeight     =   360
      ScaleWidth      =   1290
      TabIndex        =   12
      Top             =   3720
      Width           =   1350
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Text Label"
      Height          =   255
      Left            =   75
      TabIndex        =   10
      Top             =   3465
      Width           =   2025
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2565
      Top             =   2895
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Start"
      Height          =   450
      Left            =   2085
      TabIndex        =   9
      Top             =   3660
      Width           =   1680
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4305
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Graphics|*.bmp;*.gif;*.jpg;*.wmf"
   End
   Begin VB.PictureBox pC2 
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   630
      ScaleHeight     =   180
      ScaleWidth      =   1425
      TabIndex        =   8
      Top             =   3195
      Width           =   1485
   End
   Begin VB.PictureBox pC1 
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   630
      ScaleHeight     =   180
      ScaleWidth      =   1425
      TabIndex        =   7
      Top             =   2925
      Width           =   1485
   End
   Begin VB.ComboBox cmdPAppearance 
      Height          =   315
      ItemData        =   "test.frx":205E
      Left            =   3435
      List            =   "test.frx":2068
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2565
      Width           =   2130
   End
   Begin VB.ComboBox cmbAppearance 
      Height          =   315
      ItemData        =   "test.frx":207E
      Left            =   75
      List            =   "test.frx":209D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2565
      Width           =   2130
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   465
      Left            =   15
      TabIndex        =   0
      Top             =   1470
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   820
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Text Color:"
      Height          =   390
      Left            =   75
      TabIndex        =   11
      Top             =   3750
      Width           =   525
   End
   Begin VB.Label Label4 
      Caption         =   "Color2:"
      Height          =   210
      Left            =   75
      TabIndex        =   6
      Top             =   3225
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Color1:"
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   2955
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Advanced ProgresBar:"
      Height          =   195
      Left            =   30
      TabIndex        =   2
      Top             =   465
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "ProgressBar:"
      Height          =   225
      Left            =   30
      TabIndex        =   1
      Top             =   1245
      Width           =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'XP UPGRADE AUTO-INSERT
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub Command1_Click()
CD.FileName = ""
CD.ShowOpen
If CD.FileName = "" Then Exit Sub
APB.CustomFrame = LoadPicture(CD.FileName)
End Sub

Private Sub Check1_Click()
Select Case Check1.Value
Case 1 'On
APB.ShowText = True
Case Else
APB.ShowText = False
End Select
End Sub

Private Sub cmbAppearance_Click()
Select Case cmbAppearance.ListIndex
Case 0
APB.Style = Standart
Case 1
APB.Style = Smooth
Case 2
APB.Style = DoubleColor
Case 3
APB.Style = SmoothDoubleColor
Case 4
APB.Style = ValueDependant
Case 5
APB.Style = XPStyle
Case 6
APB.Style = CustomPictureShow
Case 7
APB.Style = CustomPictureStrech
Case 8
APB.Style = CustomPictureTile
End Select
End Sub

Private Sub cmd_Click()
Select Case LCase(CStr(cmd.Caption))
Case "start"
tmrMove.Enabled = True
cmd.Caption = "Stop"
Case "stop"
tmrMove.Enabled = False
cmd.Caption = "Start"
End Select
End Sub

Private Sub cmdPAppearance_Click()
Select Case cmdPAppearance.ListIndex
Case 0
PB.Scrolling = ccScrollingStandard
Case 1
PB.Scrolling = ccScrollingSmooth
End Select
End Sub

Private Sub Command2_Click()
CD.FileName = ""
CD.ShowOpen
If CD.FileName = "" Then Exit Sub
Set APB.CustomPicture = LoadPicture(CD.FileName)
End Sub

Private Sub Form_Load()
APB.Value = 0
PB.Value = 0
APB.Max = 100
PB.Max = 100
APB.ShowText = False
cmbAppearance.ListIndex = 0
cmdPAppearance.ListIndex = 0
pC1.BackColor = APB.BarColor1
pC2.BackColor = APB.BarColor2
pt.BackColor = APB.TextColor
End Sub

Private Sub pC1_Click()
CD.ShowColor
pC1.BackColor = CD.Color
APB.BarColor1 = CD.Color
End Sub

Private Sub pC2_Click()
CD.ShowColor
pC2.BackColor = CD.Color
APB.BarColor2 = CD.Color
End Sub

Private Sub pt_Click()
CD.ShowColor
pt.BackColor = CD.Color
APB.TextColor = CD.Color
End Sub

Private Sub tmrMove_Timer()
Static valu As Integer
Static boo As Boolean
If Not boo Then
PB.Value = valu
APB.Value = valu
valu = valu + 1
If valu = 101 Then valu = valu - 1: boo = True
Else
PB.Value = valu
APB.Value = valu
valu = valu - 1
If valu = -1 Then valu = 0: boo = False
End If
End Sub

Private Sub Form_Initialize()
'XP UPGRADE AUTO-INSERT
InitCommonControls
End Sub
