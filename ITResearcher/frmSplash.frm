VERSION 5.00
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akProgressBar.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   6750
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin akProgress.akProgressBar bar 
      Height          =   255
      Left            =   2085
      TabIndex        =   0
      Top             =   5280
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      BackColour      =   16777215
      FontColour      =   0
      BarColour       =   8388608
      Horizontal      =   -1  'True
      ReverseGradient =   0   'False
      Max             =   100
      Min             =   0
      GapWidth        =   2
      LineWidth       =   3
      Caption         =   1
      BorderStyle     =   0
      Margin          =   2
      Gradient        =   7
      Alignment       =   2
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   3480
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      Height          =   255
      Left            =   2865
      TabIndex        =   1
      Top             =   4200
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
mdiMain.Show
End Sub

Private Sub Timer1_Timer()
bar.Value = bar.Value + 1

If bar.Value = 20 Then
lblLoad.Caption = "Loading Database..."
ElseIf bar.Value = 40 Then
lblLoad.Caption = "Initializing Components..."
ElseIf bar.Value = 70 Then
lblLoad.Caption = "Loading Components..."
ElseIf bar.Value = 90 Then
lblLoad.Caption = "Please wait..."
ElseIf bar.Value = bar.Max Then
Unload Me
End If
End Sub
