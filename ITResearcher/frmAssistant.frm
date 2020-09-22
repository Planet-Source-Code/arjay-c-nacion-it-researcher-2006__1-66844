VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmAssistant 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin AgentObjectsCtl.Agent msaAssistant 
      Left            =   4080
      Top             =   2520
   End
End
Attribute VB_Name = "frmAssistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VR As AgentObjectsCtl.IAgentCtlCharacter
Dim CommandList As New Collection

Private Sub Form_Load()
Dim s As Variant
    Dim i As String
    msaAssistant.Characters.Load "VR", "vrgirl.acs"
    Set VR = msaAssistant.Characters("VR")
    VR.Show
    VR.MoveTo 600, 500, 300
    VR.Speak "Welcome to I.T Researcher 2006"
   
    VR.Commands.Add "Help", "Help", , True, True
    VR.Commands.Add "Time", "Tell Me the Time", , True, True
End Sub

Private Sub msaAssistant_Command(ByVal UserInput As Object)
Dim arrstrNow As Variant
  Dim arrstrTime As Variant
  Dim strOutput As String
  Dim s As Variant
  Dim strName As String * 255
  Dim hWnd As Long
  
  Select Case UserInput.Name
    Case "Time":
        arrstrNow = Split(Now, " ")
        arrstrTime = Split(arrstrNow(1), ":")
        strOutput = "Time right now is " & arrstrTime(0) & " " & arrstrNow(2) & " with " & arrstrTime(1) & IIf(Val(arrstrTime(1)) > 1, " minutes", "minute")
        VR.Speak strOutput
    Case "Help":
        VR.Speak "Opening help file"
        mdiMain.mnuHelp_Click
        'frmtour.Show
  End Select
End Sub
