VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00808080&
   Caption         =   "I.T. Researcher 2006"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9840
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":1272
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Favorites"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   12
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Artificial Intelligence"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Computer Science"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Software"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Hardware"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Data"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Graphics"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Security"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Internet and Online Services"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Open Source"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Networking"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Operating Systems"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Programming"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Extras"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Update"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Information"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Help"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "About"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00800000&
      Height          =   5685
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   2940
      TabIndex        =   1
      Top             =   660
      Width           =   3000
      Begin VB.ListBox LstSearch 
         Appearance      =   0  'Flat
         Height          =   7830
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   9000
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ALL CATEGORIES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6345
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2ACD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2B2B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2B8E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2BF8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2C666
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSys 
      Caption         =   "S&ystem"
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Web Update"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuFav 
      Caption         =   "&Favorites"
      Begin VB.Menu mnuAddFav 
         Caption         =   "&Add Favorite"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuList 
         Caption         =   "&View List"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuAllCat 
         Caption         =   "&All Categories"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAI 
         Caption         =   "Artificial Intelligence"
      End
      Begin VB.Menu mnuComsci 
         Caption         =   "Computer Science"
      End
      Begin VB.Menu mnuSoft 
         Caption         =   "Software"
      End
      Begin VB.Menu mnuHardware 
         Caption         =   "Hardware"
      End
      Begin VB.Menu mnuData 
         Caption         =   "Data"
      End
      Begin VB.Menu mnuGraphics 
         Caption         =   "Graphics"
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "Security"
      End
      Begin VB.Menu mnuInternet 
         Caption         =   "Internet and Online Services"
      End
      Begin VB.Menu mnuOpenSrc 
         Caption         =   "Open Source"
      End
      Begin VB.Menu mnuNet 
         Caption         =   "Networking"
      End
      Begin VB.Menu mnuOS 
         Caption         =   "Operating System"
      End
      Begin VB.Menu mnuProgramming 
         Caption         =   "Programming"
      End
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "&Extras"
      Begin VB.Menu mnuBuild 
         Caption         =   "&Build Your Own PC Guide"
      End
      Begin VB.Menu mnUTips 
         Caption         =   "&Tips and Tricks"
      End
      Begin VB.Menu mnuTutorials 
         Caption         =   "T&utorials"
         Begin VB.Menu mnuTutProg 
            Caption         =   "&Programming"
         End
         Begin VB.Menu mnuDatabase 
            Caption         =   "&Database Management"
         End
         Begin VB.Menu mnuWeb 
            Caption         =   "&Web Development"
         End
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Information"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileItem As New FileSystemObject

'API Declaration for reading the Listbox contents
Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" (ByVal _
        hWnd As Long, _
        ByVal wMsg As Integer, _
        ByVal wParam As String, _
        lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const LB_FINDSTRING = &H18F

Private Sub LstSearch_DblClick()
    item = LstSearch.Text
    If FileItem.FileExists(App.Path & "\Data\" & item & ".html") Then
        frmBrowser.Show
        frmBrowser.browser.Navigate (App.Path & "\Data\" & item & ".html")
        fav = item
    Else
        MsgBox "Sorry cannot find data referring to that item!", vbExclamation + vbOKOnly, "Note:"
    End If
End Sub

Private Sub MDIForm_Initialize()
If App.PrevInstance Then
   MsgBox "IT Researcher is already running.", vbExclamation + vbOKOnly, "IT Researcher 2006"
   End
   Exit Sub
End If
End Sub

Private Sub MDIForm_Load()
OpenCon
LoadAll LstSearch
Label3.Caption = "Search through " & rs.RecordCount & " items."
fav = ""
Load frmAssistant
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
con.Close
Unload frmAssistant
End
End Sub

Private Sub mnuAddFav_Click()
    If fav = "" Then
        MsgBox "No item to be added to favorites.", vbExclamation + vbOKOnly, "Error"
    Else
        frmAddFav.Show vbModal
    End If
End Sub

Private Sub mnuAI_Click()
LstSearch.Clear
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from AI", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        LstSearch.AddItem rs(0)
        .MoveNext
    Loop
End With
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "Artificial Intelligence"
rs.Close
End Sub

Private Sub mnuAllCat_Click()
LstSearch.Clear
LoadAll LstSearch
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "ALL CATEGORIES"
End Sub

Private Sub mnuComsci_Click()
LstSearch.Clear
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from Comsci", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        LstSearch.AddItem rs(0)
        .MoveNext
    Loop
End With
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "Computer Science"
rs.Close
End Sub

Private Sub mnuData_Click()
LstSearch.Clear
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from Data", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        LstSearch.AddItem rs(0)
        .MoveNext
    Loop
End With
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "Data"
rs.Close
End Sub

Private Sub mnuHardware_Click()
LstSearch.Clear
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from Hardware", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        LstSearch.AddItem rs(0)
        .MoveNext
    Loop
End With
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "Hardware"
rs.Close
End Sub

Public Sub mnuHelp_Click()
Call ShellExecute(0&, vbNullString, App.Path & "\Help\IT RESEARCHER 2006.HLP", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub mnuInternet_Click()
'Loads items belonging to Internet and Online Services Category
LstSearch.Clear
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from InternetOnline", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        LstSearch.AddItem rs(0)
        .MoveNext
    Loop
End With
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "Internet and Online Services"
rs.Close
End Sub

Private Sub mnuList_Click()
frmFavorites.Show
End Sub

Private Sub mnuSecurity_Click()
LstSearch.Clear
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from Security", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        LstSearch.AddItem rs(0)
        .MoveNext
    Loop
End With
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "Security"
rs.Close
End Sub

Private Sub mnuSoft_Click()
LstSearch.Clear
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from Software", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        LstSearch.AddItem rs(0)
        .MoveNext
    Loop
End With
Label3.Caption = "Search through " & rs.RecordCount & " items."
Label2.Caption = "Software"
rs.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err:
    Select Case Button.Index
        Case 1: mnuList_Click
        Case 3: frmExtras.Show
        Case 4: frmUpdate.Show
    End Select
err:
    Exit Sub
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Index
    Case 1:
        If ButtonMenu.Parent.Index = 2 Then mnuAI_Click
        'If ButtonMenu.Parent.Index = 5 Then mnuRQIR_Click
    Case 2:
        If ButtonMenu.Parent.Index = 2 Then mnuComsci_Click
        'If ButtonMenu.Parent.Index = 13 Then mnuRptB_Click
    Case 3:
        If ButtonMenu.Parent.Index = 2 Then mnuSoft_Click
    Case 4:
        If ButtonMenu.Parent.Index = 2 Then mnuHardware_Click
    Case 5:
        If ButtonMenu.Parent.Index = 2 Then mnuData_Click
    'Case 6:
     '   If ButtonMenu.Parent.Index = 2 Then mnuGraphics_Click
    Case 7:
        If ButtonMenu.Parent.Index = 2 Then mnuSecurity_Click
    Case 8:
        If ButtonMenu.Parent.Index = 2 Then mnuInternet_Click
    'Case 9:
     '   If ButtonMenu.Parent.Index = 2 Then mnuOpenSrc_Click
    'Case 10:
     '   If ButtonMenu.Parent.Index = 2 Then mnuNet_Click
    'Case 11:
    '    If ButtonMenu.Parent.Index = 2 Then mnuOS_Click
    'Case 12:
     '   If ButtonMenu.Parent.Index = 2 Then mnuProgramming_Click
End Select
End Sub

Private Sub txtSearch_Change()
LstSearch.ListIndex = SendMessage(LstSearch.hWnd, LB_FINDSTRING, _
                      txtSearch, ByVal txtSearch.Text)

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
        LstSearch_DblClick
End Select
End Sub
