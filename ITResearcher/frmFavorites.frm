VERSION 5.00
Begin VB.Form frmFavorites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favorites"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox lstfav 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "List of Favorites"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2325
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1725
      Picture         =   "frmFavorites.frx":1272
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If rsFav.State <> adStateClosed Then rsFav.Close
rsFav.CursorLocation = adUseClient
rsFav.Open "Select * from favorites", con, adOpenStatic, adLockReadOnly
With rsFav
    Do While Not .EOF
        lstfav.AddItem rsFav(1)
        .MoveNext
    Loop
End With
End Sub

Private Sub lstfav_DblClick()
On Error GoTo Errhandler:
    Dim cod As String
    If frmBrowser.Visible = False Then frmBrowser.Show: Exit Sub
    If rsFav.State <> adStateClosed Then rsFav.Close
    rsFav.Open "Select * from Favorites where Desc ='" & lstfav.Text & "'", con, adOpenStatic, adLockReadOnly
    cod = rsFav(0)
    frmBrowser.browser.Navigate (App.Path & "\Data\" & cod & ".html")
    rsFav.Close
    Unload Me
Errhandler:
    If err.Number = 0 Then
    Exit Sub
    Else
    MsgBox "Unable to find that item.", vbCritical + vbOKOnly, "Error"
    End If
End Sub
