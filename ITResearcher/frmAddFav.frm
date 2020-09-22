VERSION 5.00
Begin VB.Form frmAddFav 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Item to Favorites"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   Icon            =   "frmAddFav.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Description for This Page"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmAddFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim description As String
If rsFav.State <> adStateClosed Then rsFav.Close
rsFav.Open "select * from Favorites where code ='" & item & "'", con, adOpenKeyset, adLockOptimistic
    If rsFav.EOF Then
        description = txtDesc.Text
        rsFav.AddNew
        rsFav.Fields(0) = item
        rsFav.Fields(1) = description
        rsFav.Update
        MsgBox "Item added to list.", vbInformation + vbOKOnly, "Add to Favorites"
        Unload Me
    Else
        MsgBox "Item already exists in the list!", vbExclamation + vbOKOnly, "Add to Favorites"
        Unload Me
    End If
rsFav.Close
End Sub
