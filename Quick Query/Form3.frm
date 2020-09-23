VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add SQL"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

If Text1.Text = "" Then
    frmSQLServer.Caption = "Nothing to Add"
    Exit Sub
    
End If
AddSQL = Text1.Text
Unload Me


End Sub
