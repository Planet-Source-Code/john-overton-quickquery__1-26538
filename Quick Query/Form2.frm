VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "SQL Server"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form2"
   ScaleHeight     =   885
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Server Name"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdok 
         Caption         =   "OK"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
If text1.Text = "" Then
    frmSQLServer.Caption = "Server Name Incorrect"
    Exit Sub
    
End If

ServerName = text1.Text
Unload Me

End Sub

Private Sub Command1_Click()

End Sub
