VERSION 5.00
Begin VB.Form frmData2Word 
   Caption         =   "Create Word Document"
   ClientHeight    =   1425
   ClientLeft      =   210
   ClientTop       =   1725
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4245
   Begin VB.CommandButton cmdAdd2Word 
      Caption         =   "Create Word Document"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "E&xit"
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmData2Word"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd2Word_Click()

  Insert2Word
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Wrd = Nothing
    Set TheRange = Nothing
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
    Unload Me
End Sub
