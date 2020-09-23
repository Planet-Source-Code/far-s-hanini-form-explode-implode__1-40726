VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   6690
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me!"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call ImplodeForm(Me, 2, 500, 1)
    End
    Set Form1 = Nothing
End Sub

Private Sub Form_Load()
    Call ExplodeForm(Me, 500)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ImplodeForm(Me, 2, 500, 1)
End Sub


