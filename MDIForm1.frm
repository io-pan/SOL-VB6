VERSION 5.00
Begin VB.MDIForm MDIf 
   BackColor       =   &H8000000C&
   Caption         =   "Solaire"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10095
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "MDIf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ferm As Boolean

Private Sub Command1_Click()
'If Picture1.Height = 200 Then
'Picture1.Height = 500
'Else
'Picture1.Height = 200
'End If

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub
  




Private Sub MDIForm_Load()
    ferm = False
    
    Set fGbl = New fGbl
    fGbl.Show
    fGbl.WindowState = 2
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
ferm = True
End Sub
