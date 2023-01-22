VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim A(5) As Integer
Dim B(5) As Integer
Dim i, j
For i = 1 To 5
A(i) = Val(InputBox("Enter elements of first Array"))
Next
For i = 1 To 5
Print (A(i))
Next
For j = 1 To 5
B(j) = Val(InputBox("Enter elements of second Array"))
Next
For j = 1 To 5
Print (B(j))
Next
End Sub
