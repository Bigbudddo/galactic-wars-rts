VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' Highscore Table Maker

Dim module As hiscoretable

Public Sub Command1_Click()
channum = FreeFile()
Open "C:\hiscoretable.dat" For Random As channum Len = Len(module)
For Counter = 1 To 5
module.username = "AAA"
module.turns = 1
Put #channum, Counter, module
Next
MsgBox "Table sucessfully made"
End Sub
