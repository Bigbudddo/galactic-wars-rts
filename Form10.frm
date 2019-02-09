VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   Caption         =   "Form10"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form10"
   ScaleHeight     =   2205
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton load 
      Caption         =   "load"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox list 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 10 - Load Game

Dim savelist As save
Dim val As Integer

Public Sub close_Click()
Form10.Hide '' Hides the Form and goes back to Form7
val = 0 '' Resets val before Form closes
End Sub

Public Sub Form_Activate()
If val = 0 Then
listsaves
'' When the form activates or loads then if the value in val is 0 then the program will use the listsaves sub
End If
End Sub

Public Sub Form_Load()
If val = 0 Then
listsaves
'' When the form activates or loads then if the value in val is 0 then the program will use the listsaves sub
End If
End Sub

Sub listsaves()
channum = FreeFile()
Open "C:\savedgame.dat" For Random As channum Len = Len(savelist) '' Opens the file
totalrecords = LOF(channum) / Len(savelist) '' Gets the total number of records in the file
For Counter = 1 To totalrecords
Get #channum, Counter, savelist
list.AddItem savelist.savename
'' For each record in the file, the program will display the name part of the record
Next
Close #channum '' Closes the file
val = 1 '' Sets val to 1 so the sub doesn't repeat
End Sub
