VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7575
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton create 
      Caption         =   "Create Map Using Locations"
      Height          =   735
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View a Map"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label labellocation 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   8160
      Visible         =   0   'False
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As listmaps
Dim templocation As String
Dim val As Integer


Public Sub Command1_Click()
val = 0
Form1.Hide
Form3.Show
End Sub

Public Sub Command2_Click()
sel = List1.ListIndex + 1
channum = FreeFile()
Open "C:\Battlefields\list.dat" For Random As channum Len = Len(z)
total = LOF(channum) / Len(z)
Get #channum, sel, z
templocation = Trim$(z.mappath)
Close #channum
labellocation = templocation
MsgBox ("File Loaded Sucessfully")
Form3.Show
Form1.Hide
End Sub

Public Sub Command3_Click()
Close #channum
End
End Sub
Public Sub create_Click()
val = 0
Form1.Hide
Form2.Show
End Sub

Public Sub Form_Activate()
If val = 0 Then
List1.Clear
listmaps
End If
End Sub

Public Sub Form_Load()
If val = 0 Then
List1.Clear
listmaps
End If
End Sub

Sub listmaps()
channum = FreeFile()
Open "C:\Battlefields\list.dat" For Random As channum Len = Len(z)
total = LOF(channum) / Len(z)
For Counter = 1 To total
Get #channum, Counter, z
List1.AddItem Trim$(z.mapname)
Next
Close #channum
val = 1
End Sub
