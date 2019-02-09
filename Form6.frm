VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000012&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form6"
   ScaleHeight     =   8220
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   375
      Left            =   10080
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000012&
      Caption         =   "Hard"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   6120
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000012&
      Caption         =   "Medium"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000012&
      Caption         =   "Easy"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labellocation 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Choose your Battlefield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label from6 
      Height          =   15
      Left            =   2040
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label difficulty 
      Caption         =   "0"
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Select Difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 6 - Select Battle Options

Dim list As listmaps
Dim val As Integer
Dim templocation As String

Private Sub Command1_Click()
'' This command links with the 'back' button
val = 0
Form6.Hide
form2.Show
'' Returns the value in val to 0 then goes back to form2(users profile)
End Sub

Public Sub Command2_Click()
'' This command links with 'next' button

sel = List1.ListIndex + 1 '' +1 is needed as index starts at 0
'' This checks to see which 'map' has been selected in the lsitbox by using the index
If sel = 0 Then
labellocation = "C:\Battlefields\Earth (Default).dat"
End If
'' If the user has not selected any map in the listbox, a default map is selected so no errors occur

If sel > 0 Then
channum = FreeFile()
Open "C:\Battlefields\list.dat" For Random As channum Len = Len(list)
total = LOF(channum) / Len(list)
Get #channum, sel, list
templocation = Trim$(list.mappath)
Close #channum
labellocation = templocation
End If
'' When the user has selected a map, the file 'list.dat' is loaded and the appropriate map is selected based on user's choice

If Option1 = True Then
difficulty = 1
End If
If Option2 = True Then
difficulty = 2
End If
If Option3 = True Then
difficulty = 3
Else
difficulty = 0
End If
'' This checks to see which option the user has selected and assigns a value which will relate to difficulty in Form8(the game)
'' If no option is selected then the difficulty is automatically selected so no errors will occur

Form6.Hide
Form8.Show
'' After that, the form is hidden and then Form8 is shown
End Sub

Private Sub Command3_Click()
'' This command links with the 'quit' button
End
End Sub

Public Sub Form_Activate()
If val = 0 Then
listmaps
End If
'' the 'val' varibale appears here so that the sub is not reapted
End Sub

Public Sub Form_Load()
If val = 0 Then
listmaps
End If
'' the 'val' varibale appears here so that the sub is not reapted
End Sub

Sub listmaps()
'' The sub will show all maps that have been created using a seperate file that contains the names off all created maps
channum = FreeFile()
Open "C:\Battlefields\list.dat" For Random As channum Len = Len(list)
total = LOF(channum) / Len(list)
'' Opens up the list and gets the total number of records
For Counter = 1 To total
Get #channum, Counter, list
List1.AddItem Trim$(list.mapname)
Next
Close #channum
val = 1
'' After the list is displayed the val is returned as 1
End Sub
