VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form5"
   ScaleHeight     =   8265
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.ListBox List3 
      Height          =   4155
      Left            =   5520
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton findplayer 
      Caption         =   "Find Player"
      Height          =   315
      Left            =   9480
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox searchbox 
      Height          =   285
      Left            =   9000
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Return to Command HQ"
      Height          =   495
      Left            =   9600
      TabIndex        =   7
      Top             =   7560
      Width           =   2055
   End
   Begin VB.CommandButton sortdecending 
      Caption         =   "Sort Decending"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton sortascending 
      Caption         =   "Sort Ascending"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   4155
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   6720
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   5175
      TabIndex        =   10
      Top             =   5760
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Cheated?"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Score (No, Turns)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Commanders who have Conquered"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 5 - Highscore Table
Dim tprofile As users
Dim module As hiscoretable
Dim tempscore As Integer
Dim tempname As String
Dim tempcheat As String
Dim stopper As Integer
Dim searchnames(200)
Dim storenames(200) As String
Dim storescores(200) As Integer
Dim storecheated(200) As String
Dim dummyname As String
Dim dummyscore As Integer
Dim dummycheated As String
Dim records As Integer
Dim trecords As Integer
Public Sub Command4_Click()
'' Links with "Return to Command HQ" button
form2.from2 = 0
'' Value needs to be assigned for validation otherwise the form will try write to the highscore table file
form2.Show
Form5.Hide
'' Hides this current form and then shows Form2
End Sub

Public Sub findplayer_Click()
'' Count Occurences and Search algorithm
counts = 0
flag = 0
'' Sets variables back to 0

channum = FreeFile()
Open "C:\hiscoretable.dat" For Random As channum Len = Len(module)
records = LOF(channum) / Len(module)
For Counter = 1 To records
Get #channum, Counter, module
searchnames(Counter) = Trim$(module.username)
Next
Close #channum
'' Opens up the file, gets all the names from the file and puts them into an array. The file is then closed

For Counter = 1 To records
'' 'records' is taken from when the file is loaded, it gets the total number of records in the file.
'' Going from 1 to records means that it will loop for each record in the file
If searchbox = searchnames(Counter) Then
flag = 1
counts = counts + 1
'' When the search finds a name in the array matching what the user is searching it will set a flag to let the user know it has found
'' what the user is searching for and then add one onto the number of times it has found what the user is searching
If flag = 1 Then
MsgBox ("User " & searchbox & " was found at position " & Counter)
flag = 0
'' When the flag=1 then the user is told that the search has found what the user was searching for and at what position in the list
End If
End If
Next

If counts > 0 Then
MsgBox ("There was " & counts & " count(s) of that the username " & searchbox & " found in the highscore table")
Else
MsgBox ("The username could not be found")
End If
'' If the counts are greater than zero then the user is told the number of times the name the user searched for is found
'' If the counts is less than zero the user is told that there is no name that the user searched in the list

End Sub

Public Sub Form_Activate()
default
'' When the form is loaded or activated it runs the 'deafult' sub
End Sub

Public Sub Form_Load()
default
'' When the form is loaded or activated it runs the 'deafult' sub
End Sub

Sub default()
Text1 = ""

If Form8.fromgame = 2 Then
'' fromgame in form8 will return a value of 2 if the user has won the game
If stopper = 0 Then
'' The stopper varibale is used so that it wont repeat this step until the user has come from the right place (e.g- Won in the game in Form8)
win_loss
scoreintable
display
'' If the stopper is 0 and fromgame is 2 then the 'win_loss' and 'scoreintable' sub's are used
End If
End If

If Form8.fromgame = 3 Then
'' fromgame in form8 will return a value of 3 if the user has lost the game
If stopper = 0 Then
'' The stopper varibale is used so that it wont repeat this step until the user has come from the right place (e.g- Lost in the game in Form8)
win_loss
display
'' If the stopper is 0 and fromgame is 3 then the 'win_loss' and 'display' sub's are used
End If
End If

If form2.from2 = 1 Then
'' from2 is assigned as 1 if the user has come to this form from form2
display
End If

End Sub

Sub scoreintable()
'' This sub will write to the file if the user won the game in Form8
tempname = ""
tempscore = 0
tempcheat = ""
'' Returns all temporary variables to blank ready to be used
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(tprofile)
Get #channum, form1.selected, tprofile
tempname = tprofile.username
tempscore = Form8.turns
tempcheat = Form8.cheated
'' Opens up the file and gets the appropriate record based on the what was selected in Form1
'' The values of the teporary varibales change to the information of the selected record
Close #channum
channum = 0
'' The value is then closed and the channum returned to 0
channum = FreeFile()
Open "C:\hiscoretable.dat" For Random As channum Len = Len(module)
records = LOF(channum) / Len(module)
'' Another file is then opened and the total number of records is gotten
For c = records + 1 To records + 1
'' records+1 to records+1 means that a new record is added after the last record
module.username = tempname
module.turns = tempscore
module.cheated = tempcheat
Put #channum, c, module
'' The appropriate data is written into the record and saved
Next
Close #channum
'' The file is then closed
End Sub

Sub display()
List1.Clear
List2.Clear
List3.Clear
'' When the display sub is activated, the three listboxes are cleared
channum = FreeFile()
Open "C:\hiscoretable.dat" For Random As channum Len = Len(module)
records = LOF(channum) / Len(module)
'' The file hiscoretable is opened and the total number of records is taken
trecords = records
'' trecords is the varibale for how many records are in the file, this is used for sorting
For d = 1 To records
Get #channum, d, module
storenames(d) = Trim$(module.username)
storescores(d) = module.turns
storecheated(d) = Trim$(module.cheated)
'' For each record in the file the command will loop adding the information to temporary arrays which is used for searching and sorting
List1.AddItem module.username
List2.AddItem module.turns
List3.AddItem module.cheated
'' Then the information is added to the correct listbox
Next
Close #channum
End Sub

Sub win_loss()
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(tprofile)
Get #channum, form1.selected, tprofile
'' The file is opened and the correct record is selected depending on the users choice in Form1 (profile selecting)

If Form8.fromgame = 2 Then
wins = tprofile.wins + 1
''If the user has won the game in form8 then a temporary varibale is used to hold the total number of wins the user's profile has
End If

If Form8.fromgame = 3 Then
losses = tprofile.losses + 1
''If the user has lost the game in form8 then a temporary varibale is used to hold the total number of losses the user's profile has
End If

If Form8.fromgame = 2 Then
tprofile.wins = wins
Put #channum, form1.selected, tprofile
''If the user has won the game in form8 then the temporary varibale will be saved into the record in the file
End If

If Form8.fromgame = 3 Then
tprofile.losses = losses
Put #channum, form1.selected, tprofile
''If the user has lost the game in form8 then the temporary varibale will be saved into the record in the file
End If

Close #channum
stopper = 1
'' The stopper varibale is given a value to stop the sub from repeating
End Sub

Public Sub sortascending_Click()
List1.Clear
List2.Clear
List3.Clear
'' The lists needs to clear before sorting
For Counter = 1 To trecords
'' Loops for how many records there are in the file
dummyscore = 0
dummyname = ""
dummycheated = ""
'' The values in the varibales are turned to blank ready for sorting
If storescores(Counter + 1) > storescores(Counter) Then
dummyscore = storescores(Counter)
storescores(Counter) = storescores(Counter + 1)
storescores(Counter + 1) = dummyscore
'' If the If condition is true then the values are placed into 'dummy' variables and swapped by overwriting the values in each varibale
dummyname = storenames(Counter)
storenames(Counter) = storenames(Counter + 1)
storenames(Counter + 1) = dummyname
'' The swapping is also done for the username's
dummycheated = storecheated(Counter)
storecheated(Counter) = storecheated(Counter + 1)
storecheated(Counter + 1) = dummycheated
'' And then the values in the cheated variable are swapped
End If
Next
For Counter = 1 To trecords
List1.AddItem storenames(Counter)
List2.AddItem storescores(Counter)
List3.AddItem storecheated(Counter)
'' Once the values are swapped then the program can put them back into the listboxes to be displayed sorted
Next
End Sub

Public Sub sortdecending_Click()
List1.Clear
List2.Clear
List3.Clear
'' The lists needs to clear before sorting
For Counter = 1 To trecords
'' Loops for how many records there are in the file
dummyscore = 0
dummyname = ""
dummycheated = ""
'' The values in the varibales are turned to blank ready for sorting
If storescores(Counter + 1) < storescores(Counter) Then
dummyscore = storescores(Counter)
storescores(Counter) = storescores(Counter + 1)
storescores(Counter + 1) = dummyscore
'' If the If condition is true then the values are placed into 'dummy' variables and swapped by overwriting the values in each varibale
dummyname = storenames(Counter)
storenames(Counter) = storenames(Counter + 1)
storenames(Counter + 1) = dummyname
'' The swapping is also done for the username's
dummycheated = storecheated(Counter)
storecheated(Counter) = storecheated(Counter + 1)
storecheated(Counter + 1) = dummycheated
'' And then the values in the cheated variable are swapped
End If
Next
For Counter = 1 To trecords
List1.AddItem storenames(Counter)
List2.AddItem storescores(Counter)
List3.AddItem storecheated(Counter)
'' Once the values are swapped then the program can put them back into the listboxes to be displayed sorted
Next
End Sub
