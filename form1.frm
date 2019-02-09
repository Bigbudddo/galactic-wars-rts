VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00000000&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      Picture         =   "form1.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   5295
      TabIndex        =   10
      Top             =   5880
      Width           =   5295
   End
   Begin VB.CommandButton about 
      Caption         =   "About"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton exit_button 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton create 
      Caption         =   "Create New User"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton login_button 
      Caption         =   "Login Exsisting User"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox user_select 
      Height          =   315
      ItemData        =   "form1.frx":2167
      Left            =   3600
      List            =   "form1.frx":2169
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "user_select"
      ToolTipText     =   "Please Select your Username"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   6480
      Picture         =   "form1.frx":216B
      ScaleHeight     =   2535
      ScaleWidth      =   5325
      TabIndex        =   5
      Top             =   5760
      Width           =   5325
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Galactic Wars"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   2640
      TabIndex        =   9
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Stuart James"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Battle for Earth"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label selected 
      Caption         =   "HELLO!"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "I'm Invisible :P"
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 1 - Main Menu
Dim qprofile As users
Dim validation As Integer
Dim playernames(900) As String
Dim totalrecords As String

Private Sub about_Click()
'' Brings up small about box, Showing details of game
Form4.Show
End Sub

Public Sub create_Click()
'' Takes user to another form were they can create a profile
form1.Hide
Form3.Show
End Sub

Public Sub exit_button_Click()
'' Closes the game
Close #channum
End
End Sub

Public Sub Form_Activate()
'' When the Form reactivates then it will close any open file, reset user list
'' and then re-read the file
Close #channum
user_select.Clear
read_file
End Sub

Public Sub Form_Load()
'' When the Form starts up for the first time it will go to the read file sub
read_file
End Sub

Sub read_file()
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(qprofile)
totalrecords = LOF(channum) / Len(qprofile)
'' Opens the file "Userprofiles" and gets the total number of records in
'' the file
For Counter = 1 To totalrecords
Get #channum, Counter, qprofile
playernames(Counter) = Trim$(qprofile.username)
user_select.AddItem playernames(Counter)
Next
'' Goes through each record and places the name of each profile from the
'' record into a Combobox
Close #channum
End Sub

Public Sub login_button_Click()

validation = 0

If user_select = "" Then
MsgBox "Please Select a Username in order to continue"
validation = 1
End If
'' This validation bit checks that the user has selected a name from the
'' Combobox

If validation = 0 Then

For Counter = 1 To totalrecords
If playernames(Counter) = user_select Then
selected = Counter
'' This will check all the names in the file with the one the user selected
'' the position of the selected username is placed in "selected"
End If
Next


passval = 0
temppass = InputBox("Enter Password")
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(qprofile)
Get #channum, selected, qprofile
If temppass = Trim$(qprofile.password) Then
passval = 1
Else
MsgBox ("Invalid Password")
End If
'' This asks the User for a password for the profile they selected
'' the user types in a password and it will be checked with the
'' password in the file. If it's wrong, the user is notified

If passval = 1 Then
fromgame = 1
form1.Hide
form2.Show
'' If the user enters the correct password then the user is taken
'' to Form2.
End If

End If
End Sub

