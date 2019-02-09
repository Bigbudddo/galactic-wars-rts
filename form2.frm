VERSION 5.00
Begin VB.Form form2 
   BackColor       =   &H00000000&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8160
   ClientLeft      =   270
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   4155
      ItemData        =   "form2.frx":0000
      Left            =   4920
      List            =   "form2.frx":0002
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton logout_button1 
      Caption         =   "Logout"
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton highscore 
      Caption         =   "Highscores"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton start_button 
      Caption         =   "New Battle"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Commander 
      Caption         =   "Commander"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   7560
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6600
      Picture         =   "form2.frx":0004
      ScaleHeight     =   2415
      ScaleWidth      =   5295
      TabIndex        =   7
      Top             =   5880
      Width           =   5295
   End
   Begin VB.Label from2 
      Height          =   15
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Welcome To Duty Commander"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 2 - Userprofile Page
Dim userprofile As users
Dim total As Integer

Public Sub Commander_Click()
'' The commander button takes the user to Form 9 which allows the user to customise
'' there commanders attributes
Close #channum
'' Closes the file opened when the Form activated
Form9.Show
form2.Hide
End Sub

Private Sub Form_Activate()
Form7.menulabel = "2"
'' When the Form activates/loads, It will set a label in Form 7 to 2. This is for
'' validation in Form 8
read_file
'' The file is then read
End Sub

Public Sub Form_Load()
Form7.menulabel = "2"
'' When the Form activates/loads, It will set a label in Form 7 to 2. This is for
'' validation in Form 8
read_file
'' The file is then read
End Sub

Public Sub highscore_Click()
'' The command takes the user to the Highscore table Form. (Form5)
from2 = 1
'' This is used for Validation purposes, so the highscoretable file will no be overwritten in Form5
Close #channum
'' Closes any open file
Form5.Show
form2.Hide
End Sub

Public Sub logout_button1_Click()
'' This logs the user out of the profile and back to main menu
Close #channum
'' Closes any open file
List1.Clear
form1.Show
form2.Hide
'' The listbox for the commanders stats is then cleared for the next person to log in
End Sub

Sub read_file()
'' The read file sub, reads the file and then the record selected in Form1
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(userprofile)
toterecords = LOF(channum) / Len(userprofile)
Get #channum, form1.selected, userprofile
'' Opens up the file "Userprofiles" and then goes to the record position found in Form1. (selected = counter)
display_stats
'' Goes to display stats sub
End Sub

Sub display_stats()
List1.Clear
List1.AddItem "Username :-"
List1.AddItem userprofile.username
List1.AddItem ""
List1.AddItem "Wins :-"
List1.AddItem userprofile.wins
List1.AddItem "Losses :-"
List1.AddItem userprofile.losses
List1.AddItem ""
List1.AddItem "Commander Attributes :-"
List1.AddItem userprofile.ah
List1.AddItem userprofile.as
List1.AddItem userprofile.aa
List1.AddItem userprofile.al
List1.AddItem ""
List1.AddItem "Commander Ability"
'' The following list of commands, will reset the list and then display all the players stats
'' The list needs to be cleared first as another user's stats will still be there after they log off

If userprofile.Commander = 1 Then
List1.AddItem "Royal Wealth"
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\1.jpg"
End If
If userprofile.Commander = 2 Then
List1.AddItem "Increased Weapon Skills"
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\2.jpg"
End If
If userprofile.Commander = 3 Then
List1.AddItem "Power to the People"
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\3.jpg"
End If
If userprofile.Commander = 4 Then
List1.AddItem "Increased Number of Troops"
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\4.jpg"
End If
'' This adds a display onto the players list depends on what character they chose when creating there
'' profile

Picture1.picture = LoadPicture(pic)
'' Displays the user's commander picture

End Sub

Public Sub start_button_Click()
'' When this button is clicked, All open files are closed and the user is taken to Form6, where they select their map, etc
from2 = 0
Close #channum
Form6.Show
form2.Hide
End Sub
