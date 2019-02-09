VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8025
   ClientLeft      =   270
   ClientTop       =   660
   ClientWidth     =   11610
   LinkTopic       =   "Form3"
   ScaleHeight     =   8025
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7680
      TabIndex        =   31
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton strength_plus 
      Caption         =   "+"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton health_plus 
      Caption         =   "+"
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton health_minus 
      Caption         =   "-"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Finished"
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   6960
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000012&
      Caption         =   "Bad"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8640
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000012&
      Caption         =   "Good"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Look"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   240
      ScaleHeight     =   4155
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Attributes"
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   4680
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton leadership_plus 
         Caption         =   "+"
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton leadership_minus 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton armour_plus 
         Caption         =   "+"
         Height          =   375
         Left            =   1440
         TabIndex        =   24
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton armour_minus 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton strength_minus 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label armour 
         BackColor       =   &H00000000&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   29
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label leadership 
         BackColor       =   &H80000012&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000012&
         Caption         =   "Leadership"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Armour"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label strength 
         BackColor       =   &H80000012&
         Caption         =   "X"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Strength"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label health 
         BackColor       =   &H00000000&
         Caption         =   "X"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Health"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   6240
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   5415
      TabIndex        =   32
      Top             =   5520
      Width           =   5415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Enter a Password (No More than 15 Characters"
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
      Left            =   7680
      TabIndex        =   30
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "Choose you Side"
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
      Left            =   7680
      TabIndex        =   10
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Enter Player Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "points remaining."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label remaining 
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "You have"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Create new Commander"
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
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 3 - Create a User
Dim usercreater As users
Dim points As Integer
Dim choosenside As Integer
Dim appearanceset As Integer
Dim validation As Integer

Public Sub armour_minus_Click()
validation = 0
Validate_minus
If validation = 0 Then
remaining = remaining + 1
armour = armour - 1
''If the validation is sucessful then the point will be take away
End If
'' The validation comes from the validate sub's depending on if you are taking a point away or adding a point on
End Sub

Public Sub armour_plus_Click()
validation = 0
Validate_plus
If validation = 0 Then
remaining = remaining - 1
armour = armour + 1
End If
End Sub

Public Sub Command1_Click()
'' Command one links with the button on Form3 labeled "change look"
appearanceset = appearanceset + 1
If appearanceset = 5 Then
appearanceset = 1
End If
'' As there is only 5 different pictures to select, once it reaches the fifth, if the user goes further, it goes back to picture 1

If appearanceset = 1 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\1.jpg"
End If

If appearanceset = 2 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\2.jpg"
End If

If appearanceset = 3 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\3.jpg"
End If

If appearanceset = 4 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\4.jpg"
End If

Picture1.picture = LoadPicture(pic)
'' Depending on the 'appearanceset' value will display the appropriate picture to the user
End Sub

Public Sub Command2_Click()
'' This command links with the 'finished' button on Form3

validation = 0
'' The validation variable needs to be reset back to 0

If Text1 = "" Then
MsgBox "Please enter a username in order to continue"
validation = 1
End If

If Text2 = "" Then
MsgBox "Invalid Password"
validation = 1
End If

If Option1 = False And Option2 = False Then
MsgBox "Please choose a side in order to continue"
validation = 1
End If

If appearanceset = 0 Then
MsgBox "Please Choose a Commander"
validation = 1
End If
'' The two 'if' statements, are input validation so the user can't have a blank username, or a blank password
'' The other two 'if' statements are again input validation so the user can't have a blank picture or no 'side' choosen

If Option1 = True Then
choosenside = 1
End If

If Option2 = True Then
choosenside = 2
End If
'' Depends on option the user has checked on the form, will assign a value for the user's 'side'

If validation = 0 Then
'' Only when all the input validation has passed, will the file be saved

channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(usercreater)
total = LOF(channum) / Len(usercreater)
For Counter = total + 1 To total + 1
usercreater.id_number = total + 1
usercreater.username = Text1
usercreater.password = Text2
usercreater.side = choosenside
usercreater.wins = 0
usercreater.losses = 0
usercreater.Commander = appearanceset
usercreater.ah = health
usercreater.as = strength
usercreater.aa = armour
usercreater.al = leadership
usercreater.points = remaining
Put #channum, Counter, usercreater
'' All information the user has entered on the form will be saved into a file called 'userprofiles'
Next
Close #channum
'' The file will then be closed
form1.Show
Form3.Hide
'' Once the file is saved, the user is taken to the main menu, ready to log in
End If

End Sub

Public Sub Command3_Click()
form1.Show
Form3.Hide
End Sub

Private Sub Form_Activate()
defaults
End Sub

Public Sub Form_Load()
defaults
End Sub

Public Sub health_minus_Click()
validation = 0
Validate_minus
If validation = 0 Then
remaining = remaining + 1
health = health - 1
End If
End Sub

Public Sub health_plus_Click()
validation = 0
Validate_plus
If validation = 0 Then
remaining = remaining - 1
health = health + 1
End If
End Sub

Public Sub leadership_minus_Click()
validation = 0
Validate_minus
If validation = 0 Then
remaining = remaining + 1
leadership = leadership - 1
End If
End Sub

Public Sub leadership_plus_Click()
validation = 0
Validate_plus
If validation = 0 Then
remaining = remaining - 1
leadership = leadership + 1
End If
End Sub

Public Sub strength_minus_Click()
validation = 0
Validate_minus
If validation = 0 Then
remaining = remaining + 1
strength = strength - 1
End If
End Sub

Public Sub strength_plus_Click()
validation = 0
Validate_plus
If validation = 0 Then
remaining = remaining - 1
strength = strength + 1
End If
End Sub

Sub Validate_plus()

If remaining <= 0 Then
MsgBox "No more Points Remaining"
validation = 1
End If
'' This checks to see if there is any 'skill points' remaining for the user to use and if not will not allow the user
'' to add anymore points on in the other sub's

End Sub

Sub Validate_minus()

If remaining >= 40 Then
MsgBox "Cannot take anymore points off"
validation = 1
End If
'' This checks to see if there is no 'skill points' remaining for the user to take off the 'stats' of the commander and if not will not allow the user
'' to take off anymore points on in the other sub's


End Sub

Sub defaults()
Picture1.picture = LoadPicture("")
Text1 = ""
Text2 = ""
appearanceset = 0
remaining = 40
health = 0
strength = 0
armour = 0
leadership = 0
'' When the user comes to the form, if a user has already been created earlier, all the data needs to be cleared before a new user can be made
End Sub
