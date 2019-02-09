VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00000000&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   660
   ClientWidth     =   11820
   LinkTopic       =   "Form9"
   ScaleHeight     =   8040
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   240
      ScaleHeight     =   4155
      ScaleWidth      =   2595
      TabIndex        =   24
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Attributes"
      ForeColor       =   &H8000000E&
      Height          =   4095
      Left            =   9240
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
      Begin VB.TextBox password 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "If you do not wish to change the password, please leave this."
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton leadership_plus 
         Caption         =   "+"
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton leadership_minus 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton armour_plus 
         Caption         =   "+"
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton armour_minus 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton strength_plus 
         Caption         =   "+"
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton strength_minus 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton health_plus 
         Caption         =   "+"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton health_minus 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.Label label3 
         BackColor       =   &H80000012&
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label leadership 
         BackColor       =   &H80000012&
         Caption         =   "X"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000012&
         Caption         =   "Leadership"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label armour 
         BackColor       =   &H80000012&
         Caption         =   "X"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
         Caption         =   "Armour"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label strength 
         BackColor       =   &H80000012&
         Caption         =   "X"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000012&
         Caption         =   "Strength"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label health 
         BackColor       =   &H80000012&
         Caption         =   "X"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         Caption         =   "Health"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   615
      Left            =   9720
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   6720
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   5175
      TabIndex        =   25
      Top             =   5520
      Width           =   5175
   End
   Begin VB.Label username 
      BackColor       =   &H80000012&
      Caption         =   "COMMANDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "points remaining"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label remaining 
      BackColor       =   &H80000012&
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   10200
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "You currently have"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8760
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Commander Customise"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 9 - Customise Commander
'' Form9 can be shown from Form2

Dim profiles As users
Dim validation As Integer

Public Sub armour_minus_Click()
validation = 0
Validate_minus
If validation = 0 Then
remaining = remaining + 1
armour = armour - 1
End If
End Sub

Public Sub armour_plus_Click()
validation = 0
Validate_plus
If validation = 0 Then
remaining = remaining - 1
armour = armour + 1
End If
End Sub

Public Sub Command2_Click()
'' Command2 links with the 'back' button on Form9
validation = 0
'' Validation needs to be set back to 0 before continuing

If password = "" Then
MsgBox "A Password, Must be entered before proceeding"
validation = 1
'' Checks to see if the password box is empty, if it is the user is notified and the validation is set to 1
End If

If validation = 0 Then
save
Form9.Hide
form2.Show
'' When the validation is 0 then the 'save' sub is used then Form9 is hidden and form2 is shown
End If
End Sub

Public Sub Form_Activate()
display_stats
'' When the form activates or loads then the 'display_stats' sub is used
End Sub

Public Sub Form_Load()
display_stats
'' When the form activates or loads then the 'display_stats' sub is used
End Sub

Sub display_stats()
'' display_stats is linked with 'Form_Load' sub and 'Form_Activate' sub
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(profiles)
Get #channum, form1.selected, profiles
'' The file userprofiles is loaded and the appropriate record is selected based on what the user selected in Form1
remaining = profiles.points
health = profiles.ah
strength = profiles.as
armour = profiles.aa
leadership = profiles.al
username = Trim$(profiles.username)
password = Trim$(profiles.password)
'' Temporary files are loaded with the information from that record

If profiles.Commander = 1 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\1.jpg"
End If
If profiles.Commander = 2 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\2.jpg"
End If
If profiles.Commander = 3 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\3.jpg"
End If
If profiles.Commander = 4 Then
pic = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\4.jpg"
End If
'' Based on the value in the 'Commander' part of the record will asgin 'pic' a specific path to a picture

Picture1.picture = LoadPicture(pic)
'' The picture is then loaded in the Picturebox (Picture1)

Close #channum '' Closes file
End Sub

Public Sub health_minus_Click()
validation = 0
Validate_minus
If validation = 0 Then
remaining = remaining + 1
health = health - 1
''If the validation is sucessful then the point will be take away
End If
'' The validation comes from the validate sub's depending on if you are taking a point away or adding a point on
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

Sub save()
'' 'save' sub links with Command2 ('back' button)
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(profiles) '' Opens the file
profiles.points = remaining
profiles.ah = health
profiles.as = strength
profiles.aa = armour
profiles.al = leadership
profiles.password = password
'' The information in the record is overwritten by the temporary variables that the 'old' data was written into
Put #channum, form1.selected, profiles
Close #channum '' closes file
End Sub
