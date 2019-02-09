VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00000000&
   Caption         =   "Battlemenu"
   ClientHeight    =   4455
   ClientLeft      =   4635
   ClientTop       =   2340
   ClientWidth     =   3195
   LinkTopic       =   "Form7"
   ScaleHeight     =   4455
   ScaleWidth      =   3195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton load 
      Caption         =   "Load"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "Feature not added yet"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton savegame 
      Caption         =   "Save"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Feature not added yet"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton mainmenu 
      Caption         =   "Return to Menu"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton quitbutton 
      Caption         =   "Quit Game"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton resume 
      Caption         =   "Resume"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label menulabel 
      Caption         =   "0"
      Height          =   15
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Battle Menu"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form7 - Battlemenu
'' Form7 shows when menu is clicked in Form8
Dim save As save

Public Sub load_Click()
Form10.Show
'' When the load button is clicked then Form10 is shown to the user
End Sub

Public Sub mainmenu_Click()
'' Main menu button takes the user back to form8 by hiding form8 as Form8 does not hide when the button is pressed in Form8
menulabel = "1"
'' This varibale is needed so that the game takes the user to Form2 when the user comes back to the game from the menu
Form7.Hide
End Sub

Public Sub quitbutton_Click()
Close #channum
End
'' Any file opened is closed and then the program closes
End Sub

Public Sub resume_Click()
Form7.Hide
'' The form is hidden and then the user is taken back to Form8 as Form8 is not hidden when Form7 is shown
End Sub

Public Sub savegame_Click()
filename = InputBox("Please enter a name for the file")
'' The user is asked for a name so that the record has a name to be identified by

For Counter = 0 To 13
'' This loop is to check that any unoccupied picture boxes in Form8 do not have a value of 0
If Form8.locationl(Counter) = 0 Then
Form8.locationl(Counter) = 3
'' If the program finds a value of 0 it will change it to 3
End If
Next

channum = FreeFile()
Open "C:\savedgame.dat" For Random As channum Len = Len(save)
records = LOF(channum) / Len(save)
'' The file is opened and total number of records in the file is taken
For Counter = records + 1 To records + 1
'' Counter is the position of the record after the last record saved
save.savename = filename
save.mapname = Form6.labellocation
save.utroops = Form8.unitlabel(1)
save.uheavy = Form8.unitlabel(2)
save.utanks = Form8.unitlabel(3)
save.uspecial1 = Form8.unitlabel(4)
save.uspecial2 = Form8.unitlabel(5)
save.usermoney = Form8.moneylabel
save.userarea = Form8.arealabel
'' Unitbale, moneylabel and arealabel are all information on the user that is taken from Form8
save.ctroops = Form8.cunitlabel(1)
save.cheavy = Form8.cunitlabel(2)
save.ctanks = Form8.cunitlabel(3)
save.cspecial1 = Form8.cunitlabel(4)
save.cspecial2 = Form8.cunitlabel(5)
save.compmoney = Form8.cunitlabel(0)
''cunitlabel is an array that holds all the information on the computer
save.comparea = Form8.comparealabel
save.loc1 = Form8.locationl(0)
save.loc2 = Form8.locationl(1)
save.loc3 = Form8.locationl(2)
save.loc4 = Form8.locationl(3)
save.loc5 = Form8.locationl(4)
save.loc6 = Form8.locationl(5)
save.loc7 = Form8.locationl(6)
save.loc8 = Form8.locationl(7)
save.loc9 = Form8.locationl(8)
save.loc10 = Form8.locationl(9)
save.loc11 = Form8.locationl(10)
save.loc12 = Form8.locationl(11)
save.loc13 = Form8.locationl(12)
save.loc14 = Form8.locationl(13)
'' Loc# is the position of each Picutrebox in form8
Put #channum, Counter, save
'' The information is saved into the record
Next
Close #channum
MsgBox ("File Saved Successfully")
'' The user is told that the save was successful and closes the file
End Sub
