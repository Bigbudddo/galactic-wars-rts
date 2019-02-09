VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   660
   ClientWidth     =   11880
   LinkTopic       =   "Form8"
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton gotobutton 
      Caption         =   "Move"
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   7440
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   5520
      ItemData        =   "Form8.frx":0000
      Left            =   10080
      List            =   "Form8.frx":0002
      TabIndex        =   29
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10080
      TabIndex        =   28
      Top             =   5640
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   13
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   24
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   12
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   11
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   10
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   4320
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   9
      Left            =   8280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   8
      Left            =   5040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   7
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   6
      Left            =   3360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   5
      Left            =   4800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   4
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   3
      Left            =   7320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   2
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   1
      Left            =   6480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Index           =   0
      Left            =   7320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton skill 
      Caption         =   "SKILL"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton special2 
      Caption         =   "SPECIAL 2"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton special1 
      Caption         =   "SPECIAL 1"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton tank 
      Caption         =   "TANK"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton heavytroops 
      Caption         =   "HEAVY"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton troops 
      Caption         =   "TROOP"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton goldmine 
      Caption         =   "MONEY"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton menu 
      Caption         =   "Menu"
      Height          =   495
      Left            =   10080
      TabIndex        =   3
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "END Turn"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton attack 
      Caption         =   "Attack"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   0
      Picture         =   "Form8.frx":0004
      ScaleHeight     =   6015
      ScaleWidth      =   10095
      TabIndex        =   30
      Top             =   0
      Width           =   10095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Build Commands"
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   0
      TabIndex        =   31
      Top             =   5760
      Width           =   9975
      Begin VB.Label unitlabel 
         BackColor       =   &H00000000&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   40
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label unitlabel 
         BackColor       =   &H00000000&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   39
         Top             =   480
         Width           =   375
      End
      Begin VB.Label unitlabel 
         BackColor       =   &H00000000&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   38
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label unitlabel 
         BackColor       =   &H00000000&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   37
         Top             =   480
         Width           =   375
      End
      Begin VB.Label unitlabel 
         BackColor       =   &H00000000&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   36
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label moneylabel 
         BackColor       =   &H00000000&
         Caption         =   "moneylabel"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Money :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7440
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Areas Controled :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label arealabel 
         BackColor       =   &H00000000&
         Caption         =   "arealabel"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   32
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   13
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   12
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   11
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   10
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   9
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   8
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   7
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   6
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   5
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   4
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   3
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   2
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label locationl 
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cunitlabel 
      Height          =   135
      Index           =   5
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cunitlabel 
      Height          =   135
      Index           =   4
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cunitlabel 
      Height          =   135
      Index           =   3
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cunitlabel 
      Height          =   135
      Index           =   2
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cunitlabel 
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cunitlabel 
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label comparealabel 
      Height          =   135
      Left            =   0
      TabIndex        =   43
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cheated 
      Caption         =   "0"
      Height          =   135
      Left            =   0
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   135
      Left            =   360
      TabIndex        =   41
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Money"
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label trooplabel 
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label fromgame 
      Caption         =   "0"
      Height          =   135
      Left            =   360
      TabIndex        =   26
      Top             =   7800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label turns 
      Height          =   135
      Left            =   360
      TabIndex        =   25
      Top             =   7920
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As maps
Dim playersprofile As users
Dim usercontrol As Integer
Dim computercontrol As Integer
Dim usersside As Integer
Dim usermoney As Integer
Dim userarea As Integer
Dim userturns As Integer
Dim usertroops(5) As Integer
Dim moneyvalidate As Integer
Dim extra As Integer
Dim moneybutton As Integer
Dim troopbutton As Integer
Dim heavybutton As Integer
Dim tankbutton As Integer
Dim specialbutton As Integer
Dim movesbutton As Integer
Dim skillcooldown As Integer
Dim selected As Integer
Dim map(13) As Integer
Dim userpicture As String
Dim comppicture As String
Dim difficulty As Integer
Dim newsel As Integer
Dim newposition As Integer
Dim totalcompunits As Integer
Dim totaluserunits As Integer
Dim specialunits As Integer
Dim username As String
Dim computertalk(6) As String
Dim cheatvalidate As Integer
Dim cheatenabled  As Integer
Dim cheat As Integer
Dim compunits(5) As Integer
Dim comparea As Integer
Public Sub attack_Click()

totalcompunits = 0
totaluserunits = 0
specialunits = 0

totalcompunits = (compunits(1) + compunits(2) + compunits(3)) - (Int(usertroops(4)) / 2)
totaluserunits = (usertroops(1) + usertroops(2) + usertroops(3)) - (Int(compunits(4)) / 2)
specialunits = compunits(5) - usertroops(5)

If totalcompunits >= totaluserunits And specialunits >= 0 Then
MsgBox "The Computer has defended there Country well, You flee for now"
For Counter = 1 To 5
usertroops(Counter) = Int(usertroops(Counter) / 2)
unitlabel(Counter) = usertroops(Counter)
Next
End If

If totaluserunits >= totalcompunits And specialunits <= 0 Then
MsgBox "Well Done, You have defeated the enemy and gained there country and riches"
For Counter = 1 To 5
compunits(Counter) = Int(compunits(Counter) / 2) + difficulty
cunitlabel(Counter) = compunits(Counter)
Next
usermoney = usermoney + 5000
moneylabel = usermoney
userarea = userarea + 1
arealabel = userarea
comparea = comparea - 1
comparealabel = comparea
map(selected) = 1
locationl(selected) = map(selected)
Picture2(selected).picture = LoadPicture(userpicture)
End If
movesbutton = 1
wincheck
End Sub

Public Sub Command3_Click()

userturns = userturns + 1
skillcooldown = skillcooldown + 1

resetbuttons

extra = 0

enemyturn
End Sub

Public Sub Form_Activate()
If Form7.menulabel = "1" Then
clear_game
Form8.Hide
form2.Show
End If
If Form7.menulabel = "2" Then
loadmap
defaults
End If
End Sub

Public Sub menu_Click()
Close #channum
Form7.Show
Form7.menulabel = "0"
End Sub

Public Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
List1.AddItem username
List1.AddItem Text1
cheat_check
response
Text1 = ""
End If
End Sub

Public Sub Form_Load()
loadmap
defaults
End Sub

Sub side_check()
channum = FreeFile()
Open "C:\userprofiles.dat" For Random As channum Len = Len(playersprofile)
totalrecords = LOF(channum) / Len(playersprofile)
For Counter = form1.selected To form1.selected
Get #channum, Counter, playersprofile
username = Trim$(playersprofile.username) + " Says:"
If playersprofile.side = 1 Then
usersside = 1
userpicture = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\goodlocation.jpg"
comppicture = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\badlocation.jpg"
Else
usersside = 2
userpicture = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\badlocation.jpg"
comppicture = "C:\Documents and Settings\jamesst\My Documents\Stuarts Project\Pictures\goodlocation.jpg"
End If
Next
End Sub

Sub assign_buttons()

If usersside = 1 Then
goldmine.Caption = "Ore Refinery"
troops.Caption = "Soldier"
heavytroops.Caption = "Heavy Support Soldier"
tank.Caption = "Tank"
special1.Caption = "Sniper"
special2.Caption = "Demon Slayer"
End If

If usersside = 2 Then
goldmine.Caption = "Slave Mine"
troops.Caption = "Corrupted Soldier"
heavytroops.Caption = "Enraged Rocket Troop"
tank.Caption = "Tank"
special1.Caption = "Virual Scientist"
special2.Caption = "Demon"
End If

Picture2(0).picture = LoadPicture(userpicture)
map(0) = 1
locationl(0) = map(0)
Picture2(13).picture = LoadPicture(comppicture)
map(13) = 2
locationl(13) = map(13)

If playersprofile.Commander = 1 Then
skill.Caption = "Royal Wealth"
skill.ToolTipText = "Increases Money from Factory by 1000"
End If

If playersprofile.Commander = 2 Then
skill.Caption = "Weapon Skill"
skill.ToolTipText = "All Units have Stronger weapons"
End If

If playersprofile.Commander = 3 Then
skill.Caption = "People Power"
skill.ToolTipText = "All Units have there health doubled"
End If

If playersprofile.Commander = 4 Then
skill.Caption = "Conscription"
skill.ToolTipText = "All units produced will be doubled"
End If

End Sub
Public Sub goldmine_Click()

usermoney = usermoney + 1000 + extra
moneylabel = usermoney
moneybutton = 1
buttonenable

End Sub

Public Sub gotobutton_Click()

If map(selected) = 1 Then
MsgBox "Welcome home Commander, Your people have made you money"
usermoney = usermoney + 200
moneylabel = usermoney
movesbutton = 1
buttonenable
End If

If map(selected) = 0 Then
map(selected) = 1
locationl(selected) = map(selected)
userarea = userarea + 1
arealabel = userarea
Picture2(selected).picture = LoadPicture(userpicture)
movesbutton = 1
buttonenable
End If

End Sub

Public Sub heavytroops_Click()

uservalidate = 0
If usermoney < 1000 Then
MsgBox "Insufficent Funds"
uservalidate = 1
End If

If uservalidate = 0 Then
usermoney = usermoney - 1000
moneylabel = usermoney
usertroops(2) = usertroops(2) + 5
unitlabel(2) = usertroops(2)
heavybutton = heavybutton + 1
buttonenable
End If

End Sub

Public Sub Picture2_Click(Index As Integer)
gotobutton.Enabled = False
attack.Enabled = False
selected = Index
move_attack_commands
End Sub

Public Sub skill_Click()

skill.Enabled = False
skillcooldown = 0

If playersprofile.Commander = 1 Then
extra = 5000
End If

If playersprofile.Commander = 2 Then
extra = 5
End If

If playersprofile.Commander = 3 Then
extra = 100
End If

If playersprofile.Commander = 4 Then
extra = 2
End If

End Sub

Public Sub special1_Click()

uservalidate = 0
If usermoney < 5000 Then
MsgBox "Insufficent Funds"
uservalidate = 1
End If

If uservalidate = 0 Then
usermoney = usermoney - 5000
moneylabel = usermoney
usertroops(4) = usertroops(4) + 1
unitlabel(4) = usertroops(4)
specialbutton = 1
buttonenable
End If

End Sub

Public Sub special2_Click()

uservalidate = 0
If usermoney < 10000 Then
MsgBox "Insufficent Funds"
uservalidate = 1
End If

If uservalidate = 0 Then
usermoney = usermoney - 10000
moneylabel = usermoney
usertroops(5) = usertroops(5) + 1
unitlabel(5) = usertroops(5)
specialbutton = 1
buttonenable
End If

End Sub

Public Sub tank_Click()

uservalidate = 0
If usermoney < 1500 Then
MsgBox "Insufficent Funds"
uservalidate = 1
End If

If uservalidate = 0 Then
usermoney = usermoney - 1500
moneylabel = usermoney
usertroops(3) = usertroops(3) + 1
unitlabel(3) = usertroops(3)
tankbutton = 1
buttonenable
End If

End Sub

Public Sub troops_Click()

uservalidate = 0
If usermoney < 500 Then
MsgBox "Insufficent Funds"
uservalidate = 1
End If

If uservalidate = 0 Then
usermoney = usermoney - 500
moneylabel = usermoney
newtroops = 10

If extra = 2 Then
newtroops = newtroops * 2
End If

usertroops(1) = usertroops(1) + newtroops
unitlabel(1) = usertroops(1)
troopbutton = troopbutton + 1
buttonenable
End If

End Sub

Sub buttonenable()

Dim modifier As Integer

If movesbutton = 1 Then
modifier = 1
End If

If moneybutton = 1 Then
goldmine.Enabled = False
End If

If troopbutton + modifier = 3 Then
troops.Enabled = False
End If

If heavybutton + modifier = 2 Then
heavytroops.Enabled = False
End If

If tankbutton + modifier = 1 Then
tank.Enabled = False
End If

If specialbutton + modifier = 1 Then
special1.Enabled = False
special2.Enabled = False
End If

If movesbutton = 1 Then
attack.Enabled = False
gotobutton.Enabled = False
End If

End Sub

Sub resetbuttons()
moneybutton = 0
troopbutton = 0
heavybutton = 0
tankbutton = 0
specialbutton = 0
movesbutton = 0

goldmine.Enabled = True
troops.Enabled = True
heavytroops.Enabled = True
tank.Enabled = True
special1.Enabled = True
special2.Enabled = True
gotobutton.Enabled = False
attack.Enabled = False

If skillcooldown = 3 Then
skill.Enabled = True
skillcooldown = 0
End If
End Sub

Sub move_attack_commands()

newsel = 0
newsel = selected - 1

If movesbutton = 0 Then

If map(selected) = 0 And map(newsel) = 1 Then
gotobutton.Enabled = True
attack.Enabled = False
End If

If map(selected) = 1 And map(newsel) = 1 Then
gotobutton.Enabled = True
attack.Enabled = False
End If

If map(selected) = 2 And map(newsel) = 1 Then
attack.Enabled = True
gotobutton.Enabled = False
End If

End If

End Sub

Sub enemyturn()

spvalidate = 0

If skillcooldown = 2 Then
If compcommander = 1 Then
cextra = 5000
End If
If compcommander = 2 Then
cextra = 5
End If
If compcommander = 3 Then
cextra = 100
End If
If compcommander = 4 Then
cextra = 2
End If
End If

compunits(0) = compunits(0) + 1000 + cextra
cunitlabel(0) = compunits(0)

For Counter = 1 To 10 + difficulty

build = CInt(Int((5 * Rnd()) + 1))

If build = 1 Then
cvalidate = 0
If compunits(0) < 500 Then
cvalidate = 1
End If
If cvalidate = 0 Then
newunits = 10
If cextra = 2 Then
newunits = newunits * 2
End If
compunits(1) = compunits(1) + newunits + difficulty
compunits(0) = compunits(0) - 500
cunitlabel(1) = compunits(1)
cunitlabel(0) = compunits(0)
End If
End If

If build = 2 Then
cvalidate = 0
If compunits(0) < 1000 Then
cvalidate = 1
End If
If cvalidate = 0 Then
compunits(2) = compunits(2) + 5 + difficulty
compunits(0) = compunits(0) - 1000
cunitlabel(2) = compunits(2)
cunitlabel(0) = compunits(0)
End If
End If

If build = 3 Then
cvalidate = 0
If compunits(0) < 1500 Then
cvalidate = 1
End If
If cvalidate = 0 Then
compunits(3) = compunits(3) + 1 + difficulty
compunits(0) = compunits(0) - 1500
cunitlabel(3) = compunits(3)
cunitlabel(0) = compunits(0)
End If
End If

If build = 4 Then
cvalidate = 0
If compunits(0) < 5000 Then
cvalidate = 1
End If
If cvalidate = 0 And spvalidate = 0 Then
compunits(4) = compunits(4) + 1
compunits(0) = compunits(0) - 5000
cunitlabel(4) = compunits(4)
cunitlabel(0) = compunits(0)
spvalidate = 1
End If
End If

If build = 5 Then
cvalidate = 0
If compunits(0) < 1000 Then
cvalidate = 1
End If
If cvalidate = 0 And spvalidate = 0 Then
compunits(5) = compunits(5) + 1
compunits(0) = compunits(0) - 10000
cunitlabel(5) = compunits(5)
cunitlabel(0) = compunits(0)
spvalidate = 1
End If
End If

Next

If comparea >= 1 Then
newposition = 13 - comparea

If map(newposition) = 0 Then
map(newposition) = 2
locationl(newposition) = map(newposition)
comparea = comparea + 1
comparealabel = comparea
Picture2(newposition).picture = LoadPicture(comppicture)
End If

If map(newposition) = 1 Then
compattack
End If

End If

End Sub

Sub defaults()
difficulty = Form6.difficulty
List1.Clear
usermoney = 10000
moneylabel = usermoney
userarea = 1
comparea = 1
compunits(0) = 10000 + (10000 * difficulty)
arealabel = userarea
side_check
assign_buttons
gotobutton.Enabled = False
attack.Enabled = False
compcommander = CInt(Int((4 * Rnd()) + 1))
For Counter = 1 To 5
unitlabel(Counter) = usertroops(Counter)
Next
End Sub

Sub wincheck()
If comparea = 0 Then
MsgBox ("VICTORY: You have crushed your Opponent")
turns = userturns
If cheat = 1 Then
cheated = "Yes"
Else
cheated = "No"
End If
fromgame = 2
clear_game
Close #channum
Form8.Hide
Form5.Show
End If
End Sub

Sub compattack()

totalcompunits = 0
totaluserunits = 0
specialunits = 0

totalcompunits = (compunits(1) + compunits(2) + compunits(3)) - usertroops(4)
totaluserunits = (usertroops(1) + usertroops(2) + usertroops(3)) - (Int(compunits(4) / 2))
specialunits = compunits(5) - usertroops(5)

If totalcompunits >= totaluserunits And specialunits >= 0 Then
MsgBox "The Computer has taken over your country, Some of your men have been spared"
For Counter = 1 To 5
usertroops(Counter) = Int(usertroops(Counter) / 2)
unitlabel(Counter) = usertroops(Counter)
Next
map(newposition) = 2
locationl(newposition) = map(newposition)
comparea = comparea + 1
userarea = userarea - 1
userlabel = userarea
comparealabel = comparea
compunits(0) = compunits(0) + 5000
Picture2(newposition).picture = LoadPicture(comppicture)
End If

If totalcompunits <= totaluserunits And specialunits <= 0 Then
MsgBox "The Computer has failed to take your country, You see there men flee"
For Counter = 1 To 5
compunits(Counter) = Int(compunits(Counter) / 2) + difficulty
cunitlabel(Counter) = compunits(Counter)
Next
End If
failcheck
End Sub

Sub failcheck()
If userarea = 0 Then
MsgBox ("DEFEATED: The Computer has managed to Crush you")
turns = userturns
fromgame = 3
clear_game
Close #channum
Form8.Hide
Form5.Show
End If
End Sub

Sub cheat_check()

cheatenabled = 0

If Text1 = "EPICWIN" Then
cheatenabled = 1
comparea = 0
cheat = 1
wincheck
End If

If Text1 = "EPICFAIL" Then
cheatenabled = 1
userarea = 0
cheat = 1
failcheck
End If

If Text1 = "GOLDILOCKS" Then
cheatenabled = 1
usermoney = usermoney + 100
moneylabel = usermoney
cheat = 1
End If

If Text1 = "HAILHUNTERS" Then
cheatenabled = 1
usertroops(5) = usertroops(5) + 5
unitlabel(5) = usertroops(5)
cheat = 1
End If

End Sub

Sub response()

computername = "Computer" + " Says:"
computertalk(1) = "I like Cake!"
computertalk(2) = "I Shall crush you!"
computertalk(3) = "Bow before my Army!"
computertalk(4) = "I shall never ally with you"
computertalk(5) = "Scared yet?"
computertalk(6) = "Bah, You have to cheat to beat me!?"

If cheatenabled = 1 Then
sentence = 6
End If

If cheatenabled = 0 Then
sentence = CInt(Int((5 * Rnd()) + 1))
End If

List1.AddItem computername
List1.AddItem computertalk(sentence)
End Sub

Sub clear_game()
For Counter = 0 To 13
map(Counter) = 0
locationl(Counter) = map(Counter)
Picture2(Counter) = LoadPicture("")
Next

For l = 1 To 5
usertroops(l) = 0
compunits(l) = 0
Next

skill.Enabled = True
goldmine.Enabled = True

userarea = 0
comparea = 0
defaults
End Sub

Sub loadmap()
pat = Form6.labellocation
If pat = "" Then
pat = "C:\Battlefields\earth.dat"
End If

channum = FreeFile()
Open pat For Random As channum Len = Len(a)
total = LOF(channum) / Len(a)
Get #channum, 1, a
battlemap = Trim$(a.picture)
Picture1.picture = LoadPicture(battlemap)

Picture2(0).Left = a.loc1left
Picture2(0).Top = a.loc1top

Picture2(1).Left = a.loc2left
Picture2(1).Top = a.loc2top

Picture2(2).Left = a.loc3left
Picture2(2).Top = a.loc3top

Picture2(3).Left = a.loc4left
Picture2(3).Top = a.loc4top

Picture2(4).Left = a.loc5left
Picture2(4).Top = a.loc5top

Picture2(5).Left = a.loc6left
Picture2(5).Top = a.loc6top

Picture2(6).Left = a.loc7left
Picture2(6).Top = a.loc7top

Picture2(7).Left = a.loc8left
Picture2(7).Top = a.loc8top

Picture2(8).Left = a.loc9left
Picture2(8).Top = a.loc9top

Picture2(9).Left = a.loc10left
Picture2(9).Top = a.loc10top

Picture2(10).Left = a.loc11left
Picture2(10).Top = a.loc11top

Picture2(11).Left = a.loc12left
Picture2(11).Top = a.loc12top

Picture2(12).Left = a.loc13left
Picture2(12).Top = a.loc13top

Picture2(13).Left = a.loc14left
Picture2(13).Top = a.loc14top

Close #channum

End Sub
