VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11850
   LinkTopic       =   "Form3"
   ScaleHeight     =   8430
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Top             =   6480
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Index           =   1
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   480
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   360
      Width           =   10095
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   13
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   14
         Top             =   4680
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   12
         Left            =   3240
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   13
         Top             =   4440
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   11
         Left            =   5280
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   12
         Top             =   5160
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   10
         Left            =   7440
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   11
         Top             =   4440
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   9
         Left            =   5760
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   10
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   8
         Left            =   9000
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   7
         Left            =   7920
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   6
         Left            =   7920
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   7
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   5
         Left            =   5760
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   6
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   4
         Left            =   3000
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   3
         Left            =   4080
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   2
         Left            =   1080
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Index           =   0
         Left            =   360
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As maps

Public Sub Command1_Click()
Close #channum
End
End Sub

Private Sub Command2_Click()
Close #channum
Form1.Show
Form3.Hide
End Sub

Public Sub Form_Load()
pat = Form1.labellocation
channum = FreeFile()
Open pat For Random As channum Len = Len(a)
total = LOF(channum) / Len(a)
Get #channum, total, a
map = Trim$(a.picture)
Picture1.picture = LoadPicture(map)

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

MsgBox ("File Loaded Sucessfully")
End Sub
