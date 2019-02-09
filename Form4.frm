VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Galactic Wars"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   660
   ClientWidth     =   11715
   LinkTopic       =   "Form4"
   ScaleHeight     =   8070
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   240
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   8175
      ScaleWidth      =   11295
      TabIndex        =   1
      Top             =   0
      Width           =   11295
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "There is no love.... Only War....."
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   7560
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Stuart James
'' Advanced Higher Computing
'' RTS Game
'' Form 4 - About The Game

Private Sub Command2_Click()
'' Command links with the 'back' button
Close #channum
form1.Show
Form4.Hide
'' This hides the form the user is currently on and then reshows the main menu
End Sub
