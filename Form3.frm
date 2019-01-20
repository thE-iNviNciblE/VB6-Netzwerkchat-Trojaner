VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'Kein
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   1320
      Top             =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "ThE iNviNcible Presents Janschat"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   2520
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fest Einfach
      Height          =   3225
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5070
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_Click()
Form1.show
Form3.hide
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Form1.show
Form3.hide
Timer1.Enabled = False
End Sub
