VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "HILFE BITTE"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   Icon            =   "help.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Kreuz
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "help.frx":0442
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Zurück"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form1.Enabled = True
End Sub

Private Sub Form_Terminate()
Form1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub
