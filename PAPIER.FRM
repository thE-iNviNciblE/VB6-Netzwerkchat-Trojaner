VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Alles über den Papierkorb"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "papier.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   2160
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check2 
      Caption         =   "Papierkorb leeren"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Vorschau"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   1320
   End
   Begin VB.TextBox ges 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Papierkorb sichtbar"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox papierkorb 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Senden"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Sendet den kram zum anderen Rechner"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Papierkorb neuer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim papieranaus

Private Sub Check1_Click()
If Check1.Value = 1 Then
    papieranaus = "an"
End If
If Check1.Value = 0 Then
    papieranaus = "aus"
End If
End Sub

Private Sub Command1_Click()
Form1.W.SendData ges.text
Form6.Hide
End Sub

Private Sub Timer1_Timer()
ges.text = papieranaus & "|" & papierkorb & "|" & "papierkorb"
End Sub
