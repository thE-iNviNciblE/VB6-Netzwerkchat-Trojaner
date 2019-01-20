VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Netzwerkchat 2ter Teil"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10545
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   7155
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Width           =   10455
      Begin VB.CommandButton Command2 
         Caption         =   "Papierkorb"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Zeigt einem alles über den Papierkorb"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   3960
      TabIndex        =   5
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command3 
         Caption         =   "&Back"
         Height          =   495
         Left            =   720
         TabIndex        =   10
         ToolTipText     =   "Geht zurück auf die Hauptform"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3000
         Top             =   2760
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Hier wird der Text Empfangen"
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Empfangen !!!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "&Senden"
         Height          =   495
         Left            =   720
         TabIndex        =   4
         ToolTipText     =   "Sendest den Text aus der Textbox ab"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   2
         ToolTipText     =   "Hier wird der Text gesendet"
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Senden !!!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   10575
   End
   Begin VB.Menu befehle 
      Caption         =   "Befehle"
      Begin VB.Menu back 
         Caption         =   "Hauptform"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
Form1.Show
Form5.Hide
End Sub

Private Sub Command1_Click()
Form1.W.SendData Text2.Text
End Sub

Private Sub Command2_Click()
Form6.Show
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Also hier könnt ihr alles mit dem Mülleimer machen"
End Sub

Private Sub Command3_Click()
Form1.Show
Form5.Hide
End Sub

Private Sub Timer1_Timer()
Text3.Text = Form1.Text3
End Sub
