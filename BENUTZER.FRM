VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Benutzdaten"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   Icon            =   "benutzer.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "&Zurück"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Daten setzen"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Daten holen"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "windows OEM"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Benutzername"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Organisation"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Betriebsystem"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form1.W.SendData "Getall"
Form1.W.GetData Temp
Text1.Text = Temp
End Sub

