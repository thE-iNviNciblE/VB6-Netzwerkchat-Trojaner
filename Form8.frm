VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form8 
   Caption         =   "ICQ WEBPAGER !!!! NACHRICHT SENDEN"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   3750
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   600
         Top             =   2760
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Send"
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox ICQNr 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "Hier Bitte die ICQ Nummer eingeben"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox mail 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         ToolTipText     =   "Hier Bitte die E-Mail Addresse eintragen"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Name2 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Betreff 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox text 
         Height          =   1095
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "ICQ Nr."
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Betreff"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Nachricht"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ICQMessage ICQNr.text, Name2.text, mail.text, Betreff.text, text.text
Form8.Hide
Form1.Show
End Sub
