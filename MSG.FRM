VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "MSGBOX CREATOR !!!!"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "msg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   0
      Width           =   6135
      Begin VB.CheckBox Check1 
         Caption         =   "Zufall"
         Height          =   255
         Left            =   4200
         TabIndex        =   24
         Top             =   720
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.TextBox TITEL 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Prompt 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&VORSCHAU"
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Prompt"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "TITEL"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "RANDOM / ANZAHL "
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   3240
      TabIndex        =   15
      Top             =   2760
      Width           =   3135
      Begin VB.OptionButton Option10 
         Caption         =   "Random"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "MODAL ???"
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   2895
      Begin VB.OptionButton Option9 
         Caption         =   "Applicationmodal"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Systemmodal"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SYMBOL"
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
      Begin VB.OptionButton Option3 
         Caption         =   "Fragezeichen"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ausrufezeichen"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Kritisch"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCHALTFLÄCHEN"
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
      Begin VB.OptionButton Option7 
         Caption         =   "OK + Abbrechen+ Wiederholen"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton Option6 
         Caption         =   "OK + ABBRECHEN"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         Caption         =   "JA + NEIN"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "OK ONLY"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4680
      Top             =   3960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SENDEN !!!!!!!"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox ges 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Gesamt"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public la As Integer
Dim optionwert As Integer
Dim OptionSCHALt As Integer
Dim MODAL As Integer

Private Sub Check1_Click()
Check1.Value = 1
zufall = Int(12 * Rnd)
Select Case zufall
Case 2
Prompt.Text = "Windows Schwerer Ausnahmefehler"
TITEL.Text = "Fehlercode FF#34"
Option1.Value = True
Option4.Value = True
Case 3
Prompt.Text = "Zu wenig Arbeitsspeicher zum ausführen von §${[{[!§$§"
TITEL.Text = "%&&/§$)([]§"
Option1.Value = True
Option4.Value = True
Case 4
Prompt.Text = "Erforderlich DLL Datei wurde nicht" & vbCrLf & "gefunden"
TITEL.Text = "Fehler beim Starten des Programms"
Option1.Value = True
Option4.Value = True
Case 5
Prompt.Text = "Outlook meldet Virenbefall"
TITEL.Text = "Virren"
Option1.Value = True
Option4.Value = True
Case 6
Prompt.Text = "Windows Neustarten ?"
TITEL.Text = "Reboot"
Option3.Value = True
Option5.Value = True
Case 7
Prompt.Text = "Schwerer Systemfehler alle Daten " & vbCrLf & " gehen verloren"
TITEL.Text = "Kritisch !!!"
Option1.Value = True
Option4.Value = True
Case 8
Prompt.Text = "Die Eigentschaften des Objekts sind nicht verfügbar"
TITEL.Text = "Desktop"
Option1.Value = True
Option4.Value = True
Case 9
Prompt.Text = "Scandisk meldet Datenträger C:\ ist beschädigt"
TITEL.Text = "Scandisk"
Option1.Value = True
Option4.Value = True
Case 10
Prompt.Text = "Die Datei ist beschädigt"
TITEL.Text = "Explorer Fehler"
Option1.Value = True
Option4.Value = True
Case 11
Prompt.Text = "§$%§${{]][§$§!!???"
TITEL.Text = "FF#4453CC4"
Option1.Value = True
Option4.Value = True
Case 12
Prompt.Text = ""
TITEL.Text = ""
End Select
End Sub

Private Sub Command1_Click()
If TITEL.Text = "" Then
If Prompt.Text = "" Then
MsgBox "Unvollständige Angaben", vbCritical + vbSystemModal + vbOKOnly, "User Fehler 2"
Exit Sub
End If
End If
Form1.W.SendData ges.Text
Form3.Hide
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "23lj3"
End Sub

Private Sub Command2_Click()
If TITEL.Text = "" Then
If Prompt.Text = "" Then
MsgBox "Unvollständige Angaben", vbCritical + vbSystemModal + vbOKOnly, "User Fehler 2"
Exit Sub
End If
End If
''''''NNNNNEEEEEEUUUUUUU SSSSSPPPPPIILLLLLLLLTTTTTTEEEEEENNNNNN
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
'Je nachdem was man hier ändert also MsgBox paras (j - 5) , Paras (j - 2) + paras (j -1 ) + paras (j -3), Paras (j-4)
' so ändern sich auch die sachen die klappen....
'Das ganze problem liegt in dieser Zeile , ich habe das gefühl er Addiert die ganzen Paras
MsgBox Paras(j - 5), (Int(Paras(j - 1)) + Int(Paras(j - 2)) + Int(Paras(j - 3))), Paras(j - 4)

End Sub

Private Sub Form_Load()
Option1.Value = True
Option4.Value = True
Option8.Value = True

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    optionwert = vbCritical
End If
End Sub
Private Sub Option2_Click()
If Option2.Value = True Then
    optionwert = vbInformation
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    optionwert = vbQuestion
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
    OptionSCHALt = vbOKOnly
End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
    OptionSCHALt = vbYesNo
End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
    OptionSCHALt = vbOKCancel
End If
End Sub

Private Sub Option7_Click()
If Option7.Value = True Then
    OptionSCHALt = vbAbortRetryIgnore
End If
End Sub

Private Sub Option8_Click()
If Option8.Value = True Then
    MODAL = vbSystemModal
End If
End Sub

Private Sub Option9_Click()
If Option9.Value = True Then
    MODAL = vbApplicationModal
End If
End Sub

Private Sub Timer1_Timer()
ges.Text = Prompt.Text & "|" & TITEL.Text & "|" & optionwert & "|" & OptionSCHALt & "|" & MODAL & "|" & "MSG"
End Sub
