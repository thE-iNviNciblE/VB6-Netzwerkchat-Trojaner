Attribute VB_Name = "Module2"
Option Explicit

Public Sub ICQMessage(ICQNummer As String, Name As String, eMailAdresse As String, _
Subject As String, Message As String)

Dim r_name$, r_Subject$, r_Message$

'Ist die Eingegebene ICQ Nummer überhaubt eine Zahl?
'If Not IsNumeric(ICQNummer) Then
'MsgBox "Es ist eine gültige oder falsche ICQ Nummer eingegeben", vbInformation + vbOKOnly, "ICQ Pager"
'Exit Sub
'End If

'Ist überhaubt ein Name eingegeben?
'If Trim(Name) = "" Then
'MsgBox "Es ist kein Name eingegeben", vbInformation + vbOKOnly, "ICQ Pager"
'Form1.Name2.SetFocus
'Exit Sub
'End If
'Ist überhaubt eine eMailAdresse eingegeben?
'If Trim(eMailAdresse) = "" Then
'MsgBox "Es ist kein eMail Adresse eingegeben", vbInformation + vbOKOnly, "ICQ Pager"
'Form1.mail.SetFocus
'Exit Sub
'End If

'Ist überhaubt ein Subject eingegeben?
'If Trim(Subject) = "" Then
'MsgBox "Es ist kein Subject eingegeben", vbInformation + vbOKOnly, "ICQ Pager"
'Form1.Betreff.SetFocus
'Exit Sub
'End If

'Ist überhaubt eine Nachricht eingegeben?
'If Trim(Message) = "" Then
'MsgBox "Es ist keine Nachricht eingegeben", vbInformation + vbOKOnly, "ICQ Pager"
'Form1.text.SetFocus
'Exit Sub
'End If

'Jetzt werden bei den Variabelen die leerzeichen durch ein + ersetzt
r_name = ChangeSpaces(Trim(Name))
r_Subject = ChangeSpaces(Trim(Subject))
r_Message = ChangeSpaces(Trim(Message))
'ende

'Absenden der Nachricht über das Microsoft Internet Transfer Controle
Form1.Inet1.Execute "http://wwp.icq.com/scripts/WWPMsg.dll?from=" & r_name & "&fromemail=" & eMailAdresse _
& "&subject=" & r_Subject & "&body=" & r_Message & "&to=" & ICQNummer & "&Send=" & """"
End Sub


'In dieser Function werden die leerzeichen " " durch ein "+" ersetzt
Private Function ChangeSpaces(cString As String) As String
   On Error Resume Next
   Dim cChar As String
   Dim cReturn As String
   Dim nLoop As Long
   cReturn = ""
   For nLoop = 1 To Len(cString)
       cChar = Mid(cString, nLoop, 1)
      
       If cChar = " " Then
          cChar = "+"
       End If
       cReturn = cReturn + cChar
   Next
   ChangeSpaces = cReturn
End Function



