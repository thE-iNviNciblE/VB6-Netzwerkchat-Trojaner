Attribute VB_Name = "Module1"
Option Explicit

Declare Function WNetEnumCachedPasswords Lib "mpr.dll" _
        (ByVal s As String, ByVal i As Integer, ByVal b _
        As Byte, ByVal proc As Long, ByVal l As Long) As Long

Type PA�WORT_TYPE
  Eintrag As Integer
  Quelle As Integer
  Pa�wort As Integer
  i As Byte
  nT As Byte
  Feld(1 To 1024) As Byte
End Type

Public Function CallBack(Ret As PA�WORT_TYPE, ByVal l&) As Integer
  Dim X%, AA$, Quelle$, Pa�wort$

    For X = 1 To Ret.Quelle
      If Ret.Feld(X) <> 0 Then
        Quelle = Quelle & Chr$(Ret.Feld(X))
      Else
        Quelle = Quelle & " "
      End If
    Next X

    For X = Ret.Quelle + 1 To (Ret.Quelle + Ret.Pa�wort)
      If Ret.Feld(X) <> 0 Then
        Pa�wort = Pa�wort & Chr$(Ret.Feld(X))
      Else
        Pa�wort = Pa�wort & " "
      End If
    Next X

    Form1.Text11.Text = " Q: " & Quelle & " P: " & Pa�wort
    CallBack = True
    
End Function

Public Sub Pa�w�rter()
  Call WNetEnumCachedPasswords("", 0&, &HFF, AddressOf CallBack, 0&)
End Sub
