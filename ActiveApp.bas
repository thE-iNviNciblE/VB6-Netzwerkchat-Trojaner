Attribute VB_Name = "activeapp"
Option Explicit

Global CurrentApp_hWnd As Long
Global CurrentApp_Title As String
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Function GetCaption(hwnd As Long)
'Get the caption of a specified HWND
Dim hWndlength As Long, hWndTitle As String, A As Long

'Get the length of the caption
hWndlength = GetWindowTextLength(hwnd)
'Fill a string to the length of the caption
hWndTitle = String$(hWndlength, 0)
'Fill the string with the real caption
A = GetWindowText(hwnd, hWndTitle, (hWndlength + 1))
GetCaption = hWndTitle
End Function

Function TrimTime(theTime As String) As String
Dim Hour As String, AM As String, RestString As String

'Trim the hours from the rest of the time
Hour = Left(theTime, InStr(theTime, ":") - 1)
'Trim everything but hours
RestString = Mid(theTime, InStr(theTime, ":"), Len(theTime))

'Set AM
AM = " AM"

'Computers use military time..if it's later than
'Noon in military time, turn it into regular time
'and change to PM
If Hour > 12 Then Hour = Hour - 12: AM = " PM"

'Put the time back together
TrimTime = Hour & RestString & AM
End Function
