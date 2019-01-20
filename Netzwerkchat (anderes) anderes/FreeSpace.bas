Attribute VB_Name = "modFreeSpace"
Private Declare Function GetModuleHandle Lib "kernel32" _
Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias _
"LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
(ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
(ByVal hLibModule As Long) As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" _
Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As _
Long) As Long

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes _
As Currency, lpTotalNumberOfFreeBytes As Currency) As Long

Public Function GetDriveFreeSpace(Optional ByVal Drive As String = "") As Variant
   If Drive = "" Then Drive = CurDir$
   GetDriveFreeSpace = CDec(0)
   If Exported("kernel32", "GetDiskFreeSpaceExA") Then
      Dim cAvail As Currency
      Dim cTotal As Currency
      Dim cFree As Currency
      If GetDiskFreeSpaceEx(Drive, cAvail, cTotal, cFree) Then
         GetDriveFreeSpace = CDec(cAvail * 10000)
      End If
   Else
      Dim nSecPerClus As Long
      Dim nBytPerSec As Long
      Dim nFreeClus As Long
      Dim nTotalClus As Long
      If GetDiskFreeSpace(Drive, nSecPerClus, nBytPerSec, _
      nFreeClus, nTotalClus) Then
         GetDriveFreeSpace = CDec(nSecPerClus * nBytPerSec * _
         nFreeClus)
      End If
   End If
End Function

Public Function FormatFileSize(ByVal Size As Variant, Optional ByVal LongDisplay As Boolean = False) As String
   Dim sRet As String
   Const KB& = 1024
   Const MB& = KB * KB
   
   If Size < KB Then
      sRet = Format(Size, "#,##0") & " byte"
      If Size <> 1 Then sRet = sRet & "s"
   Else
      Select Case Size / KB
         Case Is < 10
            sRet = Format(Size / KB, "0.00") & " KB"
         Case Is < 100
            sRet = Format(Size / KB, "0.0") & " KB"
         Case Is < 1000
            sRet = Format(Size / KB, "0") & " KB"
         Case Is < 10000
            sRet = Format(Size / MB, "0.00") & " MB"
         Case Is < 100000
            sRet = Format(Size / MB, "0.0") & " MB"
         Case Is < 1000000
            sRet = Format(Size / MB, "0") & " MB"
         Case Is < 10000000
            sRet = Format(Size / MB / KB, "0.00") & " GB"
      End Select
   End If
   
   If LongDisplay Then
      If Size >= KB Then
         sRet = sRet & " (" & Format(Size, "#,##0") & " bytes)"
      End If
   End If
   FormatFileSize = sRet
End Function

Private Function Exported(ByVal ModuleName As String, ByVal ProcName As String) As Boolean
   Dim hModule As Long
   Dim lpProc As Long
   Dim FreeLib As Boolean
   
   hModule = GetModuleHandle(ModuleName)
   If hModule = 0 Then
      hModule = LoadLibrary(ModuleName)
      FreeLib = True
   End If
   
   If hModule Then
      lpProc = GetProcAddress(hModule, ProcName)
      Exported = (lpProc <> 0)
   End If
   
   If FreeLib Then Call FreeLibrary(hModule)
End Function

