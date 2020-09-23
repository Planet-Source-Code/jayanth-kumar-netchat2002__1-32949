Attribute VB_Name = "Module1"

Public server, user As String
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_MEMORY = &H4
Global SoundBuffer() As Byte

Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)


Private Declare Function mciSendString Lib "WINMM.DLL" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_REBOOT As Long = 2
Private Const EWX_LOGOFF As Long = 0
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long



Sub BeginPlaySound(ByVal ResourceId As Integer)
    SoundBuffer = LoadResData(ResourceId, "CUSTOM")
    sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub

Sub EndPlaySound()
    sndPlaySound ByVal vbNullString, 0&
End Sub
Function OpenCDROM()
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Function

Function CloseCDROM()
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door closed", strReturn, 127, 0)
End Function

Function ShutDown()
Dim lngresult
lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Function
Function Restart()
Dim lngresult
lngresult = ExitWindowsEx(EWX_REBOOT, 0&)
End Function
Function LogOff()
Dim lngresult
lngresult = ExitWindowsEx(EWX_LOGOFF, 0&)
End Function

Function closewin()
Dim lngresult
DestroyWindow (Form1.hWnd)
End Function


Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function


