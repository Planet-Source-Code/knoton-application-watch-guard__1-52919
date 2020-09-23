Attribute VB_Name = "KillProcessCleanShutdown"
Option Explicit

Private Type LUID
  lowpart As Long
  highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   LuidUDT As LUID
   Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const GW_HWNDNEXT = 2
Const SMTO_BLOCK = &H1
Const WM_CLOSE = &H10
Const WM_NULL = &H0

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SendMessageTimeout Lib "User32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Public Function RunFile(Path As String, SenderForm As Form)
If Path <> "" Then
    ShellExecute SenderForm.hWnd, vbNullString, Path, vbNullString, "C:\", SW_SHOWNORMAL
End If
End Function


Public Function KillProcess(ByVal hProcessID As Long, Optional ByVal ExitCode As Long) As Boolean
Dim hToken As Long
Dim hProcess As Long
Dim tp As TOKEN_PRIVILEGES

'If the system is NT set priviliges to open/terminate processes
If GetVersion() >= 0 Then
    If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then GoTo CleanUp
    If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then GoTo CleanUp
    If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then GoTo CleanUp
End If

'Get the open handle to the running process
hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)

'If sucedded kill the open handle
If hProcess Then
    KillProcess = (TerminateProcess(hProcess, ExitCode) <> 0)
    CloseHandle hProcess
End If

'If NT adjust the priviliges back
If GetVersion() >= 0 Then
    tp.Attributes = 0
    AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
     
CleanUp:
    If hToken Then CloseHandle hToken
End If
End Function

'Do a clean shutdown incase the application do somekind of saving during shutdown
Public Function CleanShutDown(ByVal hProcessID As Long) As Boolean
Dim lngResult As Long
Dim lngReturnValue As Long
Dim hwndAPP As Long
On Error GoTo ErrHandler

'Get Hwnd from PID
hwndAPP = GetWinHandle(hProcessID)

'send WM_Close message to the Process, give it 5 seconds to respond
lngReturnValue = SendMessageTimeout(hwndAPP, WM_CLOSE, 0&, 0&, SMTO_BLOCK, 5000, lngResult)
If lngReturnValue <> 0 Then CleanShutDown = True
ErrHandler:
End Function

'Get Pid from Hwnd
Public Function ProcIDFromWnd(ByVal hWnd As Long) As Long
Dim idProc As Long
GetWindowThreadProcessId hWnd, idProc
ProcIDFromWnd = idProc
End Function
   
'get Hwnd from Pid
Public Function GetWinHandle(hInstance As Long) As Long
Dim tempHwnd As Long

tempHwnd = FindWindow(vbNullString, vbNullString)

Do Until tempHwnd = 0
   If GetParent(tempHwnd) = 0 Then
      If hInstance = ProcIDFromWnd(tempHwnd) Then
         GetWinHandle = tempHwnd
         Exit Do
      End If
   End If

   tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
Loop
End Function

'Check if the application is responding, wait the selected timeout
Public Function IsAppResponding(ByVal hInstance As Long, ByVal iTimeOut As Integer) As Boolean
Dim lngReturnValue As Long, lngResult As Long, hwndAPP As Long
hwndAPP = GetWinHandle(hInstance)
lngReturnValue = SendMessageTimeout(hwndAPP, WM_NULL, 0&, 0&, SMTO_BLOCK, iTimeOut, lngResult)
IsAppResponding = lngReturnValue
End Function
