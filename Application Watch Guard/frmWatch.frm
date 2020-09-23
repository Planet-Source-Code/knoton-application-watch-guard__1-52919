VERSION 5.00
Begin VB.Form frmWatch 
   Caption         =   " Application Watch Guard"
   ClientHeight    =   2115
   ClientLeft      =   3855
   ClientTop       =   4005
   ClientWidth     =   6435
   Icon            =   "frmWatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   6435
   Begin VB.ListBox lstWatch 
      Appearance      =   0  'Flat
      Height          =   1590
      ItemData        =   "frmWatch.frx":0442
      Left            =   0
      List            =   "frmWatch.frx":0444
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   300
      Width           =   6435
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6120
      Top             =   2520
   End
   Begin VB.Line LineMenu 
      X1              =   0
      X2              =   6420
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblListWatch 
      Caption         =   "Applications to watch"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   2775
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblTimeOut 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAddWatch 
         Caption         =   "Add Watch"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuReadLog 
         Caption         =   "Read Log"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuAutostart 
         Caption         =   "Enable Autostart"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuTimeout 
         Caption         =   "Timeout"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowTo 
         Caption         =   "HowTo"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Developers Website"
      End
   End
End
Attribute VB_Name = "frmWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Path() As String        'Path to the executable to start and watch
Private hInst() As Long         'Instance handle from Shell function.
Private TimeOut As Long         'Time in milliseconds for the response timeout
Private NrOfExe As Integer      'No of watches
Private Initiated As Boolean    'If it is running or not
Private lPath As String         'Path to the inifile
Private objSettings As Settings 'Object to read and write settings
Private CD As CommonDialog      'CommonDialog object
Private blnWatch As Boolean     'Flag for if watches exist

'Get the files shortpath
Private Function GetShortPath(strFileName As String) As String
Dim ret As Long
Dim tmp As String
Dim ShortFileName As String * 260
ret = GetShortPathName(strFileName, ShortFileName, Len(ShortFileName))
GetShortPath = Left$(ShortFileName, ret)
End Function

'Rearrange the array when removing an application from the watch
Private Sub RearrangeArray(StartPathArrNr As Integer)
Dim i, a As Integer
Call WriteWatchLog(CStr(Date) & " " & CStr(Time) & " Removed " & Path(NrOfExe))
lstWatch.RemoveItem (StartPathArrNr)
a = StartPathArrNr
Timer1.Enabled = False

If a <> UBound(Path) Then
    For i = a To NrOfExe - 1
        Path(i) = Path(i + 1)
        hInst(i) = hInst(i + 1)
    Next
End If

NrOfExe = NrOfExe - 1
If NrOfExe >= 0 Then
    ReDim Preserve Path(NrOfExe)
    ReDim Preserve hInst(NrOfExe)
Else
    ReDim Path(0)
    ReDim hInst(0)
    Initiated = False
    mnuStart.Caption = "Start"
    lblStatus = "Status = Waiting"
End If

objSettings.RemoveSection lPath, "Watch"

For i = 0 To NrOfExe
    Call WriteWatchExe(Path(i))
Next

If Initiated = True Then Timer1.Enabled = True
End Sub

'Start up all the applications to be watched
Private Sub Start()
Dim i As Integer
On Error GoTo ErrHandler
For i = 0 To NrOfExe
    'Save all started applications PID
    hInst(i) = Shell(GetShortPath(Path(i)), 1)
    Call WriteWatchLog(CStr(Date) & " " & CStr(Time) & " Initialized " & Path(NrOfExe))
Next
Initiated = True
Timer1.Enabled = True

Exit Sub
ErrHandler:
Initiated = True
Call WriteWatchLog(CStr(Date) & " " & CStr(Time) & " Not found " & Path(NrOfExe))
Call RearrangeArray(i)
End Sub

'Initiate settings
Private Sub Form_Load()
Dim temp As String
Dim Autostart As Boolean
Dim intPos As Integer
Set objSettings = New Settings

lPath = App.Path & "\Watch.ini"
Me.Show

Initiated = False
Call AddSystray(Me, App.Title)

Set CD = New CommonDialog
CD.DialogTitle = "Choose the executable of the application to watch"
CD.Filter = "Executable (*.exe)|*.exe"

TimeOut = CInt(objSettings.Read(lPath, "SETTINGS", "TimeOut", "5000"))
lblTimeOut.Caption = "Timeout = " & TimeOut & " MS"

ReadWatchExe

Autostart = CBool(objSettings.Read(lPath, "SETTINGS", "Autostart", "False"))
lblStatus = "Status = Waiting"

frmWatch.Height = CInt(objSettings.Read(lPath, "POSITION", "Height", "2880"))
frmWatch.Width = CInt(objSettings.Read(lPath, "POSITION", "Width", "6555"))
frmWatch.Top = CInt(objSettings.Read(lPath, "POSITION", "Top", "200"))
frmWatch.Left = CInt(objSettings.Read(lPath, "POSITION", "Left", "200"))

If Autostart Then
    mnuAutostart.Checked = True
    mnuStart_Click
    lblStatus = "Status = Running"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rtn As Long
rtn = X / Screen.TwipsPerPixelX
Select Case rtn
        Case WM_LBUTTONDOWN
            Me.WindowState = 0
            Me.Show
            SetForegroundWindow Me.hWnd
End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide

'collect the position/size of the form to be saved later
If Me.WindowState <> 1 Then
    lstWatch.Height = frmWatch.ScaleHeight - lblListWatch.Height - lblTimeOut.Height
    lstWatch.Width = frmWatch.ScaleWidth - 10
    lblTimeOut.Top = lstWatch.Height + lblListWatch.Height + 80
    lblStatus.Top = lblTimeOut.Top
    LineMenu.X2 = frmWatch.Width
End If
End Sub

'Shutdown the watcher and save the position/size to the inifile
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Me.WindowState = 2 Then Me.WindowState = 0
Call RemoveSystray

objSettings.Save lPath, "SETTINGS", "TimeOut", CStr(TimeOut)
objSettings.Save lPath, "SETTINGS", "Autostart", CStr(mnuAutostart.Checked)

objSettings.Save lPath, "POSITION", "Height", CStr(frmWatch.Height)
objSettings.Save lPath, "POSITION", "Width", CStr(frmWatch.Width)
objSettings.Save lPath, "POSITION", "Top", CStr(frmWatch.Top)
objSettings.Save lPath, "POSITION", "Left", CStr(frmWatch.Left)

End Sub

'Delete a watch
Private Sub lstWatch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    If lstWatch.ListIndex > -1 Then
        If MsgBox("Do you want to remove selected application from the watch ?", vbYesNo) = vbYes Then
            Call RearrangeArray(lstWatch.ListIndex)
        Else
            lstWatch.ListIndex = -1
        End If
    End If
End If
End Sub

'Add a watch
Private Sub mnuAddWatch_Click()
Dim temp As String
CD.ShowOpen
temp = CD.FileName
If temp <> "" Then
    NrOfExe = NrOfExe + 1
    ReDim Preserve Path(NrOfExe)
    ReDim Preserve hInst(NrOfExe)
    Path(NrOfExe) = temp
    lstWatch.AddItem temp
    Call WriteWatchExe(temp)
    If Initiated Then
        hInst(NrOfExe) = Shell(GetShortPath(Path(NrOfExe)), 1)
        Call WriteWatchLog(CStr(Date) & " " & CStr(Time) & " Initialized " & Path(NrOfExe))
    End If
    blnWatch = True
    mnuStart.Enabled = True
End If

End Sub

'Enable/disable autostart
Private Sub mnuAutostart_Click()
mnuAutostart.Checked = Not mnuAutostart.Checked
objSettings.Save lPath, "SETTINGS", "Autostart", CStr(mnuAutostart.Checked)
End Sub

'Shutdown
Private Sub mnuExit_Click()
Unload Me
End Sub

'Read Helpfile
Private Sub mnuHowTo_Click()
Call RunFile(App.Path & "\HowTo.txt", Me)
End Sub

'Read the log
Private Sub mnuReadLog_Click()
Dim temp As String
temp = App.Path & "/WatchLog.txt"
Call RunFile(temp, Me)
End Sub

'Start/pause the watcher
Private Sub mnuStart_Click()
If blnWatch Then
If mnuStart.Caption = "Start" Then
    mnuStart.Caption = "Pause"
    lblStatus = "Status = Running"
    If Initiated = False Then
        Call Start
        Me.WindowState = 1
        Me.Hide
    Else
        Timer1.Enabled = True
    End If
Else
    Timer1.Enabled = False
    mnuStart.Caption = "Start"
    lblStatus = "Status = Paused"
End If
End If
End Sub

'Set timeout
Private Sub mnuTimeout_Click()
Dim temp As String
temp = InputBox("Set Timeout in milliseconds for the watch.", "Set Timeout")
If Trim(temp) <> "" Then
    If IsNumeric(temp) Then
        If temp <= 60000 Then
        objSettings.Save lPath, "SETTINGS", "Timeout", "5000"
        TimeOut = Int(temp)
        lblTimeOut.Caption = "Timeout = " & temp & " MS"
        Else
            MsgBox "60000 ms is maximum"
        End If
    Else
        MsgBox "the number is not numeric !"
    End If
End If
End Sub

'Goto developers homepage
Private Sub mnuWeb_Click()
Call RunFile("http://www.iklartext.com", Me)
End Sub

'Timer that checks the responding of each watch in the list
Private Sub Timer1_Timer()
Dim i As Integer
On Error GoTo ErrHandler

For i = 0 To NrOfExe
    'Check if app is responding
    If Not IsAppResponding(hInst(i), TimeOut) Then
        'If not Paused
        If Timer1.Enabled = True Then
            'Try to do a clean Shut down in case the application has somekind of saving of settings during shut down
            'If it doesn´t sucide do a kill
            If Not CleanShutDown(hInst(i)) Then KillProcess (hInst(i))
            
            'Start up the application again
            hInst(i) = Shell(GetShortPath(Path(i)), 1)
            On Error GoTo PathNotFound
            'Write to logfile
            Call WriteWatchLog(CStr(Date) & " " & CStr(Time) & " Restarted " & Path(NrOfExe))
        Else
            Exit Sub
        End If
    End If
Next

Exit Sub
PathNotFound:
Call WriteWatchLog(CStr(Date) & " " & CStr(Time) & " Not found " & Path(NrOfExe))
Call RearrangeArray(i)
ErrHandler:
End Sub

'Write to the log
Private Sub WriteWatchLog(What As String)
Dim intFilNr As Integer
intFilNr = FreeFile
Open App.Path & "/WatchLog.txt" For Append As intFilNr
Print #intFilNr, What
Close #intFilNr
End Sub

'Save a application to be watched
Private Sub WriteWatchExe(What As String)
objSettings.SaveSection lPath, "Watch", What
End Sub

'Get the applications to watch over
Private Sub ReadWatchExe()
Dim var As Variant, i As Integer
On Error GoTo NoWatch
var = Split(objSettings.ReadSection(lPath, "WATCH"), Chr(0))

If IsArray(var) Then
    ReDim Path(UBound(var))
    ReDim hInst(UBound(var))
    NrOfExe = UBound(var)
    For i = 0 To UBound(var)
        Path(i) = var(i)
        lstWatch.AddItem Path(i)
    Next
End If
blnWatch = True
mnuStart.Enabled = True
Exit Sub
NoWatch:
mnuStart.Enabled = False
End Sub
