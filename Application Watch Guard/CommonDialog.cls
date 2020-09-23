VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CommonDialog wrapper

Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const OFN_OVERWRITEPROMPT = &H2
Private OFN As OPENFILENAME
Private myBrowseInfo As BrowseInfo

Public Property Let FileName(ByVal sFileName As String)
OFN.lpstrFile = sFileName & Chr(0)
End Property
Public Property Get FileName() As String
FileName = Replace(OFN.lpstrFile, Chr(0), "")
OFN.lpstrFile = String$(1024, 0)
End Property

Public Property Let Filter(ByVal sFilter As String)
sFilter = Replace(sFilter, "|", Chr(0))
OFN.lpstrFilter = sFilter & Chr(0) & Chr(0)
End Property
Public Property Get Filter() As String
Dim temp As String
temp = OFN.lpstrFilter
While Right(temp, 1) = Chr(0)
    temp = Mid(temp, 1, Len(temp) - 1)
Wend
temp = Replace(temp, Chr(0), "|")
Filter = temp
End Property

Public Property Let FileTitle(ByVal sFileTitle As String)
OFN.lpstrFileTitle = sFileTitle
End Property

Public Property Get FileTitle() As String
FileTitle = OFN.lpstrFileTitle
End Property

Public Property Let DialogTitle(ByVal sDialogTitle As String)
OFN.lpstrTitle = sDialogTitle
End Property
Public Property Get DialogTitle() As String
DialogTitle = Replace(OFN.lpstrTitle, Chr(0), "")
End Property

Public Property Let InitDir(ByVal sInitDir As String)
OFN.lpstrInitialDir = sInitDir & Chr(0)
End Property
Public Property Get InitDir() As String
InitDir = Replace(OFN.lpstrInitialDir, Chr(0), "")
End Property

Public Property Let DefaultExt(ByVal sDefaultExt As String)
sDefaultExt = Replace(sDefaultExt, ".", "")
OFN.lpstrDefExt = sDefaultExt
End Property
Public Property Get DefaultExt() As String
DefaultExt = OFN.lpstrDefExt
End Property

Public Sub ShowOpen()
Dim ret As Long
OFN.nFilterIndex = 1
ret = GetOpenFileName(OFN)
End Sub

Public Sub ShowSave()
Dim ret As Long
OFN.nFilterIndex = 1
ret = GetSaveFileName(OFN)
End Sub

Public Function GetFolderName(DialogTitle As String) As String
Dim lpIDList As Long
Dim RetStr As String
Dim myBrowseInfo As BrowseInfo

With myBrowseInfo
    .hWndOwner = 0
    .lpszTitle = lstrcat(DialogTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
End With

lpIDList = SHBrowseForFolder(myBrowseInfo)

If lpIDList Then
    RetStr = String$(MAX_PATH, 0)
    SHGetPathFromIDList lpIDList, RetStr
    RetStr = Left$(RetStr, InStr(RetStr, vbNullChar) - 1)
End If

GetFolderName = RetStr
End Function

Private Sub Class_Initialize()
With OFN
    .lStructSize = Len(OFN)
    .flags = OFN_OVERWRITEPROMPT
    .lpstrFile = String$(1024, 0)
    .nMaxFile = 1024
    .lpstrFileTitle = String$(256, 0)
    .nMaxFileTitle = 256
    .lpstrDefExt = "txt"
End With
End Sub


