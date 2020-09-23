VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'declares for ini controlling
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'reads ini string
Public Function Read(ByVal ConfigObject As String, ByVal Section As String, ByVal Key As String, Optional ByVal strDefault As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, strDefault, RetVal, 255, ConfigObject)
Read = Left(RetVal, v)
End Function

'reads ini section
Public Function ReadSection(ByVal ConfigObject As String, ByVal Section As String) As String
Dim RetVal As String * 4096, v As Long
v = GetPrivateProfileSection(Section, RetVal, 4096, ConfigObject)
ReadSection = Left(RetVal, v - 1)
End Function

'writes ini
Public Sub Save(ByVal ConfigObject As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
WritePrivateProfileString Section, Key, Value, ConfigObject
End Sub

'writes ini section
Public Sub SaveSection(ConfigObject As String, Section As String, Value As String)
WritePrivateProfileSection Section, Value, ConfigObject
End Sub

'removes ini section
Public Sub RemoveSection(ConfigObject As String, Section As String)
WritePrivateProfileString Section, vbNullString, "", ConfigObject
End Sub

'Public Sub SetDefault()
'Dim lPath As String
'lPath = App.Path & "\Watch.ini"
'If Dir(lPath) = "" Then
'    Save lPath, "SETTINGS", "Timeout", "5000"
'    Save lPath, "SETTINGS", "Autostart", "False"
'    Save lPath, "POSITION", "Height", "2880"
'    Save lPath, "POSITION", "Width", "6555"
'    Save lPath, "POSITION", "Top", "200"
'    Save lPath, "POSITION", "Left", "200"
'End If
'End Sub
