Attribute VB_Name = "ModuleMain"
' Require Variable Declaration
Option Explicit

' INI File API Calls
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' File Open API calls and constants
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SETTEXT = &HC

' About Dialog API calls
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

' Cursor Position constants
Public Const WM_USER = &H400
Public Const EM_EXLINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB

' Keycode constants
Public Const vbKeyMenu = &H5D

' CHM Help API calls
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' INI File Name
Dim IniFile As String

' Program Entrypoint
Public Sub Main()
    ' Set the INI File Name
    IniFile = App.Path
    If Right$(App.Path, 1) <> "\" Then IniFile = IniFile & "\"
    IniFile = IniFile & "VBSE.ini"
    ' Load the editor
    Load FormMain
    ' See if we have a file to open
    If Command <> "" Then
        Dim FilePath As String
        FilePath = Command
        ' Strip double quotes if present
        If Left$(FilePath, 1) = """" Then FilePath = Right$(FilePath, Len(FilePath) - 1)
        If Right$(FilePath, 1) = """" Then FilePath = Left$(FilePath, Len(FilePath) - 1)
        If Dir(FilePath) = "" Then
            MsgBox FilePath & vbCrLf & "File not found." & vbCrLf & "Please verify the correct file name was given.", vbExclamation, "Open"
        ElseIf Not FormMain.OpenFile(FilePath) Then
            MsgBox Command & " is invalid and cannot be opened", vbExclamation, "Open"
        End If
    End If
    ' Show the editor
    FormMain.Show
End Sub

' Gets a setting from the INI file
Public Function GetSettingFromIniFile(Section As String, KeyName As String, DefaultValue As String) As String
    ' Declare Variables
    Dim ReturnValue As String * 255
    Dim ReturnValueLength As Long
    ' Get the setting from the INI file
    ReturnValueLength = GetPrivateProfileString(Section, KeyName, DefaultValue, ReturnValue, Len(ReturnValue), IniFile)
    GetSettingFromIniFile = Left$(ReturnValue, ReturnValueLength)
End Function

' Saves a setting to the INI file
Public Sub SaveSettingToIniFile(Section As String, KeyName As String, Value As String)
    ' Save the setting to the INI file
    Call WritePrivateProfileString(Section, KeyName, Value, IniFile)
End Sub

' Deletes a setting from the INI file
Public Sub DeleteSettingFromIniFile(Section As String, KeyName As String)
    ' Delete the setting from the INI file
    Call WritePrivateProfileString(Section, KeyName, vbNullString, IniFile)
End Sub
