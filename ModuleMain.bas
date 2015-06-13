Attribute VB_Name = "ModuleMain"
' Require Variable Declaration
Option Explicit

' Scroll Bar API calls and constants
Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3

' File Open API calls and constants
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SETTEXT = &HC

' About Dialog API calls
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

' Program Entrypoint
Public Sub Main()
    ' Load the editor
    Load FormMain
    ' See if we have a file to open
    If Command <> "" Then
        Dim FilePath As String
        FilePath = Command
        ' Strip double quotes if present
        If Left(FilePath, 1) = """" Then FilePath = Right(FilePath, Len(FilePath) - 1)
        If Right(FilePath, 1) = """" Then FilePath = Left(FilePath, Len(FilePath) - 1)
        If Dir(FilePath) = "" Then
            MsgBox FilePath & vbCrLf & "File not found." & vbCrLf & "Please verify the correct file name was given.", vbExclamation, "Open"
        ElseIf Not FormMain.OpenFile(FilePath) Then
            MsgBox Command & " is invalid and cannot be opened", vbExclamation, "Open"
        End If
    End If
    ' Show the editor
    FormMain.Show
End Sub
