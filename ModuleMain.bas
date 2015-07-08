Attribute VB_Name = "ModuleMain"
' Require Variable Declaration
Option Explicit

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

' Implementation of VB6's Replace function
Public Function Replace(sIn As String, sFind As String, sReplace As String, Optional nStart, Optional nCount, Optional bCompare) As String
    Dim nC As Long, nPos As Integer, sOut As String
    If IsMissing(nStart) Then
        nStart = 1
    End If
    If IsMissing(nCount) Then
        nCount = -1
    End If
    If IsMissing(bCompare) Then
        bCompare = 0
    End If
    sOut = sIn
    nPos = InStr(CLng(nStart), sOut, sFind, bCompare)
    If nPos = 0 Then GoTo EndFn:
    Do
        nC = nC + 1
        sOut = Left(sOut, nPos - 1) & sReplace & _
           Mid(sOut, nPos + Len(sFind))
        If CLng(nCount) <> -1 And nC >= CLng(nCount) Then Exit Do
        nPos = InStr(CLng(nStart), sOut, sFind, bCompare)
    Loop While nPos > 0
EndFn:
    Replace = sOut
End Function


