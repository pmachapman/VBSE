VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Navigator"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' Require Variable Declaration
Option Explicit

' App code name
Public Property Get appCodeName() As String
    appCodeName = "VBSE"
End Property

' App name
Public Property Get appName() As String
    appName = App.Title
End Property

' App version
Public Property Get appVersion() As String
    appVersion = App.Major & "." & App.Minor
End Property

' Cookies enabled
Public Property Get cookieEnabled() As Boolean
    cookieEnabled = False
End Property

' Language
Public Property Get language() As String
    language = "en"
End Property

' Browser Online
Public Property Get onLine() As Boolean
    onLine = False
End Property

' Platform
Public Property Get platform() As String
    platform = "Win32"
End Property

' User Agent
Public Property Get userAgent() As String
    userAgent = App.Title & "/" & App.Major & "." & App.Minor
End Property


