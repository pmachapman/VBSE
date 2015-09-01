VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Require Variable Declaration
Option Explicit

' window.alert() implementation
Public Sub alert(Optional message As Variant)
    If IsMissing(message) Then message = ""
    MsgBox CStr(message), vbOKOnly + vbExclamation, "Alert"
End Sub

' window.confirm() implementation
Public Function confirm(Optional message As Variant) As Boolean
    If IsMissing(message) Then message = ""
    confirm = MsgBox(CStr(message), vbOKCancel + vbQuestion, "Confirm") = vbOK
End Function

' window.prompt() implementation
Public Function prompt(Optional message As Variant, Optional defaultText As Variant) As String
    If IsMissing(message) Then message = ""
    If IsMissing(defaultText) Then defaultText = ""
    prompt = InputBox(CStr(message), "Prompt", CStr(defaultText))
End Function

