VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Window"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
' Require Variable Declaration
Option Explicit

' window.alert() implementation
Public Sub alert(message As Variant)
    MsgBox CStr(message), vbOKOnly, "Alert"
End Sub
