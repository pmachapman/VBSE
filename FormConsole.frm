VERSION 5.00
Begin VB.Form FormConsole 
   Caption         =   "Console"
   ClientHeight    =   3000
   ClientLeft      =   1215
   ClientTop       =   5625
   ClientWidth     =   5910
   Icon            =   "FormConsole.frx":0000
   LinkTopic       =   "Console"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3000
   ScaleWidth      =   5910
   Begin VB.TextBox TextOutput 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu MenuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu MenuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu MenuEditSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEditClear 
         Caption         =   "C&lear Console"
      End
   End
End
Attribute VB_Name = "FormConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Require variable declaration
Option Explicit

' Cache the output so we can have more than 64KB
Dim Output As String

' Logs the specified message
Public Sub log(message As Variant)
    ' Updated the cached output
    If Len(Output) = 0 Then
        Output = message
    Else
        Output = Output & vbCrLf & message
    End If
    ' Write it to the text box
    Call SendMessage(TextOutput.hwnd, WM_SETTEXT, 0&, ByVal Output)
    Call SetWindowText(TextOutput.hwnd, Output)
    ' Move the cursor to the end
    On Error Resume Next
    TextOutput.SelStart = Len(TextOutput.Text)
    TextOutput.SelLength = 0
End Sub

' Form Activate Event Handler
Private Sub Form_Activate()
    ' Move the cursor to the end
    On Error Resume Next
    TextOutput.SelStart = Len(TextOutput.Text)
    TextOutput.SelLength = 0
End Sub

' Form Query Unload Event Handler
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' If we are being closed by the close button, just hide the form
    If UnloadMode = 0 Then
        Me.Hide
        Cancel = 1
        FormMain.MenuViewConsole.Checked = False
    End If
End Sub

' Form Resize Event Handler
Private Sub Form_Resize()
    TextOutput.Width = Me.ScaleWidth
    TextOutput.Height = Me.ScaleHeight
End Sub

' Edit -> Clear Console Menu Click Event Handler
Private Sub MenuEditClear_Click()
    Output = ""
    TextOutput.Text = ""
End Sub

' Edit -> Copy Menu Click Event Handler
Private Sub MenuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText TextOutput.SelText
End Sub

' Edit -> Select All Menu Click Event Handler
Private Sub MenuEditSelectAll_Click()
    On Error Resume Next
    TextOutput.SelStart = 0
    TextOutput.SelLength = Len(TextOutput.Text)
End Sub

' Output Textbox Change Event Handler
Private Sub TextOutput_Change()
    ' Enable/Disable relevant menu items
    If Len(Output) = 0 Then
        MenuEditCopy.Enabled = False
        MenuEditSelectAll.Enabled = False
    Else
        MenuEditCopy.Enabled = True
        MenuEditSelectAll.Enabled = True
    End If
End Sub

' Output Textbox Key Down Event Handler
Private Sub TextOutput_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Handle Ctrl+A shortcut, as the edit menu is hiddne
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        MenuEditSelectAll_Click
    ElseIf KeyCode = vbKeyMenu Then
        TextOutput_MouseDown vbRightButton, Shift, 0, 0
    End If
End Sub

' Output Textbox Mouse Down Event Handler
Private Sub TextOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the right mouse button
    If Button = vbRightButton Then
        ' Disable the textbox
        TextOutput.Enabled = False
        ' This DoEvents seems to be optional?
        DoEvents
        ' Re-enable the control, so that it doesn't appear as grayed
        TextOutput.Enabled = True
        TextOutput.SetFocus
        ' Show the custom menu
        PopupMenu MenuEdit
    End If
End Sub
