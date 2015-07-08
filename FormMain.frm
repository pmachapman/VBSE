VERSION 4.00
Begin VB.Form FormMain 
   Caption         =   "Very Basic Script Editor"
   ClientHeight    =   2910
   ClientLeft      =   2370
   ClientTop       =   4515
   ClientWidth     =   6000
   Height          =   3720
   Icon            =   "FormMain.frx":0000
   Left            =   2310
   LinkTopic       =   "Main"
   LockControls    =   -1  'True
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Top             =   3765
   Width           =   6120
   Begin VB.TextBox TextMain 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   1
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.PictureBox PictureStatus 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
      Begin VB.Label LabelLanguage 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   50
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label LabelStatus 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ln 1, Col 1"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.TextBox TextMain 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialogMain 
      Left            =   4560
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSScriptControlCtl.ScriptControl ScriptMain 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu MenuFileSeparator1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileMRU 
         Caption         =   "MRU List"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MenuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MenuEditSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu MenuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu MenuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu MenuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuEditSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEditGoto 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
      Begin VB.Menu MenuEditSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MenuRun 
      Caption         =   "&Run"
      Begin VB.Menu MenuRunStart 
         Caption         =   "&Start"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "&Language"
      Begin VB.Menu MenuLanguageText 
         Caption         =   "&Text"
      End
      Begin VB.Menu MenuLanguageJScript 
         Caption         =   "&JScript"
      End
      Begin VB.Menu MenuLanguageVBScript 
         Caption         =   "&VBScript"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu MenuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu MenuFormatWordWrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu MenuFormatFont 
         Caption         =   "&Font..."
      End
   End
   Begin VB.Menu MenuView 
      Caption         =   "&View"
      Begin VB.Menu MenuViewConsole 
         Caption         =   "&Console"
      End
      Begin VB.Menu MenuViewStatusBar 
         Caption         =   "&Status Bar"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MenuHelpScript 
         Caption         =   "Windows &Script Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuHelpUpdates 
         Caption         =   "&Check For Updates"
      End
      Begin VB.Menu MenuHelpSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' Require Variable Declaration
Option Explicit
' Scripting Objects
Dim Navigator As New Navigator
Dim Window As New Window
' Editor Variables
Dim CurrentTextBox As Integer
Dim FilePath As String
Dim RestoreStatusBar As Boolean
Dim TextChanged As Boolean
Dim UndoText As String
Dim UndoStart As Long
Dim UndoLength As Long
' MRU List Constants and Variables
Private Const MaxMRU = 4    ' Maximum number of MRUs in list (-1 for no limit)
Private Const NotFound = -1 ' Indicates a duplicate entry was not found
Private Const NoMRUs = -1   ' Indicates no MRUs are currently defined
Private MRUCount As Long    ' Maintains a count of MRUs defined

' Add Menu Item Routine
Private Sub AddMenuElement(NewItem As String)
    ' Declare variables
    Dim i As Long
    ' Handle error when using a removed MRU item
    On Error GoTo AlreadyLoaded
    ' Check that we will not exceed maximum MRUs
    If (MRUCount < (MaxMRU - 1)) Or (MaxMRU = -1) Then
        ' Increment the menu count
        MRUCount = MRUCount + 1
        ' Show the separator
        MenuFileSeparator1.Visible = True
        ' Check if this is the first item
        If MRUCount <> 0 Then
            ' Add a new element to the menu
            Load MenuFileMRU(MRUCount)
        End If
AlreadyLoaded:
        ' Make new element visible
        MenuFileMRU(MRUCount).Visible = True
    End If
    ' Shift items to maintain most recent to least recent
    For i = MRUCount To 1 Step -1
        ' Set the captions
        MenuFileMRU(i).Caption = MenuFileMRU(i - 1).Caption
    Next i
    ' Set caption for new item
    MenuFileMRU(0).Caption = NewItem
End Sub

' Add MRU Item Routine
Private Sub AddMRUItem(NewItem As String)
   Dim result As Long
   ' Call sub to check for duplicates
   result = CheckForDuplicateMRU(NewItem)
   ' Handle case if duplicate found
   If result <> NotFound Then
      ' Call sub to reorder MRU list
      ReorderMRUList NewItem, result
   Else
      ' Call sub to add new item to MRU menu
      AddMenuElement NewItem
   End If
End Sub

' Check For Duplicate MRU Item Function
Private Function CheckForDuplicateMRU(ByVal NewItem As String) As Long
    Dim i As Long
    ' Uppercase newitem for string comparisons
    NewItem = UCase(NewItem)
    ' Check all existing MRUs for duplicate
    For i = 0 To MRUCount
        If UCase(MenuFileMRU(i).Caption) = NewItem Then
            ' Duplicate found, return the location of the duplicate
            CheckForDuplicateMRU = i
            ' Stop searching
            Exit Function
        End If
    Next i
    ' No duplicate found, so return -1
    CheckForDuplicateMRU = -1
End Function

' Form Activate Event Handler
Private Sub Form_Activate()
    ' Process the word wrap value
    If MenuFormatWordWrap.Checked And CurrentTextBox = 0 Then
        CurrentTextBox = 1
        Call SendMessage(TextMain(CurrentTextBox).hwnd, WM_SETTEXT, 0&, ByVal TextMain(0).Text)
        Call SetWindowText(TextMain(CurrentTextBox).hwnd, TextMain(0).Text)
        TextMain(CurrentTextBox).SelStart = TextMain(0).SelStart
        TextMain(CurrentTextBox).SelLength = TextMain(0).SelLength
    ElseIf CurrentTextBox = 1 Then
        CurrentTextBox = 0
        Call SendMessage(TextMain(CurrentTextBox).hwnd, WM_SETTEXT, 0&, ByVal TextMain(1).Text)
        Call SetWindowText(TextMain(CurrentTextBox).hwnd, TextMain(1).Text)
        TextMain(CurrentTextBox).SelStart = TextMain(1).SelStart
        TextMain(CurrentTextBox).SelLength = TextMain(1).SelLength
    End If
    ' Set focus throws error 5 if the console has focus
    On Error Resume Next
    TextMain(CurrentTextBox).SetFocus
    TextMain(CurrentTextBox).ZOrder
    On Error GoTo 0
    ' Enable/Disable the Go To... menu if word wrapper
    MenuEditGoto.Enabled = Not MenuFormatWordWrap.Checked
    ' Hide the status bar menu and status bar if word wrapped
    If MenuFormatWordWrap.Checked Then
        PictureStatus.Visible = False
        RestoreStatusBar = MenuViewStatusBar.Checked
        MenuViewStatusBar.Checked = False
        MenuViewStatusBar.Enabled = False
    ElseIf RestoreStatusBar Then
        PictureStatus.Visible = True
        MenuViewStatusBar.Checked = True
        MenuViewStatusBar.Enabled = True
    Else
        MenuViewStatusBar.Enabled = True
    End If
    ' Resize the form elements
    Form_Resize
End Sub

' Form Load Event Handler
Private Sub Form_Load()
    ' Set the file path
    FilePath = "Untitled"
    ' Set the window caption
    Me.Caption = FilePath & " - " & App.Title & " " & App.Major & "." & App.Minor
    ' Load Settings
    Me.Left = CLng(GetSetting("Peter Chapman", "VBSE", "Left", Me.Left))
    Me.Top = CLng(GetSetting("Peter Chapman", "VBSE", "Top", Me.Top))
    Me.Width = CLng(GetSetting("Peter Chapman", "VBSE", "Width", Me.Width))
    Me.Height = CLng(GetSetting("Peter Chapman", "VBSE", "Height", Me.Height))
    Me.WindowState = CLng(GetSetting("Peter Chapman", "VBSE", "WindowState", Me.WindowState))
    MenuFormatWordWrap.Checked = CBool(GetSetting("Peter Chapman", "VBSE", "WordWrap", MenuFormatWordWrap.Checked))
    MenuViewConsole.Checked = CBool(GetSetting("Peter Chapman", "VBSE", "Console", MenuViewConsole.Checked))
    FormConsole.Visible = MenuViewConsole.Checked
    FormConsole.Left = CLng(GetSetting("Peter Chapman", "VBSE", "ConsoleLeft", FormConsole.Left))
    FormConsole.Top = CLng(GetSetting("Peter Chapman", "VBSE", "ConsoleTop", FormConsole.Top))
    FormConsole.Width = CLng(GetSetting("Peter Chapman", "VBSE", "ConsoleWidth", FormConsole.Width))
    FormConsole.Height = CLng(GetSetting("Peter Chapman", "VBSE", "ConsoleHeight", FormConsole.Height))
    FormConsole.WindowState = CLng(GetSetting("Peter Chapman", "VBSE", "ConsoleWindowState", FormConsole.WindowState))
    MenuViewStatusBar.Checked = CBool(GetSetting("Peter Chapman", "VBSE", "StatusBar", MenuViewStatusBar.Checked))
    Dim language As String
    language = CStr(GetSetting("Peter Chapman", "VBSE", "Language", "VBScript"))
    If language = "Text" Then
        MenuLanguageText_Click
    ElseIf language = "JScript" Then
        MenuLanguageJScript_Click
    Else
        MenuLanguageVBScript_Click
    End If
    TextMain(0).FontName = CStr(GetSetting("Peter Chapman", "VBSE", "FontName", TextMain(0).FontName))
    TextMain(0).FontSize = CInt(GetSetting("Peter Chapman", "VBSE", "FontSize", TextMain(0).FontSize))
    TextMain(0).FontBold = CBool(GetSetting("Peter Chapman", "VBSE", "FontBold", TextMain(0).FontBold))
    TextMain(0).FontItalic = CBool(GetSetting("Peter Chapman", "VBSE", "FontItalic", TextMain(0).FontItalic))
    TextMain(0).FontUnderline = CBool(GetSetting("Peter Chapman", "VBSE", "FontUnderline", TextMain(0).FontUnderline))
    TextMain(0).FontStrikethru = CBool(GetSetting("Peter Chapman", "VBSE", "FontStrikethru", TextMain(0).FontStrikethru))
    TextMain(1).FontName = TextMain(0).FontName
    TextMain(1).FontSize = TextMain(0).FontSize
    TextMain(1).FontBold = TextMain(0).FontBold
    TextMain(1).FontItalic = TextMain(0).FontItalic
    TextMain(1).FontUnderline = TextMain(0).FontUnderline
    TextMain(1).FontStrikethru = TextMain(0).FontStrikethru
    FormConsole.TextOutput.FontName = TextMain(0).FontName
    FormConsole.TextOutput.FontSize = TextMain(0).FontSize
    FormConsole.TextOutput.FontBold = TextMain(0).FontBold
    FormConsole.TextOutput.FontItalic = TextMain(0).FontItalic
    FormConsole.TextOutput.FontUnderline = TextMain(0).FontUnderline
    FormConsole.TextOutput.FontStrikethru = TextMain(0).FontStrikethru
    ' Initialize the count of MRUs
    MRUCount = NoMRUs
    ' Call sub to retrieve the MRU filenames
    GetMRUFileList
    ' Resize the text box to suit the window
    Form_Resize
End Sub

' Form OLE Drag Drop Event Handler
' Requires Visual Basic 5
'Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
'    If Data.GetFormat(vbCFFiles) = True Then
'        If Not OpenFile(Data.Files.Item(1)) Then
'            MsgBox Data.Files.Item(1) & " is invalid and cannot be opened.", vbExclamation, "Open"
'        End If
'    End If
'End Sub

' Form Query Unload Event Handler
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' See if there is something to save
    If TextChanged And Not (TextMain(CurrentTextBox).Text = "" And FilePath = "Untitled") Then
        Select Case MsgBox("The text in the file " & FilePath & " has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbYesNoCancel + vbExclamation, App.Title)
            Case vbYes
                MenuFileSave_Click
                ' If the save did not take place, cancel exiting
                If TextChanged = True Then
                    Cancel = True
                End If
            Case vbCancel
                Cancel = True
        End Select
    End If
End Sub

' Form Resize Event
Private Sub Form_Resize()
    On Error Resume Next
    ' Set up the text boxes correctly
    TextMain(0).Top = 0
    TextMain(0).Left = 0
    TextMain(1).Top = 0
    TextMain(1).Left = 0
    ' If we have the status bar showing
    If MenuViewStatusBar.Checked Then
        ' Position the status bar
        PictureStatus.Top = Me.ScaleHeight - PictureStatus.Height
        PictureStatus.Left = 0
        PictureStatus.Width = Me.ScaleWidth
        LabelStatus.Width = PictureStatus.ScaleWidth - 100
        LabelLanguage.Width = LabelStatus.Width
        ' The text box fills the entire window, less the status bar
        TextMain(0).Width = Me.ScaleWidth
        TextMain(0).Height = Me.ScaleHeight - PictureStatus.Height
    Else
        ' The text box fills the entire window
        TextMain(0).Width = Me.ScaleWidth
        TextMain(0).Height = Me.ScaleHeight
    End If
    ' Set the other text box's dimensions
    TextMain(1).Width = TextMain(0).Width
    TextMain(1).Height = TextMain(0).Height
    ' Show/Hide the language label
    If Me.ScaleWidth > 150 Then
        LabelLanguage.Visible = True
    Else
        LabelLanguage.Visible = False
    End If
End Sub

' Form Unload Event Handler
Private Sub Form_Unload(Cancel As Integer)
    ' Save Settings
    SaveSetting "Peter Chapman", "VBSE", "WindowState", Me.WindowState
    If Me.WindowState <> vbNormal Then Me.WindowState = vbNormal
    SaveSetting "Peter Chapman", "VBSE", "Left", Me.Left
    SaveSetting "Peter Chapman", "VBSE", "Top", Me.Top
    SaveSetting "Peter Chapman", "VBSE", "Width", Me.Width
    SaveSetting "Peter Chapman", "VBSE", "Height", Me.Height
    SaveSetting "Peter Chapman", "VBSE", "WordWrap", MenuFormatWordWrap.Checked
    SaveSetting "Peter Chapman", "VBSE", "Console", MenuViewConsole.Checked
    SaveSetting "Peter Chapman", "VBSE", "ConsoleWindowState", FormConsole.WindowState
    If FormConsole.WindowState <> vbNormal Then
        If Not FormConsole.Visible Then FormConsole.Show
        FormConsole.WindowState = vbNormal
    End If
    SaveSetting "Peter Chapman", "VBSE", "ConsoleLeft", FormConsole.Left
    SaveSetting "Peter Chapman", "VBSE", "ConsoleTop", FormConsole.Top
    SaveSetting "Peter Chapman", "VBSE", "ConsoleWidth", FormConsole.Width
    SaveSetting "Peter Chapman", "VBSE", "ConsoleHeight", FormConsole.Height
    SaveSetting "Peter Chapman", "VBSE", "StatusBar", MenuViewStatusBar.Checked
    If MenuLanguageText.Checked Then
        SaveSetting "Peter Chapman", "VBSE", "Language", "Text"
    ElseIf MenuLanguageJScript.Checked Then
        SaveSetting "Peter Chapman", "VBSE", "Language", "JScript"
    Else
        SaveSetting "Peter Chapman", "VBSE", "Language", "VBScript"
    End If
    SaveSetting "Peter Chapman", "VBSE", "FontName", TextMain(CurrentTextBox).FontName
    SaveSetting "Peter Chapman", "VBSE", "FontSize", TextMain(CurrentTextBox).FontSize
    SaveSetting "Peter Chapman", "VBSE", "FontBold", TextMain(CurrentTextBox).FontBold
    SaveSetting "Peter Chapman", "VBSE", "FontItalic", TextMain(CurrentTextBox).FontItalic
    SaveSetting "Peter Chapman", "VBSE", "FontUnderline", TextMain(CurrentTextBox).FontUnderline
    SaveSetting "Peter Chapman", "VBSE", "FontStrikethru", TextMain(CurrentTextBox).FontStrikethru
    ' Call sub to save the MRU filenames
    SaveMRUFileList
    ' Exit the program
    End
End Sub

' Gets the cursor co-ordinates, and updates the status bar
Private Sub GetCursorCoordinates()
    ' Make sure word wrap is off
    If Not MenuFormatWordWrap.Checked Then
        ' Declare variables
        Dim LineNumber As Long
        Dim Column As Long
        Dim Start As Long
        ' Get the co-ordinates
        Start = TextMain(CurrentTextBox).SelStart
        LineNumber = SendMessage(TextMain(CurrentTextBox).hwnd, EM_EXLINEFROMCHAR, -1, ByVal 0&)
        Column = SendMessage(TextMain(CurrentTextBox).hwnd, EM_LINEINDEX, ByVal LineNumber, ByVal CLng(0))
        ' Update the status bar
        LabelStatus.Caption = "Ln " + CStr(LineNumber + 1) & ", Col " & CStr(Start - Column + 1)
    End If
End Sub

' Gets The File MRU List
Private Sub GetMRUFileList()
    Dim i As Long        ' Loop control variable
    Dim result As String ' Name of MRU from registry
    Dim results(MaxMRU) As String
    ' Loop through all entries
    Do
        ' Retrieve entry from registry
        result = GetSetting("Peter Chapman", "VBSE", "MRUFile" & Trim(CStr(i)), "")
        
        ' Check if a value was returned
        If result <> "" Then
            results(i) = result
        End If
        
        ' Increment counter
        i = i + 1
    Loop Until (result = "")
    ' Add each MRU item
    For i = (MaxMRU - 1) To 0 Step -1
        If results(i) <> "" Then
            ' Call sub to additem to MRU list
            AddMRUItem results(i)
        End If
    Next i
End Sub

' Gets the file name from a path
Function GetFileNameFromPath(ByVal Path As String) As String
    If Right(Path, 1) <> "\" And Len(Path) > 0 Then
        GetFileNameFromPath = GetFileNameFromPath(Left(Path, Len(Path) - 1)) + Right(Path, 1)
    End If
End Function

' Initialises Scripting
Private Sub InitialiseScripting()
    ' Reset objects and code
    ScriptMain.Reset
    
    ' Add some VBA objects
    ' Upper case to keep separate from the web browser objects
    ScriptMain.AddObject "App", App, True
    ScriptMain.AddObject "Clipboard", Clipboard, True
    ScriptMain.AddObject "Me", Me, True
    ScriptMain.AddObject "Printer", Printer, True
    ScriptMain.AddObject "Printers", Printers, True
    ScriptMain.AddObject "Screen", Screen, True
    
    ' Add some web browser related objects
    ' Names are lowercase for JScript compatibility
    ScriptMain.AddObject "console", FormConsole, True
    ScriptMain.AddObject "navigator", Navigator, True
    ScriptMain.AddObject "window", Window, True
End Sub

' Edit Menu Click Event Handler
Private Sub MenuEdit_Click()
    ' Disable/Enable the menu items as required
    MenuEditUndo.Enabled = UndoText <> TextMain(CurrentTextBox).Text
    MenuEditCut.Enabled = TextMain(CurrentTextBox).SelLength > 0
    MenuEditCopy.Enabled = TextMain(CurrentTextBox).SelLength > 0
    MenuEditPaste.Enabled = Clipboard.GetFormat(vbCFText) And Len(Clipboard.GetText()) > 0
    MenuEditDelete.Enabled = TextMain(CurrentTextBox).SelLength > 0
End Sub

' Edit -> Copy Menu Click Event Handler
Private Sub MenuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText TextMain(CurrentTextBox).SelText
End Sub

' Edit -> Cut Menu Click Event Handler
Private Sub MenuEditCut_Click()
    ' Copy then delete
    MenuEditCopy_Click
    MenuEditDelete_Click
End Sub

' Edit -> Delete Menu Click Event Handler
Private Sub MenuEditDelete_Click()
    ' Store the undo value
    UndoStart = TextMain(CurrentTextBox).SelStart
    UndoLength = TextMain(CurrentTextBox).SelLength
    UndoText = TextMain(CurrentTextBox).Text
    ' Clear the selected text
    TextMain(CurrentTextBox).SelText = ""
End Sub

' Edit -> Go To Menu Click Event Handler
Private Sub MenuEditGoto_Click()
    ' Declare Variables
    Dim LineNumber As String
    Dim CharacterIndex As Long
    ' Get the current line number
    LineNumber = CStr(SendMessage(TextMain(CurrentTextBox).hwnd, EM_EXLINEFROMCHAR, -1, ByVal 0&) + 1)
    ' Get the new line number
    LineNumber = InputBox("Line Number:", "Go To Line", LineNumber)
    ' If we have a valid number
    If LineNumber <> "" And IsNumeric(LineNumber) And CLng(Val(LineNumber)) > 0 Then
        CharacterIndex = CLng(LineNumber) - 1
        CharacterIndex = SendMessage(TextMain(CurrentTextBox).hwnd, EM_LINEINDEX, ByVal CharacterIndex, ByVal CLng(0))
        TextMain(CurrentTextBox).SetFocus
        If CharacterIndex <> -1 Then
            TextMain(CurrentTextBox).SelStart = CharacterIndex
        End If
    End If
End Sub

' Edit -> Paste Menu Click Event Handler
Private Sub MenuEditPaste_Click()
    ' Store the undo value
    UndoStart = TextMain(CurrentTextBox).SelStart
    UndoLength = TextMain(CurrentTextBox).SelLength
    UndoText = TextMain(CurrentTextBox).Text
    ' Replace the selected text with the clipboard
    TextMain(CurrentTextBox).SelText = Clipboard.GetText()
End Sub

' Edit -> Select All Menu Click Event Handler
Private Sub MenuEditSelectAll_Click()
    ' Select all text
    TextMain(CurrentTextBox).SelStart = 0
    TextMain(CurrentTextBox).SelLength = Len(TextMain(CurrentTextBox).Text)
End Sub

' Edit -> Undo Menu Click Event Handler
Private Sub MenuEditUndo_Click()
    ' Store the redo text and position
    Dim RedoText As String
    Dim RedoStart As Long
    Dim RedoLength As Long
    RedoText = TextMain(CurrentTextBox).Text
    RedoStart = TextMain(CurrentTextBox).SelStart
    RedoLength = TextMain(CurrentTextBox).SelLength
    ' Undo the text, and return the cursor to the appropriate position
    TextMain(CurrentTextBox).Text = UndoText
    TextMain(CurrentTextBox).SelStart = UndoStart
    TextMain(CurrentTextBox).SelLength = UndoLength
    ' Store the redo text and cursor as the undo text
    UndoText = RedoText
    UndoStart = RedoStart
    UndoLength = RedoLength
End Sub

' File -> Exit Menu Click Event Handler
Private Sub MenuFileExit_Click()
    Unload Me
End Sub

' File -> MRU Menu Click Event Handler
Private Sub MenuFileMRU_Click(Index As Integer)
    ' Open the file
    If Dir(MenuFileMRU(Index).Caption) = "" Then
        If MsgBox(MenuFileMRU(Index).Caption & " could not be not found." & vbCrLf & vbCrLf & "Would you like to remove this item from the menu?", vbQuestion + vbYesNo, "Open") = vbYes Then
            RemoveMenuElement MenuFileMRU(Index).Caption
        End If
    ElseIf Not OpenFile(MenuFileMRU(Index).Caption) Then
        If MsgBox(MenuFileMRU(Index).Caption & " is invalid and cannot be opened." & vbCrLf & "vbCrLf & Would you like to remove this item from the menu?", vbQuestion + vbYesNo, "Open") = vbYes Then
            RemoveMenuElement MenuFileMRU(Index).Caption
        End If
    End If
End Sub

' File -> New Menu Click Event Handler
Private Sub MenuFileNew_Click()
    ' Show the save prompt if the text has changed
    If TextChanged And Not (TextMain(CurrentTextBox).Text = "" And FilePath = "Untitled") Then
        Select Case MsgBox("The text in the file " & FilePath & " has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbYesNoCancel + vbExclamation, App.Title)
            Case vbYes
                MenuFileSave_Click
            Case vbCancel
                Exit Sub
        End Select
    End If
    ' Clear the text box, undo, and changed values
    TextMain(CurrentTextBox).Text = ""
    UndoText = ""
    UndoStart = 0
    UndoLength = 0
    TextChanged = False
    ' Reset the file path
    FilePath = "Untitled"
    ' Set the window caption
    Me.Caption = FilePath & " - " & App.Title & " " & App.Major & "." & App.Minor
End Sub

' File -> Open Menu Click Event Handler
Private Sub MenuFileOpen_Click()
    ' Show the save prompt if the text has changed
    If TextChanged And Not (TextMain(CurrentTextBox).Text = "" And FilePath = "Untitled") Then
        Select Case MsgBox("The text in the file " & FilePath & " has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbYesNoCancel + vbExclamation, "Open")
            Case vbYes
                MenuFileSave_Click
            Case vbCancel
                Exit Sub
        End Select
    End If
    ' Set up the common dialog
    CommonDialogMain.CancelError = True
    CommonDialogMain.Filter = "All Files (*.*)|*.*|JScript Files (*.js)|*.js|Text Files (*.txt)|*.txt|VBScript Files (*.vbs)|*.vbs|"
    If MenuLanguageJScript.Checked Then
        CommonDialogMain.FilterIndex = 2
    ElseIf MenuLanguageText.Checked Then
        CommonDialogMain.FilterIndex = 3
    Else
        CommonDialogMain.FilterIndex = 4
    End If
    ' Show the dialog
    On Error GoTo CancelOpen
    CommonDialogMain.ShowOpen
    If CommonDialogMain.FileName <> "" Then
        If Dir(CommonDialogMain.FileName) = "" Then
            MsgBox CommonDialogMain.FileName & vbCrLf & "File not found." & vbCrLf & "Please verify the correct file name was given.", vbExclamation, "Open"
        ElseIf Not OpenFile(CommonDialogMain.FileName) Then
            MsgBox CommonDialogMain.FileName & " is invalid and cannot be opened.", vbExclamation, "Open"
        End If
    End If
CancelOpen:
End Sub

' File -> Save Menu Click Event Handler
Private Sub MenuFileSave_Click()
    SaveFile False
End Sub

' File -> Save As Menu Click Event Handler
Private Sub MenuFileSaveAs_Click()
    SaveFile True
End Sub

' Format -> Font Menu Click Event Handler
Private Sub MenuFormatFont_Click()
    ' Set up the font dialog
    CommonDialogMain.FontName = TextMain(CurrentTextBox).FontName
    CommonDialogMain.FontSize = TextMain(CurrentTextBox).FontSize
    CommonDialogMain.FontBold = TextMain(CurrentTextBox).FontBold
    CommonDialogMain.FontItalic = TextMain(CurrentTextBox).FontItalic
    CommonDialogMain.FontUnderline = TextMain(CurrentTextBox).FontUnderline
    CommonDialogMain.FontStrikethru = TextMain(CurrentTextBox).FontStrikethru
    CommonDialogMain.CancelError = False
    CommonDialogMain.Flags = cdlCFANSIOnly + cdlCFBoth
    ' Show the font dialog
    On Error GoTo CancelFont
    CommonDialogMain.ShowFont
    ' Update the font
    TextMain(0).FontName = CommonDialogMain.FontName
    TextMain(0).FontSize = CommonDialogMain.FontSize
    TextMain(0).FontBold = CommonDialogMain.FontBold
    TextMain(0).FontItalic = CommonDialogMain.FontItalic
    TextMain(0).FontUnderline = CommonDialogMain.FontUnderline
    TextMain(0).FontStrikethru = CommonDialogMain.FontStrikethru
    TextMain(1).FontName = TextMain(0).FontName
    TextMain(1).FontSize = TextMain(0).FontSize
    TextMain(1).FontBold = TextMain(0).FontBold
    TextMain(1).FontItalic = TextMain(0).FontItalic
    TextMain(1).FontUnderline = TextMain(0).FontUnderline
    TextMain(1).FontStrikethru = TextMain(0).FontStrikethru
    FormConsole.TextOutput.FontName = TextMain(0).FontName
    FormConsole.TextOutput.FontSize = TextMain(0).FontSize
    FormConsole.TextOutput.FontBold = TextMain(0).FontBold
    FormConsole.TextOutput.FontItalic = TextMain(0).FontItalic
    FormConsole.TextOutput.FontUnderline = TextMain(0).FontUnderline
    FormConsole.TextOutput.FontStrikethru = TextMain(0).FontStrikethru
CancelFont:
End Sub

' Format -> Word Wrap Menu Click Event Handler
Private Sub MenuFormatWordWrap_Click()
    ' Update the menu
    MenuFormatWordWrap.Checked = Not MenuFormatWordWrap.Checked
    ' Update the window
    Form_Activate
End Sub

' Help -> About Menu Click Event Handler
Private Sub MenuHelpAbout_Click()
    Call ShellAbout(Me.hwnd, "Windows", App.Title & " " & App.Major & "." & App.Minor & vbCrLf & App.LegalCopyright, Me.Icon)
End Sub

' Help -> Windows Script Help Menu Click Event Handler
Private Sub MenuHelpScript_Click()
    ' Get the path to the help file
    Dim HelpPath As String
    HelpPath = App.Path
    If Right(HelpPath, 1) <> "\" Then
        HelpPath = HelpPath & "\"
    End If
    HelpPath = HelpPath & "script56.chm"
    ' Check the help file exists
    If Dir(HelpPath) = "" Then
        ' Ask if the user wants to download the help file
        If MsgBox("The file ""script56.chm"" was not found in the program directory" & vbCrLf & vbCrLf & "Do you want to download this file?", vbYesNo + vbQuestion) = vbYes Then
            Call ShellExecute(Me.hwnd, "Open", "https://www.microsoft.com/en-nz/download/details.aspx?id=2764", "", App.Path, 1)
        End If
    Else
        Call ShellExecute(Me.hwnd, "Open", HelpPath, "", App.Path, 1)
    End If
End Sub

' Help -> Check For Updates Menu Click Event Handler
Private Sub MenuHelpUpdates_Click()
    Call ShellExecute(Me.hwnd, "Open", "https://github.com/pmachapman/VBSE/releases", "", App.Path, 1)
End Sub

' Language -> JScript Menu Click Event Handler
Private Sub MenuLanguageJScript_Click()
    MenuLanguageJScript.Checked = True
    MenuLanguageText.Checked = False
    MenuLanguageVBScript.Checked = False
    MenuRun.Enabled = True
    ScriptMain.language = "JScript"
    InitialiseScripting
    LabelLanguage.Caption = "JScript"
End Sub

' Language -> Text Menu Click Event Handler
Public Sub MenuLanguageText_Click()
    MenuLanguageJScript.Checked = False
    MenuLanguageText.Checked = True
    MenuLanguageVBScript.Checked = False
    MenuRun.Enabled = False
    LabelLanguage.Caption = "Text"
End Sub

' Language -> VBScript Menu Click Event Handler
Private Sub MenuLanguageVBScript_Click()
    MenuLanguageJScript.Checked = False
    MenuLanguageText.Checked = False
    MenuLanguageVBScript.Checked = True
    MenuRun.Enabled = True
    ScriptMain.language = "VBScript"
    InitialiseScripting
    LabelLanguage.Caption = "VBScript"
End Sub

' Start -> Run Menu Click Event Handler
Private Sub MenuRunStart_Click()
    On Error Resume Next
    InitialiseScripting
    ScriptMain.AddCode TextMain(CurrentTextBox).Text
End Sub
' View -> Console Menu Click Event Handler
Private Sub MenuViewConsole_Click()
    MenuViewConsole.Checked = Not MenuViewConsole.Checked
    If MenuViewConsole.Checked Then
        FormConsole.Show
    Else
        FormConsole.Hide
    End If
End Sub

' View -> Status Bar Menu Click Event Handler
Private Sub MenuViewStatusBar_Click()
    MenuViewStatusBar.Checked = Not MenuViewStatusBar.Checked
    RestoreStatusBar = MenuViewStatusBar.Checked
    PictureStatus.Visible = MenuViewStatusBar.Checked
    Form_Resize
End Sub

' Open File Function
Public Function OpenFile(FileName As String) As Boolean
    On Error GoTo OpenFileError
    Dim F As Integer
    Dim S As String
    If FileName <> "" Then
        ' Get Text into Memory
        F = FreeFile
        Open FileName For Input As F
        S = Input$(LOF(F), F)
        Close F
        ' Put it into Text Box
        ' Only works properly under NT\2000\XP
        Call SendMessage(TextMain(CurrentTextBox).hwnd, WM_SETTEXT, 0&, ByVal S)
        Call SetWindowText(TextMain(CurrentTextBox).hwnd, S)
        OpenFile = True
        ' Update the file path
        FilePath = FileName
        ' Handle the file language
        UpdateFileLanguage
        ' Update the window caption
        Me.Caption = GetFileNameFromPath(FilePath) & " - " & App.Title & " " & App.Major & "." & App.Minor
        ' Reset the undo and changed values
        UndoText = S
        UndoStart = 0
        UndoLength = 0
        TextChanged = False
        ' Call sub to add this file as an MRU
        On Error GoTo 0
        AddMRUItem FilePath
    Else
        OpenFile = False
    End If
    Exit Function
OpenFileError:
    OpenFile = False
End Function

' Remove Menu Item Routine
Private Sub RemoveMenuElement(RemoveItem As String)
    Dim i As Long
    Dim result As Long
    ' Only do this if we have more than one item
    If MRUCount > 0 Then
        ' Call sub to check for duplicates
        result = CheckForDuplicateMRU(RemoveItem)
        ' Call sub to reorder MRU list
        ReorderMRUList RemoveItem, result
        ' Shift items up to the top of the list
        For i = 1 To MRUCount
            ' Set the captions
            MenuFileMRU(i - 1).Caption = MenuFileMRU(i).Caption
        Next i
    Else
        ' Hide the separator
        MenuFileSeparator1.Visible = False
    End If
    ' Remove the last item
    MenuFileMRU(MRUCount).Visible = False
    MRUCount = MRUCount - 1
End Sub

' Reorder MRU List Routine
Private Sub ReorderMRUList(DuplicateMRU As String, DuplicateLocation As Long)
    Dim i As Long
    ' Move entries previously "more recent" than the
    ' duplicate down one in the MRU list
    For i = DuplicateLocation To 1 Step -1
        MenuFileMRU(i).Caption = MenuFileMRU(i - 1).Caption
    Next i
    ' Set the caption of new item
    MenuFileMRU(0).Caption = DuplicateMRU
End Sub

' Save File Routine
Private Sub SaveFile(SaveAs As Boolean)
    ' If a file is not open, or we are saving as, show the save dialog
    If FilePath = "Untitled" Or SaveAs Then
        ' Set up the common dialog
        CommonDialogMain.CancelError = True
        CommonDialogMain.Filter = "All Files (*.*)|*.*|JScript Files (*.js)|*.js|Text Files (*.txt)|*.txt|VBScript Files (*.vbs)|*.vbs|"
        If MenuLanguageJScript.Checked Then
            CommonDialogMain.FilterIndex = 2
        ElseIf MenuLanguageText.Checked Then
            CommonDialogMain.FilterIndex = 3
        Else
            CommonDialogMain.FilterIndex = 4
        End If
        ' Show the dialog
        On Error GoTo CancelSave
        CommonDialogMain.ShowSave
        ' Take action based on the dialog's result
        If CommonDialogMain.FileName = "" Then
            Exit Sub
        ElseIf Dir(CommonDialogMain.FileName) = "" Then
            FilePath = CommonDialogMain.FileName
        ElseIf MsgBox(CommonDialogMain.FileName & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Save As") = vbYes Then
            FilePath = CommonDialogMain.FileName
        Else
            Exit Sub
        End If
    End If
    ' Save the file
    Dim F As Integer
    F = FreeFile
    Open FilePath For Output As F
    Print #F, TextMain(CurrentTextBox).Text;
    Close F
    ' Update the window caption
    Me.Caption = GetFileNameFromPath(FilePath) & " - " & App.Title & " " & App.Major & "." & App.Minor
    ' Reset the changed flag
    TextChanged = False
    ' Call sub to add this file as an MRU
    AddMRUItem FilePath
    ' Update the language menu
    UpdateFileLanguage
CancelSave:
End Sub

' Save MRU List Routine
Private Sub SaveMRUFileList()
    Dim i As Long ' Loop control variable
    ' Loop through all MRU
    For i = 0 To MRUCount
        ' Write MRU to registry with key as its position in list
        SaveSetting "Peter Chapman", "VBSE", "MRUFile" & Trim(CStr(i)), MenuFileMRU(i).Caption
    Next i
    ' Loop through any missing MRU
    On Error GoTo NoMoreToDelete
    For i = MRUCount + 1 To MaxMRU - 1
        ' Delete the removed MRU item
        DeleteSetting "Peter Chapman", "VBSE", "MRUFile" & Trim(CStr(i))
    Next i
NoMoreToDelete:
End Sub

' Scripting Error Handler
Private Sub ScriptMain_Error()
    MsgBox "Error " & ScriptMain.Error.Number & ": " & ScriptMain.Error.Description & vbCrLf & "On Line: " & ScriptMain.Error.Line & vbCrLf & vbCrLf & ScriptMain.Error.Text, vbCritical, "Script Error"
End Sub

' Textbox Change Event Handler
Private Sub TextMain_Change(Index As Integer)
    ' Update the changed flag
    TextChanged = True
    ' Update the status bar
    GetCursorCoordinates
End Sub

' Textbox Key Down Event Handler
Private Sub TextMain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    ' Update the status bar
    GetCursorCoordinates
    ' If selected text is being overwritten
    If KeyCode = vbKeyMenu Then
        TextMain_MouseDown Index, vbRightButton, Shift, 0, 0
    ElseIf Shift = 0 And TextMain(Index).SelLength > 0 Then
        ' Store the undo value
        UndoStart = TextMain(Index).SelStart
        UndoLength = TextMain(Index).SelLength
        UndoText = TextMain(Index).Text
    End If
End Sub

' Textbox Mouse Down Event Handler
Private Sub TextMain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the right mouse button
    If Button = vbRightButton Then
        ' Disable the textbox
        TextMain(Index).Enabled = False
        ' This DoEvents seems to be optional?
        DoEvents
        ' Re-enable the control, so that it doesn't appear as grayed
        TextMain(Index).Enabled = True
        TextMain(Index).SetFocus
        ' Show the custom menu
        PopupMenu MenuEdit
    ElseIf Button = vbLeftButton Then
        ' Update the status bar
        GetCursorCoordinates
    End If
End Sub

' Updates the language manu based on the open file
Private Sub UpdateFileLanguage()
    If FilePath = "Untitled" Then
        Exit Sub
    ElseIf Len(FilePath) > 4 And LCase(Right(FilePath, 4)) = ".vbs" Then
        MenuLanguageVBScript_Click
    ElseIf Len(FilePath) > 3 And LCase(Right(FilePath, 3)) = ".js" Then
        MenuLanguageJScript_Click
    Else
        MenuLanguageText_Click
    End If
End Sub
