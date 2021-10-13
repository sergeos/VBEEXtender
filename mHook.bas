Attribute VB_Name = "mHook"
Option Explicit

Private Const HC_ACTION           As Long = 0
Private Const HC_NOREMOVE         As Long = 3
Private Const MOD_CONTROL         As Long = &H2
Private Const WH_KEYBOARD         As Long = 2
Private Const VK_CONTROL          As Long = &H11
Private Const VK_LCONTROL         As Long = &HA2
Private Const VK_RCONTROL         As Long = &HA3
Private Const MAPVK_VK_TO_CHAR    As Long = 2

Private Declare Function MapVirtualKey Lib "user32" _
                         Alias "MapVirtualKeyW" ( _
                         ByVal wCode As Long, _
                         ByVal wMapType As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" _
                         Alias "SetWindowsHookExW" ( _
                         ByVal idHook As Long, _
                         ByVal lpfn As Long, _
                         ByVal hmod As Long, _
                         ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                         ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" ( _
                         ByVal hHook As Long, _
                         ByVal ncode As Long, _
                         ByVal wParam As Long, _
                         ByRef lParam As Any) As Long
Private Declare Function GetAsyncKeyState Lib "user32" ( _
                         ByVal vKey As Long) As Long

Private m_hHook     As Long     ' // Hook handle
Private m_bSubMode  As Boolean
Public IDE          As VBIDE.VBE


' // Enable hotkey Ctrl+K
Public Function EnableHotkey() As Boolean
    
    If m_hHook = 0 Then
    
        m_hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KbdProc, 0, App.ThreadID)
    
        If m_hHook = 0 Then
            Exit Function
        End If
        
        m_bSubMode = False
        
    End If
    
    EnableHotkey = True
    
End Function

' // Disable hotkey Ctrl+K and disable subcommand mode (if enabled)
Public Function DisableHotkey() As Boolean

    If m_hHook Then
        UnhookWindowsHookEx m_hHook
        m_hHook = 0
    End If
    
    m_bSubMode = False
    
    DisableHotkey = True
    
End Function

Private Function KbdProc( _
                 ByVal lCode As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long) As Long
    Dim sKeyName    As String
    Dim lChar       As Long
    Dim bCtrl       As Boolean
    
    If (lCode = HC_ACTION) And lParam >= 0 Then
        
        Select Case wParam
        Case VK_CONTROL, VK_LCONTROL, VK_RCONTROL
            ' // Ignored keycodes
        Case Else

            ' // Get CTRL state
            bCtrl = GetAsyncKeyState(VK_CONTROL) And &H8000&
            
            If Not m_bSubMode Then
                
                If wParam = vbKeyK And bCtrl Then
                
                    m_bSubMode = True
                    
                    'frmMain.lblMode.BackColor = vbGreen
                    'frmMain.lblMode.Caption = "Wait for subcommand"
                    
                    'MsgBox "Wait for subcommand"

                End If
            
            Else
                
'                If bCtrl Then
'                    sKeyName = "Ctrl+"
'                End If
'
                lChar = MapVirtualKey(wParam, MAPVK_VK_TO_CHAR)
                
                If lChar <= 0 Then
                    sKeyName = sKeyName & "[" & CStr(wParam) & "]"
                Else
                    sKeyName = sKeyName & ChrW$(lChar)
                End If
                
                'frmMain.lstEvents.AddItem sKeyName
                'MsgBox sKeyName
                
                ' // Break subcommand sequence
                m_bSubMode = False
                
                
                If bCtrl Then
                    Select Case wParam
                        Case 67 'comment
                            SmartCommentBlock False
                        Case 85 'uncomment
                            SmartCommentBlock True
                    End Select
                End If
                
                'frmMain.lblMode.BackColor = vbCyan
                'frmMain.lblMode.Caption = "Wait for command"
                'MsgBox "Wait for command"
''                With IDE.ActiveCodePane.CodeModule
''                    .CodePane.SetSelection 8, 22, 8, 31
'''                    .CodePane.
''                End With
''
''                Dim startRow    As Long
''                Dim endRow      As Long
''                Dim startCol    As Long
''                Dim endCol      As Long
''
''                IDE.ActiveCodePane.GetSelection startRow, startCol, endRow, endCol
''
''                sKeyName = IDE.ActiveCodePane.CodeModule.Lines(startRow, 1)
''                Debug.Print startRow, startCol, endRow, endCol
                
'                IDE.CommandBars("edit").Controls(IIf(InStr(1, sKeyName, "'") > 0, "Uncomment Block", "Comment Block")).Execute
                
            End If
            
        End Select
        
    End If
                     
    KbdProc = CallNextHookEx(0, lCode, wParam, ByVal lParam)
                     
End Function


'================================================
'Un/Comment all lines in Block
'================================================
Private Sub SmartCommentBlock(bRemove As Boolean)
On Error GoTo Err:
    Dim tLines$(), i&
    'dim bRemove As Boolean
    Dim aux$, lpos&, FirstPos&
    Dim eta$, sofsCol&, eofsCol&
    'Dim VBCM As VBIDE.CodeModule
    Dim pre$, Pos$
    Dim lngStartLine&, lngEndLine&, lngStartColumn&, lngEndColumn&

    'Set VBCM = VBInstance.ActiveCodePane.CodeModule
    
    'Get cursor position
    IDE.ActiveCodePane.GetSelection lngStartLine, lngStartColumn, lngEndLine, lngEndColumn
    'Fix line not fully selected
    If (lngEndColumn = 1) And (lngStartLine <> lngEndLine) Then lngEndLine = lngEndLine - 1
    'Get all the lines
    tLines$ = Split(IDE.ActiveCodePane.CodeModule.Lines(lngStartLine, (lngEndLine - lngStartLine) + 1), vbCrLf)

    'Exit on empty lines
    If (UBound(tLines$) < 0) Then Exit Sub
    
    'Detect if we need to add or remove comment
'    If (Left$(Trim$(tLines$(0)), 1) = "'") Then bRemove = True

    For i = lngStartLine To lngEndLine
        aux$ = tLines$(i - lngStartLine)
        eta$ = aux$
        lpos& = CountStartingSpacesTabs(aux$)
        If (bRemove) Then
            'skip empty lines
            If (aux$ <> "") Then
                'Check if line got ' as first text
                If (Mid$(aux$, lpos + 1, 1) = "'") Then
                    'Remove it
                    pre$ = Left$(aux$, lpos)
                    Pos$ = Mid$(aux$, lpos + 2)
                    aux$ = pre$ & Pos$
                    IDE.ActiveCodePane.CodeModule.ReplaceLine i, aux$
                    If i = lngStartLine Then
                        If eta$ <> aux$ Then
                            sofsCol = -1
                        End If
                    ElseIf i = lngEndLine Then
                        If eta$ <> aux$ Then
                            eofsCol = -1
                        End If
                    End If
                End If
            End If
        Else
            'Add comment
            
            'Keep the beginning of text of the first comment line
            If (i = lngStartLine) Then
                FirstPos& = lpos&
            Else
                'Try to put a comment character in the same column for every other line
                'we check if the number of spaces at the begging of the other lines (lpos)
                'is bigger than the first=> split there
                If (lpos > FirstPos) Then lpos = FirstPos
            End If
            pre$ = Left$(aux$, lpos)
            Pos$ = Mid$(aux$, lpos + 1)
            aux$ = pre$ & "'" & Pos$
            IDE.ActiveCodePane.CodeModule.ReplaceLine i, aux$
            If i = lngStartLine Then
                sofsCol = 1
            ElseIf i = lngEndLine Then
                eofsCol = 1
            End If
        End If
    
    Next
    
    'Set selection to the new block position
    'IDE.ActiveCodePane.SetSelection lngStartLine, 1, lngEndLine, 999
    IDE.ActiveCodePane.SetSelection lngStartLine, lngStartColumn + sofsCol, lngEndLine, lngEndColumn + eofsCol
    
    Exit Sub
Err:
    Debug.Print "SmartCommentBlock Err: " & Err.Description
End Sub




'===================================================
'Count the number of spaces and tabs at the
'beggining of a string
'===================================================
Private Function CountStartingSpacesTabs(ByVal s$) As Long
On Error GoTo Err:
Dim i&
    
    If (s$ = "") Then CountStartingSpacesTabs = 0: Exit Function
    i = 0
    s$ = Replace(s$, vbTab, String$(4, " "))
    Do While (Left(s$, 1) = " ")
        i = i + 1
        s = Mid$(s$, 2)
    Loop
    CountStartingSpacesTabs = i
    Exit Function
Err:
    Debug.Print "CountStartingSpacesTabs Err: " & Err.Description
End Function












'Option Explicit
'
'Private Const MSG_WINDOW_CLASS  As String = "ExtHotKey_123423908493"
'
'Private Const HC_ACTION           As Long = 0
'Private Const HC_NOREMOVE         As Long = 3
'Private Const HWND_MESSAGE        As Long = -3
'Private Const WM_NCCREATE         As Long = &H81
'Private Const WM_HOTKEY           As Long = &H312
'Private Const MOD_CONTROL         As Long = &H2
'Private Const WH_KEYBOARD         As Long = 2
'Private Const VK_CONTROL          As Long = &H11
'Private Const VK_LCONTROL         As Long = &HA2
'Private Const VK_RCONTROL         As Long = &HA3
'Private Const MAPVK_VK_TO_CHAR    As Long = 2
'
'Private Type WNDCLASSEX
'    cbSize          As Long
'    style           As Long
'    lpfnwndproc     As Long
'    cbClsextra      As Long
'    cbWndExtra2     As Long
'    hInstance       As Long
'    hIcon           As Long
'    hCursor         As Long
'    hbrBackground   As Long
'    lpszMenuName    As Long
'    lpszClassName   As Long
'    hIconSm         As Long
'End Type
'
'Private Declare Function MapVirtualKey Lib "user32" _
'                         Alias "MapVirtualKeyW" ( _
'                         ByVal wCode As Long, _
'                         ByVal wMapType As Long) As Long
'Private Declare Function RegisterHotKey Lib "user32" ( _
'                         ByVal hwnd As Long, _
'                         ByVal id As Long, _
'                         ByVal fsModifiers As Long, _
'                         ByVal vk As Long) As Long
'Private Declare Function UnregisterHotKey Lib "user32" ( _
'                         ByVal hwnd As Long, _
'                         ByVal id As Long) As Long
'Private Declare Function RegisterClassEx Lib "user32" _
'                         Alias "RegisterClassExW" ( _
'                         ByRef pcWndClassEx As WNDCLASSEX) As Integer
'Private Declare Function UnregisterClass Lib "user32" _
'                         Alias "UnregisterClassW" ( _
'                         ByVal lpClassName As Long, _
'                         ByVal hInstance As Long) As Long
'Private Declare Function CreateWindowEx Lib "user32" _
'                         Alias "CreateWindowExW" ( _
'                         ByVal dwExStyle As Long, _
'                         ByVal lpClassName As Long, _
'                         ByVal lpWindowName As Long, _
'                         ByVal dwStyle As Long, _
'                         ByVal x As Long, _
'                         ByVal y As Long, _
'                         ByVal nWidth As Long, _
'                         ByVal nHeight As Long, _
'                         ByVal hWndParent As Long, _
'                         ByVal hMenu As Long, _
'                         ByVal hInstance As Long, _
'                         ByRef lpParam As Any) As Long
'Private Declare Function SetWindowsHookEx Lib "user32" _
'                         Alias "SetWindowsHookExW" ( _
'                         ByVal idHook As Long, _
'                         ByVal lpfn As Long, _
'                         ByVal hmod As Long, _
'                         ByVal dwThreadId As Long) As Long
'Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
'                         ByVal hHook As Long) As Long
'Private Declare Function DestroyWindow Lib "user32" ( _
'                         ByVal hwnd As Long) As Long
'Private Declare Function CallNextHookEx Lib "user32" ( _
'                         ByVal hHook As Long, _
'                         ByVal ncode As Long, _
'                         ByVal wParam As Long, _
'                         ByRef lParam As Any) As Long
'Private Declare Function GetAsyncKeyState Lib "user32" ( _
'                         ByVal vKey As Long) As Long
'
'Private m_iClassAtom    As Integer  ' // Class atom of window
'Private m_hWndMsg       As Long     ' // Messagewindow handle
'Private m_hHook         As Long     ' // Hook handle
'
'' // Enable hotkey Ctrl+K
'Public Function EnableHotkey() As Boolean
'
'    If m_hWndMsg = 0 Then
'        Exit Function
'    End If
'
'    EnableHotkey = RegisterHotKey(m_hWndMsg, 1, MOD_CONTROL, vbKeyK)
'
'End Function
'
'' // Disable hotkey Ctrl+K and disable subcommand mode (if enabled)
'Public Function DisableHotkey() As Boolean
'
'    If m_hWndMsg = 0 Then
'        Exit Function
'    End If
'
'    If m_hHook Then
'        UnhookWindowsHookEx m_hHook
'        m_hHook = 0
'    End If
'
'    DisableHotkey = UnregisterHotKey(m_hWndMsg, 1)
'
'End Function
'
'' // Init message window
'Public Function InitMessageWindow() As Boolean
'    Dim tWndClass   As WNDCLASSEX
'
'    If m_iClassAtom = 0 Then
'
'        With tWndClass
'
'            .cbSize = LenB(tWndClass)
'            .hInstance = App.hInstance
'            .lpfnwndproc = FAR_PROC(AddressOf MsgWndProc)
'            .lpszClassName = StrPtr(MSG_WINDOW_CLASS)
'
'        End With
'
'        m_iClassAtom = RegisterClassEx(tWndClass)
'
'        If m_iClassAtom = 0 Then
'            Exit Function
'        End If
'
'    End If
'
'    m_hWndMsg = CreateWindowEx(0, StrPtr(MSG_WINDOW_CLASS), 0, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, ByVal 0&)
'
'    If m_hWndMsg = 0 Then
'
'        UnregisterClass StrPtr(MSG_WINDOW_CLASS), App.hInstance
'        m_iClassAtom = 0
'        Exit Function
'
'    End If
'
'    InitMessageWindow = True
'
'End Function
'
'Public Sub UninitMessageWindow()
'
'    DisableHotkey
'
'    If m_hWndMsg Then
'        DestroyWindow m_hWndMsg
'        m_hWndMsg = 0
'    End If
'
'    If m_iClassAtom Then
'        UnregisterClass StrPtr(MSG_WINDOW_CLASS), App.hInstance
'        m_iClassAtom = 0
'    End If
'
'End Sub
'
'Private Function SwitchToSubCommandMode() As Boolean
'
'    If m_hHook = 0 Then
'
'        m_hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KbdProc, 0, App.ThreadID)
'
'        If m_hHook = 0 Then
'            Exit Function
'        End If
'
'    End If
'
'    'frmMain.lblMode.BackColor = vbGreen
'    Debug.Print "Wait for subcommand"
'
'    SwitchToSubCommandMode = True
'
'End Function
'
'Private Function KbdProc( _
'                 ByVal lCode As Long, _
'                 ByVal wParam As Long, _
'                 ByVal lParam As Long) As Long
'    Dim sKeyName    As String
'    Dim lChar       As Long
'
'    If (lCode = HC_ACTION Or lCode = HC_NOREMOVE) And lParam >= 0 Then
'
'        Select Case wParam
'        Case VK_CONTROL, VK_LCONTROL, VK_RCONTROL
'            ' // Ignored keycodes
'        Case Else
'
'            ' // Get CTRL state
'            If GetAsyncKeyState(VK_CONTROL) Then
'                sKeyName = "Ctrl+"
'            End If
'
'            lChar = MapVirtualKey(wParam, MAPVK_VK_TO_CHAR)
'
'            If lChar <= 0 Then
'                sKeyName = sKeyName & "[" & CStr(wParam) & "]"
'            Else
'                sKeyName = sKeyName & ChrW$(lChar)
'            End If
'
'            Debug.Print sKeyName
'
'            ' // Break subcommand sequence
'            UnhookWindowsHookEx m_hHook
'            m_hHook = 0
'
'            'frmMain.lblMode.BackColor = vbCyan
'            Debug.Print "Wait for command"
'
'        End Select
'
'    End If
'
'    KbdProc = CallNextHookEx(0, lCode, wParam, ByVal lParam)
'
'End Function
'
'Private Function MsgWndProc( _
'                 ByVal hwnd As Long, _
'                 ByVal lMsg As Long, _
'                 ByVal wParam As Long, _
'                 ByVal lParam As Long) As Long
'
'    Select Case lMsg
'    Case WM_NCCREATE
'        MsgWndProc = 1
'    Case WM_HOTKEY
'        If wParam = 1 And m_hHook = 0 Then
'            If Not SwitchToSubCommandMode Then
'                MsgBox "Unable to switch to subcommand mode", vbCritical
'            End If
'        End If
'    End Select
'
'End Function
'
'Private Function FAR_PROC( _
'                 ByVal pfn As Long) As Long
'    FAR_PROC = pfn
'End Function
'
'
'


