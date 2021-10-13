VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} AddIn 
   ClientHeight    =   9480
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   22560
   _ExtentX        =   39793
   _ExtentY        =   16722
   _Version        =   393216
   Description     =   "This add-in extend standart VBE CodePane to receiving shortcuts."
   DisplayName     =   "VBE Extender"
   AppName         =   "Visual Basic for Applications IDE"
   AppVer          =   "6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0"
End
Attribute VB_Name = "AddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public IDE As VBIDE.VBE
Public MenuItem As Office.CommandBarButton
Public WithEvents MenuItemHandler As VBIDE.CommandBarEvents
Attribute MenuItemHandler.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    'On Error Resume Next

    Set IDE = Application
    
    Dim AddInsMenu As Office.CommandBar
    Set AddInsMenu = IDE.CommandBars("Add-Ins")
    
    If Not (AddInsMenu Is Nothing) Then
        'Set MenuItem = AddInsMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        'MenuItem.Caption = AddInName & "..."
    
'        Set MenuItemHandler = IDE.Events.CommandBarEvents(MenuItem)
    End If
    
    If Not EnableHotkey() Then
        MsgBox "Unable to initialize hotkey", vbCritical
    Else
    '    MsgBox "Enable HK"
    End If

    
    
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    MenuItem.Delete
    DisableHotkey
End Sub

Private Sub MenuItemHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next

'    Dim Dlg As New ChooseForm
'    Dlg.Initialize IDE
'    Dlg.Show 1
End Sub

