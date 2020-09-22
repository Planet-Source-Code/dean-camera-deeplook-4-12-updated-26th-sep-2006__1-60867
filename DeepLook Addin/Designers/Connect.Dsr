VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7605
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10080
   _ExtentX        =   17780
   _ExtentY        =   13414
   _Version        =   393216
   Description     =   "Allows one-click scanning of open projects with DeepLook."
   DisplayName     =   "DeepLook Helper Addin"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================================================
Public VBInstance                           As VBIDE.VBE
Private OptionsForm                         As New frmOptions
Private AboutForm                           As New frmAbout
Private DLToolbar                           As CommandBar
Private ScanProjectButton                   As CommandBarControl
Private OptionsButton                       As CommandBarControl
Private AboutButton                         As CommandBarControl
Private WithEvents ScanProjectButtonEvents  As CommandBarEvents
Attribute ScanProjectButtonEvents.VB_VarHelpID = -1
Private WithEvents OptionsButtonEvents      As CommandBarEvents
Attribute OptionsButtonEvents.VB_VarHelpID = -1
Private WithEvents AboutButtonEvents        As CommandBarEvents
Attribute AboutButtonEvents.VB_VarHelpID = -1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
'========================================================================================================

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set VBInstance = Application

    CreateDLToolBar
    CheckIfEXEFound

    Set ScanProjectButtonEvents = VBInstance.Events.CommandBarEvents(ScanProjectButton)
    Set OptionsButtonEvents = VBInstance.Events.CommandBarEvents(OptionsButton)
    Set AboutButtonEvents = VBInstance.Events.CommandBarEvents(AboutButton)
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    OptionsForm.Hide
    AboutForm.Hide
    Unload OptionsForm
    Unload AboutForm

    SaveSetting "DLAddin", "Options", "Position", DLToolbar.Position
    SaveSetting "DLAddin", "Options", "Top", DLToolbar.Top
    SaveSetting "DLAddin", "Options", "Left", DLToolbar.Left
    SaveSetting "DLAddin", "Options", "RowIndex", DLToolbar.RowIndex

    Set OptionsForm = Nothing
    Set AboutForm = Nothing

    ScanProjectButton.Delete
    OptionsButton.Delete
    AboutButton.Delete
    DLToolbar.Delete
End Sub

Private Sub OptionsButtonEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    OptionsForm.LoadOptions
    OptionsForm.Show
    CheckIfEXEFound
End Sub

Private Sub AboutButtonEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    AboutForm.Show
End Sub

Private Sub ScanProjectButtonEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim EXEPath As String
    Dim Count As Long
    Dim DirtyProjects As Boolean
    Dim ContinueScan As Integer

    EXEPath = GetSetting("DLAddin", "Options", "EXEpath", vbNullString)

    For Count = 1 To VBInstance.VBProjects.Count
        If VBInstance.VBProjects(Count).IsDirty Then DirtyProjects = True
    Next

    If DirtyProjects And GetSetting("DLAddin", "Options", "ShowWarning", 1) Then ' Dirty means a project has been changed since last save
        ContinueScan = MsgBox("One or more open projects have been modified since last change." & vbNewLine & "DeepLook results will be based on the files as they were last saved." & vbNewLine & vbNewLine & "Continue with scan?", vbQuestion + vbYesNo, "DeepLook Addin")
        If ContinueScan = vbNo Then Exit Sub
    End If

    If Not ((EXEPath = vbNullString) Or (Dir(EXEPath) = vbNullString)) Then
        If VBInstance.VBProjects.Count > 1 Then ' Group file opened
            If InStr(VBInstance.VBProjects.FileName, ":\") > 0 Then ' Group project saved
                ShellExecute frmOptions.hwnd, "Open", EXEPath, VBInstance.VBProjects.FileName, "C:\", SW_SHOWNORMAL
            Else
                MsgBox "Error: The current group file has not been saved." & vbNewLine & "DeepLook scan cannot commence until the group file and it's associated projects have been saved.", vbCritical
            End If
        Else ' Single project opened
            If InStr(VBInstance.ActiveVBProject.FileName, ":\") > 0 Then ' File saved
                ShellExecute frmOptions.hwnd, "Open", EXEPath, VBInstance.ActiveVBProject.FileName, "C:\", SW_SHOWNORMAL
            Else
                MsgBox "Error: The current project file has not been saved." & vbNewLine & "DeepLook scan cannot commence until the project has been saved.", vbCritical
            End If
        End If
    End If
End Sub

Private Sub CreateDLToolBar()
    If GetSetting("DLAddin", "Options", "Position", vbNullString) <> vbNullString Then ' Addin has been opened before
        Set DLToolbar = VBInstance.CommandBars.Add("DeepLook Toolbar", GetSetting("DLAddin", "Options", "Position"))
    
        DLToolbar.Top = GetSetting("DLAddin", "Options", "Top")
        DLToolbar.Left = GetSetting("DLAddin", "Options", "Left")
        DLToolbar.RowIndex = GetSetting("DLAddin", "Options", "RowIndex")
    Else
        Set DLToolbar = VBInstance.CommandBars.Add("DeepLook Toolbar", msoBarNoChangeDock)
    End If

    DLToolbar.Visible = True

    Set ScanProjectButton = DLToolbar.Controls.Add(msoControlButton)
    Set OptionsButton = DLToolbar.Controls.Add(msoControlButton)
    Set AboutButton = DLToolbar.Controls.Add(msoControlButton)

    With ScanProjectButton
        .Caption = "Scan with DeepLook"
        .FaceId = 526
    End With
    With OptionsButton
        .Caption = "DeepLook Addin Options"
        .FaceId = 991
    End With
    With AboutButton
        .Caption = "About DeepLook Addin"
        .FaceId = 1014
    End With
End Sub

Public Sub CheckIfEXEFound()
    Dim EXEPath As String

    EXEPath = GetSetting("DLAddin", "Options", "EXEpath", vbNullString)

    If (EXEPath = vbNullString) Or (Dir(EXEPath) = vbNullString) Then
        ScanProjectButton.Enabled = False
    End If
End Sub
