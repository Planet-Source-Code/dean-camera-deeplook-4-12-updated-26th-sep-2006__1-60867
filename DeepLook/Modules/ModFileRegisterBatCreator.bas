Attribute VB_Name = "ModFileRegisterBatCreator"
'  .======================================.
' /         DeepLook Project Scanner       \
' |       By Dean Camera, 2003 - 2005      |
' \  Visual Basic Project Scanning Engine  /
'  '======================================'
' / Most of this project is now commented  \
' \           to help developers.          /
'  '======================================'

Option Explicit

'-----------------------------------------------------------------------------------------------
Dim BATFileNum As Integer
'-----------------------------------------------------------------------------------------------

Public Sub CreateBatHeader(FileName As String)
    BATFileNum = FreeFile
    Open FileName For Output As #BATFileNum

    Print #BATFileNum, "@echo off" & vbCrLf & "echo                    En-Tech DeepLook Project Scanner" & vbCrLf & _
        "echo             *** Automatic file register batch script file ***" & vbCrLf & _
        "echo -------------------------------------------------------------------------------" & vbCrLf & _
        "echo You must be using WinME/98/95 and have the RegSvr32.exe in your windows folder." & vbNewLine & "echo." & _
        vbCrLf & "pause" & vbCrLf & "cls"
End Sub

Public Sub AddBatRegAndCopyFile(FileName As String, Findex As Long, Fmax As Long)
    Print #BATFileNum, "echo *** Copying File #" & Findex & " of " & Fmax & " (" & FileName & ")..."
    Print #BATFileNum, "echo." & vbCrLf & "copy """ & FileName & """, """ & "%WINDIR%\System\" & FileName & """"
    Print #BATFileNum, "echo *** Registering File #" & Findex & " of " & Fmax & " (" & FileName & ")..."
    Print #BATFileNum, "%WINDIR%\System\Regsvr32.exe ""%WINDIR%\System\" & FileName & """ /s"
    Print #BATFileNum, "wait 1" & vbCrLf & "cls"
End Sub

Public Sub AddBatFooter(FileName As String)
    Print #BATFileNum, "echo." & vbCrLf & "echo." & vbCrLf & "echo File copy/registration complete." & vbCrLf & "pause"
    
    Close #BATFileNum
End Sub
