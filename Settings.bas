Attribute VB_Name = "Settings"
Option Explicit

' Open the settings file.
Public Sub OpenSettings(fileSettings As File, fileMode As FileModeEnum, fileAccess As FileAccessEnum)
    Err = 0
    On Error Resume Next

    fileSettings.Open App.Path & "\settings.ini", fileMode, fileAccess

    If fileAccess = fsAccessWrite Then
        If Err Then
            MsgBox "ERROR: " & Err.Number & " - " & Err.Description
        End If
    Else
        If Err Then
            If Err = -2147024894 Then 'File not found
                ' Defaults.
                SaveSettings fileSettings, settEnterBehaviourSendCRLF, settMonitorTypeRecv
                OpenSettings fileSettings, fsModeInput, fsAccessRead
            Else
                MsgBox "ERROR: " & Err.Number & " - " & Err.Description
            End If
        End If
    End If
End Sub

' Close the settings file.
Public Sub CloseSettings(fileSettings As File)
    fileSettings.Close
End Sub

' Write the settings to the file.
Public Sub SaveSettings(fileSettings As File, enterBehaviour As String, monitorType As String)
    OpenSettings fileSettings, fsModeOutput, fsAccessWrite
    
    fileSettings.LinePrint enterBehaviour
    fileSettings.LinePrint monitorType

    CloseSettings fileSettings
End Sub

' Grabs the settings from the file.
Public Function GetSettings(fileSettings As File) As Variant
    Dim arrSettings(1) As String

    OpenSettings fileSettings, fsModeInput, fsAccessRead
    arrSettings(0) = fileSettings.LineInputString
    arrSettings(1) = fileSettings.LineInputString
    CloseSettings fileSettings
    
    GetSettings = arrSettings
End Function

