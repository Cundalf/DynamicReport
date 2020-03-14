Attribute VB_Name = "mdlMain"
Option Explicit

Sub Main()
    PATH_INI = Environ$("USERPROFILE") & "\CADR_config.ini"
    InitConfig
    mdiMain.Show
End Sub

Private Sub InitConfig()
    ' Generate a configuration file with default values.

    On Error GoTo error
    If Not FileExists(PATH_INI) Then
        Open PATH_INI For Output As #1
        Close #1
    End If
    
    If IniGet(PATH_INI, "DB", "IP") = "" Then IniWrite PATH_INI, "DB", "IP", "IP"
    If IniGet(PATH_INI, "DB", "NAMESPACE") = "" Then IniWrite PATH_INI, "DB", "NAMESPACE", "NAMESPACE"
    If IniGet(PATH_INI, "CONFIG", "EXPORT") = "" Then IniWrite PATH_INI, "CONFIG", "EXPORT", "3"
    
    Exit Sub
error:
    MsgBox "The required configuration could not be generated.", vbCritical + vbOKOnly, App.ProductName
    End
End Sub
