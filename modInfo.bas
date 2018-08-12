Attribute VB_Name = "modInfo"
Option Explicit
Option Private Module

'Dieses Modul beinhaltet Funktionen des Info-Popups "frmInfo"

Public Sub infoINFO() 'Öffnen bzw. schließen des Popups
    If frmInfo.visible = False Then
        Load frmInfo
        frmInfo.Show
    Else
        Unload frmInfo
    End If
End Sub

Public Sub infoSOURCECODEURL() 'URL im Browser aufrufen
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:=xlef_strGitHub
    On Error GoTo 0
End Sub

Public Sub infoSPENDELINKURL() 'URL im Browser aufrufen
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:=xlef_strSpendenURL
    On Error GoTo 0
End Sub

Public Sub infoDOWNLOADURL() 'URL im Browser aufrufen
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:=xlef_strDownload
    On Error GoTo 0
End Sub


Public Sub infoFEEDBACK() 'Feedback E-Mail
    On Error Resume Next
        Dim objMail As Object 'Shell-Objekt für E-Mail
        Set objMail = CreateObject("Shell.Application")
        objMail.ShellExecute "mailto:" & xlef_strKontakt1 _
            & "&subject=" & "Feedback: " & xlef_strToolname & " Version " & xlef_strVersion & " / " _
            & Application.OperatingSystem & " / Excel-Version " & Application.Version
    On Error GoTo 0
End Sub
