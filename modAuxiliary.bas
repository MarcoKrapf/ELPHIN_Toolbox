Attribute VB_Name = "modAuxiliary"
Option Explicit
Option Private Module

'Modul mit Hilfsfunktionen
'=========================

'Sprache ändern
Public Sub Sprache()
    If xlef_strSprache = 2 Then
        xlef_strSprache = 3
    Else
        xlef_strSprache = 2
    End If
    MyRibbon.Invalidate
End Sub

'Textbausteine holen
Public Function GetText(id As String) As String
    Dim r As Long
    For r = 1 To 10000 'xlef_wksTexte.Cells(Rows.Count, 1).End(xlUp).row '--> stürzt ab(?!)
        If xlef_wksTexte.Cells(r, 1).Value = id Then
            GetText = xlef_wksTexte.Cells(r, xlef_strSprache).Value
            Exit Function
        End If
    Next
End Function

'UserForm in der Mitte platzieren
Public Sub PlaceUserFormInCenter(frmMe As Object)
    With frmMe
        .StartUpPosition = 0
        .Top = ActiveWindow.Top + ((ActiveWindow.Height - frmMe.Height) / 2)
        .Left = ActiveWindow.Left + ((ActiveWindow.Width - frmMe.Width) / 2)
    End With
End Sub

'UsedRange selektieren
Public Sub SelectUsedRange()
    ActiveWorkbook.ActiveSheet.UsedRange.Select
End Sub

'UsedRange einlesen
Public Sub TakeUsedRange()
    Set xlef_Sel = Selection 'Selektierten Bereich in Variable einlesen
End Sub

'Aktionen beim Starten einer Funktion
Public Sub Aktion()

    xlef_blnDo = True
    
    Set xlef_wksTarget = ActiveWorkbook.ActiveSheet
    Set xlef_Sel = Selection 'Selektierten Bereich in Variable einlesen
    Call SelectionSave 'Aktuellen Zustand merken für Wiederherstellung
    
End Sub
