Attribute VB_Name = "modWorksheetFunctions"
Option Explicit
Option Private Module

'Worksheet-Funktionen
'====================

'Als Value: Erster Buchstabe jedes Wortes groß (=GROSS2)
Public Sub valuePROPER()
    xlef_art = "cell"
    Call Aktion   'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnDo = False Then Exit Sub
    
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    On Error GoTo FEHLER
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = WorksheetFunction.Proper(xlef_rngCell.Value)
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
    
    Exit Sub
FEHLER:
    Call ErrorPopup(err)
End Sub

'Als Funktion: Erster Buchstabe jedes Wortes groß (=GROSS2)
Public Sub funcPROPER()
    Call FunktionBau("PROPER")
End Sub

'Als Value: Überflüssige Leerzeichen entfernen (=GLÄTTEN)
Public Sub valueTRIM_WKS()
    xlef_art = "cell"
    Call Aktion   'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnDo = False Then Exit Sub
    
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    On Error GoTo FEHLER
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = WorksheetFunction.Trim(xlef_rngCell.Value)
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
    
    Exit Sub
FEHLER:
    Call ErrorPopup(err)
End Sub

'Als Funktion: Überflüssige Leerzeichen entfernen (=GLÄTTEN)
Public Sub funcTRIM_WKS()
    Call FunktionBau("TRIM")
End Sub

'Zufallszahl
Public Sub fnRAND()
    MsgBox "Macht das Sinn? Wenn ja, dann mit Unter/Obergrenze und Anzahl Nachkommastellen. Nur hart oder auch als Funktion?"
    xlef_art = "cell"
    Call Aktion   'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnDo = False Then Exit Sub
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = WorksheetFunction.RandBetween(0, 100)
    Next
End Sub


