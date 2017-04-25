Attribute VB_Name = "modVBAFunctions"
Option Explicit
Option Private Module

'VBA-Funktionen
'==============

'Als Value: Konvertieren in Großbuchstaben
Public Sub valueUCASE()
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    On Error Resume Next 'Fehler überspringen (z.B. wenn Zelle Fehler hat)
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = UCase(xlef_rngCell.Value)
    Next
End Sub

'Als Funktion: Konvertieren in Großbuchstaben
Public Sub funcUCASE()
    Call FunktionBau("UPPER")
End Sub

'Als Value: Konvertieren in Kleinbuchstaben
Public Sub valueLCASE()
    xlef_art = "cell"
    Call Aktion          'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    On Error Resume Next 'Fehler überspringen (z.B. wenn Zelle Fehler hat)
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = LCase(xlef_rngCell.Value)
    Next
End Sub

'Als Funktion: Konvertieren in Kleinbuchstaben
Public Sub funcLCASE()
    Call FunktionBau("LOWER")
End Sub


'Als Value: Leerzeichen am Anfang der Zelle entfernen
Public Sub valueLTRIM()
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    On Error Resume Next 'Fehler überspringen (z.B. wenn Zelle Fehler hat)
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = LTrim(xlef_rngCell.Value)
    Next
End Sub

'Als Value: Leerzeichen am Ende der Zelle entfernen
Public Sub valueRTRIM()
    xlef_art = "cell"
    Call Aktion         'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    On Error Resume Next 'Fehler überspringen (z.B. wenn Zelle Fehler hat)
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = RTrim(xlef_rngCell.Value)
    Next
End Sub

'Als Value: Leerzeichen am Anfang und Ende der Zelle entfernen
Public Sub valueTRIM()
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    On Error Resume Next 'Fehler überspringen (z.B. wenn Zelle Fehler hat)
    
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = Trim(xlef_rngCell.Value)
    Next
End Sub

'Als Value: Absolutwert einer Zahl (Zahl ohne Vorzeichen)
Public Sub valueABS()
    xlef_art = "cell"
    Call Aktion          'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    On Error Resume Next 'Fehler überspringen (z.B. wenn Zelle keine Zahl ist)
    
    For Each xlef_rngCell In xlef_Sel
        If xlef_rngCell.Value <> "" Then xlef_rngCell.Value = Abs(xlef_rngCell.Value)
    Next
End Sub

'Als Funktion: Absolutwert einer Zahl (Zahl ohne Vorzeichen)
Public Sub funcABS()
    Call FunktionBau("ABS")
End Sub


