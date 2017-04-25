Attribute VB_Name = "modWorksheetFunctions"
Option Explicit
Option Private Module

'Worksheet-Funktionen
'====================

'Als Value: Erster Buchstabe jedes Wortes gro� (=GROSS2)
Public Sub valuePROPER()
    xlef_art = "cell"
    Call Aktion   'Selektion einlesen und Zellinhalte merken f�r Wiederherstellung
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = WorksheetFunction.Proper(xlef_rngCell.Value)
    Next
End Sub

'Als Funktion: Erster Buchstabe jedes Wortes gro� (=GROSS2)
Public Sub funcPROPER()
    Call FunktionBau("PROPER")
End Sub

'Als Value: �berfl�ssige Leerzeichen entfernen (=GL�TTEN)
Public Sub valueTRIM_WKS()
    xlef_art = "cell"
    Call Aktion   'Selektion einlesen und Zellinhalte merken f�r Wiederherstellung
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = WorksheetFunction.Trim(xlef_rngCell.Value)
    Next
End Sub

'Als Funktion: �berfl�ssige Leerzeichen entfernen (=GL�TTEN)
Public Sub funcTRIM_WKS()
    Call FunktionBau("TRIM")
End Sub

'Zufallszahl
Public Sub fnRAND()
    MsgBox "Macht das Sinn? Wenn ja, dann mit Unter/Obergrenze und Anzahl Nachkommastellen. Nur hart oder auch als Funktion?"
    xlef_art = "cell"
    Call Aktion   'Selektion einlesen und Zellinhalte merken f�r Wiederherstellung
    For Each xlef_rngCell In xlef_Sel
        xlef_rngCell.Value = WorksheetFunction.RandBetween(0, 100)
    Next
End Sub


