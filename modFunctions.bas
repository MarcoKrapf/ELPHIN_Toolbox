Attribute VB_Name = "modFunctions"
Option Explicit
Option Private Module

'Zellinhalt in Worksheet-Funktion einbauen
Public Sub FunktionBau(funktion As String)
    Dim strFormel As String  'Variable f�r zurechtgestutzte Formel/Funktion
    
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken f�r Wiederherstellung
    
    On Error Resume Next 'Fehler �berspringen (z.B. wenn Zelle Fehler hat)
    
    For Each xlef_rngCell In xlef_Sel
        If xlef_rngCell.Value <> "" Then 'nur wenn Zelle nicht leer ist
            If xlef_rngCell.HasFormula Then
                strFormel = xlef_rngCell.Formula 'wenn Zelle eine Formel/Funktion enth�lt
            Else
                strFormel = xlef_rngCell.Value 'wenn Zelle keine Formel/Funktion enth�lt
            End If

            'Formel extrahieren ('=' und '+' am Anfang entfernen)
            strFormel = FormelExtrakt(strFormel)
            'Formel neu zusammenbauen und in Zelle schreiben
            If xlef_rngCell.HasFormula Then
                xlef_rngCell.Formula = FormelNeu(funktion, strFormel) 'wenn Zelle eine Formel/Funktion enth�lt
            Else
                xlef_rngCell.Formula = FormelNeuText(funktion, strFormel) 'wenn Zelle keine Formel/Funktion enth�lt
            End If

        End If
    Next
End Sub

'Formel/Funktion stutzen ('=' und '+' am Anfang entfernen)
Public Function FormelExtrakt(str As String) As String
    Do While Left(str, 1) = "=" Or Left(str, 1) = "+"
        str = Right(str, Len(str) - 1)
    Loop
    FormelExtrakt = str
End Function

'Formel/Funktion neu zusammenbauen (Argument auf dem Worksheet ohne Anf�hrungszeichen "")
Public Function FormelNeu(fn As String, strFormel As String) As String
    FormelNeu = "=" & fn & "(" & strFormel & ")"
End Function

'Formel/Funktion neu zusammenbauen (Argument auf dem Worksheet in Anf�hrungszeichen "")
Public Function FormelNeuText(fn As String, strFormel As String) As String
    FormelNeuText = "=" & fn & "(" & Chr(34) & strFormel & Chr(34) & ")"       ' Chr(34) erzeugt ein Anf�hrungszeichen "
End Function

