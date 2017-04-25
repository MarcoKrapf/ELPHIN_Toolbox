Attribute VB_Name = "DEV_modSPIELPLATZ"
Option Explicit

'Modul zum Ausprobieren während Entwicklung und Test
'===================================================

'-->modVBAFunctions
'Liest Anzahl der Zeichen in einer Zelle aus (nur Buchstaben und Leerzeichen, ???keine Punkte(???????)) (=LÄNGE)
Public Sub fnLEN()
    MsgBox "Was machen wir damit? Als MsgBox lassen?"
    Set xlef_Sel = Selection 'Selektierten Bereich in Variable einlesen
    On Error Resume Next 'Fehler überspringen (z.B. wenn Zelle keine Zahl ist)
    For Each xlef_rngCell In xlef_Sel
        If MsgBox("Zelle " & xlef_rngCell.Address(False, False) & ": " & xlef_rngCell.Value & vbCrLf & _
            "Länge: " & Len(xlef_rngCell.Value), vbInformation + vbOKCancel) = vbCancel Then Exit Sub 'Prüfung abbrechen wenn CANCEL geklickt wird
    Next
End Sub

'Sind zwei Zellen identisch (=DBANZAHL)
Public Sub fnDCOUNT()
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    For Each xlef_rngCell In xlef_Sel
        'xlef_rngCell.Value = WorksheetFunction.DCount
        MsgBox "Länge: " & WorksheetFunction.DbCount '(xlef_rngCell.Value)
    Next
End Sub


'Inhalt aus Windows-Zwischenablage ausgeben
Public Sub fnGETCLIPBOARD()
    MsgBox "Brauchen wir das?"
    Dim clip As Variant
    
    Set xlef_objClip = New DataObject
    
    xlef_objClip.GetFromClipboard
    clip = xlef_objClip.GetText
    
    Range(ActiveCell.Address).Value = clip
End Sub


'SVERWEIS vereinfachen,(=SVERWEIS(Suchkriterium,Spalte,Rückgabespalte), sodass Eingangsparater Matrix nicht benötigt wird, da bei großen Tabellen kompliziert Rückgabespalte auszurechnen
Public Function fnSVERWEIS(Suchkriterium As Variant, Spalte As Variant, Rückgabespalte As Variant) As Variant

End Function



'Führende Nullen löschen
Function fnFNullenWeg(strText As String) As String
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If Left$(strText, 1) = "0" Then
            strText = Right$(strText, Len(strText) - 1)
          Else
            Exit For
        End If
    Next i
    fnFNullenWeg = strText
End Function

Sub Test()
    Dim OutApp As Object, Mail As Object, i
    Dim Nachricht

' nachfolgend den gewünschten Tabellenbereich einstellen
    Range(Selection.Address).Copy

' Öffnen der Mail
        Set OutApp = CreateObject("Outlook.Application")
        Set Nachricht = OutApp.CreateItem(0)
        With Nachricht
            .Subject = ActiveSheet.Name
            .To = "@"
            .Display
        End With
        Set OutApp = Nothing
        Set Nachricht = Nothing

'Kurz warten, damit die Mail Zeit zum Öffnen hat
'        Application.Wait (Now + TimeValue("0:00:01"))

' Dann die Zwischenablage einfügen
        Application.SendKeys ("^v") ' Strg-V
 
    
End Sub

'Bedingte Spaltenlöschung in gesamter Tabelle
Public Sub BEDCOLCLEAR()
 Dim luc As Integer
 Dim i As Integer
 Dim intRowInd As Integer
 intRowInd = InputBox(prompt:="Bitte definieren Sie den Spaltenindex des Suchwerts.", Title:="Spaltenindex des Suchwerts")
 strSearchContent = InputBox(prompt:="Bitte definieren Sie das Suchwort.", Title:="Suchwert eingeben")
    
luc = Cells(1, Columns.Count).End(xlToLeft).Columns.Column 'Ermittlung der letzten Spalte in Zeile 1
'** Durchlauf aller Zeile
For i = luc To 1 Step -1 'Zählung rückwärts bis Spalte 1

If Cells(intRowInd, i).Value = strSearchContent Then 'Abfragen, ob in der vorher definierten Spalte der vorher definierter Suchbegriff steht
Columns(i).Delete Shift:=xlToLeft
End If
Next s
End Sub
'Wechseln-Wenn
Public Function WECHSELNWENN(Formel As Variant, Wert As Variant, Ersatz As Variant) As Variant
If Formel = Wert Then
WECHSELNWENN = Ersatz
Else
WECHSELNWENN = Formel
End If
End Function

'Wenn Formelergebnis 0 ergibt, dann Zelle leer darstellen
Public Sub fnLEERWENNNULL_WW()
    Dim strFormel As String  'Variable für zurechtgestutzte Formel
    
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    For Each xlef_rngCell In xlef_Sel
        If Not IsError(xlef_rngCell.Value) Then
            If xlef_rngCell.Value = 0 Then 'nur wenn Zelle eine Formel mit dem Ergebnis 0 enthält
                If xlef_rngCell.HasFormula Then
                    strFormel = xlef_rngCell.Formula
                    'Formel extrahieren ('=' und '+' am Anfang entfernen)
                    strFormel = FormelExtrakt(strFormel)
                    'Formel neu zusammenbauen und in Zelle schreiben
                    xlef_rngCell.Formula = "=WECHSELNWENN(" & strFormel & ",0,"""")"   ' Chr(34) erzeugt ein Anführungszeichen "
                Else
                    xlef_rngCell.Value = ""
                End If
            End If
        End If
        
    Next
End Sub

'Fehler löschen
Public Sub ERRORCLEAN()
    
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    For Each xlef_rngCell In xlef_Sel
        If IsError(xlef_rngCell.Value) Then 'Wenn Fehler dann "" anzeigen
            If xlef_rngCell.HasFormula Then
                strFormel = xlef_rngCell.Formula
                'Formel extrahieren ('=' und '+' am Anfang entfernen)
                strFormel = FormelExtrakt(strFormel)
                'Formel neu zusammenbauen und in Zelle schreiben
                xlef_rngCell.Formula = "=IFERROR(" & strFormel & ","""")"
            Else
                xlef_rngCell.Value = ""
            End If
        End If
    Next
End Sub

Sub Test2()
ActiveSheet.UsedRange.
End Sub
