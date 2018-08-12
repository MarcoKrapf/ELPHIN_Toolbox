Attribute VB_Name = "modCustomFunctions"
Option Explicit
Option Private Module

'Alle leeren Zeilen löschen
Sub fnLEEREZEILEN()
    Dim i As Integer 'Zählvariable für Schleife

    xlef_art = "row"
    Call Aktion   'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnProgressBar And xlef_coll.Count > 0 Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_coll.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If

    For i = xlef_coll.Count To 1 Step -1 'Collection rückwärts durchlaufen
        ActiveSheet.Rows(xlef_coll(i)).Delete Shift:=xlUp 'unterste leere Zeile löschen und hochschieben
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next i

    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Alle leeren Spalten löschen
Sub fnLEERESPALTEN()
    Dim i As Integer 'Zählvariable für Schleife

    xlef_art = "col"
    Call Aktion   'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnProgressBar And xlef_coll.Count > 0 Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_coll.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If

    For i = xlef_coll.Count To 1 Step -1 'Collection rückwärts durchlaufen
        ActiveSheet.Columns(xlef_coll(i)).Delete Shift:=xlUp 'rechteste leere Spalte löschen und nach links schieben
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next i

    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Formel als Wert einfügen
Public Sub valueFORMELzuWERT()
    xlef_art = "cell"
    Call Aktion   'Selektion einlesen und Formeln merken für Wiederherstellung
    
    If xlef_blnDo = False Then Exit Sub
    
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    For Each xlef_rngCell In xlef_Sel
        With xlef_rngCell
            If .HasFormula Then 'nur wenn die Zelle eine Formel enthält
                .Copy
                .PasteSpecial Paste:=xlValues
            End If
        End With
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Markierten Bereich in geichem Format in neue E-Mail kopieren // http://www.online-excel.de/excel/singsel_vba.php?f=28
Public Sub fnMAILausSELECTION()
'    Application.Dialogs(xlDialogSendMail).Show 'kann die Zeile gebraucht werden?? (Marco / 30.6.17)

    Dim objMail As Object 'Shell-Objekt für E-Mail

    Range(Selection.Address).Copy 'Selektion in die Windows-Zwischenablage speichern
    
    On Error Resume Next
        Set objMail = CreateObject("Shell.Application")
        objMail.ShellExecute "mailto:" & "@" _
            & "&subject=" & "Gesendet aus Excel Workbook: " & ActiveWorkbook.Name & " // Worksheet: " & ActiveSheet.Name _
                & " // Bereich: " & Selection.Address(False, False) & " // " & Now _
            & "&body="
'        objMail.ShellExecute "mailto:" & CStr(Chr(32) _
'            & "&subject=" & "Gesendet aus Excel Workbook: " & ActiveWorkbook.Name & " // Worksheet: " & ActiveSheet.Name _
'                & " // Bereich: " & Selection.Address(False, False) & " // " & Now _
'            & "&body="
        Application.Wait (Now + TimeValue("0:00:03"))
        Application.SendKeys ("^v")

    On Error GoTo 0
        
End Sub

'Spaltennummer in Windows-Zwischenablage kopieren
Public Sub fnSPALTEMERKEN()
    Dim col As Long
    
    Set xlef_objClip = Nothing
    Set xlef_objClip = New DataObject 'Neues Objekt für die Windows-Zwischenablage
    col = Range(ActiveCell.Address).Column
    
    xlef_objClip.SetText col
    xlef_objClip.PutInClipboard
    
End Sub

'Zeilennummer in Windows-Zwischenablage kopieren
Public Sub fnZEILEMERKEN()
    Dim row As Long
    
    Set xlef_objClip = Nothing
    Set xlef_objClip = New DataObject 'Neues Objekt für die Windows-Zwischenablage
    row = Range(ActiveCell.Address).row
    
    xlef_objClip.SetText row
    xlef_objClip.PutInClipboard
    
End Sub

'Vergleich von 2 Worksheets: Nur Werte der Zellen vergleichen, nicht Formate usw.
Public Sub fnWORKSHEETVERGLEICH()
    If ActiveWorkbook.Worksheets.Count > 1 Then 'Nur wenn mehr als 1 Worksheet
        
        Dim wks As Worksheet
        
        'ListBoxen leeren
        frmWorksheetVergleich.ListBoxWks1.Clear
        frmWorksheetVergleich.ListBoxWks2.Clear
        
        For Each wks In ActiveWorkbook.Worksheets
            frmWorksheetVergleich.ListBoxWks1.AddItem wks.Name 'Worksheet-Namen in ListBox1 eintragen
            frmWorksheetVergleich.ListBoxWks2.AddItem wks.Name 'Worksheet-Namen in ListBox2 eintragen
        Next
        
        frmWorksheetVergleich.Show
    Else
        MsgBox GetText("ELP_001"), , GetText("ELP_002")
    End If
End Sub

'Wenn Formelergebnis 0 ergibt, dann Zelle leer darstellen
Public Sub fnLEERWENNNULL()
    Dim strFormel As String  'Variable für zurechtgestutzte Formel
    
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnDo = False Then Exit Sub
    
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    For Each xlef_rngCell In xlef_Sel
        If xlef_rngCell.HasFormula And xlef_rngCell.Value = 0 Then 'nur wenn Zelle eine Formel mit dem Ergebnis 0 enthält
            strFormel = xlef_rngCell.Formula
            'Formel extrahieren ('=' und '+' am Anfang entfernen)
            strFormel = FormelExtrakt(strFormel)
            'Formel neu zusammenbauen und in Zelle schreiben
            xlef_rngCell.Formula = "=IF(" & strFormel & "=0," & Chr(34) & Chr(34) & "," & strFormel & ")"   ' Chr(34) erzeugt ein Anführungszeichen "
        End If
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Prüfen auf geschützte Leerzeichen (Zeichencode 160)
Sub fnGESCHLEERZEICHENpruefen()
    Dim i As Integer 'Variable für Schleife
    Dim zeichen As String 'Variable für ein einzelnes Zeichen innerhalb einer Zelle
    
    Set xlef_Sel = Selection 'Selektierten Bereich in Variable einlesen
    
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    For Each xlef_rngCell In xlef_Sel
        For i = 1 To Len(xlef_rngCell)
            zeichen = Mid(xlef_rngCell, i, 1)
            Select Case Asc(zeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                Case 160 'geschütztes Leerzeichen gefunden
                    If MsgBox(GetText("ELP_035") & vbCrLf & GetText("ELP_036") & " " & xlef_rngCell.Address(False, False), _
                        vbExclamation + vbOKCancel) = vbCancel Then Exit Sub 'Prüfung abbrechen wenn CANCEL geklickt wird
                    Exit For
            End Select
        Next
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Geschützte Leerzeichen (Zeichencode 160) durch normales (Zeichencode 32) ersetzen
Sub fnGESCHLEERZEICHENaustauschen()
    Dim i As Integer 'Variable für Schleife
    Dim zeichen As String 'Variable für ein einzelnes Zeichen innerhalb einer Zelle
    
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnDo = False Then Exit Sub
        
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    For Each xlef_rngCell In xlef_Sel
        For i = 1 To Len(xlef_rngCell)
            zeichen = Mid(xlef_rngCell, i, 1)
            Select Case Asc(zeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                Case 160 'Geschütztes Leerzeichen gefunden
                    xlef_rngCell.Value = Application.WorksheetFunction _
                        .Replace(xlef_rngCell.Value, i, 1, Chr(32)) 'Geschütztes Leerzeichen durch normales ersetzen
            End Select
        Next
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Prüfen auf Steuerzeichen (Zeichencodes 1-31, 127, 129, 141, 143, 144, 157)
Sub fnSTEUERZEICHENpruefen()
    Dim i As Integer 'Variable für Schleife
    Dim zeichen As String 'Variable für ein einzelnes Zeichen innerhalb einer Zelle
    
    Set xlef_Sel = Selection 'Selektierten Bereich in Variable einlesen
    
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    For Each xlef_rngCell In xlef_Sel
        For i = 1 To Len(xlef_rngCell)
            zeichen = Mid(xlef_rngCell, i, 1)
            Select Case Asc(zeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                Case 1 To 31, 127, 129, 141, 143, 144, 157 'Steuerzeichen gefunden
                    If MsgBox(GetText("ELP_037") & vbCrLf & GetText("ELP_036") & " " & xlef_rngCell.Address(False, False), _
                        vbExclamation + vbOKCancel) = vbCancel Then Exit Sub 'Prüfung abbrechen wenn CANCEL geklickt wird
                    Exit For
            End Select
        Next
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Steuerzeichen entfernen (Zeichencodes 1-31, 127, 129, 141, 143, 144, 157)
Sub fnSTEUERZEICHENentfernen()
    Dim i As Integer 'Variable für Schleife
    Dim zeichen As String 'Variable für ein einzelnes Zeichen innerhalb einer Zelle
    Dim fund As Boolean 'Kennzeichen für Fund
    
    xlef_art = "cell"
    Call Aktion           'Selektion einlesen und Zellinhalte merken für Wiederherstellung
    
    If xlef_blnDo = False Then Exit Sub
    
    If xlef_blnProgressBar Then 'Wenn Fortschrittsbalken angehakt ist
        xlef_dblBalkenAnteil = frmFortschritt.lblFortschrittBalken.Width / xlef_Sel.Count 'Schrittweite
        Call FortschrittON("Aktion ausführen")
    End If
    
    For Each xlef_rngCell In xlef_Sel
        fund = 0 'Markierung zurücksetzen
        For i = 1 To Len(xlef_rngCell)
            zeichen = Mid(xlef_rngCell, i, 1)
            Select Case Asc(zeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                Case 1 To 31, 127, 129, 141, 143, 144, 157 'Steuerzeichen gefunden
                    xlef_rngCell.Value = Application.WorksheetFunction _
                        .Replace(xlef_rngCell.Value, i, 1, Chr(9)) 'Steuerzeichen durch horizontalen Tab ersetzen
                fund = True 'Markieren, dass die Zelle bereinigt werden muss
            End Select
        Next
        If fund = True Then
            xlef_rngCell.Value = Application.WorksheetFunction.Clean(xlef_rngCell.Value) 'Horizontale Tabs entfernen
        End If
        If xlef_blnProgressBar Then 'Fortschrittsbalken aktualisieren
            xlef_dblBalkenAktuell = xlef_dblBalkenAktuell + xlef_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call BalkenAkt(xlef_dblBalkenAktuell)
        End If
    Next
    
    If xlef_blnProgressBar Then Unload frmFortschritt
End Sub

'Bedingte Zeilenlöschung in gesamter Tabellle  - bisher kein Rückganig möglich
 Public Sub fnBEDZEILCLEAR_A()
 Dim lur As Integer             'ermittelt last used row in der Tabelle in Spalte 1 --> könnte problematisch sein, wenn nicht in jeder Spalte gleichviele EInträge vorhanden sind
 Dim i As Integer               ' Zählervariable für Schleife
 Dim intColInd As Integer       'Spaltenindex in der gesucht werden soll
 Dim strSearchContent As String 'Suchwert
 Dim X As Variant
 
 
    X = InputBox(prompt:="Bitte definieren Sie den Spaltenindex des Suchwerts.", Title:="Spaltenindex des Suchwerts")
    If X = vbCancel Then
        Exit Sub
    Else
        intColInd = Int(X)
    End If
    
    X = InputBox(prompt:="Bitte definieren Sie das Suchwort.", Title:="Suchwert eingeben")
    If X = vbCancel Then
    Exit Sub
    Else
        strSearchContent = X
    End If
    

    lur = Cells(Rows.Count, 1).End(xlUp).Rows.row 'Ermittlung der letzten Zeile in Spalte A
    
'** Durchlauf aller Zeilen
    For i = lur To 2 Step -1 'Zählung rückwärts bis Zeile 2
     
     If Cells(i, intColInd).Value = strSearchContent Then 'Abfragen, ob in der vorher definierten Spalte der vorher definierter Suchbegriff steht
     Rows(i).Delete Shift:=xlUp
     End If
    Next i

End Sub
'Bedingte Zeilen in Selektion löschen - bisher kein Rückganig möglich
Public Sub fnBEDZEILCLEAR_S()

 Dim i As Integer               ' Zählervariable für Schleife
 Dim intColIndSel As Integer       'Spaltenindex in der gesucht werden soll
 Dim strSearchContent As String 'Suchwert
 Dim X As Variant
 
 Call Aktion
 
 If xlef_blnDo = False Then Exit Sub
 
    X = InputBox(prompt:="Bitte definieren Sie den Spaltenindex  des Suchwerts innerhalb der Selektion.", Title:="Spaltenindex des Suchwerts")
    If X = vbCancel Then
        Exit Sub
    Else
        intColIndSel = Int(X) + Selection.Column - 1 'um Spaltenverschiebung nach rechts festzustellen
    End If
    
    X = InputBox(prompt:="Bitte definieren Sie das Suchwort.", Title:="Suchwert eingeben")
    If X = vbCancel Then
    Exit Sub
    Else
        strSearchContent = X
    End If
    

    
'** Durchlauf aller Zeilen
    For i = Selection.row + Selection.Rows.Count To Selection.row Step -1 'Zählung rückwärts bis Zeile 2
     
     If Cells(i, intColIndSel).Value = strSearchContent Then 'Abfragen, ob in der vorher definierten Spalte der vorher definierter Suchbegriff steht
     Rows(i).Delete Shift:=xlUp
     End If
    Next i

End Sub
