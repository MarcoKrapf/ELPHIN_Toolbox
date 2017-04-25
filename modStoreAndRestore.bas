Attribute VB_Name = "modStoreAndRestore"
Option Explicit
Option Private Module

'Beschreibung..............

Public Sub SelectionSave()
    Dim i As Long
    
    Select Case xlef_art
    
        Case "cell" 'Inhalte der selektierten Zellen merken
            'Array neu dimensionieren (zweidimensional)
            ReDim xlef_arrOrg(xlef_Sel.Count - 1, 1)
            'Array füllen
            For Each xlef_rngCell In xlef_Sel 'Alle Zellen der Selektion durchlaufen
                xlef_arrOrg(i, 0) = xlef_rngCell.Address 'Adresse der Zelle
                If xlef_rngCell.HasFormula Then
                    xlef_arrOrg(i, 1) = xlef_rngCell.Formula 'Formel der Zelle
                Else
                    xlef_arrOrg(i, 1) = xlef_rngCell.Value 'Wert der Zelle
                End If
                i = i + 1
            Next
            
        Case "row" 'Zeilennummern merken
            Set xlef_coll = Nothing 'Collection leermachen
            Set xlef_coll = New Collection 'Neues Collection-Objekt generieren
            
            If xlef_Sel.Count > 1 Then 'Wenn mehr als eine Zelle selektiert ist, dann nur die Zeilen innerhalb dieses Bereichs durchsuchen (nur Area 1)
                For i = xlef_Sel.row To xlef_Sel.row + xlef_Sel.Rows.Count - 1 'oberste bis unterste Zeile der Selection
                    If ActiveSheet.Cells(i, Range("1:1").Columns.Count).End(xlToLeft).Column = 1 And ActiveSheet.Cells(i, 1).Value = "" Then
                        xlef_coll.Add i 'Nummer der leeren Zeile an Collection anfügen
                    End If
                Next i
            Else 'Wenn kein Bereich, also nur 1 Zelle selektiert ist, dann das ganze benutzte Worksheet durchsuchen
                For i = 1 To ActiveSheet.UsedRange.Rows.Count 'alle Zeilen des benutzten Bereichs durchlaufen
                    If ActiveSheet.Cells(i, Range("1:1").Columns.Count).End(xlToLeft).Column = 1 And ActiveSheet.Cells(i, 1).Value = "" Then
                        xlef_coll.Add i 'Nummer der leeren Zeile an Collection anfügen
                    End If
                Next i
            End If
            
        Case "col" 'Spaltennummern merken
            Set xlef_coll = Nothing 'Collection leermachen
            Set xlef_coll = New Collection 'Neues Collection-Objekt generieren
            
            If xlef_Sel.Count > 1 Then 'Wenn mehr als eine Spalte selektiert ist, dann nur die Spalten innerhalb dieses Bereichs durchsuchen (nur Area 1)
                For i = xlef_Sel.Column To xlef_Sel.Column + xlef_Sel.Columns.Count - 1 'linkeste bis rechteste Spalte der Selection
                    If ActiveSheet.Cells(Range("A:A").Rows.Count, i).End(xlUp).row = 1 And ActiveSheet.Cells(1, i).Value = "" Then
                        xlef_coll.Add i 'Nummer der leeren Zeile an Collection anfügen
                    End If
                Next i
            Else 'Wenn kein Bereich, also nur 1 Spalte selektiert ist, dann das ganze benutzte Worksheet durchsuchen
                For i = 1 To ActiveSheet.UsedRange.Columns.Count 'alle Spalten des benutzten Bereichs durchlaufen
                    If ActiveSheet.Cells(Range("A:A").Rows.Count, i).End(xlUp).row = 1 And ActiveSheet.Cells(1, i).Value = "" Then
                        xlef_coll.Add i 'Nummer der leeren Spalte an Collection anfügen
                    End If
                Next i
            End If
            
    End Select
End Sub


'Letzten Zustand wiederherstellen
Public Sub SelectionRestore()
    Dim i As Long
    
    If xlef_wksTarget Is Nothing Then Set xlef_wksTarget = ActiveWorkbook.ActiveSheet 'Wenn noch kein Worksheet in der Variablen steht
    xlef_wksTarget.Activate 'Worksheet anzeigen, auf dem wiederhergestellt wird
    
    Application.ScreenUpdating = False 'Bildschirmaktualisierung ausschalten (Performance)
    
    Select Case xlef_art
        Case "cell" 'Inhalte der Zellen wiederherstellen
            For i = 0 To UBound(xlef_arrOrg)
                xlef_wksTarget.Range(xlef_arrOrg(i, 0)).Value = xlef_arrOrg(i, 1)
            Next
        Case "row" 'Zeilen wiederherstellen
            For i = 1 To xlef_coll.Count
                xlef_wksTarget.Rows(xlef_coll(i)).Insert
            Next i
        Case "col" 'Spalten wiederherstellen
            For i = 1 To xlef_coll.Count
                ActiveSheet.Columns(xlef_coll(i)).Insert
            Next i
        Case "cellC" 'Farben der Zellen wiederherstellen
            For i = 1 To xlef_coll.Count
                With xlef_wksTarget.Range(xlef_coll(i)(0))
                    .Interior.Color = xlef_coll(i)(1)
                    .Font.Color = xlef_coll(i)(2)
                End With
            Next i
    End Select
    
    Application.ScreenUpdating = True 'Bildschirmaktualisierung wieder einschalten
End Sub
