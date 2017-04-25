VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWorksheetVergleich 
   Caption         =   "[TITEL]"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5400
   OleObjectBlob   =   "frmWorksheetVergleich.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmWorksheetVergleich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnStart_Click()
        
    Dim wks1 As Integer, wks2 As Integer
    Dim row As Long, col As Long
    Dim rowMax As Long, colMax As Long
    Dim colSave(2) As Variant
    
    If ListBoxWks1.ListIndex > -1 And ListBoxWks2.ListIndex > -1 Then 'nur starten wenn beide Seiten ausgewählt
        
        xlef_art = "cellC"
    
        Set xlef_coll = Nothing 'Collection leermachen
        Set xlef_coll = New Collection 'Neues Collection-Objekt generieren
        
        wks1 = ListBoxWks1.ListIndex + 1
        wks2 = ListBoxWks2.ListIndex + 1
        rowMax = WorksheetFunction.Max(Sheets(wks1).UsedRange.Rows.row + Sheets(wks1).UsedRange.Rows.Count - 1, Sheets(wks2).UsedRange.Rows.row + Sheets(wks2).UsedRange.Rows.Count - 1)
        colMax = WorksheetFunction.Max(Sheets(wks1).UsedRange.Columns.Column + Sheets(wks1).UsedRange.Columns.Count - 1, Sheets(wks2).UsedRange.Columns.Column + Sheets(wks2).UsedRange.Columns.Count - 1)

        Set xlef_wksTarget = ActiveWorkbook.Worksheets(ListBoxWks2.ListIndex + 1)
        
        xlef_wksTarget.Activate 'zweites Worksheet anzeigen
        Application.ScreenUpdating = False 'Bildschirmaktualisierung ausschalten (Performance)
        
        On Error Resume Next 'Fehler tritt auf, wenn in einer Zelle ein Fehler ist
        
        'Nur den maximal benutzen Bereich durchlaufen (Performance)
        For row = 1 To rowMax 'Zeilen durchlaufen
            For col = 1 To colMax 'Spalten durchlaufen
                If Worksheets(wks1).Cells(row, col).Value <> Worksheets(wks2).Cells(row, col).Value Then 'Zellen vergleichen
                    'Originalfarben merken
                    colSave(0) = Worksheets(wks2).Cells(row, col).Address
                    If Worksheets(wks2).Cells(row, col).Interior.Color = 16777215 Then
                        colSave(1) = xlNone
                    Else
                        colSave(1) = Worksheets(wks2).Cells(row, col).Interior.Color
                    End If
                    colSave(2) = Worksheets(wks2).Cells(row, col).Font.Color
                    xlef_coll.Add colSave
                    
                    'Zellen neu färben
                    With Worksheets(wks2).Cells(row, col)
                        .Interior.Color = 16711935
                        .Font.Color = 65535
                    End With
                End If
            Next col
        Next row
        
        Application.ScreenUpdating = True 'Bildschirmaktualisierung wieder einschalten
        frmWorksheetVergleich.Hide 'Popup ausblenden
    End If
End Sub


Private Sub ListBoxWks1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If ListBoxWks1.ListIndex = ListBoxWks2.ListIndex Then
        ListBoxWks1.ListIndex = -1
    End If
End Sub

Private Sub ListBoxWks2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If ListBoxWks2.ListIndex = ListBoxWks1.ListIndex Then
        ListBoxWks2.ListIndex = -1
    End If
End Sub

Private Sub UserForm_Initialize()
    frmWorksheetVergleich.Caption = "Worksheets vergleichen"
    frmWorksheetVergleich.btnStart.Caption = "Vergleich starten"
End Sub
