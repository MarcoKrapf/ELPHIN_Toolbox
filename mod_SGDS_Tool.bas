Attribute VB_Name = "mod_SGDS_Tool"
Option Explicit
Option Private Module
            
' Variablen deklarieren
Public g_sgds_intAusgabeAbXmal As Integer 'Ab dem wievielten Vorkommen ein Datensatz ausgegeben wird
Public g_sgds_varFilesSelected As Variant 'Ausgewählte Dateien, die eingelesen werden
Dim sgds_collFilesSelected As Collection 'Ausgewählte Dateien (kompletter Pfad)
Dim sgds_strPath As String 'Pfadname des ausgewählten Ordners
Dim sgds_collHeaders As Collection 'Collection mit den Spaltenüberschriften
Dim sgds_collColumnsSelected As Collection 'Collection mit den ausgewählten Spalten zum Vergleich
Dim sgds_arrayContent() As Variant 'Array mit den Inhalten aus den eingelesenen Excel-Dateien
Dim sgds_lngLinesFilled As Long 'Anzahl der im Array "sgds_arrayContent" gespeicherten Datensätze
Dim sgds_arrayOutputCount() As Variant 'Array mit 2 Spalten: Vergleichsstring und Anzahl der Treffer
Dim sgds_lngLinesOutput As Long 'Anzahl der im Array "sgds_arrayOutputCount" gespeicherten Datensätze
Dim sgds_collColumnsOutput As Collection 'Collection mit den ausgewählten Spalten für die Ausgabe
Dim sgds_arrayOutput() As Variant 'Array mit den Daten für die Ausgabe
Dim sgds_objLabel As MSForms.Label 'Label für die dynamische Ausgabe der Datensätze im UserForm

Public Sub openGUI()
    Cells(1, 1).Interior.Color = Cells(1, 1).Interior.Color 'Absturz verhindern
    frm_SGDS_GUI.Show (vbModeless)
End Sub

'Select the files in a folder
Public Sub SelectFiles()

    g_sgds_varFilesSelected = Application.GetOpenFilename(FileFilter:=GetText("SG_034"), _
                                        FilterIndex:=2, _
                                        Title:=GetText("SG_033"), _
                                        MultiSelect:=True)
    On Error GoTo ERRORHANDLER 'If nofile is selected
    
    'Button
    sgds_strPath = Left$(g_sgds_varFilesSelected(1), InStrRev(g_sgds_varFilesSelected(1), Application.PathSeparator)) 'extract the path name
    With frm_SGDS_GUI.CommandButtonSelectFolder
        .Caption = GetText("SG_035") & ": " & sgds_strPath & vbNewLine _
                    & GetText("SG_036") & ": " & UBound(g_sgds_varFilesSelected)
    End With
    
    'Clear the ListBox and the ComboBox with the files
    With frm_SGDS_GUI
        .ListBoxFiles.Clear
        .ComboBoxFileForColumns.Clear
    End With
    
    'Change button colours
    If UBound(g_sgds_varFilesSelected) > 0 Then
        frm_SGDS_GUI.CommandButtonSelectFolder.BackColor = &H8000000F
        frm_SGDS_GUI.ListBoxFiles.BackColor = &H80FFFF
        frm_SGDS_GUI.ComboBoxFileForColumns.BackColor = &H80FFFF
    Else
        frm_SGDS_GUI.CommandButtonSelectFolder.BackColor = &H80FFFF
        frm_SGDS_GUI.ListBoxFiles.BackColor = &H8000000F
        frm_SGDS_GUI.ComboBoxFileForColumns.BackColor = &H8000000F
    End If
    
    Exit Sub
ERRORHANDLER:
    
End Sub

Public Sub ResetOutputPage()
    'Change button states and colours
    With frm_SGDS_GUI
        .ListBoxPreview.Clear
        .CommandButtonStart.Enabled = False
        .CommandButtonStart.BackColor = &H8000000F
        .CommandButtonPreview.Enabled = False
        .CommandButtonPreview.BackColor = &H8000000F
    End With
End Sub

Public Sub PopulateBoxesWithFiles()

    Dim file As Variant
    Dim i As Long
    
    On Error GoTo ERRORHANDLER
    
    'Reset the Collection
    Set sgds_collFilesSelected = New Collection
    
    'Populate the Collection
    For i = 1 To UBound(g_sgds_varFilesSelected)
        sgds_collFilesSelected.Add g_sgds_varFilesSelected(i) 'Write selected files with full path
    Next i

    'Populate ListBox and ComboBox
    For Each file In sgds_collFilesSelected
        With frm_SGDS_GUI
            .ListBoxFiles.AddItem Mid$(file, InStrRev(file, Application.PathSeparator) + 1)
            .ComboBoxFileForColumns.AddItem Mid$(file, InStrRev(file, Application.PathSeparator) + 1)
        End With
    Next file
    
    Exit Sub
ERRORHANDLER:
        
End Sub

Public Sub GetColumnHeaders()

    Dim fullpath As String 'path and file name
    Dim row As Long 'row with the column headers
    Dim i As Long
    
    fullpath = sgds_strPath & frm_SGDS_GUI.ComboBoxFileForColumns.Value 'full path of the selected file
    row = frm_SGDS_GUI.SpinButtonRow.Value 'row with the headers dependent on the spin button
    
    With frm_SGDS_GUI
        .ListBoxColumnHeaders.Clear
        .ListBoxColumnsOutput.Clear
    End With
    
    frm_SGDS_GUI.ListBoxColumnsOutput.AddItem (GetText("SG_037"))
    
    With frm_SGDS_GUI
        If .ComboBoxFileForColumns.Value <> "" Then
            Call ReadHeaders(fullpath, row)
            For i = 1 To sgds_collHeaders.Count
                .ListBoxColumnHeaders.AddItem sgds_collHeaders(i)
                .ListBoxColumnsOutput.AddItem sgds_collHeaders(i)
            Next
        End If
    End With
    
End Sub

Public Sub ReadHeaders(ByVal file As String, row As Long)
    Dim i As Long, j As Long
    Dim wks As Worksheet
    Dim fileName As String
    Dim cols As Integer

    Set sgds_collHeaders = New Collection 'reset Collection
    
    Application.EnableEvents = False
    Workbooks.Open file
    
    Set wks = ActiveWorkbook.Worksheets(1) 'take the first worksheet
    
    fileName = Mid$(file, InStrRev(file, Application.PathSeparator) + 1)
    cols = wks.Cells(row, Columns.Count).End(xlToLeft).Column 'count the columns
    
    'Populate the Collection
    For i = 1 To cols
        sgds_collHeaders.Add wks.Cells(row, i).Value
    Next
    
    ActiveWorkbook.Close savechanges:=True
    Application.EnableEvents = True
    
End Sub

Public Sub SelectAllColumns(bln As Boolean)
    
    Dim i As Long
    
    With frm_SGDS_GUI.ListBoxColumnHeaders
        For i = 0 To .ListCount - 1
            .Selected(i) = bln 'select if true, deselect if false
        Next
    End With
    
End Sub

Public Sub SelectAllColumnsOutput(bln As Boolean)
    
    Dim i As Long
    
    With frm_SGDS_GUI.ListBoxColumnsOutput
        For i = 0 To .ListCount - 1
            .Selected(i) = bln 'select if true, deselect if false
        Next
    End With
    
End Sub

Public Sub CreateSearchArrays()
    
    With frmFortschritt
        .lblFortschrittBalken.Width = .lblFortschrittbalkenLeer.Width / 5 * 1
        .ForeColor = vbWhite
        .Font.Bold = True
        .Caption = GetText("SG_080")
    End With
    DoEvents
    
    Dim i As Long
    
    Set sgds_collColumnsSelected = New Collection 'reset the collection for the selected columns to compare
    Set sgds_collColumnsOutput = New Collection 'reset the collection for the selected columns for output

    Erase sgds_arrayContent
    ReDim sgds_arrayContent(0 To sgds_collHeaders.Count + 1, 0 To 0)
    
    'Loop through all columns of the selected file
    With frm_SGDS_GUI.ListBoxColumnHeaders
        For i = 0 To .ListCount - 1
            'Write all column headers in the Array
            sgds_arrayContent(i + 1, 0) = .List(i)
            'Write the selected columns for the comparison in a Collection
            If .Selected(i) = True Then sgds_collColumnsSelected.Add .List(i)
        Next
    End With

    'Get the columns for output
    With frm_SGDS_GUI.ListBoxColumnsOutput
        For i = 0 To .ListCount - 1
            'Write the selected columns for the comparison in a Collection
            If .Selected(i) = True Then sgds_collColumnsOutput.Add .List(i)
            
        Next
    End With
    
    sgds_arrayContent(0, 0) = GetText("SG_037")
    sgds_arrayContent(sgds_collHeaders.Count + 1, 0) = GetText("SG_038")
    
    sgds_lngLinesFilled = 0 'Reset the number of filled lines in the Array (2nd dimension)
    
    'Loop through all files
    For i = 1 To sgds_collFilesSelected.Count
        Call AccessSpreadsheet(sgds_collFilesSelected(i))
    Next
    
#If Debugging Then
Debug.Print "Aktuelle Anzahl Datensätze: " & sgds_lngLinesFilled
#End If
    
End Sub

Public Sub AccessSpreadsheet(ByVal file As String)
    Dim fileName As String
    Dim fileRows As Long
    Dim fileHeaderRow As Long
    Dim currentColumn As Integer
    Dim i As Long, j As Long, k As Integer, m As Integer
    
    Application.EnableEvents = False
    Workbooks.Open file
    
    fileName = Mid$(file, InStrRev(file, Application.PathSeparator) + 1)
    fileHeaderRow = FindHeaderRow(sgds_collHeaders(1))
    fileRows = Cells(Rows.Count, 1).End(xlUp).row - fileHeaderRow 'number of rows in this file
  
#If Debugging Then
Debug.Print "Aktuelle Anzahl Datensätze: " & sgds_lngLinesFilled
Debug.Print "   ausgewählt: " & frm_SGDS_GUI.ComboBoxFileForColumns.Value & vbNewLine & _
                "   aktuell   : " & fileName
#End If
    
    ReDim Preserve sgds_arrayContent(0 To sgds_collHeaders.Count + 1, sgds_lngLinesFilled + fileRows) 'extend Array

    'Loop through the rows in the selected file
    For i = fileHeaderRow To fileHeaderRow + fileRows - 1
        'Loop through all columns in the selected file
        For j = 1 To frm_SGDS_GUI.ListBoxColumnHeaders.ListCount
                'Find the corresponding column
                For k = 1 To ActiveSheet.Columns.Count
                    If Cells(fileHeaderRow, k).Value = sgds_arrayContent(j, 0) Then
                        sgds_arrayContent(j, sgds_lngLinesFilled + 1) = Cells(i + 1, k).Value
                        For m = 1 To sgds_collColumnsSelected.Count
                            If sgds_collColumnsSelected(m) = sgds_arrayContent(j, 0) Then
                                sgds_arrayContent(sgds_collHeaders.Count + 1, sgds_lngLinesFilled + 1) _
                                            = sgds_arrayContent(sgds_collHeaders.Count + 1, sgds_lngLinesFilled + 1) & sgds_arrayContent(j, sgds_lngLinesFilled + 1)
                            End If
                        Next m
                        Exit For
                    End If
                Next k
        Next j
        sgds_arrayContent(0, sgds_lngLinesFilled + 1) = fileName 'file name
        sgds_lngLinesFilled = sgds_lngLinesFilled + 1
    Next i

    ActiveWorkbook.Close savechanges:=True
    Application.EnableEvents = True
    
End Sub

'Return the row with the headers
Public Function FindHeaderRow(str As String) As Long
    Dim i As Long, j As Integer
    'Loop through all rows
    For i = 1 To Rows.Count
        'Loop through all columns
        For j = 1 To Columns.Count
            If Cells(i, j).Value = str Then
                FindHeaderRow = i 'return value
                Exit Function 'exit function if the row is found
            End If
        Next j
    Next i
End Function

'Return the column number
Public Function FindColumn(str As String) As Integer
    Dim i As Integer
    For i = 1 To Cells(frm_SGDS_GUI.SpinButtonRow.Value, Columns.Count).End(xlToLeft).Column
        If Cells(frm_SGDS_GUI.SpinButtonRow.Value, i).Value = str Then Exit For 'exit loop if the column is found
    Next i
    FindColumn = i 'return the column number
End Function

Public Sub CreateOutputCountArray()
    Dim i As Long, j As Long
    Dim found As Boolean

    With frmFortschritt
        .lblFortschrittBalken.Width = .lblFortschrittbalkenLeer.Width / 5 * 2
        .ForeColor = vbWhite
        .Font.Bold = True
        .Caption = GetText("SG_081")
    End With
    DoEvents
    
    'Fill the Array for the output
    sgds_lngLinesOutput = 0
    Erase sgds_arrayOutputCount
    ReDim sgds_arrayOutputCount(0 To 1, 0 To sgds_lngLinesOutput)

    'Loop through the Array with all content
    For i = 1 To UBound(sgds_arrayContent, 2)
        found = False 'reset the check variable
        'Loop through the Array for the output
        For j = 0 To UBound(sgds_arrayOutputCount, 2)
            If sgds_arrayContent(sgds_collHeaders.Count + 1, i) = sgds_arrayOutputCount(0, j) Then 'check whether the entry already exists
                sgds_arrayOutputCount(1, j) = sgds_arrayOutputCount(1, j) + 1 'add one more entry to the counter
                found = True
                Exit For
            End If
        Next j
        If found = False Then 'if no entry exists
            sgds_lngLinesOutput = sgds_lngLinesOutput + 1
            ReDim Preserve sgds_arrayOutputCount(0 To 1, 0 To sgds_lngLinesOutput) 'enlarge the Array
            sgds_arrayOutputCount(0, sgds_lngLinesOutput) = sgds_arrayContent(sgds_collHeaders.Count + 1, i)  'add a new entry
            sgds_arrayOutputCount(1, sgds_lngLinesOutput) = 1  'first entry
        End If
    Next i
    
End Sub

Public Sub Output()

    With frmFortschritt
        .lblFortschrittBalken.Width = .lblFortschrittbalkenLeer.Width / 5 * 4
        .ForeColor = vbWhite
        .Font.Bold = True
        .Caption = GetText("SG_083")
    End With
    DoEvents
    
    Select Case frm_SGDS_GUI.ComboBoxAusgabe.ListIndex
        Case 0
            Call OutputScreen
        Case 1
            Call OutputExcel
        Case 2
            Call OutputTXT("CSV")
        Case 3
            Call OutputTXT("TXT")
        Case 4
            Call OutputWORD
    End Select
End Sub

Private Sub OutputScreen()
    Dim i As Long, j As Long
    
    Unload frm_SGDS_OutScreen
    
    Application.ScreenUpdating = False
    
    With frm_SGDS_OutScreen
        'Spaltenüberschriften
        For i = 1 To sgds_collColumnsOutput.Count
            Set sgds_objLabel = .Controls.Add("Forms.Label.1", , True)
            With sgds_objLabel
                .Caption = sgds_collColumnsOutput(i)
                .Font.Bold = True
                .Top = 10
                .Left = 10 + ((i - 1) * 80)
            End With
        Next
        .Height = 60
        
        'Datensätze
        For j = 1 To UBound(sgds_arrayOutput, 1)
            For i = 1 To UBound(sgds_arrayOutput, 2)
                Set sgds_objLabel = .Controls.Add("Forms.Label.1", , True)
                With sgds_objLabel
                    .Caption = sgds_arrayOutput(j, i)
                    .Top = 10 + j * 12
                    .Height = 12
                    .Left = 10 + ((i - 1) * 80)
                End With
            Next
            .Height = .Height + sgds_objLabel.Height
        Next
        
        'Popup anpassen
        .Caption = GetText("SG_039") & ": " & SearchString & " (" & UBound(sgds_arrayOutput) & " " & GetText("SG_040") & " " & GetText("SG_026") & " " & sgds_collColumnsOutput.Count & " " & GetText("SG_048") & ")"
        .Width = 30 + sgds_collColumnsOutput.Count * 80
        If .Height > 440 Then
            .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
            .ScrollHeight = .Height 'height of the vertical scrolling
            .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
            .Height = 440
        End If
        .Show (vbModeless)
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub OutputExcel()
    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim i As Long, j As Long

    'Neues Workbook öffnen
    Set wkb = Workbooks.Add
    Set wks = wkb.Worksheets(1)
    wks.Activate
    
    Application.ScreenUpdating = False

    'Überschriften
    For i = 1 To sgds_collColumnsOutput.Count
        wks.Cells(1, i).Value = sgds_collColumnsOutput(i)
    Next

    'Datensätze
    For j = 1 To UBound(sgds_arrayOutput, 1)
        For i = 1 To UBound(sgds_arrayOutput, 2)
            wks.Cells(1 + j, i).Value = sgds_arrayOutput(j, i)
        Next
    Next
    
    Application.ScreenUpdating = True
 
End Sub

Private Sub OutputTXT(filetyp As String)
    Dim varSave As Variant 'Infos aus Dialogbox zum Speichern
    Dim intFileNr As Integer 'Nächste freie Nummer beim Export
    Dim strLineOutput As String 'Zusammengesetzter String für eine Zeile
    Dim strFilter As String, strTitel As String, strSep As String 'Variablen für TXT bzw. CSV
    Dim i As Long, j As Long
    
    Select Case filetyp
        Case "TXT":
            strFilter = GetText("SG_041")
            strTitel = GetText("SG_042")
            strSep = " "
        Case "CSV":
            strFilter = GetText("SG_043")
            strTitel = GetText("SG_044")
            strSep = ";"
    End Select
    
    intFileNr = FreeFile 'Nächste freie Nummer zuweisen
    
    'Speichern unter - Dialog
    varSave = Application.GetSaveAsFilename( _
        InitialFileName:=GetText("SG_045"), _
        FileFilter:=strFilter, _
        Title:=strTitel)
        
    'Wenn Dialog nicht abgebrochen wurde, dann exportieren
            
    If varSave <> False Then
        Open varSave For Output As #intFileNr 'Ausgangskanal öffnen
        
            'Titelinformationen
            Print #intFileNr, GetText("SG_039") & ": " & SearchString & " (" & UBound(sgds_arrayOutput) & " " & GetText("SG_040") & ")"
            Print #intFileNr,
            
            'Spaltenüberschriften
            strLineOutput = ""
            For i = 1 To sgds_collColumnsOutput.Count
                strLineOutput = strLineOutput & sgds_collColumnsOutput(i) & strSep
            Next
            Print #intFileNr, strLineOutput
            
            'Datensätze
            For j = 1 To UBound(sgds_arrayOutput, 1)
                strLineOutput = ""
                For i = 1 To UBound(sgds_arrayOutput, 2)
                    strLineOutput = strLineOutput & sgds_arrayOutput(j, i) & strSep
                Next
                Print #intFileNr, strLineOutput
            Next

        Close #intFileNr 'Ausgangskanal schließen
    End If
        
End Sub

Private Sub OutputWORD()
    MsgBox "Ausgabe in Word"
End Sub

Private Function SearchString() As String
    Dim i As Integer
    
    For i = 1 To sgds_collColumnsSelected.Count
        SearchString = SearchString & sgds_collColumnsSelected(i) & " + "
    Next
    SearchString = Left(SearchString, Len(SearchString) - 3)
End Function

Public Sub CreateOutputArray()
    Dim i As Long, j As Long, k As Integer
    Dim outRow As Integer, outCol As Integer

    With frmFortschritt
        .lblFortschrittBalken.Width = .lblFortschrittbalkenLeer.Width / 5 * 3
        .ForeColor = vbWhite
        .Font.Bold = True
        .Caption = GetText("SG_082")
    End With
    DoEvents
    
    'Create an array for the output data
        'Loop through the array for the output
            For i = 1 To UBound(sgds_arrayOutputCount, 2)
                'Check for output
                If sgds_arrayOutputCount(1, i) >= g_sgds_intAusgabeAbXmal Then
                    j = j + sgds_arrayOutputCount(1, i) 'number of rows needed for the array
                End If
            Next
            
    Erase sgds_arrayOutput
    ReDim sgds_arrayOutput(0 To j, 1 To sgds_collColumnsOutput.Count)
    
    'Start value for sgds_arrayOutput
    outRow = 1
    
    'Populate the output array
        'Loop through the array for the output
        For i = 1 To UBound(sgds_arrayOutputCount, 2)
            'Check for output
            If sgds_arrayOutputCount(1, i) >= g_sgds_intAusgabeAbXmal Then
#If Debugging Then
Debug.Print "output sgds_arrayOutputCount: " & sgds_arrayOutputCount(0, i)
#End If
                'Loop through the array with all content
                For j = 1 To UBound(sgds_arrayContent, 2)
                    If sgds_arrayContent(sgds_collHeaders.Count + 1, j) = sgds_arrayOutputCount(0, i) Then
                        outCol = 1 'reset column number
#If Debugging Then
Debug.Print "output sgds_arrayContent(" & sgds_collHeaders.Count + 1 & "," & j & "):  " & sgds_arrayContent(sgds_collHeaders.Count + 1, j)
#End If
                        For k = 0 To UBound(sgds_arrayContent, 1) - 1
                            If ColInOutput(sgds_arrayContent(k, 0)) = True Then
#If Debugging Then
Debug.Print "output sgds_arrayContent col: " & sgds_arrayContent(k, j)
#End If
                                sgds_arrayOutput(outRow, outCol) = sgds_arrayContent(k, j)
                                outCol = outCol + 1
                            End If
                        Next
                        outRow = outRow + 1 'next row
                    End If
                Next
            End If
        Next
        
    'Anzahl der Datensätze zur Ausgabe aktualisieren
    frm_SGDS_GUI.FrameOutput.Caption = GetText("SG_032") & ": " & UBound(sgds_arrayOutput) & " " & GetText("SG_040") & " " & GetText("SG_026") & " " & sgds_collColumnsOutput.Count & " " & GetText("SG_048")

End Sub

Private Function ColInOutput(ByVal col As Variant) As Boolean
    Dim i As Long
    For i = 1 To sgds_collColumnsOutput.Count
        If sgds_collColumnsOutput(i) = col Then
            ColInOutput = True
            Exit Function
        End If
    Next i
    ColInOutput = False
End Function

Public Sub OutputPreview()
    Dim i As Integer, j As Integer

    With frmFortschritt
        .lblFortschrittBalken.Width = .lblFortschrittbalkenLeer.Width / 5 * 4
        .ForeColor = vbWhite
        .Font.Bold = True
        .Caption = GetText("SG_084")
    End With
    DoEvents
    
    With frm_SGDS_GUI.ListBoxPreview
        .Clear
        .ColumnCount = sgds_collColumnsOutput.Count
        
        'Überschriften
        .AddItem
        For i = 1 To sgds_collColumnsOutput.Count
            .List(0, i - 1) = UCase(sgds_collColumnsOutput(i))
            If i > 9 Then Exit For 'max. 10 Spalten
        Next
        
        'Max. 20 Datensätze
        For i = 1 To UBound(sgds_arrayOutput)
            .AddItem
            For j = 1 To sgds_collColumnsOutput.Count
                .List(i, j - 1) = sgds_arrayOutput(i, j)
                'Max. 10 Spalten
                If j > 9 Then Exit For
            Next
            If i > 19 Then Exit For
        Next
    End With
    
End Sub
