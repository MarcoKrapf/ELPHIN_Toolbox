VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_SGDS_GUI 
   Caption         =   "[Titel]"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8460
   OleObjectBlob   =   "frm_SGDS_GUI.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_SGDS_GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = GetText("SG_002")
        .lblTitelVersion = GetText("SG_047")
        'MultiPage
            .MultiPage1.Value = 0 'Activate the first page
            .MultiPage1.Pages(0).Caption = GetText("SG_003")
            .MultiPage1.Pages(1).Caption = GetText("SG_004")
            .MultiPage1.Pages(2).Caption = GetText("SG_005")
        'Page "Datenquelle"
            'Section "Auswahl der Dateien"
                .FrameSourceFiles.Caption = GetText("SG_006")
                With .CommandButtonSelectFolder
                    .Caption = GetText("SG_007")
                    .ControlTipText = GetText("SG_049")
                End With
            'Section "Dateien"
                .ListBoxFiles.ControlTipText = GetText("SG_050")
            'Section "Spalten"
                .FrameSourceColumns.Caption = GetText("SG_009")
                .FrameSourceColumnsFile.Caption = GetText("SG_010")
                .FrameSourceColumnsCol.Caption = GetText("SG_011")
                With .ComboBoxFileForColumns
                    .Style = fmStyleDropDownList 'allow only values from the item list, no free entries
                    .ControlTipText = GetText("SG_050")
                End With
                .lblRowHeaders.TextAlign = fmTextAlignCenter
                .lblRowHeaders.ControlTipText = GetText("SG_051")
                .CommandButtonRowHeader1.Caption = GetText("SG_012")
                .CommandButtonRowHeader1.ControlTipText = GetText("SG_052")
                With .SpinButtonRow
                    .Orientation = fmOrientationVertical
                    .Value = 1
                    .Min = 1
                    .Max = ThisWorkbook.Worksheets(1).Rows.Count - 1
                End With
                .CommandButtonSelectAllColumns.Caption = GetText("SG_013")
                .CommandButtonSelectNoColumns.Caption = GetText("SG_014")

                With .ListBoxColumnHeaders
                    .MultiSelect = fmMultiSelectExtended 'multiple columns can be selected
                    .ControlTipText = GetText("SG_053")
                End With
            CommandButtonInputOK.Caption = GetText("SG_015")
            
        'Page "Verarbeitung und Ausgabe"
            'Section "Ausgabeoptionen"
                .FrameOutputOptions.Caption = GetText("SG_016")
                .lblSpaltenAusgabe.Caption = GetText("SG_017")
                .CommandButtonSelectAllColumnsOutput.Caption = GetText("SG_013")
                .CommandButtonSelectNoColumnsOutput.Caption = GetText("SG_014")
                With .ListBoxColumnsOutput
                    .MultiSelect = fmMultiSelectExtended 'multiple columns can be selected
                    .ControlTipText = GetText("SG_055")
                End With
            'Section "Vorschau"
                .CommandButtonPreview.Caption = GetText("SG_020")
                .CommandButtonPreview.Enabled = False
                .ListBoxPreview.ControlTipText = GetText("SG_056")
            'Section "Ausgabe"
                .CommandButtonStart.Caption = GetText("SG_021")
                .CommandButtonStart.BackColor = &H8000000F
                .CommandButtonStart.Enabled = False
        'Page "Info"
            .lblInfo1.Caption = GetText("SG_060") & vbNewLine & vbNewLine _
                                & GetText("SG_061") & vbNewLine & vbNewLine _
                                & GetText("SG_062") & vbNewLine & vbNewLine _
                                & GetText("SG_063") & vbNewLine & vbNewLine _
                                & GetText("SG_064") & vbNewLine & vbNewLine _
                                & GetText("SG_065") & vbNewLine & vbNewLine _
                                & GetText("SG_066") & vbNewLine & vbNewLine _
                                & GetText("SG_067") & vbNewLine & vbNewLine _
                                & GetText("SG_068") & vbNewLine & vbNewLine _
                                & GetText("SG_069") & vbNewLine & vbNewLine _
                                & GetText("SG_070")
    End With

    With ComboBoxAusgabe
        .ControlTipText = GetText("SG_057")
        .Style = fmStyleDropDownList 'allow only values from the item list, no free entries
        .AddItem GetText("SG_022")
        .AddItem GetText("SG_023")
        .AddItem GetText("SG_024")
        .AddItem GetText("SG_025")
'        .AddItem "Word-Datei (.docx)"
        .ListIndex = 1 'Startwert "Bildschirm"
    End With

    With SpinButtonVorkommen 'Ausgabe ab dem X. Vorkommen
        .Orientation = fmOrientationVertical
        .Min = 1
'        .Max = 1000 'wenn kein Max angegeben ist der Wert anscheinend 100
    End With
    
    Call modAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub


'Elements on page "Datenquelle"
'------------------------------

Private Sub CommandButtonSelectFolder_Click() 'click on the button "Ordner öffnen und Dateien auswählen"
    Call SelectFiles
    Call PopulateBoxesWithFiles
    FrameSourceColumns.Caption = GetText("SG_009")
    FrameSourceColumnsFile.Caption = GetText("SG_010")
    Call CommandButtonRowHeader1_Click
    Call ResetOutputPage
End Sub

Private Sub ListBoxFiles_Click() 'click on the ListBox with the file names
    ComboBoxFileForColumns.ListIndex = ListBoxFiles.ListIndex 'selected file = file for columns
    FrameSourceColumns.Caption = GetText("SG_009") & " - " _
                                & GetText("SG_027") & ": " & ComboBoxFileForColumns.Value
    FrameSourceColumnsFile.Caption = GetText("SG_027")
    Call PageOutputActivation
    Call ResetOutputPage
End Sub

Private Sub ComboBoxFileForColumns_Change() 'change of the file for the column headers
    ListBoxFiles.ListIndex = ComboBoxFileForColumns.ListIndex 'file for columns = selected file
    ComboBoxFileForColumns.BackColor = &H8000000F 'grau
    ListBoxFiles.BackColor = &H8000000F 'grau
    ListBoxColumnHeaders.BackColor = &H80FFFF 'gelb
    Call GetColumnHeaders
    Call PageOutputActivation
    Call ResetOutputPage
End Sub

Private Sub SpinButtonRow_Change() 'change of the row with for column headers
    On Error Resume Next
    
    lblRowHeaders.Caption = GetText("SG_028") & ": " & vbNewLine & SpinButtonRow.Value
    Call GetColumnHeaders
    
    On Error GoTo 0
End Sub

Private Sub CommandButtonRowHeader1_Click() 'click on the button "Zeile 1"
    SpinButtonRow.Value = 1
End Sub

Private Sub CommandButtonSelectAllColumns_Click() 'click on the button "Alle Spalten selektieren"
    Call SelectAllColumns(True)
End Sub

Private Sub CommandButtonSelectNoColumns_Click() 'click on the button "Keine Spalte selektieren"
    Call SelectAllColumns(False)
End Sub

Private Sub CommandButtonInputOK_Click() 'click on the button "Weiter zur Ausgabe"
    MultiPage1.Value = 1
End Sub


'Elements on page "Datenausgabe"
'------------------------------

Private Sub CommandButtonSelectAllColumnsOutput_Click() 'click on the button "Alle Spalten selektieren"
    Call SelectAllColumnsOutput(True)
End Sub

Private Sub CommandButtonSelectNoColumnsOutput_Click() 'click on the button "Keine Spalte selektieren"
    Call SelectAllColumnsOutput(False)
End Sub

Private Sub CommandButtonPreview_Click() 'click on the button "Vorschau Aktualisieren"

    frmFortschritt.Show (vbModeless)
    
    Call Build
    Call OutputPreview
    
    Unload frmFortschritt
    
End Sub

Private Sub CommandButtonStart_Click() 'click on the button "Datensätze ausgeben"

    frmFortschritt.Show (vbModeless)
    
    Call Build
    Call Output
    
    Unload frmFortschritt
        
End Sub

Private Sub Build()
    Call CreateSearchArrays
    Call CreateOutputCountArray
    Call CreateOutputArray
End Sub

Private Sub SpinButtonVorkommen_Change()
    lblVorkommen.Caption = GetText("SG_029") & " " & SpinButtonVorkommen.Value & GetText("SG_030")
    lblVorkommen.ControlTipText = GetText("SG_054")
    g_sgds_intAusgabeAbXmal = SpinButtonVorkommen.Value
    Call NumberOutputReset
End Sub

'Aktualisieren dar Anzeige, wie viele Datensätze ausgegeben werden
Private Sub NumberOutputReset()
    FrameOutputPreview.Caption = GetText("SG_031")
    FrameOutput.Caption = GetText("SG_032")
    ListBoxPreview.Clear
End Sub

'Aktivierung der Buttons
Private Sub ListBoxColumnsOutput_Change()
    Dim i As Integer
    CommandButtonPreview.Enabled = False
    CommandButtonStart.Enabled = False
    CommandButtonStart.BackColor = &H8000000F
        For i = 0 To frm_SGDS_GUI.ListBoxColumnsOutput.ListCount - 1
            If frm_SGDS_GUI.ListBoxColumnsOutput.Selected(i) = True Then
                CommandButtonPreview.Enabled = True
                CommandButtonStart.Enabled = True
                CommandButtonStart.BackColor = &H80FF80
                Exit Sub
            End If
        Next
End Sub

Private Sub ListBoxColumnHeaders_Change()
    Call PageOutputActivation
    Call NumberOutputReset
End Sub

'Aktivieren bzw. Deaktivieren der Registerkarte "Ausgabe"
Private Sub PageOutputActivation()
    Dim i As Integer
    CommandButtonInputOK.Enabled = False
    CommandButtonInputOK.BackColor = &H8000000F 'grau
    MultiPage1.Pages(1).Enabled = False
        For i = 0 To frm_SGDS_GUI.ListBoxColumnHeaders.ListCount - 1
            If frm_SGDS_GUI.ListBoxColumnHeaders.Selected(i) = True Then
                CommandButtonInputOK.Enabled = True
                CommandButtonInputOK.BackColor = &H80FF80 'grün
                MultiPage1.Pages(1).Enabled = True
                Exit Sub
            End If
        Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload frm_SGDS_OutScreen
End Sub
