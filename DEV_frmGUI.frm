VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DEV_frmGUI 
   Caption         =   "GUI für Entwicklung und Test"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11880
   OleObjectBlob   =   "DEV_frmGUI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "DEV_frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFortschrittsbalken_Click()
    Call optionsProgressBar
End Sub

Private Sub chkScreenUpdate_Click()
    Call optionsScreenUpdate
End Sub

Private Sub CommandButton1_Click()
    Call valueUCASE
End Sub

Private Sub CommandButton11_Click()
    Call valueTRIM_WKS
End Sub

Private Sub CommandButton12_Click()
'    Call fnLEN
End Sub

Private Sub CommandButton13_Click()
    Call fnLEEREZEILEN
End Sub

Private Sub CommandButton14_Click()
    Call fnLEERESPALTEN
End Sub

Private Sub CommandButton15_Click()
    Call valueFORMELzuWERT
End Sub

Private Sub CommandButton17_Click()
    Call fnMAILausSELECTION
End Sub

Private Sub CommandButton18_Click()
    Call fnLEERWENNNULL
End Sub

Private Sub CommandButton19_Click()
    Call fnSPALTEMERKEN
End Sub

Private Sub CommandButton2_Click()
    Call SelectionRestore
End Sub

Private Sub CommandButton20_Click()
    Call fnZEILEMERKEN
End Sub

Private Sub CommandButton21_Click()
'    Call fnGETCLIPBOARD
End Sub

Private Sub CommandButton22_Click()
    Call fnWORKSHEETVERGLEICH 'Nur Werte der Zellen vergleichen, nicht Formate usw.
End Sub

Private Sub CommandButton23_Click()
    Call valueABS
End Sub

Private Sub CommandButton24_Click()
    Call fnGESCHLEERZEICHENpruefen
End Sub

Private Sub CommandButton25_Click()
    Call fnSTEUERZEICHENpruefen
End Sub

Private Sub CommandButton26_Click()
    Call fnGESCHLEERZEICHENaustauschen
End Sub

Private Sub CommandButton27_Click()
    Call fnSTEUERZEICHENentfernen
End Sub

Private Sub CommandButton28_Click()
    Call funcABS
End Sub

Private Sub CommandButton29_Click()
    Call funcUCASE
End Sub

Private Sub CommandButton3_Click()
    Call valueLCASE
End Sub

Private Sub CommandButton30_Click()
    Call funcLCASE
End Sub

Private Sub CommandButton34_Click()
    Call funcPROPER
End Sub

Private Sub CommandButton35_Click()
    Call funcTRIM_WKS
End Sub

Private Sub CommandButton36_Click()
 Call fnBEDZEILCLEAR_A
End Sub

Private Sub CommandButton38_Click()
    Call infoSOURCECODEURL
End Sub

Private Sub CommandButton39_Click()
    Call infoINFO
End Sub

Private Sub CommandButton4_Click()
    Call valueLTRIM
End Sub

Private Sub CommandButton40_Click()
    Call fnBEDZEILCLEAR_S
End Sub

Private Sub CommandButton43_Click()
    Call infoFEEDBACK
End Sub

Private Sub CommandButton5_Click()
    Call valueRTRIM
End Sub

Private Sub CommandButton6_Click()
    Call valueTRIM
End Sub

Private Sub CommandButton7_Click()
    Call valuePROPER
End Sub

Private Sub CommandButton9_Click()
    Call fnRAND
End Sub

Public Sub ToggleButton1_Click()
    If ToggleButton1 Then
        ToggleButton1.Caption = "Schreibschutz empfehlen beim Speichern"
        xlef_SchreibschutzEmpfehlen = True
    Else
        ToggleButton1.Caption = "Schreibschutz NICHT empfehlen beim Speichern"
        xlef_SchreibschutzEmpfehlen = False
    End If
End Sub

Private Sub UserForm_Initialize()
    CommandButton1.ControlTipText = "Zellinhalte der Selection zu Großbuchstaben"
    CommandButton3.ControlTipText = "Zellinhalte der Selection zu Kleinbuchstaben"
    CommandButton4.ControlTipText = "Leerzeichen links in Zellen der Selection entfernen"
    CommandButton5.ControlTipText = "Leerzeichen rechts in Zellen der Selection entfernen"
    CommandButton6.ControlTipText = "Leerzeichen links und rechts in Zellen der Selection entfernen"
    CommandButton7.ControlTipText = "Jedes Wort in Zellen der Selection mit Großbuchstaben beginnen"
    CommandButton11.ControlTipText = "Leerzeichen links und rechts und mehrfache mittendrin in Zellen der Selection entfernen"
    CommandButton13.ControlTipText = "Leere Zeilen entfernen: Wenn nur 1 Zelle selektiert, dann auf ganzem Worksheet, wenn Bereich selektiert dann nur in Selection"
    CommandButton14.ControlTipText = "Leere Spalten entfernen: Wenn nur 1 Zelle selektiert, dann auf ganzem Worksheet, wenn Bereich selektiert dann nur in Selection"
    CommandButton15.ControlTipText = "Formeln in Zellen der Selection zu absoluten Werten umwandeln"
    CommandButton19.ControlTipText = "Nummer der Spalte in Zwischenablage kopieren (z.B. zum Einfügen in SVERWEIS)"
    CommandButton20.ControlTipText = "Nummer der Zeile in Zwischenablage kopieren"
    CommandButton17.ControlTipText = "E-Mail generieren mit Inhalt der Selection als Body"
    CommandButton21.ControlTipText = "[evtl. nur für Testzwecke]"
    CommandButton9.ControlTipText = "Zufallszahl zwischen 0 und 100 (macht das Sinn???)"
    CommandButton12.ControlTipText = "MsgBox mit Länge jeder Zelle der Selection"
    CommandButton22.ControlTipText = "Vergleich von 2 Worksheets auf Unterschiede in den Zellen (nur Werte, nicht Formate usw.)"
    CommandButton23.ControlTipText = "Liefert den Absolutwert einer Zahl. Der Absolutwert einer Zahl ist die Zahl ohne ihr Vorzeichen"
    CommandButton24.ControlTipText = "Prüfen auf geschützte Leerzeichen (Zeichencode 160)"
    CommandButton25.ControlTipText = "Prüfen auf Steuerzeichen (Zeichencodes 1-31, 127, 129, 141, 143, 144, 157)"
    CommandButton26.ControlTipText = "Geschützte Leerzeichen entfernen (Zeichencode 160)"
    CommandButton27.ControlTipText = "Steuerzeichen entfernen (Zeichencodes 1-31, 127, 129, 141, 143, 144, 157)"
    ToggleButton1.ControlTipText = "Wenn gedrückt, dann wird beim Speichern die Schreibschutzempfehlung empfohlen"
End Sub
