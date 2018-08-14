VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    xlef_strSprache = 2 '(2 = deutsch, 3 = englisch --> Spalte im Worksheet ELPHIN_Texte)
    Set xlef_wksTexte = ELPHIN_txt 'Bessere Alternative zu ThisWorkbook.Worksheets("ELPHIN_Texte")
    xlex_blnUNDO = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Select Case True And True
        Case xlef_SchreibschutzEmpfehlen And ThisWorkbook.ReadOnlyRecommended 'Button: Schreibschutz AKTIVIEREN / Excel: Schreibschutz AKTIV

        Case xlef_SchreibschutzEmpfehlen And Not ThisWorkbook.ReadOnlyRecommended 'Button: Schreibschutz AKTIVIEREN / Excel: Schreibschutz NICHT AKTIV
            If MsgBox("Speichern mit SCHREIBSCHUTZ EMPFEHLEN aktivieren?", vbQuestion + vbYesNo, "SCHREIBSCHUTZ EMPFEHLEN ist inaktiv") = vbYes Then
                ThisWorkbook.ReadOnlyRecommended = True
            End If
        Case Not xlef_SchreibschutzEmpfehlen And ThisWorkbook.ReadOnlyRecommended 'Button: Schreibschutz NICHT AKTIVIEREN / Excel: Schreibschutz AKTIV
            If MsgBox("Speichern mit SCHREIBSCHUTZ EMPFEHLEN deaktivieren?", vbQuestion + vbYesNo, "SCHREIBSCHUTZ EMPFEHLEN ist aktiv") = vbYes Then
                ThisWorkbook.ReadOnlyRecommended = False
            End If
        Case Not xlef_SchreibschutzEmpfehlen And Not ThisWorkbook.ReadOnlyRecommended 'Button: Schreibschutz NICHT AKTIVIEREN / Excel: Schreibschutz NICHT AKTIV
    
    End Select
End Sub
