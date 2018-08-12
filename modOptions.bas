Attribute VB_Name = "modOptions"
Option Explicit
Option Private Module

'Checkbox Screenupdate
Public Sub optionsScreenUpdate()
    If xlef_blnScreenUpdate Then
        xlef_blnScreenUpdate = False
    Else
        xlef_blnScreenUpdate = True
    End If
End Sub

'Checkbox Fortschrittsbalken
Public Sub optionsProgressBar()
    If xlef_blnProgressBar Then
        xlef_blnProgressBar = False
    Else
        xlef_blnProgressBar = True
    End If
End Sub

'Screenupdate an ----> geht nicht richtig !!!!!!!!!!
Public Sub optionsScreenUpdateON()
'    Application.ScreenUpdating = True
End Sub

'Screenupdate aus ----> geht nicht richtig !!!!!!!!!!
Public Sub optionsScreenUpdateOFF()
'    Application.ScreenUpdating = False
End Sub


'Fortschrittsbalken
'------------------

Public Sub FortschrittON(strText As String) 'Fortschrittsbalken einblenden
    xlef_dblBalkenAktuell = 0
    Load frmFortschritt
    With frmFortschritt
        .StartUpPosition = 1 'Zentriert im Element, zu dem das UserForm-Objekt gehört
        .Caption = strText 'Überschrift
        .lblFortschrittBalken.Width = 0 'Balken zurücksetzen
        .lblFortschrittBalken.BackColor = &H800080    'Farbe setzen
        .Show
    End With
End Sub

Public Sub BalkenAkt(dblProzent As Double) 'Fortschrittsbalken aktualisieren
    With frmFortschritt
        .lblFortschrittBalken.Width = CInt(dblProzent) 'Breite des Balkens
    End With
    DoEvents 'neu zeichnen
End Sub
