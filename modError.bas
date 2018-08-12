Attribute VB_Name = "modError"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Fehlerbehandlung, damit das Tool nicht abstürzt

Public Sub ErrorPopup(err As ErrObject)
    If xlef_blnProgressBar Then Unload frmFortschritt
    
    MsgBox (GetText("ERR_003") & vbCrLf & vbCrLf & _
        err.Description & vbCrLf & _
        "(" & GetText("ERR_005") & " " & err.Number & ")"), _
        vbExclamation, GetText("ERR_004") & "!"
End Sub

