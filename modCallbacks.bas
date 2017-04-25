Attribute VB_Name = "modCallbacks"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Hier werden die Klick-Ereignisse aus dem Ribbon verarbeitet

'Callback for dropdown3 onAction
Sub drpCELLTEXT_onAction(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "drp6"
            Call valueUCASE
        Case "drp7"
            Call funcUCASE
        Case "drp8"
            Call valueLCASE
        Case "drp9"
            Call funcLCASE
        Case "drp10"
            Call valuePROPER
        Case "drp11"
            Call funcPROPER
        Case Else
            
    End Select
End Sub

'Callback for dropdown6 onAction
Sub drpCELLTRIM_onAction(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "drp12"
            Call valueLTRIM
        Case "drp13"
            Call valueRTRIM
        Case "drp14"
            Call valueTRIM 'Leerzeichen links und rechts entfernen (als Wert)
        Case "drp15"
            Call valueTRIM_WKS 'Alle mehrfachen Leerzeichen entfernen (als Wert)
        Case "drp16"
            Call funcTRIM_WKS 'Alle mehrfachen Leerzeichen entfernen (als Funktion)
        Case "drp23"
            Call fnGESCHLEERZEICHENpruefen
        Case "drp24"
            Call fnGESCHLEERZEICHENaustauschen
        Case "drp25"
            Call fnSTEUERZEICHENpruefen
        Case "drp26"
            Call fnSTEUERZEICHENentfernen
        Case Else
            
    End Select
End Sub

'Callback for dropdown5 onAction
Sub drpCELLMATH_onAction(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "drp17"
            Call valueABS
        Case "drp18"
            Call funcABS
        Case Else
            
    End Select
End Sub

'Callback for dropdown7 onAction
Sub drpCELLFORMFUNC_onAction(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "drp20"
            Call valueFORMELzuWERT
        Case "drp21"
            Call fnLEERWENNNULL
        Case Else
            
    End Select
End Sub

'Callback for dropdown2 onAction
Sub drpROW_onAction(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "drp3"
            Call fnLEEREZEILEN
        Case "drp2"
            Call fnZEILEMERKEN
        Case Else
            
    End Select
End Sub

'Callback for dropdown4 onAction
Sub drpCOL_onAction(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "drp4"
            Call fnLEERESPALTEN
        Case "drp1"
            Call fnSPALTEMERKEN
        Case Else
            
    End Select
End Sub

'Callback for btn11 onAction
Sub btnWKSVGL_onAction(control As IRibbonControl)
    Call fnWORKSHEETVERGLEICH 'Nur Werte der Zellen vergleichen, nicht Formate usw.
End Sub

'Callback for cbox1 onAction
Sub cboxSCHREIBSCHUTZ_onAction(control As IRibbonControl, pressed As Boolean)
    If pressed = True Then
        xlef_SchreibschutzEmpfehlen = True
        ThisWorkbook.ReadOnlyRecommended = True
    Else
        xlef_SchreibschutzEmpfehlen = False
        ThisWorkbook.ReadOnlyRecommended = False
    End If
End Sub

'Callback for btn1 onAction
Sub btnMAIL_onAction(control As IRibbonControl)
    dosomething ("Mail mit Clipboard")
End Sub

'Callback for btn2 onAction
Sub btnUNDO_onAction(control As IRibbonControl)
    Call SelectionRestore
End Sub

'Callback for btn3 onAction
Sub btnINFO_onAction(control As IRibbonControl)
    Call infoINFO
End Sub

'Für Tests von neuen Ribbon-Items, aufruf in der Callback-Prozedur z.B. mit dosomething ("Worksheets vergleichen")
Private Sub dosomething(txt As String)
    MsgBox txt
End Sub


