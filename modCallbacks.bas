Attribute VB_Name = "modCallbacks"
Option Explicit
'Option Private Module

'Modulbeschreibung:
'Hier werden die Klick-Ereignisse aus dem Ribbon verarbeitet

Public Sub MyAddInInitialize(Ribbon As IRibbonUI) 'Excel-Menüband in Variable einlesen
    Set MyRibbon = Ribbon
    xlef_strVersion = GetText("ELP_000")
End Sub

Public Sub ribbonNEW() 'Excel-Menüband neu aufbauen
    MyRibbon.Invalidate
End Sub

'CALLBACKS FÜR BUTTONS
'=====================

'Callbacks der Buttons im Menu "Text konvertieren"
Public Sub xlef_btn_TextKonv01(control As IRibbonControl)
    Call valueUCASE
End Sub
Public Sub xlef_btn_TextKonv02(control As IRibbonControl)
    Call funcUCASE
End Sub
Public Sub xlef_btn_TextKonv03(control As IRibbonControl)
    Call valueLCASE
End Sub
Public Sub xlef_btn_TextKonv04(control As IRibbonControl)
    Call funcLCASE
End Sub
Public Sub xlef_btn_TextKonv05(control As IRibbonControl)
    Call valuePROPER
End Sub
Public Sub xlef_btn_TextKonv06(control As IRibbonControl)
    Call funcPROPER
End Sub

'Callbacks der Buttons im Menu "Zeichen entfernen"
Public Sub xlef_btn_Zeichen_01(control As IRibbonControl)
    Call valueLTRIM
End Sub
Public Sub xlef_btn_Zeichen_02(control As IRibbonControl)
    Call valueRTRIM
End Sub
Public Sub xlef_btn_Zeichen_03(control As IRibbonControl)
    Call valueTRIM 'Leerzeichen links und rechts entfernen (als Wert)
End Sub
Public Sub xlef_btn_Zeichen_04(control As IRibbonControl)
    Call valueTRIM_WKS 'Alle mehrfachen Leerzeichen entfernen (als Wert)
End Sub
Public Sub xlef_btn_Zeichen_05(control As IRibbonControl)
    Call funcTRIM_WKS 'Alle mehrfachen Leerzeichen entfernen (als Funktion)
End Sub
Public Sub xlef_btn_Zeichen_06(control As IRibbonControl)
    Call fnGESCHLEERZEICHENpruefen
End Sub
Public Sub xlef_btn_Zeichen_07(control As IRibbonControl)
    Call fnGESCHLEERZEICHENaustauschen
End Sub
Public Sub xlef_btn_Zeichen_08(control As IRibbonControl)
    Call fnSTEUERZEICHENpruefen
End Sub
Public Sub xlef_btn_Zeichen_09(control As IRibbonControl)
    Call fnSTEUERZEICHENentfernen
End Sub

'Callbacks der Buttons im Menu "Mathematik"
Public Sub xlef_btn_Math01(control As IRibbonControl)
    Call valueABS
End Sub
Public Sub xlef_btn_Math02(control As IRibbonControl)
    Call funcABS
End Sub

'Callbacks der Buttons im Menu "Formeln/Funktionen"
Public Sub xlef_btn_Funk01(control As IRibbonControl)
    Call valueFORMELzuWERT
End Sub
Public Sub xlef_btn_Funk02(control As IRibbonControl)
    Call fnLEERWENNNULL
End Sub

'Callbacks der Buttons im Menu "Zeilen"
Public Sub xlef_btn_Zeilen01(control As IRibbonControl)
    Call fnLEEREZEILEN
End Sub
Public Sub xlef_btn_Zeilen02(control As IRibbonControl)
    Call fnZEILEMERKEN
End Sub

'Callbacks der Buttons im Menu "Spalten"
Public Sub xlef_btn_Spalten01(control As IRibbonControl)
    Call fnLEERESPALTEN
End Sub
Public Sub xlef_btn_Spalten02(control As IRibbonControl)
    Call fnSPALTEMERKEN
End Sub

'Callback für Button "Worksheets vergleichen" onAction
Sub btnWKSVGL_onAction(control As IRibbonControl)
    Call fnWORKSHEETVERGLEICH 'Nur Werte der Zellen vergleichen, nicht Formate usw.
End Sub

'Callback für Checkbox "Schreibschutz" onAction
Sub cboxSCHREIBSCHUTZ_onAction(control As IRibbonControl, pressed As Boolean)
    If pressed = True Then
        xlef_SchreibschutzEmpfehlen = True
        ThisWorkbook.ReadOnlyRecommended = True
    Else
        xlef_SchreibschutzEmpfehlen = False
        ThisWorkbook.ReadOnlyRecommended = False
    End If
End Sub

'Callback für Button "E-Mail" onAction
Sub btnMAIL_onAction(control As IRibbonControl)
    dosomething ("Mail mit Clipboard")
End Sub

'Callbacks für Buttons "Rückgängig" onAction
Sub btnUNDO_onAction(control As IRibbonControl)
    Call SelectionRestore
End Sub
Sub btnUNDO_onoff_onAction(control As IRibbonControl)
    If xlex_blnUNDO = True Then
        xlex_blnUNDO = False
    Else
        xlex_blnUNDO = True
    End If
    Call ribbonNEW
End Sub

'Callback für Button "UsedRange" onAction
Sub btn_usedrange_onAction(control As IRibbonControl)
    Call modAuxiliary.SelectUsedRange
End Sub

'Callback für Button "Info" onAction
Sub btnINFO_onAction(control As IRibbonControl)
    Call infoINFO
End Sub

'Callback für Button "Sprache" onAction
Sub xlef_btn_sprache_onAction(control As IRibbonControl)
    Call Sprache
End Sub

'Callback für Button "SGDS Tool" onAction
Sub btn_sgds_onAction(control As IRibbonControl)
    Call mod_SGDS_Tool.openGUI
End Sub

'CALLBACKS FÜR LABELS
'====================

Public Sub AI_GetLabel(control As IRibbonControl, ByRef returnedVal)
    Select Case control.id
    
        Case "xlef_group1"
            returnedVal = GetText("GRP001") '" "
                Case "xlef_btn_undo"
                    If xlex_blnUNDO = True Then
                        returnedVal = GetText("BTN001a")
                    Else
                        returnedVal = GetText("BTN001b")
                    End If
                Case "xlef_btn_undo_onoff"
                    If xlex_blnUNDO = True Then
                        returnedVal = GetText("BTN030b")
                    Else
                        returnedVal = GetText("BTN030a")
                    End If
        Case "xlef_group2"
            returnedVal = GetText("GRP002")
                Case "xlef_menu_TextKonv"
                    returnedVal = GetText("MENU001")
                        Case "xlef_btn_TextKonv01"
                            returnedVal = GetText("BTN002")
                        Case "xlef_btn_TextKonv02"
                            returnedVal = GetText("BTN003")
                        Case "xlef_btn_TextKonv03"
                            returnedVal = GetText("BTN004")
                        Case "xlef_btn_TextKonv04"
                            returnedVal = GetText("BTN005")
                        Case "xlef_btn_TextKonv05"
                            returnedVal = GetText("BTN006")
                        Case "xlef_btn_TextKonv06"
                            returnedVal = GetText("BTN007")
                Case "xlef_menu_Zeichen"
                    returnedVal = GetText("MENU002")
                        Case "xlef_btn_Zeichen_01"
                            returnedVal = GetText("BTN008")
                        Case "xlef_btn_Zeichen_02"
                            returnedVal = GetText("BTN009")
                        Case "xlef_btn_Zeichen_03"
                            returnedVal = GetText("BTN010")
                        Case "xlef_btn_Zeichen_04"
                            returnedVal = GetText("BTN011")
                        Case "xlef_btn_Zeichen_05"
                            returnedVal = GetText("BTN012")
                        Case "xlef_btn_Zeichen_06"
                            returnedVal = GetText("BTN013")
                        Case "xlef_btn_Zeichen_07"
                            returnedVal = GetText("BTN014")
                        Case "xlef_btn_Zeichen_08"
                            returnedVal = GetText("BTN015")
                        Case "xlef_btn_Zeichen_09"
                            returnedVal = GetText("BTN016")
                Case "xlef_menu_Math"
                    returnedVal = GetText("MENU003")
                        Case "xlef_btn_Math01"
                            returnedVal = GetText("BTN017")
                        Case "xlef_btn_Math02"
                            returnedVal = GetText("BTN018")
                Case "xlef_menu_Funk"
                    returnedVal = GetText("MENU004")
                        Case "xlef_btn_Funk01"
                            returnedVal = GetText("BTN019")
                        Case "xlef_btn_Funk02"
                            returnedVal = GetText("BTN020")
            Case "xlef_group3"
                returnedVal = GetText("GRP003")
                    Case "xlef_menu_Zeilen"
                        returnedVal = GetText("MENU005")
                            Case "xlef_btn_Zeilen01"
                                returnedVal = GetText("BTN021")
                            Case "xlef_btn_Zeilen02"
                                returnedVal = GetText("BTN022")
                    Case "xlef_menu_Spalten"
                        returnedVal = GetText("MENU006")
                            Case "xlef_btn_Spalten01"
                                returnedVal = GetText("BTN023")
                            Case "xlef_btn_Spalten02"
                                returnedVal = GetText("BTN024")
        Case "xlef_group4"
            returnedVal = GetText("GRP004")
                Case "xlef_btn_WksVgl"
                    returnedVal = GetText("BTN025")
        Case "xlef_group5"
            returnedVal = GetText("GRP005")
                Case "xlef_btn_INFO"
                    returnedVal = GetText("BTN026")
        Case "xlef_group6"
            returnedVal = GetText("GRP006")
                Case "xlef_btn_spracheDE"
                    returnedVal = GetText("BTN027")
                Case "xlef_btn_spracheEN"
                    returnedVal = GetText("BTN028")
        Case "xlef_group7"
            returnedVal = GetText("GRP007")
                Case "xlef_btn_sgds"
                    returnedVal = GetText("BTN029")
        Case "xlef_group8"
            returnedVal = GetText("GRP008")
                Case "xlef_btn_usedrange"
                    returnedVal = GetText("BTN031")

    End Select
End Sub


'CALLBACKS FÜR SCREENTIPS
'========================

Public Sub AI_GetScreentip(control As IRibbonControl, ByRef screentip)

    Select Case control.id
        Case "xlef_btn_undo"
            screentip = GetText("SCRTIPBTN001")
        Case "xlef_btn_usedrange"
            screentip = GetText("SCRTIPBTN031")
        Case "xlef_menu_TextKonv"
            screentip = GetText("SCRTIPMENU001")
                Case "xlef_btn_TextKonv01"
                    screentip = GetText("SCRTIPBTN002")
                Case "xlef_btn_TextKonv02"
                    screentip = GetText("SCRTIPBTN003")
                Case "xlef_btn_TextKonv03"
                    screentip = GetText("SCRTIPBTN004")
                Case "xlef_btn_TextKonv04"
                    screentip = GetText("SCRTIPBTN005")
                Case "xlef_btn_TextKonv05"
                    screentip = GetText("SCRTIPBTN006")
                Case "xlef_btn_TextKonv06"
                    screentip = GetText("SCRTIPBTN007")
        Case "xlef_menu_Zeichen"
            screentip = GetText("SCRTIPMENU002")
                Case "xlef_btn_Zeichen_01"
                    screentip = GetText("SCRTIPBTN008")
                Case "xlef_btn_Zeichen_02"
                    screentip = GetText("SCRTIPBTN009")
                Case "xlef_btn_Zeichen_03"
                    screentip = GetText("SCRTIPBTN010")
                Case "xlef_btn_Zeichen_04"
                    screentip = GetText("SCRTIPBTN011")
                Case "xlef_btn_Zeichen_05"
                    screentip = GetText("SCRTIPBTN012")
                Case "xlef_btn_Zeichen_06"
                    screentip = GetText("SCRTIPBTN013")
                Case "xlef_btn_Zeichen_07"
                    screentip = GetText("SCRTIPBTN014")
                Case "xlef_btn_Zeichen_08"
                    screentip = GetText("SCRTIPBTN015")
                Case "xlef_btn_Zeichen_09"
                    screentip = GetText("SCRTIPBTN016")
        Case "xlef_menu_Math"
            screentip = GetText("SCRTIPMENU003")
                Case "xlef_btn_Math01"
                    screentip = GetText("SCRTIPBTN017")
                Case "xlef_btn_Math02"
                    screentip = GetText("SCRTIPBTN018")
        Case "xlef_menu_Funk"
            screentip = GetText("SCRTIPMENU004")
                Case "xlef_btn_Funk01"
                    screentip = GetText("SCRTIPBTN019")
                Case "xlef_btn_Funk02"
                    screentip = GetText("SCRTIPBTN020")
        Case "xlef_menu_Zeilen"
            screentip = GetText("SCRTIPMENU005")
                Case "xlef_btn_Zeilen01"
                    screentip = GetText("SCRTIPBTN021")
                Case "xlef_btn_Zeilen02"
                    screentip = GetText("SCRTIPBTN022")
        Case "xlef_menu_Spalten"
            screentip = GetText("SCRTIPMENU006")
                Case "xlef_btn_Spalten01"
                    screentip = GetText("SCRTIPBTN023")
                Case "xlef_btn_Spalten02"
                    screentip = GetText("SCRTIPBTN024")
        Case "xlef_btn_WksVgl"
            screentip = GetText("SCRTIPBTN025")
        Case "xlef_btn_INFO"
            screentip = GetText("SCRTIPBTN026")
        Case "xlef_btn_sgds"
            screentip = GetText("SG_002")
    End Select

End Sub


'CALLBACKS FÜR SUPERTIPS
'=======================

Public Sub AI_GetSupertip(control As IRibbonControl, ByRef supertip)
   
    Select Case control.id
        Case "xlef_btn_undo"
            supertip = GetText("SUPTIPBTN001")
            Case "xlef_menu_undo"
            supertip = GetText("SUPTIPBTN030")
            Case "xlef_btn_undo_onoff"
            supertip = GetText("SUPTIPBTN030")
        Case "xlef_btn_usedrange"
            supertip = GetText("SUPTIPBTN031")
        Case "xlef_menu_TextKonv"
            supertip = GetText("SUPTIPMENU001")
            Case "xlef_btn_TextKonv01"
                supertip = GetText("SUPTIPBTN002")
            Case "xlef_btn_TextKonv02"
                supertip = GetText("SUPTIPBTN003")
            Case "xlef_btn_TextKonv03"
                supertip = GetText("SUPTIPBTN004")
            Case "xlef_btn_TextKonv04"
                supertip = GetText("SUPTIPBTN005")
            Case "xlef_btn_TextKonv05"
                supertip = GetText("SUPTIPBTN006")
            Case "xlef_btn_TextKonv06"
                supertip = GetText("SUPTIPBTN007")
        Case "xlef_menu_Zeichen"
            supertip = GetText("SUPTIPMENU002")
                Case "xlef_btn_Zeichen_01"
                    supertip = GetText("SUPTIPBTN008")
                Case "xlef_btn_Zeichen_02"
                    supertip = GetText("SUPTIPBTN009")
                Case "xlef_btn_Zeichen_03"
                    supertip = GetText("SUPTIPBTN010")
                Case "xlef_btn_Zeichen_04"
                    supertip = GetText("SUPTIPBTN011")
                Case "xlef_btn_Zeichen_05"
                    supertip = GetText("SUPTIPBTN012")
                Case "xlef_btn_Zeichen_06"
                    supertip = GetText("SUPTIPBTN013")
                Case "xlef_btn_Zeichen_07"
                    supertip = GetText("SUPTIPBTN014")
                Case "xlef_btn_Zeichen_08"
                    supertip = GetText("SUPTIPBTN015")
                Case "xlef_btn_Zeichen_09"
                    supertip = GetText("SUPTIPBTN016")
        Case "xlef_menu_Math"
            supertip = GetText("SUPTIPMENU003")
                Case "xlef_btn_Math01"
                    supertip = GetText("SUPTIPBTN017")
                Case "xlef_btn_Math02"
                    supertip = GetText("SUPTIPBTN018")
        Case "xlef_menu_Funk"
            supertip = GetText("SUPTIPMENU004")
                Case "xlef_btn_Funk01"
                    supertip = GetText("SUPTIPBTN019")
                Case "xlef_btn_Funk02"
                    supertip = GetText("SUPTIPBTN020")
        Case "xlef_menu_Zeilen"
            supertip = GetText("SUPTIPMENU005")
                Case "xlef_btn_Zeilen01"
                    supertip = GetText("SUPTIPBTN021")
                Case "xlef_btn_Zeilen02"
                    supertip = GetText("SUPTIPBTN022")
        Case "xlef_menu_Spalten"
            supertip = GetText("SUPTIPMENU006")
                Case "xlef_btn_Spalten01"
                    supertip = GetText("SUPTIPBTN023")
                Case "xlef_btn_Spalten02"
                    supertip = GetText("SUPTIPBTN024")
        Case "xlef_btn_WksVgl"
            supertip = GetText("SUPTIPBTN025")
        Case "xlef_btn_INFO"
            supertip = GetText("SUPTIPBTN026")
        Case "xlef_btn_sgds"
            supertip = GetText("SG_047")
    End Select

End Sub

'CALLBACKS FÜR SICHTBARKEIT
'==========================
Sub GetVisible(control As IRibbonControl, ByRef visible)

    Select Case control.id
        Case "xlef_btn_spracheDE"
            If xlef_strSprache = 2 Then
                visible = True
            Else
                visible = False
            End If
        Case "xlef_btn_spracheEN"
            If xlef_strSprache = 3 Then
                visible = True
            Else
                visible = False
            End If
    End Select
                    
 End Sub
 
'CALLBACKS FÜR AKTIV/INAKTIV
'===========================
Sub IsButtonEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.id
        Case "xlef_btn_undo"
            returnedVal = xlex_blnUNDO
    End Select
End Sub

'Für Tests von neuen Ribbon-Items, aufruf in der Callback-Prozedur z.B. mit dosomething ("Worksheets vergleichen")
Private Sub dosomething(txt As String)
    MsgBox txt
End Sub
