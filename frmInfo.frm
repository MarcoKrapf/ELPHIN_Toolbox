VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfo 
   Caption         =   "[TITEL]"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7695
   OleObjectBlob   =   "frmInfo.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Texte und Aktionen für die Infos, die beim Aufrufen gezogen werden
'sowie Klick-Ereignisse
'-------------------------------------------------------------------

Private Sub UserForm_Initialize()

    With frmInfo
    
        .Caption = xlef_strToolname & " - " & GetText("ELP_004")
        
        .btnGitHub.Caption = GetText("ELP_005")
        .btnGitHub.ControlTipText = GetText("ELP_006")
        .btnFeedback.Caption = GetText("ELP_007")
        .btnFeedback.ControlTipText = GetText("ELP_008")
        
        .lblInfo.Caption = xlef_strToolname & vbNewLine & _
            GetText("ELP_009") & " " & xlef_strVersion & vbNewLine & vbNewLine & _
            GetText("ELP_018") & vbNewLine & vbNewLine & _
            GetText("ELP_010") & vbNewLine & _
            GetText("ELP_011") & ": " & xlef_strKontakt1
        
        .lblName.Caption = GetText("ELP_030")
            
        .imgInfo.ControlTipText = GetText("ELP_024")
        
        .imgKleinerHeld.ControlTipText = GetText("ELP_012")
        
        .lblSpende.Caption = GetText("ELP_013")
        .lblSpende1.Caption = GetText("ELP_014")
                    
        .lblNutzung.Caption = GetText("ELP_015")
        .lblNutzung1.Caption = GetText("ELP_016") & xlef_strToolname & GetText("ELP_017")
            
        .frameVersHist.Caption = GetText("ELP_019")
        .lblVersHist.Caption = GetText("ELP_022") & vbNewLine & GetText("ELP_023") & vbNewLine _
                        & GetText("ELP_020") & vbNewLine & GetText("ELP_021")
        .lblVersHist.Width = 340
        .lblVersHist.AutoSize = True

            With .frameVersHist
                .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                .ScrollHeight = .lblVersHist.Height + 25 'height of the vertical scrolling
                .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
            End With
    End With

    Call modAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

Private Sub btnGitHub_Click()
    Call infoSOURCECODEURL 'GitHub Repository im Browser aufrufen
End Sub

Private Sub btnFeedback_Click()
    Call infoFEEDBACK 'E-Mail generieren
End Sub

Private Sub imgKleinerHeld_Click()
    Call infoSPENDELINKURL 'Seite der Stiftung im Browser aufrufen
End Sub

Private Sub imgInfo_Click()
    Call infoDOWNLOADURL 'Download-Seite aufrufen
End Sub
