VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfo 
   Caption         =   "[TITEL]"
   ClientHeight    =   9360
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


Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblFeatures_Click()

End Sub

Private Sub lblFeatures1_Click()

End Sub

Private Sub lblV_Click()

End Sub

Private Sub lblV1_Click()

End Sub

'Modulbeschreibung:
'Texte und Aktionen f�r die Infos, die beim Aufrufen gezogen werden
'-------------------------------------------------------------------

Private Sub UserForm_Initialize()

    With frmInfo
    
        .Caption = xlef_strToolname & " - Info"
        
        .btnGitHub.Caption = "GitHub Repository"
        .btnGitHub.ControlTipText = "GitHub Repository mit dem Quellcode im Internetbrowser �ffnen"
        .btnFeedback.Caption = "Feedback"
        .btnFeedback.ControlTipText = "Feedback per E-Mail an die Entwickler senden"
        
        .lblInfo.Caption = xlef_strToolname & vbNewLine & _
            "Version " & xlef_strVersion & vbNewLine & vbNewLine & _
            "Autoren: Juliane Held und Marco Krapf" & vbNewLine & _
            "Kontakt: " & xlef_strKontakt1
            
        .imgKleinerHeld.ControlTipText = "Aufrufen der Website der Stiftung 'Gro�e Hilfe f�r kleine Helden' im Internetbrowser"
        .imgQRCode.ControlTipText = "QR-Code scannen zum Aufruf des Online-Spendenformulars"
        
        .lblSpende.Caption = "Spende"
        .lblSpende1.Caption = "Das Excel-Add-in '" & xlef_strToolname & "' wird privat entwickelt und unter " & _
            "http://marco-krapf.de/excel/ kostenlos zum Download angeboten. " & vbNewLine & _
            "�ber eine kleine Spende an die Stiftung 'Gro�e Hilfe f�r kleine Helden' f�r kranke Kinder " & _
            "in der Region Heilbronn w�rden wir uns sehr freuen."
        .lblSpende2.Caption = "Website der Stiftung"
                    
        .lblNutzung.Caption = "Nutzungsbedingungen"
        .lblNutzung1.Caption = "Das Excel-Add-In '" & xlef_strToolname & "' darf ohne Einschr�nkung privat und " & _
            "gewerblich verwendet werden. " & _
            "Die Software wird mit gr��tm�glicher Sorgfalt entwickelt und getestet. " & _
            "F�r Fehler im Code, die unkorrekte Ergebnisse liefern, Abst�rze des Programms oder des Systems " & _
            "verursachen k�nnen, sowie f�r eventuellen Datenverlust durch Anwendung der Tools wird keine " & _
            "Haftung �bernommen."
            
        .lblV.Caption = "Versionshistorie"
        .lblV1.Caption = "Version 1.0 (25.04.2017)" & vbNewLine & _
            "- Alles in Gro�buchstaben (Wert/Funktion) " & vbNewLine & _
            "- Alles in Kleinbuchstaben (Wert/Funktion)" & vbNewLine & _
            "- Jedes Wort gro� schreiben (Wert/Funktion)" & vbNewLine & _
            "- Leerzeichen links entfernen (Wert)" & vbNewLine & _
            "- Leerzeichen rechts entfernen (Wert)" & vbNewLine & _
            "- Leerzeichen links und rechts entfernen (Wert)" & vbNewLine & _
            "- Alle mehrfahen Leerzeichen entfernen (Wert/Funktion)" & vbNewLine & _
            "- Auf gesch�tzte Leerzeichen �berpr�fen" & vbNewLine & _
            "- Gesch�tzte Leerzeichen ersetzen" & vbNewLine & _
            "- Auf Steuerzeichen �berpr�fen" & vbNewLine & _
            "- Steuerzeichen entfernen" & vbNewLine & _
            "- Absoluter Wert (Wert/Funktion)" & vbNewLine & _
            "- Formel zu Wert" & vbNewLine & _
            "- Formelergebniss null (0) ausblenden" & vbNewLine & _
            "- Leere Zeilen entfernen" & vbNewLine & _
            "- Leere Spalten entfernen" & vbNewLine & _
            "- Zeilenindex merken" & vbNewLine & _
            "- Spaltenindex merken" & vbNewLine & _
            "- Worksheets vergleichen"

    End With

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

Private Sub lblSpende2_Click()
    Call infoSPENDELINKURL 'Seite der Stiftung im Browser aufrufen
End Sub
