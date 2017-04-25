Attribute VB_Name = "modStart"
Option Explicit
Option Private Module

'*****************************************************************************************
'*** Das Excel Add-in 'ELPHIN Toolbox' fügt eine neue Registerkarte in das Ribbon ein ****
'*** mit verschiedenen Funktionen. Teilweise werden Worksheet-Funktionen direkt auf ******
'*** eine markierte Zelle angewandt, teilweise werden VBA-Funktionen für das Worksheet ***
'*** verfügbar gemacht, und teilweise werden neue Funktionen zur Verfügung gestellt. *****
'***                                                                                   ***
'*** Version 1.0 (April 2017) - Excel 2007-2016 (optimiert für Windows)                ***
'*** Autoren: Juliane Held / Marco Krapf - mailto:excel@marco-krapf.de                 ***
'*****************************************************************************************

'FEATURES IMPLEMENTIERT
'======================
    'VBA-Funktionen:
        'UCASE (Konvertieren in Großbuchstaben) - als Value und als Funktion
        'LCASE (Konvertieren in Kleinbuchstaben) - als Value und als Funktion
        'LTRIM (Leerzeichen am Anfang der Zelle entfernen) - als Value
        'RTRIM (Leerzeichen am Ende der Zelle entfernen) - als Value
        'TRIM (Leerzeichen am Anfang und Ende der Zelle entfernen) - als Value
        'ABS (Absoluter Wert einer Zahl, also Zahl ohne Vorzeichen) - als Value und als Funktion
    'Worksheet-Funktionen:
        '=PROPER (=GROSS2 --> den ersten Buchstabe jedes Wortes großschreiben) - als Value und als Funktion
        '=TRIM (=GLÄTTEN --> Leerzeichen am Anfang und Ende der Zelle und mehrfache mittendrin entfernen) - als Value und als Funktion
    'Eigene Funktionen:
        'Leere Zeilen löschen
        'Leere Spalten löschen
        'Auf Steuerzeichen prüfen
        'Steuerzeichen entfernen (im Endeffekt Worksheet-Funktion =CLEAN (=SÄUBERN)) - als Value
        'Auf geschützte Leerzeichen prüfen
        'Geschützte Leerzeichen durch normales ersetzen - als Value
        'Formel/Funktion durch Wert ersetzen - als Value
        'Spaltennummer in Zwischenablage ablegen
        'Zeilennummer in Zwischenablage ablegen
        'Wenn Formel/Funktion 0 ergibt, dann Zelle leer darstellen -> in Funktion einbauen
        'Worksheet-Vergleich: 2 Tabellen auf unterschiedliche Werte vergleichen und Unterschiede farblich markieren
    'Letzten Schritt rückgängig machen

'IDEEN
'=====
'VBA-Funktionen für das Tabellenblatt verfügbar machen
    '
'Worksheet-Funktionen grafisch verfügbar machen
    '
'Eigene Funktionen
    'Führende Nullen weg (evtl. Text zu Zahl, dann gehen die alleine weg?)
    'Verketten: Marco Krapf -> Marco & Krapf -> MarcoKrapf // Mit oder ohne Leerzeichen? 1 Zielzelle für verschiedene Zellen?
    'SVERWEIS vereinfachen,(=SVERWEIS(Suchkriterium,Spalte,Rückgabespalte), sodass Eingangsparater Matrix nicht benötigt wird, da bei großen Tabellen kompliziert Rückgabespalte auszurechnen
    'SVERWEIS mit zwei Bedingungen SSVERWEIS (=SSVERWEIS(Suchkriterium1, Spalte1, Suchkriterium2, Spalte2, Rückgabespalte)
        'Alternative: Spaltennummer in Zwischenablage / auch Zeilennummer, AdressLocal, Worksheet, Workbook...
    

'WEITERENTWICKLUNG
'=================
'Sind unsere Verweise alle Standard oder müssen wir die dynamisch einbauen?
'Empfehlung für Schreibschutz aktivieren/deaktivieren (Logik passt nicht)
'Mehrere Schritte rückwärts (z.B. das Array mit den Wiederherstellungsinfos in eine Collection packen
'Wenn 0 als Wert in Zelle, dann --> "" oder IF(Zelle.Value=0;"";Zelle.Value)
'Funktionen auch in Worksheet-Kontextmenü (Rechtsklick) einbauen --> Reset des Kontextmenüs bei WorkbookClose? // oder http://www.rholtz-office.de/ribbonx/das-kontextmenue
'Info über die Zelle (Adresse, Inhalt, Datentyp, HasFormula...)
'Vergleich, erkennt Text/Zahlen auch in anderer Reihenfolge am besten eine Spalte mit einer anderen Spalte in Tabelle vergleichen (nach Spaltenindex fragen)
'Worksheets vergleichen --> nicht nur Werte, auch Formeln oder beides
'Toolsprache EN zusätzlich (Ribbon und Formulare)
'Selektierten Bereich in gleichem Format in neue E-Mail kopieren ---> ist schon 90% fertig
'Screenupdating ein/ausschalten bei allen Funktionen? --> Performance // Optional per ToggleButton?
'Dynamischer Tooltip bei "Rückgängig" --> Name der zuvor ausgeführten Aktion
'Dynamisches Ribbon
    'Sprache anpassen
    'Anzeige von Zeilen-/Spaltennummer, Adresse usw.
'Funktion 0 = "" auch wenn Ergebnis sich später ändert und dann 0 wird
'Ordentliche Fehlermeldungen wenn Hyperlink oder Mail Fehler wirft (aktuell wird nur übersprungen)
'Info-Popup wenn Worksheetvergleich angeklickt und nur 1 Worksheet da
'Info-Popup wenn Steuerzeichen prüfen / Gesch. Leerzeichen prüfen keine Funde gibt --> "sauber"
'Zeilen löschen, die "..." enthalten --> darf Juli machen, grobe Orientierung an Sub fnLEEREZEILEN()
    'Bedingte Zeilenlöschung (kompletter Wert --> value) in gesamter Tabelle/Selektion
    'Bedingte Zeilenlöschung ("Teilwert" --> instring) in gesamter Tabelle/Selektion
    'Bedingte Spaltenlöschung (kompletter Wert --> value) in gesamter Tabelle/Selektion
    'Bedingte Spaltenlöschung ("Teilwert" --> instring) in gesamter Tabelle/Selektion
'Checkbox: Passwortschutz beim Speichern
'Worksheets vergleichen --> UserForm schick machen, auch mit Icon
'Worksheets vergleichen --> Tooltips


'Bekannte Bugs
'=============
'Funktionen in Dropdownmenüs können nach Betätigung des Rückgängigbuttons nicht wiederholt ausgeführt werden (Problem mit Callbacks)


'TO-DOs
'======
    
'Auslieferung
    'M Excel-->Datei-->Beschreibung (vgl. andere Tools)
    'M VBA Projekt Beschreibung (vgl. andere Tools)
    'M xlam-Datei erstellen
    'M xmlUI Ribbon erstellen
        'Tool-Logo in Ribbon rein
    'Z Installer-Datei
    'Z Code auf GitHub
    'Z Upload auf Download-Seite
        
    
'========================================================================================================

'Globale Konstanten und Variablen deklarieren
Public Const xlef_strToolname As String = "ELPHIN Toolbox" 'Tool-Name
Public Const xlef_strVersion As String = "1.0 (April 2017)" 'Tool-Version
Public Const xlef_strKontakt1 As String = "excel@marco-krapf.de" 'Kontakt-E-Mail-Adresse
Public Const xlef_strGitHub As String = "https://github.com/MarcoKrapf/ELPHIN_Toolbox" 'GitHub-Repository
Public Const xlef_strSpendenURL As String = "http://www.ghfkh.de/" 'Spenden-Link
Public xlef_wksTarget As Worksheet 'Worksheet, auf dem die Aktion ausgeführt wird
Public xlef_Sel As Variant 'Selektierter Bereich auf dem Tabellenblatt
Public xlef_art As String 'Objekt auf das die Aktion angewendet wird (für Merken und Wiederherstellen)
Public xlef_rngCell As Range 'Einzelne Zelle in For-Each-Schleifen
Public xlef_arrOrg() As Variant 'Array zum Merken der originalen Zellinhalte
Public xlef_coll As Collection 'Collection zum Merken der leeren Zeilen
Public xlef_objClip As DataObject 'Objekt zum Ablegen von Daten in die Windows-Zwischenablage
Public xlef_SchreibschutzEmpfehlen As Boolean

'Startprozedur für Test und Entwicklung
Public Sub DEV_GUI_Start()

    xlef_SchreibschutzEmpfehlen = ThisWorkbook.ReadOnlyRecommended
    DEV_frmGUI.Show
    
End Sub

'Aktionen beim Starten einer Funktion
Public Sub Aktion()
    Set xlef_wksTarget = ActiveWorkbook.ActiveSheet
    Set xlef_Sel = Selection 'Selektierten Bereich in Variable einlesen
    Call SelectionSave 'Aktuellen Zustand merken für Wiederherstellung
End Sub
