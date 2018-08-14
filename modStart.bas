Attribute VB_Name = "modStart"
Option Explicit
Option Private Module

'*****************************************************************************************
'*** Das Excel Add-in 'ELPHIN Toolbox' fügt eine neue Registerkarte in das Ribbon ein ****
'*** mit verschiedenen Funktionen. Teilweise werden Worksheet-Funktionen direkt auf ******
'*** eine markierte Zelle angewandt, teilweise werden VBA-Funktionen für das Worksheet ***
'*** verfügbar gemacht, und teilweise werden neue Funktionen zur Verfügung gestellt. *****
'***                                                                                   ***
'*** Version 2.0 (August 2018) - Excel 2010-2016 (optimiert für Windows)               ***
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
        'Geschützte Leerzeichen durch normale ersetzen - als Value
        'Formel/Funktion durch Wert ersetzen - als Value
        'Spaltennummer in Zwischenablage ablegen
        'Zeilennummer in Zwischenablage ablegen
        'Wenn Formel/Funktion 0 ergibt, dann Zelle leer darstellen -> in Funktion einbauen
        'Worksheet-Vergleich: 2 Tabellen auf unterschiedliche Werte vergleichen und Unterschiede farblich markieren
    'Eigene Tools:
        'SG DataSet Finder - Tool zum Finden von mehrfach vorkommenden Datensätzen in verschiedenen Dateien
    'Letzten Schritt rückgängig machen

    
'KOMMENTARE 10.04.2018
'=====================
'- (erledigt) Dropdown-Problem: Ausprobieren mit "menu" anstelle "dropDown" item // oder menu mit buttons machen??!
'- (erledigt) Ribbon zweisprachig (EN/DE) machen
'- (erledigt) QR-Code rausmachen (veraltet)

'KOMMENTARE 24.07.2018
'=====================
'- (erledigt) Dropdown-Problem gelöst durch "menu" mit buttons anstelle "dropDown" item
'- (erledigt) Mehrsprachigkeit vorbereitet durch Callbacks, die Texte aus dem Worksheet "ELPHIN_Texte" ziehen

'KOMMENTARE 05.08.2018
'=====================
'- (erledigt) SG DataSet Finder (HHN-Tool) eingebaut
'- (erledigt) UserForms werden beim Aufruf zentriert platziert
'- (erledigt) Popup "Info" aktualisiert
'- (erledigt) EN-Übersetzungen aller Texte

'KOMMENTARE 10.08.2018
'=====================
'- Rückgängig-Button optional
'- Scrollbar in Info-Popup eingebaut

'KOMMENTARE 11.08.2018
'=====================
'- Warnmeldung wenn komplettes Worksheet markiert ist

'KOMMENTARE 12.08.2018
'=====================
'- Warnmeldung wenn ganze Spalte markiert ist

'Bekannte Bugs
'=============
'- keine

'TO-DOs
'======
'- SGTool: "Output-Puffer" einbauen (siehe dort)
'- SGTool: Fortschrittsbalken optimieren
'- "Rückgängig" optimieren
'- Button fürBegrenzung wenn ganze Spalten markiert sind (Logik siehe xlDuMa)
            
'Auslieferung
    'Excel-->Datei-->Beschreibung (vgl. andere Tools)
    'VBA Projekt Beschreibung (vgl. andere Tools)
    'Worksheets und Code für Entwicklung entfernen
    'xmlUI Ribbon erstellen
    'xlam-Datei erstellen
    'Installer-Datei
    'Code auf GitHub
    'Upload auf Download-Seite
    
'IDEEN
'=====
'VBA-Funktionen für das Tabellenblatt verfügbar machen
    '
'Worksheet-Funktionen grafisch verfügbar machen
    '
'Eigene Funktionen
    'Zellinhalte tauschen (z.B. zwei Zellen oder Bereiche anklicken, dass dann die Inhalte getauscht werden)
    'Führende Nullen weg (evtl. Text zu Zahl, dann gehen die alleine weg?)
    'Verketten: Marco Krapf -> Marco & Krapf -> MarcoKrapf // Mit oder ohne Leerzeichen? 1 Zielzelle für verschiedene Zellen?
    'SVERWEIS vereinfachen,(=SVERWEIS(Suchkriterium,Spalte,Rückgabespalte), sodass Eingangsparater Matrix nicht benötigt wird, da bei großen Tabellen kompliziert Rückgabespalte auszurechnen
    'SVERWEIS mit zwei Bedingungen SSVERWEIS (=SSVERWEIS(Suchkriterium1, Spalte1, Suchkriterium2, Spalte2, Rückgabespalte)
        'Alternative: Spaltennummer in Zwischenablage / auch Zeilennummer, AdressLocal, Worksheet, Workbook...
    '(Alle Zellen hervorheben, die einen bestimmten Inhalt haben (exakt oder zum Teil) -->Spielplatz: "Textstellen suchen")

'WEITERENTWICKLUNG
'=================
'unbedingt machen bei "Zeichen entfernen": ALLE unerwünschten Zeichen mit einem Klick
'Checkboxen zum Anhaken im Ribbon
    'Fortschrittsbalken anzeigen (Prototyp vorhanden)
    'Application.Screenupdate aus/ein  ---> evtl. wieder rückbauen?? geht nicht korrekt
'evtl. Begrenzung wenn ganze Spalte(n) markiert sind?
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
'Ribbon: Kleine Icons vor Dropdowns
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
'Worksheets vergleichen: Wenn ein Bereich selektiert ist, dann Vergleich nur für diesen Bereich (ähnlich wie bei Zeilen/Spalten entfernen)
    
'========================================================================================================

'Globale Konstanten und Variablen deklarieren
Public MyRibbon As IRibbonUI 'Excel-Menüband
Public Const xlef_strToolname As String = "ELPHIN Toolbox" 'Tool-Name
Public xlef_strVersion As String 'Tool-Version
Public Const xlef_strKontakt1 As String = "excel@marco-krapf.de" 'Kontakt-E-Mail-Adresse
Public Const xlef_strGitHub As String = "https://github.com/MarcoKrapf/ELPHIN_Toolbox" 'GitHub-Repository
Public Const xlef_strDownload As String = "https://marco-krapf.de/add-in-elphin-toolbox/" 'Download-Seite
Public Const xlef_strSpendenURL As String = "https://www.grosse-hilfe.de/" 'Spenden-Link
Public xlef_strSprache As Integer 'Sprache DE oder EN '(2 = deutsch, 3 = englisch --> Spalte im Worksheet ELPHIN_Texte)
Public xlef_wksTexte As Worksheet 'Worksheet mit den Textkomponenten
Public xlef_wksTarget As Worksheet 'Worksheet, auf dem die Aktion ausgeführt wird
Public xlef_Sel As Variant 'Selektierter Bereich auf dem Tabellenblatt
Public xlef_art As String 'Objekt auf das die Aktion angewendet wird (für Merken und Wiederherstellen)
Public xlef_rngCell As Range 'Einzelne Zelle in For-Each-Schleifen
Public xlef_arrOrg() As Variant 'Array zum Merken der originalen Zellinhalte
Public xlef_coll As Collection 'Collection zum Merken der leeren Zeilen
Public xlef_objClip As DataObject 'Objekt zum Ablegen von Daten in die Windows-Zwischenablage
Public xlef_SchreibschutzEmpfehlen As Boolean
Public xlef_blnScreenUpdate As Boolean 'Bildschirmaktualisierung während Aktion ausgeführt wird?
Public xlef_blnProgressBar As Boolean 'Fortschrittsbalken während Aktion ausgeführt wird?
Public xlef_dblBalkenAnteil As Double 'Stückelung des Fortschrittsbalkens
Public xlef_dblBalkenAktuell As Double 'Aktuelle Breite des Fortschrittsbalkens4
Public xlex_blnUNDO As Boolean 'Rückgängig-Button aktiv wenn true
Public xlef_blnDo As Boolean 'Aktion wird nur ausgeführt wenn true
