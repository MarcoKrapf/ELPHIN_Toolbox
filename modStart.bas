Attribute VB_Name = "modStart"
Option Explicit
Option Private Module

'*****************************************************************************************
'*** Das Excel Add-in 'ELPHIN Toolbox' f�gt eine neue Registerkarte in das Ribbon ein ****
'*** mit verschiedenen Funktionen. Teilweise werden Worksheet-Funktionen direkt auf ******
'*** eine markierte Zelle angewandt, teilweise werden VBA-Funktionen f�r das Worksheet ***
'*** verf�gbar gemacht, und teilweise werden neue Funktionen zur Verf�gung gestellt. *****
'***                                                                                   ***
'*** Version 2.0 (August 2018) - Excel 2010-2016 (optimiert f�r Windows)               ***
'*** Autoren: Juliane Held / Marco Krapf - mailto:excel@marco-krapf.de                 ***
'*****************************************************************************************

'FEATURES IMPLEMENTIERT
'======================
    'VBA-Funktionen:
        'UCASE (Konvertieren in Gro�buchstaben) - als Value und als Funktion
        'LCASE (Konvertieren in Kleinbuchstaben) - als Value und als Funktion
        'LTRIM (Leerzeichen am Anfang der Zelle entfernen) - als Value
        'RTRIM (Leerzeichen am Ende der Zelle entfernen) - als Value
        'TRIM (Leerzeichen am Anfang und Ende der Zelle entfernen) - als Value
        'ABS (Absoluter Wert einer Zahl, also Zahl ohne Vorzeichen) - als Value und als Funktion
    'Worksheet-Funktionen:
        '=PROPER (=GROSS2 --> den ersten Buchstabe jedes Wortes gro�schreiben) - als Value und als Funktion
        '=TRIM (=GL�TTEN --> Leerzeichen am Anfang und Ende der Zelle und mehrfache mittendrin entfernen) - als Value und als Funktion
    'Eigene Funktionen:
        'Leere Zeilen l�schen
        'Leere Spalten l�schen
        'Auf Steuerzeichen pr�fen
        'Steuerzeichen entfernen (im Endeffekt Worksheet-Funktion =CLEAN (=S�UBERN)) - als Value
        'Auf gesch�tzte Leerzeichen pr�fen
        'Gesch�tzte Leerzeichen durch normale ersetzen - als Value
        'Formel/Funktion durch Wert ersetzen - als Value
        'Spaltennummer in Zwischenablage ablegen
        'Zeilennummer in Zwischenablage ablegen
        'Wenn Formel/Funktion 0 ergibt, dann Zelle leer darstellen -> in Funktion einbauen
        'Worksheet-Vergleich: 2 Tabellen auf unterschiedliche Werte vergleichen und Unterschiede farblich markieren
    'Eigene Tools:
        'SG DataSet Finder - Tool zum Finden von mehrfach vorkommenden Datens�tzen in verschiedenen Dateien
    'Letzten Schritt r�ckg�ngig machen

    
'KOMMENTARE 10.04.2018
'=====================
'- (erledigt) Dropdown-Problem: Ausprobieren mit "menu" anstelle "dropDown" item // oder menu mit buttons machen??!
'- (erledigt) Ribbon zweisprachig (EN/DE) machen
'- (erledigt) QR-Code rausmachen (veraltet)

'KOMMENTARE 24.07.2018
'=====================
'- (erledigt) Dropdown-Problem gel�st durch "menu" mit buttons anstelle "dropDown" item
'- (erledigt) Mehrsprachigkeit vorbereitet durch Callbacks, die Texte aus dem Worksheet "ELPHIN_Texte" ziehen

'KOMMENTARE 05.08.2018
'=====================
'- (erledigt) SG DataSet Finder (HHN-Tool) eingebaut
'- (erledigt) UserForms werden beim Aufruf zentriert platziert
'- (erledigt) Popup "Info" aktualisiert
'- (erledigt) EN-�bersetzungen aller Texte

'KOMMENTARE 10.08.2018
'=====================
'- R�ckg�ngig-Button optional
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
'- "R�ckg�ngig" optimieren
'- Button f�rBegrenzung wenn ganze Spalten markiert sind (Logik siehe xlDuMa)
            
'Auslieferung
    'Excel-->Datei-->Beschreibung (vgl. andere Tools)
    'VBA Projekt Beschreibung (vgl. andere Tools)
    'Worksheets und Code f�r Entwicklung entfernen
    'xmlUI Ribbon erstellen
    'xlam-Datei erstellen
    'Installer-Datei
    'Code auf GitHub
    'Upload auf Download-Seite
    
'IDEEN
'=====
'VBA-Funktionen f�r das Tabellenblatt verf�gbar machen
    '
'Worksheet-Funktionen grafisch verf�gbar machen
    '
'Eigene Funktionen
    'Zellinhalte tauschen (z.B. zwei Zellen oder Bereiche anklicken, dass dann die Inhalte getauscht werden)
    'F�hrende Nullen weg (evtl. Text zu Zahl, dann gehen die alleine weg?)
    'Verketten: Marco Krapf -> Marco & Krapf -> MarcoKrapf // Mit oder ohne Leerzeichen? 1 Zielzelle f�r verschiedene Zellen?
    'SVERWEIS vereinfachen,(=SVERWEIS(Suchkriterium,Spalte,R�ckgabespalte), sodass Eingangsparater Matrix nicht ben�tigt wird, da bei gro�en Tabellen kompliziert R�ckgabespalte auszurechnen
    'SVERWEIS mit zwei Bedingungen SSVERWEIS (=SSVERWEIS(Suchkriterium1, Spalte1, Suchkriterium2, Spalte2, R�ckgabespalte)
        'Alternative: Spaltennummer in Zwischenablage / auch Zeilennummer, AdressLocal, Worksheet, Workbook...
    '(Alle Zellen hervorheben, die einen bestimmten Inhalt haben (exakt oder zum Teil) -->Spielplatz: "Textstellen suchen")

'WEITERENTWICKLUNG
'=================
'unbedingt machen bei "Zeichen entfernen": ALLE unerw�nschten Zeichen mit einem Klick
'Checkboxen zum Anhaken im Ribbon
    'Fortschrittsbalken anzeigen (Prototyp vorhanden)
    'Application.Screenupdate aus/ein  ---> evtl. wieder r�ckbauen?? geht nicht korrekt
'evtl. Begrenzung wenn ganze Spalte(n) markiert sind?
'Sind unsere Verweise alle Standard oder m�ssen wir die dynamisch einbauen?
'Empfehlung f�r Schreibschutz aktivieren/deaktivieren (Logik passt nicht)
'Mehrere Schritte r�ckw�rts (z.B. das Array mit den Wiederherstellungsinfos in eine Collection packen
'Wenn 0 als Wert in Zelle, dann --> "" oder IF(Zelle.Value=0;"";Zelle.Value)
'Funktionen auch in Worksheet-Kontextmen� (Rechtsklick) einbauen --> Reset des Kontextmen�s bei WorkbookClose? // oder http://www.rholtz-office.de/ribbonx/das-kontextmenue
'Info �ber die Zelle (Adresse, Inhalt, Datentyp, HasFormula...)
'Vergleich, erkennt Text/Zahlen auch in anderer Reihenfolge am besten eine Spalte mit einer anderen Spalte in Tabelle vergleichen (nach Spaltenindex fragen)
'Worksheets vergleichen --> nicht nur Werte, auch Formeln oder beides
'Toolsprache EN zus�tzlich (Ribbon und Formulare)
'Selektierten Bereich in gleichem Format in neue E-Mail kopieren ---> ist schon 90% fertig
'Screenupdating ein/ausschalten bei allen Funktionen? --> Performance // Optional per ToggleButton?
'Dynamischer Tooltip bei "R�ckg�ngig" --> Name der zuvor ausgef�hrten Aktion
'Ribbon: Kleine Icons vor Dropdowns
'Dynamisches Ribbon
    'Sprache anpassen
    'Anzeige von Zeilen-/Spaltennummer, Adresse usw.
'Funktion 0 = "" auch wenn Ergebnis sich sp�ter �ndert und dann 0 wird
'Ordentliche Fehlermeldungen wenn Hyperlink oder Mail Fehler wirft (aktuell wird nur �bersprungen)
'Info-Popup wenn Worksheetvergleich angeklickt und nur 1 Worksheet da
'Info-Popup wenn Steuerzeichen pr�fen / Gesch. Leerzeichen pr�fen keine Funde gibt --> "sauber"
'Zeilen l�schen, die "..." enthalten --> darf Juli machen, grobe Orientierung an Sub fnLEEREZEILEN()
    'Bedingte Zeilenl�schung (kompletter Wert --> value) in gesamter Tabelle/Selektion
    'Bedingte Zeilenl�schung ("Teilwert" --> instring) in gesamter Tabelle/Selektion
    'Bedingte Spaltenl�schung (kompletter Wert --> value) in gesamter Tabelle/Selektion
    'Bedingte Spaltenl�schung ("Teilwert" --> instring) in gesamter Tabelle/Selektion
'Checkbox: Passwortschutz beim Speichern
'Worksheets vergleichen --> UserForm schick machen, auch mit Icon
'Worksheets vergleichen --> Tooltips
'Worksheets vergleichen: Wenn ein Bereich selektiert ist, dann Vergleich nur f�r diesen Bereich (�hnlich wie bei Zeilen/Spalten entfernen)
    
'========================================================================================================

'Globale Konstanten und Variablen deklarieren
Public MyRibbon As IRibbonUI 'Excel-Men�band
Public Const xlef_strToolname As String = "ELPHIN Toolbox" 'Tool-Name
Public xlef_strVersion As String 'Tool-Version
Public Const xlef_strKontakt1 As String = "excel@marco-krapf.de" 'Kontakt-E-Mail-Adresse
Public Const xlef_strGitHub As String = "https://github.com/MarcoKrapf/ELPHIN_Toolbox" 'GitHub-Repository
Public Const xlef_strDownload As String = "https://marco-krapf.de/add-in-elphin-toolbox/" 'Download-Seite
Public Const xlef_strSpendenURL As String = "https://www.grosse-hilfe.de/" 'Spenden-Link
Public xlef_strSprache As Integer 'Sprache DE oder EN '(2 = deutsch, 3 = englisch --> Spalte im Worksheet ELPHIN_Texte)
Public xlef_wksTexte As Worksheet 'Worksheet mit den Textkomponenten
Public xlef_wksTarget As Worksheet 'Worksheet, auf dem die Aktion ausgef�hrt wird
Public xlef_Sel As Variant 'Selektierter Bereich auf dem Tabellenblatt
Public xlef_art As String 'Objekt auf das die Aktion angewendet wird (f�r Merken und Wiederherstellen)
Public xlef_rngCell As Range 'Einzelne Zelle in For-Each-Schleifen
Public xlef_arrOrg() As Variant 'Array zum Merken der originalen Zellinhalte
Public xlef_coll As Collection 'Collection zum Merken der leeren Zeilen
Public xlef_objClip As DataObject 'Objekt zum Ablegen von Daten in die Windows-Zwischenablage
Public xlef_SchreibschutzEmpfehlen As Boolean
Public xlef_blnScreenUpdate As Boolean 'Bildschirmaktualisierung w�hrend Aktion ausgef�hrt wird?
Public xlef_blnProgressBar As Boolean 'Fortschrittsbalken w�hrend Aktion ausgef�hrt wird?
Public xlef_dblBalkenAnteil As Double 'St�ckelung des Fortschrittsbalkens
Public xlef_dblBalkenAktuell As Double 'Aktuelle Breite des Fortschrittsbalkens4
Public xlex_blnUNDO As Boolean 'R�ckg�ngig-Button aktiv wenn true
Public xlef_blnDo As Boolean 'Aktion wird nur ausgef�hrt wenn true
