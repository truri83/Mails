Option Compare Database
Option Explicit

' ===========================================================================
' modSchemaTools - Intelligente DDL-Werkzeuge fuer Schema-Verwaltung
' ===========================================================================
' Zentrale Helfer fuer sicheres Erstellen und Aendern von Tabellen, Spalten,
' Indizes und Defaults. Beruecksichtigt automatisch ob eine Tabelle lokal
' oder als Linked Table im Backend liegt.
'
' KERNKONZEPT:
'   RunDDL_Smart()  erkennt ob Tabelle lokal oder verlinkt ist und fuehrt
'                   DDL auf der richtigen Datenbank aus.
'                   -> Kein manuelles "ist das FE oder BE?" noetig!
'
' Access DDL-Restriktionen (automatisch beruecksichtigt):
'   - Kein DEFAULT in CREATE TABLE  -> SetzeDefaultWert()
'   - Kein IF NOT EXISTS            -> TabelleExistiert() / FeldExistiert()
'   - Kein ALTER COLUMN             -> Spalte nur adden, nicht aendern
'
' Abhaengigkeiten: modGlobals (Konstanten), modLogging, modBackend (IstFETabelle)
'
' Aufruf-Beispiele:
'   EnsureColumn "tblEmails", "NeuesFeld", "TEXT(100)"
'   EnsureIndex "tblEmails", "idx_Email_Neu", "NeuesFeld"
'   EnsureUniqueIndex "tblEmails", "idx_Email_Unique", "NeuesFeld"
'   SetzeDefaultWert db, "tblEmails", "NeuesFeld", "''"
'   RunDDL_Smart "tblEmails", "ALTER TABLE [tblEmails] ADD COLUMN ..."
' ===========================================================================


' ---------------------------------------------------------------------------
' TABELLEN-PRUEFUNG
' ---------------------------------------------------------------------------

' Prueft ob eine Tabelle in der aktuellen DB (lokal oder linked) existiert
Public Function TabelleExistiert(ByVal strName As String) As Boolean
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Set db = CurrentDb
    TabelleExistiert = False
    For Each td In db.TableDefs
        If td.Name = strName Then
            TabelleExistiert = True
            Exit For
        End If
    Next td
    Set db = Nothing
End Function

' Prueft ob eine Tabelle eine verlinkte Backend-Tabelle ist
Public Function IstLinkedTable(ByVal strTabelle As String) As Boolean
    On Error Resume Next
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Set db = CurrentDb
    Set td = db.TableDefs(strTabelle)
    If Err.Number <> 0 Then
        IstLinkedTable = False
        Set db = Nothing
        Exit Function
    End If
    IstLinkedTable = (Len(Nz(td.Connect, "")) > 0)
    Set db = Nothing
    On Error GoTo 0
End Function

' Prueft ob eine Tabelle lokal (nicht verlinkt) ist
Public Function IstLokaleTabelle(ByVal strTabelle As String) As Boolean
    If Not TabelleExistiert(strTabelle) Then
        IstLokaleTabelle = False
        Exit Function
    End If
    IstLokaleTabelle = Not IstLinkedTable(strTabelle)
End Function


' ---------------------------------------------------------------------------
' FELD-PRUEFUNG
' ---------------------------------------------------------------------------

' Prueft ob ein Feld in einer Tabelle existiert (lokal oder im Backend)
Public Function FeldExistiert(ByVal strTabelle As String, ByVal strFeld As String) As Boolean
    On Error GoTo FallbackBackend

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim strTest As String

    Set db = CurrentDb
    Set td = db.TableDefs(strTabelle)
    strTest = td.Fields(strFeld).Name

    FeldExistiert = True
    Set td = Nothing
    Set db = Nothing
    Exit Function

FallbackBackend:
    On Error GoTo ErrExit

    If Not IstLinkedTable(strTabelle) Then
        FeldExistiert = False
        Exit Function
    End If

    Dim dbBE As DAO.Database
    Dim strSource As String
    strSource = LinkedSourceTableName(strTabelle)

    Set dbBE = HoleBackendDB(strTabelle)
    If dbBE Is Nothing Then
        FeldExistiert = False
        Exit Function
    End If

    Err.Clear
    strTest = dbBE.TableDefs(strSource).Fields(strFeld).Name
    FeldExistiert = (Err.Number = 0)

    dbBE.Close: Set dbBE = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Function

ErrExit:
    FeldExistiert = False
    On Error Resume Next
    Set td = Nothing
    Set db = Nothing
    If Not dbBE Is Nothing Then dbBE.Close: Set dbBE = Nothing
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' INDEX-PRUEFUNG
' ---------------------------------------------------------------------------

' Prueft ob ein Index in einer Tabelle existiert (lokal oder im Backend)
Public Function IndexExistiert(ByVal strTabelle As String, ByVal strIndex As String) As Boolean
    On Error GoTo ErrExit

    If IstLinkedTable(strTabelle) Then
        Dim dbLink As DAO.Database
        Dim tdfLink As DAO.TableDef
        Dim idxLink As DAO.Index
        Dim dbBE As DAO.Database
        Dim tdfBE As DAO.TableDef
        Dim idxBE As DAO.Index
        Dim strSource As String

        Set dbLink = CurrentDb
        Set tdfLink = dbLink.TableDefs(strTabelle)
        For Each idxLink In tdfLink.Indexes
            If StrComp(idxLink.Name, strIndex, vbTextCompare) = 0 Then
                IndexExistiert = True
                Set tdfLink = Nothing
                Set dbLink = Nothing
                Exit Function
            End If
        Next idxLink

        strSource = LinkedSourceTableName(strTabelle)
        Set dbBE = HoleBackendDB(strTabelle)
        If dbBE Is Nothing Then
            IndexExistiert = False
            Exit Function
        End If

        Set tdfBE = dbBE.TableDefs(strSource)
        For Each idxBE In tdfBE.Indexes
            If StrComp(idxBE.Name, strIndex, vbTextCompare) = 0 Then
                IndexExistiert = True
                Set tdfLink = Nothing
                Set dbLink = Nothing
                dbBE.Close: Set dbBE = Nothing
                Exit Function
            End If
        Next idxBE

        IndexExistiert = False
        Set tdfLink = Nothing
        Set dbLink = Nothing
        dbBE.Close: Set dbBE = Nothing
        Exit Function
    End If

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index

    ' Lokal: db-Variable halten damit tdf.Indexes zuverlaessig funktioniert
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTabelle)

    For Each idx In tdf.Indexes
        If StrComp(idx.Name, strIndex, vbTextCompare) = 0 Then
            IndexExistiert = True
            Set db = Nothing
            Exit Function
        End If
    Next idx

    IndexExistiert = False
    Set db = Nothing
    Exit Function

ErrExit:
    IndexExistiert = False
    On Error Resume Next
    Set tdfLink = Nothing
    Set dbLink = Nothing
    If Not db Is Nothing Then
        Set db = Nothing
    End If
    If Not dbBE Is Nothing Then dbBE.Close: Set dbBE = Nothing
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' SMART DDL: Fuehrt SQL auf der richtigen DB aus (lokal oder Backend)
' ---------------------------------------------------------------------------

' Erkennt ob strTabelle lokal oder verlinkt ist,
' oeffnet bei Bedarf die Backend-DB und fuehrt SQL dort aus.
' Bei verknuepften Tabellen wird danach RefreshLink aufgerufen.
'
' VERWENDUNG:
'   RunDDL_Smart "tblEmails", "ALTER TABLE [tblEmails] ADD COLUMN [Feld] TEXT(50)"
'   RunDDL_Smart "tblEmails", "CREATE INDEX [idx_X] ON [tblEmails] ([Feld])"
'   RunDDL_Smart "tblConfig", "CREATE TABLE [tblConfig] (...)"  -> lokal
Public Sub RunDDL_Smart(ByVal strTabelle As String, ByVal strSql As String)
    On Error GoTo ErrHandler

    LogDebug "DDL: " & Left$(strSql, 120), "SCHEMA"

    If IstLinkedTable(strTabelle) Then
        ' Backend: SQL in physischer DB ausfuehren
        Dim dbBE As DAO.Database
        Set dbBE = HoleBackendDB(strTabelle)
        If dbBE Is Nothing Then
            LogError "DDL fehlgeschlagen - Backend nicht erreichbar fuer: " & strTabelle, "SCHEMA"
            Exit Sub
        End If

        dbBE.Execute strSql, dbFailOnError
        dbBE.Close: Set dbBE = Nothing

        ' Link aktualisieren damit Access die Aenderung sieht
        On Error Resume Next
        CurrentDb.TableDefs(strTabelle).RefreshLink
        On Error GoTo 0
    Else
        ' Lokal: direkt ausfuehren
        CurrentDb.Execute strSql, dbFailOnError
    End If

    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 3380  ' Feld existiert bereits
            LogDebug "DDL: Feld existiert bereits (" & strTabelle & ")", "SCHEMA"
        Case 3283  ' Index existiert bereits
            LogDebug "DDL: Index existiert bereits (" & strTabelle & ")", "SCHEMA"
        Case 3284  ' Index existiert bereits (Variante)
            LogDebug "DDL: Index existiert bereits (" & strTabelle & ")", "SCHEMA"
        Case Else
            LogError "DDL Error " & Err.Number & ": " & Err.Description & _
                     " | SQL: " & Left$(strSql, 200), "SCHEMA"
    End Select
End Sub


' ---------------------------------------------------------------------------
' ENSURE-FUNKTIONEN: Idempotent (Skip wenn schon vorhanden)
' ---------------------------------------------------------------------------

' Fuegt eine Spalte hinzu, wenn sie noch nicht existiert.
' Funktioniert fuer lokale UND verknuepfte Tabellen.
'
' Beispiel:
'   EnsureColumn "tblEmails", "NeuesFeld", "TEXT(100)"
'   EnsureColumn "tblEmails", "Counter", "LONG"
Public Sub EnsureColumn(ByVal strTabelle As String, ByVal strFeld As String, _
                        ByVal strTypSQL As String)
    If Not TabelleExistiert(strTabelle) Then
        LogWarn "EnsureColumn: Tabelle existiert nicht: " & strTabelle, "SCHEMA"
        Exit Sub
    End If

    If FeldExistiert(strTabelle, strFeld) Then Exit Sub

    LogDebug "  + Spalte: " & strTabelle & "." & strFeld & " (" & strTypSQL & ")", "SCHEMA"
    RunDDL_Smart strTabelle, "ALTER TABLE [" & strTabelle & "] ADD COLUMN [" & strFeld & "] " & strTypSQL
End Sub


' Erstellt einen Index, wenn er noch nicht existiert.
' strFelder: Komma-getrennt fuer Multi-Column-Index
'
' Beispiel:
'   EnsureIndex "tblEmails", "idx_Email_Datum", "EmpfangenAm"
'   EnsureIndex "tblEmails", "idx_Multi", "OrdnerID,ThreadID"
Public Sub EnsureIndex(ByVal strTabelle As String, ByVal strIndexName As String, _
                       ByVal strFelder As String)
    If Not TabelleExistiert(strTabelle) Then Exit Sub
    If IndexExistiert(strTabelle, strIndexName) Then Exit Sub

    Dim strFeldListe As String
    strFeldListe = FormatiereFeldListe(strFelder)

    LogDebug "  + Index: " & strIndexName & " ON " & strTabelle & " (" & strFelder & ")", "SCHEMA"
    RunDDL_Smart strTabelle, "CREATE INDEX [" & strIndexName & "] ON [" & strTabelle & "] (" & strFeldListe & ")"
End Sub


' Erstellt einen UNIQUE Index, wenn er noch nicht existiert.
Public Sub EnsureUniqueIndex(ByVal strTabelle As String, ByVal strIndexName As String, _
                             ByVal strFelder As String)
    If Not TabelleExistiert(strTabelle) Then Exit Sub
    If IndexExistiert(strTabelle, strIndexName) Then Exit Sub

    Dim strFeldListe As String
    strFeldListe = FormatiereFeldListe(strFelder)

    LogDebug "  + Unique Index: " & strIndexName & " ON " & strTabelle & " (" & strFelder & ")", "SCHEMA"
    RunDDL_Smart strTabelle, "CREATE UNIQUE INDEX [" & strIndexName & "] ON [" & strTabelle & "] (" & strFeldListe & ")"
End Sub


' Setzt den Default-Wert eines Feldes per DAO (Access DDL kennt kein DEFAULT)
' Funktioniert fuer lokale UND verknuepfte Tabellen.
'
' strDefault-Formate:
'   Numerisch:  "0", "1", "500"
'   Text:       "'Gestartet'" (mit einfachen Quotes!)
'   Boolean:    "True" / "False"
'   Datum:      "=Now()"
'
' Beispiel:
'   SetzeDefaultSmart "tblEmails", "Status", "'Neu'"
'   SetzeDefaultSmart "tblEmails", "Groesse", "0"
Public Sub SetzeDefaultSmart(ByVal strTabelle As String, ByVal strFeld As String, _
                             ByVal strDefault As String)
    On Error GoTo ErrHandler

    If Not TabelleExistiert(strTabelle) Then Exit Sub

    If IstLinkedTable(strTabelle) Then
        Dim dbBE As DAO.Database
        Dim strSource As String
        strSource = LinkedSourceTableName(strTabelle)
        Set dbBE = HoleBackendDB(strTabelle)
        If dbBE Is Nothing Then Exit Sub
        dbBE.TableDefs(strSource).Fields(strFeld).DefaultValue = strDefault
        dbBE.Close: Set dbBE = Nothing
    Else
        CurrentDb.TableDefs(strTabelle).Fields(strFeld).DefaultValue = strDefault
    End If
    Exit Sub

ErrHandler:
    LogWarn "SetzeDefaultSmart: " & strTabelle & "." & strFeld & " = " & strDefault & _
            " -> " & Err.Description, "SCHEMA"
End Sub


' ---------------------------------------------------------------------------
' DROP-FUNKTIONEN
' ---------------------------------------------------------------------------

' Loescht eine Tabelle (lokal oder Backend). Tut nichts wenn nicht vorhanden.
Public Sub DropTableSmart(ByVal strTabelle As String)
    If Not TabelleExistiert(strTabelle) Then Exit Sub

    On Error GoTo ErrHandler

    If IstLinkedTable(strTabelle) Then
        ' Backend: Tabelle in physischer DB loeschen + Link entfernen
        Dim dbBE As DAO.Database
        Set dbBE = HoleBackendDB(strTabelle)
        If Not dbBE Is Nothing Then
            dbBE.Execute "DROP TABLE [" & strTabelle & "]", dbFailOnError
            dbBE.Close: Set dbBE = Nothing
        End If
        ' Link im Frontend entfernen
        CurrentDb.TableDefs.Delete strTabelle
        CurrentDb.TableDefs.Refresh
    Else
        ' Lokal: direkt loeschen
        CurrentDb.Execute "DROP TABLE [" & strTabelle & "]"
    End If

    LogDebug "  - DROP TABLE: " & strTabelle, "SCHEMA"
    Exit Sub

ErrHandler:
    LogWarn "DropTableSmart fehlgeschlagen: " & strTabelle & " - " & Err.Description, "SCHEMA"
End Sub


' ---------------------------------------------------------------------------
' SCHEMA-MIGRATION: Tabelle um fehlende Spalten/Indizes erweitern
' ---------------------------------------------------------------------------

' Stellt sicher, dass eine Tabelle alle erwarteten Spalten hat.
' arrSpalten: Array von "Feldname|SQL_TYPE" Strings
'
' Beispiel:
'   EnsureColumns "tblEmails", Array("NeuesFeld|TEXT(100)", "Counter|LONG")
Public Sub EnsureColumns(ByVal strTabelle As String, arrSpalten As Variant)
    If Not IsArray(arrSpalten) Then Exit Sub
    If Not TabelleExistiert(strTabelle) Then Exit Sub

    Dim i As Long
    Dim parts() As String

    For i = LBound(arrSpalten) To UBound(arrSpalten)
        parts = Split(CStr(arrSpalten(i)), "|")
        If UBound(parts) >= 1 Then
            EnsureColumn strTabelle, Trim$(parts(0)), Trim$(parts(1))
        End If
    Next i
End Sub


' Stellt sicher, dass alle angegebenen Defaults gesetzt sind.
' arrDefaults: Array von "Feldname|DefaultWert" Strings
'
' Beispiel:
'   EnsureDefaults "tblEmails", Array("Status|'Neu'", "Groesse|0")
Public Sub EnsureDefaults(ByVal strTabelle As String, arrDefaults As Variant)
    If Not IsArray(arrDefaults) Then Exit Sub
    If Not TabelleExistiert(strTabelle) Then Exit Sub

    Dim i As Long
    Dim parts() As String

    For i = LBound(arrDefaults) To UBound(arrDefaults)
        parts = Split(CStr(arrDefaults(i)), "|")
        If UBound(parts) >= 1 Then
            SetzeDefaultSmart strTabelle, Trim$(parts(0)), Trim$(parts(1))
        End If
    Next i
End Sub


' Stellt sicher, dass alle angegebenen Indizes existieren.
' arrIndizes: Array von "IndexName|Feld[,Feld2]" Strings
' blnUnique: True -> UNIQUE Index
'
' Beispiel:
'   EnsureIndexes "tblEmails", Array("idx_Hash|UniqueHash", "idx_Multi|OrdnerID,ThreadID"), True
Public Sub EnsureIndexes(ByVal strTabelle As String, arrIndizes As Variant, _
                         Optional ByVal blnUnique As Boolean = False)
    If Not IsArray(arrIndizes) Then Exit Sub
    If Not TabelleExistiert(strTabelle) Then Exit Sub

    Dim i As Long
    Dim parts() As String

    For i = LBound(arrIndizes) To UBound(arrIndizes)
        parts = Split(CStr(arrIndizes(i)), "|")
        If UBound(parts) >= 1 Then
            If blnUnique Then
                EnsureUniqueIndex strTabelle, Trim$(parts(0)), Trim$(parts(1))
            Else
                EnsureIndex strTabelle, Trim$(parts(0)), Trim$(parts(1))
            End If
        End If
    Next i
End Sub


' ---------------------------------------------------------------------------
' SCHEMA-INFO
' ---------------------------------------------------------------------------

' Gibt die Anzahl Felder einer Tabelle zurueck
Public Function AnzahlFelder(ByVal strTabelle As String) As Long
    On Error Resume Next
    If IstLinkedTable(strTabelle) Then
        Dim dbBE As DAO.Database
        Dim strSource As String
        strSource = LinkedSourceTableName(strTabelle)
        Set dbBE = HoleBackendDB(strTabelle)
        If Not dbBE Is Nothing Then
            AnzahlFelder = dbBE.TableDefs(strSource).Fields.Count
            dbBE.Close: Set dbBE = Nothing
        End If
    Else
        AnzahlFelder = CurrentDb.TableDefs(strTabelle).Fields.Count
    End If
    If Err.Number <> 0 Then AnzahlFelder = 0: Err.Clear
    On Error GoTo 0
End Function

' Gibt den Connect-String einer verlinkten Tabelle zurueck (leer wenn lokal)
Public Function GetConnectString(ByVal strTabelle As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    GetConnectString = Nz(db.TableDefs(strTabelle).Connect, "")
    If Err.Number <> 0 Then GetConnectString = "": Err.Clear
    Set db = Nothing
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' HILFSFUNKTIONEN (Private)
' ---------------------------------------------------------------------------

' Oeffnet die physische Backend-DB einer verlinkten Tabelle
Private Function HoleBackendDB(ByVal strTabelle As String) As DAO.Database
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Set db = CurrentDb
    Set td = db.TableDefs(strTabelle)

    Dim strConnect As String
    strConnect = Nz(td.Connect, "")
    If strConnect = "" Then
        Set HoleBackendDB = Nothing
        Exit Function
    End If

    ' Connect-String parsen: ";DATABASE=\\Server\Share\Backend.accdb"
    Dim strPfad As String
    Dim lngPos As Long
    lngPos = InStr(1, strConnect, "DATABASE=", vbTextCompare)
    If lngPos = 0 Then
        Set HoleBackendDB = Nothing
        Exit Function
    End If

    strPfad = Mid$(strConnect, lngPos + 9)
    If InStr(strPfad, ";") > 0 Then strPfad = Left$(strPfad, InStr(strPfad, ";") - 1)
    strPfad = Trim$(strPfad)

    Set HoleBackendDB = DBEngine.OpenDatabase(strPfad)
    Set td = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    LogWarn "HoleBackendDB: Kann Backend nicht oeffnen fuer " & strTabelle & _
            " - " & Err.Description, "SCHEMA"
    Set td = Nothing
    Set db = Nothing
    Set HoleBackendDB = Nothing
End Function

Private Function LinkedSourceTableName(ByVal strTabelle As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim strSource As String

    Set db = CurrentDb
    Set td = db.TableDefs(strTabelle)
    strSource = Nz(td.SourceTableName, "")
    If strSource = "" Then strSource = strTabelle

    LinkedSourceTableName = strSource
    Set td = Nothing
    Set db = Nothing
    On Error GoTo 0
End Function


' Formatiert Komma-getrennte Feldliste in [Feld1],[Feld2] Format
Private Function FormatiereFeldListe(ByVal strFelder As String) As String
    Dim parts() As String
    parts = Split(strFelder, ",")

    Dim i As Long, strResult As String
    For i = LBound(parts) To UBound(parts)
        If Len(strResult) > 0 Then strResult = strResult & ","
        strResult = strResult & "[" & Trim$(parts(i)) & "]"
    Next i

    FormatiereFeldListe = strResult
End Function


