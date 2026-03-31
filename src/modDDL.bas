Attribute VB_Name = "modDDL"
Option Compare Database
Option Explicit

' ===========================================================================
' modDDL - Zentrales DDL- und Schema-Basismodul (Backend-transparent)
' ===========================================================================
' Alle Tabellen-/Spalten-/Index-Operationen laufen ueber dieses Modul.
' Es erkennt automatisch ob eine Tabelle lokal oder im Backend liegt und
' fuehrt DDL/DAO-Operationen gegen die richtige Datenbank aus.
'
' DESIGN-PRINZIPIEN:
'   - Jede Funktion ist idempotent (mehrfach aufrufbar, kein Fehler)
'   - Linked Tables werden transparent behandelt
'   - DefaultValue wird via DAO gesetzt (da DDL ALTER bei bestehenden
'     Feldern in Access nicht funktioniert)
'   - Alle Operationen sind failsafe mit Logging
'
' OEFFENTLICHE API:
'   Existenz-Pruefungen:
'     DDL_TabelleExistiert(strTabelle)         -> Boolean
'     DDL_FeldExistiert(strTabelle, strFeld)    -> Boolean
'     DDL_IndexExistiert(strTabelle, strIndex)  -> Boolean
'     DDL_IstVerknuepft(strTabelle)             -> Boolean
'
'   Tabellen-Operationen:
'     DDL_ErstelleTabelle(strTabelle, strSQL)   -> Boolean
'     DDL_LoescheTabelle(strTabelle)            -> Boolean
'
'   Spalten-Operationen:
'     DDL_SichereSpalte(strTabelle, strFeld, strTypSQL)
'     DDL_SetzeFeldDefault(strTabelle, strFeld, strDefault)
'     DDL_SichereJaNeinSpalte(strTabelle, strFeld)
'
'   Index-Operationen:
'     DDL_SichererIndex(strTabelle, strIndex, strFelder, [bUnique])
'
'   DDL-Ausfuehrung (generisch):
'     DDL_Ausfuehren(strTabelle, strSQL)        -> Boolean
'
' ABHAENGIGKEITEN: modLogging (LogInfo, LogWarn, LogDebug)
' ===========================================================================

Private Const MODUL_NAME As String = "modDDL"


' ===========================================================================
' EXISTENZ-PRUEFUNGEN
' ===========================================================================

' Prueft ob eine Tabelle existiert (lokal oder als Link)
Public Function DDL_TabelleExistiert(ByVal strTabelle As String) As Boolean
    On Error Resume Next
    Dim s As String
    s = CurrentDb.TableDefs(strTabelle).Name
    DDL_TabelleExistiert = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Prueft ob ein Feld existiert (geht bei Linked Tables ueber die echte Backend-DB)
Public Function DDL_FeldExistiert(ByVal strTabelle As String, ByVal strFeld As String) As Boolean
    On Error GoTo Fehlschlag
    Dim tdf As DAO.TableDef
    Set tdf = HoleEchteTableDef(strTabelle)
    If tdf Is Nothing Then
        DDL_FeldExistiert = False
        Exit Function
    End If

    Dim s As String
    On Error Resume Next
    s = tdf.Fields(strFeld).Name
    DDL_FeldExistiert = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    Exit Function
Fehlschlag:
    DDL_FeldExistiert = False
End Function

' Prueft ob ein Index existiert (bei Linked Tables gegen die Backend-DB)
Public Function DDL_IndexExistiert(ByVal strTabelle As String, ByVal strIndex As String) As Boolean
    On Error GoTo Fehlschlag
    Dim tdf As DAO.TableDef
    Set tdf = HoleEchteTableDef(strTabelle)
    If tdf Is Nothing Then
        DDL_IndexExistiert = False
        Exit Function
    End If

    Dim idx As DAO.Index
    For Each idx In tdf.Indexes
        If StrComp(idx.Name, strIndex, vbTextCompare) = 0 Then
            DDL_IndexExistiert = True
            Exit Function
        End If
    Next idx
    DDL_IndexExistiert = False
    Exit Function
Fehlschlag:
    DDL_IndexExistiert = False
End Function

' Prueft ob Tabelle eine Linked Table (Backend-Verknuepfung) ist
Public Function DDL_IstVerknuepft(ByVal strTabelle As String) As Boolean
    On Error Resume Next
    Dim strConnect As String
    strConnect = CurrentDb.TableDefs(strTabelle).Connect
    DDL_IstVerknuepft = (Len(Nz(strConnect, "")) > 0)
    Err.Clear
    On Error GoTo 0
End Function


' ===========================================================================
' TABELLEN-OPERATIONEN
' ===========================================================================

' Erstellt Tabelle wenn nicht vorhanden. Gibt True zurueck bei Neuerstellung.
' Bei Linked Tables: KEINE Erstellung (Tabelle muss im Backend existieren).
' Fuer lokale FE-Tabellen: Erstellung via CurrentDb.
Public Function DDL_ErstelleTabelle(ByVal strTabelle As String, ByVal strSQL As String) As Boolean
    If DDL_TabelleExistiert(strTabelle) Then
        LogDebug "  [SKIP] " & strTabelle & " (existiert bereits)", MODUL_NAME
        DDL_ErstelleTabelle = False
        Exit Function
    End If

    On Error GoTo ErrHandler
    CurrentDb.Execute strSQL
    LogDebug "  [OK  ] " & strTabelle & " erstellt", MODUL_NAME
    DDL_ErstelleTabelle = True
    Exit Function
ErrHandler:
    LogWarn "  [FAIL] " & strTabelle & " - " & Err.Description, MODUL_NAME
    DDL_ErstelleTabelle = False
End Function

' Loescht eine Tabelle sicher (Link + Backend-Tabelle oder lokal)
Public Function DDL_LoescheTabelle(ByVal strTabelle As String) As Boolean
    If Not DDL_TabelleExistiert(strTabelle) Then
        DDL_LoescheTabelle = True
        Exit Function
    End If

    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Set db = CurrentDb

    If DDL_IstVerknuepft(strTabelle) Then
        ' Bei Linked Table: Backend-Tabelle loeschen, dann Link entfernen
        Dim strPfad As String
        strPfad = HoleBackendPfad(strTabelle)
        If Len(strPfad) > 0 Then
            Dim dbBE As DAO.Database
            Set dbBE = DBEngine.OpenDatabase(strPfad)
            On Error Resume Next
            dbBE.Execute "DROP TABLE [" & strTabelle & "]"
            On Error GoTo ErrHandler
            dbBE.Close
            Set dbBE = Nothing
        End If
        ' Link aus Frontend entfernen
        db.TableDefs.Delete strTabelle
        db.TableDefs.Refresh
    Else
        ' Lokale Tabelle direkt loeschen
        db.Execute "DROP TABLE [" & strTabelle & "]"
    End If

    Set db = Nothing
    LogDebug "  [DROP] " & strTabelle, MODUL_NAME
    DDL_LoescheTabelle = True
    Exit Function
ErrHandler:
    LogWarn "  [FAIL] DROP " & strTabelle & " - " & Err.Description, MODUL_NAME
    DDL_LoescheTabelle = False
End Function


' ===========================================================================
' SPALTEN-OPERATIONEN
' ===========================================================================

' Fuegt eine Spalte hinzu wenn sie nicht existiert (Backend-transparent)
Public Sub DDL_SichereSpalte(ByVal strTabelle As String, ByVal strFeld As String, ByVal strTypSQL As String)
    If Not DDL_TabelleExistiert(strTabelle) Then Exit Sub
    If DDL_FeldExistiert(strTabelle, strFeld) Then Exit Sub

    Dim strSQL As String
    strSQL = "ALTER TABLE [" & strTabelle & "] ADD COLUMN [" & strFeld & "] " & strTypSQL

    If DDL_Ausfuehren(strTabelle, strSQL) Then
        LogDebug "  [+FLD] " & strTabelle & "." & strFeld, MODUL_NAME
    End If
End Sub

' Setzt DefaultValue direkt via DAO (Backend-transparent)
' WICHTIG: DDL "ALTER TABLE ... ALTER COLUMN ... DEFAULT ..." funktioniert in
'          Access/Jet NICHT fuer bestehende Felder. DefaultValue muss ueber
'          die DAO Field.DefaultValue-Eigenschaft gesetzt werden.
Public Sub DDL_SetzeFeldDefault(ByVal strTabelle As String, ByVal strFeld As String, ByVal strDefault As String)
    On Error GoTo ErrHandler
    If Not DDL_FeldExistiert(strTabelle, strFeld) Then Exit Sub

    Dim tdf As DAO.TableDef
    Dim dbTarget As DAO.Database
    Dim bMussSchliessen As Boolean

    Set dbTarget = HoleZielDatenbank(strTabelle, bMussSchliessen)
    If dbTarget Is Nothing Then Exit Sub

    Set tdf = dbTarget.TableDefs(strTabelle)
    tdf.Fields(strFeld).DefaultValue = strDefault

    ' Bei Linked Table: Link refreshen damit Frontend die Aenderung sieht
    If bMussSchliessen Then
        dbTarget.Close
        Set dbTarget = Nothing
        RefreshTableLink strTabelle
    End If

    LogDebug "  [DFLT] " & strTabelle & "." & strFeld & " = " & strDefault, MODUL_NAME
    Exit Sub
ErrHandler:
    If bMussSchliessen And Not dbTarget Is Nothing Then dbTarget.Close
    LogWarn "  [WARN] Default " & strTabelle & "." & strFeld & " - " & Err.Description, MODUL_NAME
End Sub

' Stellt sicher dass ein YESNO-Feld existiert und Default 0 (Nein) hat
Public Sub DDL_SichereJaNeinSpalte(ByVal strTabelle As String, ByVal strFeld As String)
    DDL_SichereSpalte strTabelle, strFeld, "YESNO"
    DDL_SetzeFeldDefault strTabelle, strFeld, "0"
End Sub


' ===========================================================================
' INDEX-OPERATIONEN
' ===========================================================================

' Erstellt Index falls nicht vorhanden (Backend-transparent)
' strFelder: einzelnes Feld oder kommagetrennte Liste fuer Composite-Index
Public Sub DDL_SichererIndex(ByVal strTabelle As String, ByVal strIndex As String, ByVal strFelder As String, Optional ByVal bUnique As Boolean = False)
    If Not DDL_TabelleExistiert(strTabelle) Then Exit Sub
    If DDL_IndexExistiert(strTabelle, strIndex) Then Exit Sub

    ' Feldliste aufbauen (Komma-separiert -> einzelne [Feld]-Eintraege)
    Dim strFeldListe As String
    If InStr(strFelder, ",") > 0 Then
        Dim arrParts() As String
        arrParts = Split(strFelder, ",")
        Dim i As Long
        For i = LBound(arrParts) To UBound(arrParts)
            If i > LBound(arrParts) Then strFeldListe = strFeldListe & ", "
            strFeldListe = strFeldListe & "[" & Trim$(arrParts(i)) & "]"
        Next i
    Else
        strFeldListe = "[" & strFelder & "]"
    End If

    Dim strSQL As String
    If bUnique Then
        strSQL = "CREATE UNIQUE INDEX [" & strIndex & "] ON [" & strTabelle & "] (" & strFeldListe & ")"
    Else
        strSQL = "CREATE INDEX [" & strIndex & "] ON [" & strTabelle & "] (" & strFeldListe & ")"
    End If

    If DDL_Ausfuehren(strTabelle, strSQL) Then
        LogDebug "  [+IDX] " & strIndex & " ON " & strTabelle, MODUL_NAME
    End If
End Sub


' ===========================================================================
' GENERISCHE DDL-AUSFUEHRUNG (Backend-transparent)
' ===========================================================================

' Fuehrt beliebiges DDL-SQL gegen die richtige Datenbank aus.
' Bei Linked Tables -> Backend-DB oeffnen, DDL dort ausfuehren, Link refreshen.
' Bei lokalen Tabellen -> CurrentDb.Execute.
Public Function DDL_Ausfuehren(ByVal strTabelle As String, ByVal strSQL As String) As Boolean
    On Error GoTo ErrHandler

    If DDL_IstVerknuepft(strTabelle) Then
        ' Linked Table: DDL gegen Backend-DB ausfuehren
        Dim strPfad As String
        strPfad = HoleBackendPfad(strTabelle)
        If Len(strPfad) = 0 Then
            LogWarn "Backend-Pfad nicht ermittelbar fuer " & strTabelle, MODUL_NAME
            DDL_Ausfuehren = False
            Exit Function
        End If

        Dim dbBE As DAO.Database
        Set dbBE = DBEngine.OpenDatabase(strPfad)
        dbBE.Execute strSQL
        dbBE.Close
        Set dbBE = Nothing

        ' Link refreshen damit Frontend die Aenderung sieht
        RefreshTableLink strTabelle
    Else
        ' Lokale Tabelle: direkt ausfuehren
        CurrentDb.Execute strSQL
    End If

    DDL_Ausfuehren = True
    Exit Function
ErrHandler:
    ' Typische "existiert bereits"-Fehler nicht als Warnung loggen
    Select Case Err.Number
        Case 3283, 3284  ' Index existiert bereits / Feld existiert bereits
            DDL_Ausfuehren = True
        Case 3376  ' Table already exists
            DDL_Ausfuehren = False
        Case Else
            LogWarn "DDL-Fehler " & Err.Number & ": " & Err.Description & " | SQL: " & Left$(strSQL, 200), MODUL_NAME
            DDL_Ausfuehren = False
    End Select
End Function


' ===========================================================================
' INTERNE HELFER (Private)
' ===========================================================================

' Gibt die echte TableDef zurueck (bei Linked Table aus der Backend-DB)
' WICHTIG: Die zugehoerige Backend-DB bleibt offen und wird vom Aufrufer
'          NICHT geschlossen - nur fuer kurzlebige Existenz-Pruefungen verwenden.
Private Function HoleEchteTableDef(ByVal strTabelle As String) As DAO.TableDef
    On Error GoTo Fehlschlag

    If Not DDL_TabelleExistiert(strTabelle) Then
        Set HoleEchteTableDef = Nothing
        Exit Function
    End If

    Dim tdf As DAO.TableDef
    Set tdf = CurrentDb.TableDefs(strTabelle)

    If Len(Nz(tdf.Connect, "")) > 0 Then
        ' Linked Table: TableDef aus Backend-DB holen
        Dim strPfad As String
        strPfad = HoleBackendPfad(strTabelle)
        If Len(strPfad) = 0 Then
            Set HoleEchteTableDef = Nothing
            Exit Function
        End If
        Dim dbBE As DAO.Database
        Set dbBE = DBEngine.OpenDatabase(strPfad)
        Set HoleEchteTableDef = dbBE.TableDefs(strTabelle)
        ' Hinweis: dbBE bleibt offen, wird aber durch Jet/ACE intern verwaltet
    Else
        ' Lokale Tabelle
        Set HoleEchteTableDef = tdf
    End If
    Exit Function
Fehlschlag:
    Set HoleEchteTableDef = Nothing
End Function

' Oeffnet die Ziel-Datenbank fuer DAO-Operationen (DefaultValue setzen etc.)
' Bei Linked Tables: oeffnet Backend-DB, bMussSchliessen = True
' Bei lokalen Tabellen: gibt CurrentDb zurueck, bMussSchliessen = False
Private Function HoleZielDatenbank(ByVal strTabelle As String, ByRef bMussSchliessen As Boolean) As DAO.Database
    On Error GoTo Fehlschlag
    bMussSchliessen = False

    If DDL_IstVerknuepft(strTabelle) Then
        Dim strPfad As String
        strPfad = HoleBackendPfad(strTabelle)
        If Len(strPfad) = 0 Then
            Set HoleZielDatenbank = Nothing
            Exit Function
        End If
        Set HoleZielDatenbank = DBEngine.OpenDatabase(strPfad)
        bMussSchliessen = True
    Else
        Set HoleZielDatenbank = CurrentDb
        bMussSchliessen = False
    End If
    Exit Function
Fehlschlag:
    Set HoleZielDatenbank = Nothing
    bMussSchliessen = False
End Function

' Extrahiert den Backend-DB-Pfad aus der Connect-Eigenschaft einer Linked Table
Private Function HoleBackendPfad(ByVal strTabelle As String) As String
    On Error GoTo Fehlschlag
    Dim strConnect As String
    strConnect = CurrentDb.TableDefs(strTabelle).Connect

    Dim lngPos As Long
    lngPos = InStr(1, strConnect, "DATABASE=", vbTextCompare)
    If lngPos = 0 Then
        HoleBackendPfad = ""
        Exit Function
    End If

    Dim strPfad As String
    strPfad = Mid$(strConnect, lngPos + 9)
    If InStr(strPfad, ";") > 0 Then strPfad = Left$(strPfad, InStr(strPfad, ";") - 1)
    HoleBackendPfad = Trim$(strPfad)
    Exit Function
Fehlschlag:
    HoleBackendPfad = ""
End Function

' Aktualisiert den Link einer verknuepften Tabelle
Private Sub RefreshTableLink(ByVal strTabelle As String)
    On Error Resume Next
    CurrentDb.TableDefs(strTabelle).RefreshLink
    CurrentDb.TableDefs.Refresh
    Err.Clear
    On Error GoTo 0
End Sub
