# OutlookSync – Outlook-Mails in Access-Datenbank synchronisieren

> Vollständige Dokumentation: [KONZEPT.md](KONZEPT.md)

## Projektstruktur

```
├── KONZEPT.md              Architektur, Tabellenschema, Datenfluss
├── modOutlookTest.bas      Standalone-Testmodul (Outlook-Zugriff testen)
└── src/
    ├── modSchema.bas       Tabellen-DDL (10 Tabellen + Indizes + Config)
    ├── modGlobals.bas      Konstanten, MAPI-Tags, globale Objekte, Init
    ├── modCrypto.bas       SHA256-Hash (Windows CryptoAPI)
    ├── modLogging.bas      Logging (Debug.Print + Logfile)
    ├── modStringUtils.bas  String-Bereinigung, Pfade, E-Mail-Validierung
    ├── modOutlookConnect.bas  Outlook/Redemption Connect, SMTP-Auflösung
    ├── modDAO.bas          Datenzugriffsschicht (alle Tabellen)
    └── modSync.bas         Sync-Orchestrierung (Hauptlogik)
```

## Voraussetzungen

| Bedingung | Detail |
|-----------|--------|
| MS Access | 64-Bit Office |
| Outlook | geöffnet, eingeloggt (Exchange-Profil) |
| Redemption DLL | `D:\Redemption64.dll` (64-Bit) |
| DLL registriert | einmalig `regsvr32 "D:\Redemption64.dll"` als Admin |

## Ersteinrichtung

```
1. Alle .bas-Dateien aus src/ in Access importieren (ALT+F11 → Datei → Importieren)
2. Direktbereich (STRG+G):
   ErstelleAlleTabellen       ' Erstellt alle 10 Tabellen + Indizes + Config
3. Sync starten:
   SyncPosteingang "MeinProjekt", "Phase1"
```

## Schnellreferenz (Direktbereich `STRG+G`)

```vba
' === Schema ===
ErstelleAlleTabellen                    ' Tabellen erstellen (einmalig)
LoescheAlleTabellen                     ' ACHTUNG: Löscht alle Daten!

' === Sync ===
SyncPosteingang                         ' Posteingang (Defaults)
SyncPosteingang "FLIWAS", "Test", 50    ' Max 50 Mails
SyncOrdner "Postfach\Posteingang\FLIWAS", "FLIWAS", "Prod"

' === Ordnerstruktur ===
SyncOrdnerStruktur 3                    ' Alle Ordner bis Tiefe 3 in DB

' === Testmodul (standalone) ===
CheckOutlookAccessMethods               ' Schnelltest Zugriffswege
CheckOutlookDeepAccess                  ' 11-Test Suite
AnalyzeSingleEmail                      ' Mail vollständig sezieren
ShowFolderTree 3                        ' Ordnerstruktur anzeigen
```

## Tabellen (10 Stück)

| Tabelle | Zweck |
|---------|-------|
| `tblConfig` | Anwendungskonfiguration |
| `tblSyncLauf` | Sync-Protokoll je Durchlauf |
| `tblKontakte` | Kontakte (Absender + Empfänger) |
| `tblOutlookOrdner` | Ordnerstruktur |
| `tblEmailThreads` | Konversations-Threads |
| `tblEmails` | Haupt-Mail-Tabelle (Metadaten) |
| `tblEmailContent` | Mail-Inhalt (HTML + Plaintext, separiert) |
| `tblEmailEmpfaenger` | Empfänger je Mail (To/CC/BCC) |
| `tblEmailAnhaenge` | Anhang-Metadaten + Dateipfade |
| `tblEmailStatus` | Status-Historie |

## Konfiguration (tblConfig)

| Schlüssel | Standard | Beschreibung |
|-----------|----------|-------------|
| ExportBasisPfad | `%USERPROFILE%\OutlookSync\` | Basis für MSG + Anhänge |
| MaxMailsProSync | 500 | Max. Mails pro Durchlauf |
| AnhaengeExtrahieren | 1 | Anhänge auf Festplatte (1=Ja) |
| MSGExportieren | 1 | MSG-Dateien exportieren (1=Ja) |
| SignaturBilderFiltern | 1 | Signatur-Bilder überspringen |
| LogLevel | 3 | 0=Aus, 1=Error, 2=Warn, 3=Info, 4=Debug |

## Testergebnisse (modOutlookTest)

```
Outlook 16.0.0.17932 / Redemption 6.7.0.6412 / 64-Bit / Exchange
CheckOutlookDeepAccess: 35 OK / 6 FAIL (85%)
Alle FAIL = OOM Security Guard (erwartetes Verhalten)
```
| PR_DISPLAY_TO | `&HE04001E` | An-Feld |
| PR_DISPLAY_CC | `&HE03001E` | CC-Feld |

---

## Anhang-Filterlogik (`ExtractAttachmentsViaRedemption`)

```
objAtt.Hidden = True   → Signatur-Bild / Inline-Bild → ÜBERSPRINGEN
objAtt.Type   <> 1     → kein normaler Dateianhang   → ÜBERSPRINGEN
Sonst                  → echte Datei                 → SPEICHERN
```

Attachment-Typen: `1` = normale Datei · `5` = eingebettetes OLE · `6` = weitergeleitete Mail