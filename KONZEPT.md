# OutlookSync – Konzept & Architektur (v0.3)

## 1. Projektziel

Synchronisation von Microsoft Outlook-Mails in eine Access-Datenbank mit:
- **Frontend/Backend-Architektur** (FE lokal, BE auf Netzlaufwerk)
- **Extract-Release-Process** Pattern (Outlook schnell entlasten)
- **Schreib-Puffer** mit Transaktions-Batches (Performance)
- **Datei-Queue** mit Retry-Logik (Netzwerk-Robustheit)
- **Multi-Store/Archive** Unterstützung (Exchange, Online-Archive, Shared Mailboxes)
- **Selektiver Sync** über Profile (Ordnerauswahl speicherbar)
- SHA256-Duplikat-Erkennung (Batch-Prüfung per SQL IN-Clause)
- Thread/Konversations-Erkennung via Betreff + In-Reply-To
- Kontakt-Verwaltung mit Namens-Parsing und Domain-Lernen

---

## 2. Architektur-Überblick

```
┌─────────────────────────────────────────────────────────────────────┐
│                          OUTLOOK / EXCHANGE                        │
│   Postfächer · Online-Archive · Shared Mailboxes · Unterordner     │
└──────────────┬──────────────────────────────────────────────────────┘
               │ Redemption RDO (Late Binding)
               ▼
┌──────────────────────────┐
│  1. EXTRACT              │  modMailExtract.ExtrahiereKomplett()
│  Alle Daten lesen:       │  → TypMailKomplett (VBA Struct)
│  Metadaten, Content,     │  → Anhänge → lokaler Temp-Ordner
│  Empfänger, Anhänge      │  → MSG → lokaler Temp-Ordner
└──────────────┬───────────┘
               │ Set objItem = Nothing  (COM sofort freigeben!)
               ▼
┌──────────────────────────┐
│  2. BUFFER               │  modAsyncBuffer.BufferHinzufuegen()
│  In Speicher sammeln     │  → m_aBuffer() As TypMailKomplett
│  (Standard: 25 Mails)    │  → Auto-Flush wenn voll
└──────────────┬───────────┘
               │ BufferIstVoll() → BufferFlush()
               ▼
┌──────────────────────────┐
│  3. FLUSH (Transaktion)  │  modAsyncBuffer.BufferFlush()
│  Batch-Duplikatprüfung   │  → SELECT ... WHERE Hash IN (...)
│  BeginTrans              │  → Kontakte, Threads, Emails
│   → DB-INSERTs           │  → Content, Empfänger, Anhänge
│  CommitTrans             │  → DateiQueue-Einträge erstellen
└──────────────┬───────────┘
               │ DateiQueueVerarbeiten()
               ▼
┌──────────────────────────┐
│  4. COPY (Retry)         │  FileCopy mit exponent. Backoff
│  Temp → Netzlaufwerk     │  → MSG-Dateien + Anhänge
│  3 Versuche, 2s/4s/8s    │  → DB-Pfade aktualisieren
└──────────────────────────┘
```

---

## 3. Technische Basis

| Komponente | Detail |
|---|---|
| MS Access | 64-Bit, VBA 7.x |
| Outlook | 16.x, Exchange, OOM Security Guard aktiv |
| Redemption | 6.7.x (`D:\Redemption64.dll`), registriert via `regsvr32` |
| Zugriff | Late Binding (`CreateObject`), kein Verweis nötig |
| DAO | Built-in Access DAO (Datenzugriff auf Tabellen) |
| Crypto | Windows CryptoAPI (`advapi32.dll`) für SHA256 |
| Backend | Verknüpfte Tabellen auf Netzlaufwerk (.accdb) |
| Transaktionen | `DBEngine.Workspaces(0).BeginTrans/CommitTrans` |

---

## 4. Frontend/Backend-Architektur

### 4.1 Frontend (.accdb – lokal)
- VBA-Code (alle Module)
- Formulare (zukünftig)
- `tblConfig` – Anwendungskonfiguration
- `tblSyncProfil` – Gespeicherte Sync-Profile
- `tblSyncProfilOrdner` – Ordner je Profil

### 4.2 Backend (.accdb – Netzlaufwerk)
- `tblSyncLauf` – Sync-Protokoll
- `tblKontakte` – Kontakte
- `tblOutlookOrdner` – Ordnerstruktur
- `tblEmailThreads` – Konversationen
- `tblEmails` – Haupt-Mail-Tabelle
- `tblEmailContent` – Mail-Inhalt (HTML/Plain)
- `tblEmailEmpfaenger` – Empfänger
- `tblEmailAnhaenge` – Anhang-Metadaten
- `tblEmailStatus` – Status-Historie

### 4.3 Verknüpfung

```vba
' Backend auf Netzlaufwerk einrichten (einmalig):
VerknuepfeBackend "\\Server\Share\OutlookSync_BE.accdb"

' Status prüfen:
? BackendStatus()   ' → "Backend OK: \\Server\Share\OutlookSync_BE.accdb"

' Wieder auf lokal umstellen:
TrenneBackend
```

Ohne Backend-Konfiguration arbeitet alles lokal (abwärtskompatibel).

---

## 5. Tabellenschema (12 Tabellen)

### 5.1 `tblConfig` – Anwendungskonfiguration (FE)
| Feld | Typ | Beschreibung |
|---|---|---|
| ConfigID | AutoNumber PK | |
| Schluessel | Text(100) UNIQUE | Konfigurationsschlüssel |
| Wert | Text(255) | Wert |
| Beschreibung | Text(255) | Erklärung |

### 5.2 `tblSyncProfil` – Gespeicherte Sync-Profile (FE)
| Feld | Typ | Beschreibung |
|---|---|---|
| ProfilID | AutoNumber PK | |
| ProfilName | Text(100) UNIQUE | Profilname |
| Beschreibung | Text(255) | |
| IstAktiv | YesNo | |
| Projekt | Text(100) | Projekt-Zuordnung |
| Phase | Text(100) | Phase |
| MaxMailsProOrdner | Long | Max. Mails pro Ordner |
| MaxTiefe | Short | Subfolder-Tiefe |
| ExportPfad | Text(255) | Override Export-Pfad |
| ErstelltAm | DateTime | |

### 5.3 `tblSyncProfilOrdner` – Ordner je Profil (FE)
| Feld | Typ | Beschreibung |
|---|---|---|
| ID | AutoNumber PK | |
| ProfilID | Long FK | → tblSyncProfil |
| OrdnerPfad | Text(255) | Outlook-Ordner-Pfad |
| PostfachName | Text(255) | Store-DisplayName |
| IstAktiv | YesNo | |

### 5.4 `tblSyncLauf` – Sync-Protokoll (BE)
| Feld | Typ | Beschreibung |
|---|---|---|
| SyncLaufID | AutoNumber PK | |
| StartZeit | DateTime | |
| EndeZeit | DateTime | |
| Status | Text(20) | Gestartet / Abgeschlossen / Fehler / Abgebrochen |
| AnzahlGelesen | Long | Gelesene Mails |
| AnzahlNeu | Long | Neu gespeicherte Mails |
| AnzahlDuplikate | Long | Übersprungene Duplikate |
| AnzahlFehler | Long | Fehlerhafte Mails |
| OrdnerPfad | Text(255) | Outlook-Ordnerpfad |
| Projekt | Text(100) | Projektzuordnung |
| Phase | Text(100) | Phasenzuordnung |

### 5.5 `tblKontakte` – Kontakte (BE)
| Feld | Typ | Beschreibung |
|---|---|---|
| KontaktID | AutoNumber PK | |
| Anzeigename | Text(255) | Original-Anzeigename |
| Email | Text(255) | SMTP-Adresse |
| EmailTyp | Text(10) | SMTP / EX |
| Vorname | Text(100) | Geparst aus Anzeigename |
| Nachname | Text(100) | Geparst aus Anzeigename |
| Titel | Text(50) | Dr., Prof. Dr. etc. |
| Namenszusatz | Text(100) | Mittlere Namensteile |
| Institution | Text(255) | Aus Klammer oder E-Mail-Domain |
| Sortiername | Text(255) | "Nachname, Vorname" |
| KontaktTyp | Text(20) | Intern/Extern/System |
| ErstelltAm | DateTime | |
| AktualisiertAm | DateTime | |

### 5.6 `tblOutlookOrdner` – Ordnerstruktur (BE)
| Feld | Typ | Beschreibung |
|---|---|---|
| OrdnerID | AutoNumber PK | |
| OrdnerName | Text(255) | Anzeigename |
| OrdnerPfad | Text(255) UNIQUE | Vollständiger Pfad |
| ParentID | Long | FK → tblOutlookOrdner (0 = Root) |
| PostfachName | Text(255) | Store-DisplayName |
| StoreID | Text(255) | Eindeutige Store-ID (für Multi-Store) |
| ElementAnzahl | Long | Anzahl Elemente im Ordner |
| LetzterSync | DateTime | Letzter Sync-Zeitpunkt |

### 5.7 `tblEmailThreads` – Konversations-Threads (BE)
| Feld | Typ | Beschreibung |
|---|---|---|
| ThreadID | AutoNumber PK | |
| ThreadBetreff | Text(255) | Bereinigter Betreff |
| ThreadIdentifier | Text(255) UNIQUE | MessageID oder Betreff |
| Antwortanzahl | Long | |
| ErsterAbsender | Text(255) | |
| ErstesMailDatum | DateTime | |
| LetztesMailDatum | DateTime | |
| ErstelltAm | DateTime | |

### 5.8 `tblEmails` – Haupt-Mail-Tabelle (BE)
| Feld | Typ | Beschreibung |
|---|---|---|
| EmailID | AutoNumber PK | |
| OutlookEntryID | Text(255) | |
| UniqueHash | Text(64) UNIQUE | SHA256-Duplikat-Hash |
| ThreadID | Long FK | |
| OrdnerID | Long FK | |
| KontaktID_Absender | Long FK | |
| SyncLaufID | Long FK | |
| Betreff | Text(255) | |
| BetreffBereinigt | Text(255) | |
| AbsenderName | Text(255) | |
| AbsenderEmail | Text(255) | |
| EmpfangenAm | DateTime | |
| GesendetAm | DateTime | |
| Groesse | Long | |
| Wichtigkeit | Short | |
| Gelesen | YesNo | |
| HatAnhaenge | YesNo | |
| AnhangAnzahl | Short | |
| MessageClass | Text(50) | |
| InternetMessageID | Text(255) | |
| MSGDateiPfad | Text(255) | |
| Status | Text(20) | |
| ErstelltAm | DateTime | |

### 5.9–5.11 Weitere BE-Tabellen
- **tblEmailContent** (EmailID UNIQUE FK, HTMLBody MEMO, PlainTextBody MEMO, HatHTML, Größen)
- **tblEmailEmpfaenger** (EmailID FK, KontaktID FK, Typ To/CC/BCC, Name, Email)
- **tblEmailAnhaenge** (EmailID FK, Dateiname, Größe, MimeType, AnhangTyp, IstVersteckt, DateiPfad)
- **tblEmailStatus** (EmailID FK, Status, GeaendertVon, Bemerkung, GeaendertAm)

---

## 6. Modulstruktur (12 Module)

| Modul | Zeilen* | Aufgabe |
|---|---|---|
| **modMailExtract** | ~350 | **NEU v0.3**: Datentypen (TypMailKomplett etc.), COM-Extraktion, Temp-Dateien, Header-Parser |
| **modBackend** | ~330 | **NEU v0.3**: FE/BE-Verknüpfung, Reconnect, Backend-DB-Erstellung, Datenmigration |
| **modAsyncBuffer** | ~400 | **NEU v0.3**: Schreib-Puffer, Batch-Flush, Datei-Queue, Retry-Logik |
| **modSchema** | ~500 | Tabellen-DDL, Indizes, Konfiguration (+tblSyncProfil, +tblSyncProfilOrdner) |
| **modGlobals** | ~130 | Konstanten, MAPI-Tags, globale Objekte, Init/Cleanup |
| **modCrypto** | ~160 | SHA256-Hash via Windows CryptoAPI |
| **modLogging** | ~110 | SchreibeLog, LogInfo/Warn/Error/Debug, Logfile |
| **modStringUtils** | ~400 | Betreff, Dateinamen, Pfade, E-Mail, Capitalize, Institution |
| **modKontakte** | ~440 | Namens-Parsing, Geschlecht, Anrede, Domain-Lernen |
| **modOutlookConnect** | ~300 | Outlook/RDO Connect, SMTP-Auflösung, Ordner öffnen |
| **modDAO** | ~580 | Datenzugriff: Kontakte, Emails, Threads, Ordner, Anhänge |
| **modSync** | ~530 | Sync-Orchestrierung, Profil-Sync, Subfolder, Postfach |

\* Ungefähre Zeilenanzahl. Zusätzlich: `modOutlookTest.bas` (Standalone-Testmodul)

---

## 7. Datenfluss (v0.3)

```
Outlook-Ordner
    │
    ▼
modSync.SyncFolder(objFolder, "Projekt", "Phase", 500)
    │
    ├─ IstBackendVerfuegbar()           → Netzwerk-Check
    ├─ modDAO.StarteSyncLauf()          → tblSyncLauf
    │
    ├─ BufferInit + BufferSetzeKontext  → Puffer vorbereiten
    │
    ├─ Für jede Mail (IPM.Note) im Ordner:
    │   │
    │   ├─ 1. EXTRACT: modMailExtract.ExtrahiereKomplett()
    │   │      ├── Basisdaten (Betreff, Absender, Datum...)
    │   │      ├── SMTP-Auflösung via modOutlookConnect
    │   │      ├── Hash via modCrypto.GeneriereMailHash()
    │   │      ├── HTML + Plaintext Content
    │   │      ├── Empfänger → TypEmpfaengerDaten()
    │   │      ├── Anhänge → Temp-Ordner (lokal, schnell)
    │   │      └── MSG → Temp-Ordner (lokal, schnell)
    │   │
    │   ├─ 2. RELEASE: Set objItem = Nothing  ← COM sofort frei!
    │   │
    │   ├─ 3. BUFFER: BufferHinzufuegen(mk)
    │   │
    │   └─ 4. FLUSH (wenn Puffer voll):
    │          ├── Batch-Hash-Prüfung (SQL IN-Clause)
    │          ├── BeginTrans
    │          │   ├── modKontakte.ParseKontaktName() + LerneVonDomain()
    │          │   ├── modDAO.GetOderErstelleKontakt()  → tblKontakte
    │          │   ├── modDAO.GetOderErstelleThread()   → tblEmailThreads
    │          │   ├── modDAO.SpeichereEmail()          → tblEmails
    │          │   ├── modDAO.SpeichereEmailContent()   → tblEmailContent
    │          │   ├── modDAO.SpeichereEmpfaenger()     → tblEmailEmpfaenger
    │          │   ├── modDAO.SpeichereAnhangMetadaten() → tblEmailAnhaenge
    │          │   └── DateiQueue-Einträge erstellen
    │          └── CommitTrans
    │
    ├─ BufferFlush()                    → Restliche Daten
    │
    ├─ DateiQueueVerarbeiten()          → Temp → Netzwerk (mit Retry)
    │   ├── FileCopy + exponent. Backoff (2s → 4s → 8s)
    │   ├── AktualisiereAnhangPfad()    → tblEmailAnhaenge
    │   └── SetzeEmailMSGPfad()         → tblEmails
    │
    ├─ BereinigeTempDateien()           → Temp-Ordner aufräumen
    │
    └─ modDAO.BeendeSyncLauf()          → tblSyncLauf (Status + Zähler)
```

---

## 8. Abhängigkeiten

```
modSync (Orchestrierung)
  ├── modMailExtract  (Types + Extraktion)
  ├── modAsyncBuffer  (Puffer + Flush + DateiQueue)
  │     ├── modDAO
  │     │     ├── modCrypto
  │     │     ├── modStringUtils
  │     │     ├── modKontakte
  │     │     └── modLogging
  │     ├── modKontakte
  │     └── modStringUtils
  ├── modBackend      (FE/BE Status)
  ├── modOutlookConnect
  ├── modDAO
  └── modLogging

modMailExtract (Extraktion)
  ├── modOutlookConnect (SMTP-Auflösung)
  ├── modCrypto         (Hash-Generierung)
  ├── modStringUtils    (Pfade, Bereinigung)
  └── modGlobals        (Konstanten, MAPI-Tags)

modBackend (FE/BE-Verwaltung)
  ├── modSchema (TabelleExistiert, ErstelleAlleTabellen)
  ├── modStringUtils (NormalisierePfad, ErstelleOrdner)
  └── modLogging

modKontakte
  ├── modStringUtils
  └── modLogging

modSchema (standalone)
modGlobals (standalone)
modCrypto (standalone - nur advapi32.dll)
modLogging (standalone - nur Debug.Print + Dateisystem)
```

---

## 9. Export-Verzeichnisstruktur

```
{ExportBasisPfad}\            (lokal oder Netzlaufwerk)
  └── {Projekt}\
      └── {Phase}\
          ├── MSG\
          │   ├── 20260304_1423_Betreff_der_Mail.msg
          │   └── ...
          └── Anhaenge\
              ├── EmailID_42\
              │   ├── Rechnung_2026.pdf
              │   └── Anhang.xlsx
              └── EmailID_43\
                  └── Bericht.docx

{TempPfad}\                   (%TEMP%\OutlookSync\ oder konfiguriert)
  ├── att_20260304_142355_0001.pdf    (temporär, wird nach Kopie gelöscht)
  ├── msg_20260304_142355_0002.msg
  └── ...
```

---

## 10. Konfiguration (tblConfig)

| Schlüssel | Standard | Beschreibung |
|---|---|---|
| ExportBasisPfad | `%USERPROFILE%\OutlookSync\` | Basis für MSG/Anhänge |
| MaxMailsProSync | `500` | Max. Mails pro Ordner |
| AnhaengeExtrahieren | `1` | Anhänge speichern |
| MSGExportieren | `1` | MSG-Dateien exportieren |
| SignaturBilderFiltern | `1` | Hidden/Inline Signatur-Bilder überspringen |
| LogLevel | `3` | 0=Aus 1=Error 2=Warn 3=Info 4=Debug 5=Trace |
| **BackendPfad** | *(leer)* | Pfad zur Backend-.accdb (leer = lokal) |
| **TempPfad** | *(leer)* | Temp-Verzeichnis (leer = %TEMP%\OutlookSync\) |
| **BufferGroesse** | `25` | Mails im Puffer vor Flush (5-500) |
| **NetzwerkRetries** | `3` | Wiederholungen bei Datei-Kopie-Fehler |
| **NetzwerkRetryPause** | `2000` | Millisekunden Basis-Pause |

---

## 11. Import & Ersteinrichtung

```
1. Alle .bas-Dateien in Access importieren (ALT+F11 → Datei → Importieren):
   - modSchema.bas         (Tabellen-DDL)
   - modGlobals.bas        (Konstanten)
   - modCrypto.bas         (SHA256)
   - modLogging.bas        (Logging)
   - modStringUtils.bas    (String-Helfer)
   - modKontakte.bas       (Kontaktlogik)
   - modOutlookConnect.bas (Outlook/Redemption)
   - modMailExtract.bas    (NEU: Extraktion + Types)
   - modBackend.bas        (NEU: FE/BE-Verwaltung)
   - modAsyncBuffer.bas    (NEU: Schreib-Puffer)
   - modDAO.bas            (Datenzugriff)
   - modSync.bas           (Sync-Orchestrierung)

2. Tabellen erstellen (Direktbereich STRG+G):
   ErstelleAlleTabellen

3. Optional: Backend auf Netzlaufwerk einrichten:
   VerknuepfeBackend "\\Server\Share\OutlookSync_BE.accdb"
   ' oder:
   VerknuepfeBackend "S:\Datenbank\OutlookSync_BE.accdb"

4. Erster Sync:
   SyncPosteingang "MeinProjekt", "Phase1"

WICHTIG bei Upgrade von v0.2:
   LoescheAlleTabellen    ' Schema hat sich geändert (+StoreID, +Profile)
   ErstelleAlleTabellen   ' Neu erstellen
```

---

## 12. Aufruf-Beispiele (Direktbereich `STRG+G`)

```vba
' === Schema ===
ErstelleAlleTabellen              ' Alle Tabellen erstellen (einmalig)
LoescheAlleTabellen               ' ACHTUNG: Löscht alle Daten!

' === Backend ===
VerknuepfeBackend "\\Server\Share\OutlookSync_BE.accdb"
TrenneBackend
? BackendStatus()
? IstBackendVerfuegbar()

' === Einfacher Sync ===
SyncPosteingang                              ' Posteingang (Standard)
SyncPosteingang "FLIWAS", "Test", 50         ' 50 Mails, Projekt=FLIWAS
SyncPosteingang "FLIWAS", "Test", 50, True   ' Mit Unterordnern
SyncOrdner "Torsten.Kugler@rps.bwl.de\Posteingang\FLIWAS", "FLIWAS", "Prod"

' === Postfach-Sync ===
SyncPostfach "Torsten.Kugler@rps.bwl.de"    ' Ganzes Postfach

' === Profil-Sync ===
Dim id As Long
id = ErstelleSyncProfil("FLIWAS_Prod", "FLIWAS", "Produktion")
ProfilOrdnerHinzufuegen id, "Torsten.Kugler@rps.bwl.de\Posteingang\FLIWAS"
ProfilOrdnerHinzufuegen id, "Torsten.Kugler@rps.bwl.de\Gesendete Elemente"
SyncMitProfil "FLIWAS_Prod"

' === Ordnerstruktur ===
SyncOrdnerStruktur 3              ' Alle Ordner bis Tiefe 3 einlesen

' === Konfiguration ===
? LeseConfig("ExportBasisPfad")
? LeseConfig("BufferGroesse")
SchreibeConfig "BufferGroesse", "50"
```

---

## 13. Datentypen (modMailExtract)

```vba
' Komplettes Mail-Paket nach Extraktion:
Public Type TypMailKomplett
    Mail                As TypMailDaten       ' 19 skalare Felder
    Empfaenger()        As TypEmpfaengerDaten ' dynamisches Array
    EmpfaengerAnzahl    As Integer
    Anhaenge()          As TypAnhangDaten     ' dynamisches Array
    AnhangAnzahl        As Integer
    MSGTempPfad         As String             ' lokaler Temp-Pfad
    UniqueHash          As String             ' SHA256
End Type

' Datei-Operation für Queue:
Public Type TypDateiOperation
    QuellPfad       As String   ' Temp
    ZielPfad        As String   ' Netzwerk
    OperationsTyp   As String   ' "MSG" / "Anhang"
    EmailID         As Long     ' FK
    AnhangID        As Long     ' FK
    Versuche        As Integer  ' Retry-Counter
End Type
```

---

## 14. Roadmap

| Phase | Beschreibung | Status |
|---|---|---|
| **v0.1** | Schema + Grundmodule (8 .bas) | ✅ |
| **v0.2** | Kontakt-Erweiterung, Namens-Parsing, Subfolder-Sync | ✅ |
| **v0.3** | **FE/BE-Architektur, Extract-Buffer-Flush, Datei-Queue, Profil-Sync** | ← AKTUELL |
| v0.4 | CID/Inline-Bild-Auflösung, HTML-Content-Bereinigung | |
| v0.5 | StatusPanel (Fortschrittsanzeige via Access-Formular) | |
| v0.6 | Inkrementeller Sync (nur neue Mails seit LetzterSync) | |
| v0.7 | Formular: Sync-Steuerung + Ordnerauswahl + Profilmanager | |
| v1.0 | Produktionsreif | |
