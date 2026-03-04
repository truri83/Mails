# OutlookSync – Konzept & Architektur

## 1. Projektziel

Synchronisation von Microsoft Outlook-Mails in eine Access-Datenbank mit:
- Vollständige Metadaten (Betreff, Absender, Empfänger, Datum, Größe)
- Mail-Inhalt (HTML + Plaintext) in separater Tabelle (Performance)
- Anhang-Extraktion (Signatur-Bilder gefiltert) auf Festplatte
- MSG-Datei-Export auf Festplatte
- Thread/Konversations-Erkennung via Betreff + In-Reply-To
- Duplikat-Erkennung via SHA256-Hash
- Ordnerstruktur-Abbildung in DB
- Kontakt-Verwaltung (Absender + Empfänger)

---

## 2. Technische Basis

| Komponente | Detail |
|---|---|
| MS Access | 64-Bit, VBA 7.x |
| Outlook | 16.x, Exchange, OOM Security Guard aktiv |
| Redemption | 6.7.x (`D:\Redemption64.dll`), registriert via `regsvr32` |
| Zugriff | Late Binding (`CreateObject`), kein Verweis nötig |
| DAO | Built-in Access DAO (Datenzugriff auf Tabellen) |
| Crypto | Windows CryptoAPI (`advapi32.dll`) für SHA256 |

---

## 3. Tabellenschema (10 Tabellen)

### 3.1 `tblConfig` – Anwendungskonfiguration
| Feld | Typ | Beschreibung |
|---|---|---|
| ConfigID | AutoNumber PK | |
| Schluessel | Text(100) UNIQUE | Konfigurationsschlüssel |
| Wert | Text(255) | Wert |
| Beschreibung | Text(255) | Erklärung |

### 3.2 `tblSyncLauf` – Sync-Protokoll je Durchlauf
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

### 3.3 `tblKontakte` – Kontakte (Absender + Empfänger)
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
| KontaktTyp | Text(20) | Optional: Intern/Extern/System |
| ErstelltAm | DateTime | |
| AktualisiertAm | DateTime | |

### 3.4 `tblOutlookOrdner` – Ordnerstruktur
| Feld | Typ | Beschreibung |
|---|---|---|
| OrdnerID | AutoNumber PK | |
| OrdnerName | Text(255) | Anzeigename |
| OrdnerPfad | Text(255) UNIQUE | Vollständiger Pfad |
| ParentID | Long | FK → tblOutlookOrdner (0 = Root) |
| PostfachName | Text(255) | Store-DisplayName |
| ElementAnzahl | Long | Anzahl Elemente im Ordner |
| LetzterSync | DateTime | Letzter Sync-Zeitpunkt |

### 3.5 `tblEmailThreads` – Konversations-Threads
| Feld | Typ | Beschreibung |
|---|---|---|
| ThreadID | AutoNumber PK | |
| ThreadBetreff | Text(255) | Bereinigter Betreff (ohne RE:/FW:) |
| ThreadIdentifier | Text(255) UNIQUE | MessageID oder bereinigter Betreff |
| Antwortanzahl | Long | Anzahl Mails im Thread |
| ErsterAbsender | Text(255) | |
| ErstesMailDatum | DateTime | |
| LetztesMailDatum | DateTime | |
| ErstelltAm | DateTime | |

### 3.6 `tblEmails` – Haupt-Mail-Tabelle
| Feld | Typ | Beschreibung |
|---|---|---|
| EmailID | AutoNumber PK | |
| OutlookEntryID | Text(255) | Outlook EntryID |
| UniqueHash | Text(64) UNIQUE | SHA256-Duplikat-Hash |
| ThreadID | Long | FK → tblEmailThreads |
| OrdnerID | Long | FK → tblOutlookOrdner |
| KontaktID_Absender | Long | FK → tblKontakte |
| SyncLaufID | Long | FK → tblSyncLauf |
| Betreff | Text(255) | Original-Betreff |
| BetreffBereinigt | Text(255) | Ohne RE:/FW:/AW:/WG: etc. |
| AbsenderName | Text(255) | |
| AbsenderEmail | Text(255) | SMTP-Adresse |
| EmpfangenAm | DateTime | |
| GesendetAm | DateTime | |
| Groesse | Long | Bytes |
| Wichtigkeit | Short | 0=Niedrig, 1=Normal, 2=Hoch |
| Gelesen | YesNo | |
| HatAnhaenge | YesNo | |
| AnhangAnzahl | Short | |
| MessageClass | Text(50) | IPM.Note etc. |
| InternetMessageID | Text(255) | Message-ID Header |
| MSGDateiPfad | Text(255) | Pfad zur .msg-Datei |
| Status | Text(20) | Neu / Verarbeitet / Fehler |
| ErstelltAm | DateTime | |

### 3.7 `tblEmailContent` – Mail-Inhalt (separiert für Performance)
| Feld | Typ | Beschreibung |
|---|---|---|
| ContentID | AutoNumber PK | |
| EmailID | Long UNIQUE | FK → tblEmails (1:1) |
| HTMLBody | Memo | |
| PlainTextBody | Memo | |
| HatHTML | YesNo | |
| GroesseHTML | Long | |
| GroesseText | Long | |

### 3.8 `tblEmailEmpfaenger` – Empfänger je Mail
| Feld | Typ | Beschreibung |
|---|---|---|
| EmpfaengerID | AutoNumber PK | |
| EmailID | Long | FK → tblEmails |
| KontaktID | Long | FK → tblKontakte |
| Typ | Text(5) | To / CC / BCC |
| Anzeigename | Text(255) | |
| Email | Text(255) | SMTP-Adresse |

### 3.9 `tblEmailAnhaenge` – Anhang-Metadaten + Dateipfade
| Feld | Typ | Beschreibung |
|---|---|---|
| AnhangID | AutoNumber PK | |
| EmailID | Long | FK → tblEmails |
| Dateiname | Text(255) | Original-Dateiname |
| DateinameBereinigt | Text(255) | Bereinigt |
| Erweiterung | Text(20) | pdf, docx, etc. |
| Groesse | Long | Bytes |
| MimeType | Text(100) | |
| AnhangTyp | Short | 1=Datei, 5=OLE, 6=angehängte Mail |
| IstVersteckt | YesNo | Signatur-Bild (inline) |
| IstGespeichert | YesNo | Datei auf Festplatte vorhanden? |
| DateiPfad | Text(255) | Pfad auf Festplatte |
| ErstelltAm | DateTime | |

### 3.10 `tblEmailStatus` – Status-Historie
| Feld | Typ | Beschreibung |
|---|---|---|
| StatusID | AutoNumber PK | |
| EmailID | Long | FK → tblEmails |
| Status | Text(50) | |
| GeaendertVon | Text(100) | |
| Bemerkung | Text(255) | |
| GeaendertAm | DateTime | |

---

## 4. Modulstruktur (8 Module)

| Modul | Zeilen* | Aufgabe |
|---|---|---|
| **modSchema** | ~430 | Tabellen-DDL (`CREATE TABLE`), Indizes, Standardkonfiguration |
| **modGlobals** | ~130 | Konstanten, MAPI-Tags, globale Objekte, `Sleep`-API, Init/Cleanup |
| **modCrypto** | ~160 | SHA256-Hash via Windows CryptoAPI + Mail-Hash-Generierung |
| **modLogging** | ~110 | `SchreibeLog`, `LogInfo/Warn/Error/Debug`, Logfile (tagesbasiert) |
| **modStringUtils** | ~400 | Betreff, Dateinamen, Pfade, E-Mail, Capitalize, IsAlpha, Institution |
| **modKontakte** | ~320 | Namens-Parsing, Geschlecht, Anrede, Institution, Domain-Lernen |
| **modOutlookConnect** | ~300 | Outlook/RDO Connect/Disconnect, SMTP-Auflösung, Ordner öffnen |
| **modDAO** | ~580 | Datenzugriff: Kontakte, Emails, Threads, Ordner, Anhänge, Status |
| **modSync** | ~760 | Sync-Orchestrierung, Subfolder-Rekursion, Postfach-Sync |

\* Ungefähre Zeilenanzahl

Zusätzlich existiert `modOutlookTest.bas` als unabhängiges Standalone-Testmodul.

---

## 5. Datenfluss

```
Outlook-Ordner
    │
    ▼
modSync.SyncOrdner("Postfach\Ordner", "Projekt", "Phase")
    │
    ├─ modOutlookConnect.ConnectRDO()
    │
    ├─ modDAO.StarteSyncLauf()           → tblSyncLauf
    │
    ├─ Für jede Mail (IPM.Note) im Ordner:
    │   │
    │   ├─ modCrypto.GeneriereMailHash()  → Duplikat?
    │   │   └── modDAO.ExistiertMailHash()  → SKIP wenn ja
    │   │
    │   ├─ modOutlookConnect.GetAbsenderSMTP()
    │   │
    │   ├─ modDAO.GetOderErstelleKontakt()  → tblKontakte
    │   │
    │   ├─ modDAO.GetOderErstelleThread()   → tblEmailThreads
    │   │
    │   ├─ modDAO.SpeichereEmail()          → tblEmails
    │   ├─ modDAO.SpeichereEmailContent()   → tblEmailContent
    │   │
    │   ├─ modSync.VerarbeiteEmpfaenger()   → tblEmailEmpfaenger
    │   │   ├── modDAO.GetOderErstelleKontakt() je Empfänger
    │   │   │     └── modKontakte.ParseKontaktName() + LerneVonDomain()
    │   │   └── modKontakte.AktualisiereKontaktEmail() bei "Unbekannt"
    │   │
    │   ├─ modSync.VerarbeiteAnhaenge()     → tblEmailAnhaenge + Dateien
    │   │   └── Filter: Hidden=True ODER Type≠1 → nur Metadaten
    │   │
    │   └─ modSync.ExportiereMSG()          → .msg-Datei + tblEmails.MSGDateiPfad
    │
    └─ modDAO.BeendeSyncLauf()              → tblSyncLauf (Status + Zähler)
```

---

## 6. Abhängigkeiten

```
modSync
  ├── modOutlookConnect
  ├── modDAO
  │     ├── modCrypto
  │     ├── modStringUtils
  │     ├── modKontakte
  │     └── modLogging
  ├── modKontakte  (NEU v0.2)
  ├── modStringUtils
  ├── modCrypto
  └── modLogging

modKontakte (NEU v0.2)
  ├── modStringUtils (CapitalizeWord, IsAlphaOnly, ExtrahiereInstitution...)
  └── modLogging

modOutlookConnect
  ├── modGlobals
  └── modLogging

modSchema (standalone – keine Abhängigkeiten)
modGlobals (standalone – nur modSchema.LeseConfig optional)
modOutlookTest (standalone – unabhängiges Testmodul)
```

---

## 7. Export-Verzeichnisstruktur

```
{ExportBasisPfad}\
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
```

---

## 8. Import & Ersteinrichtung

```
1. Alle .bas-Dateien in Access importieren (ALT+F11 → Datei → Importieren):
   - modSchema.bas
   - modGlobals.bas
   - modCrypto.bas
   - modLogging.bas
   - modStringUtils.bas
   - modKontakte.bas   (NEU in v0.2)
   - modOutlookConnect.bas
   - modDAO.bas
   - modSync.bas

2. Direktbereich (STRG+G):
   ErstelleAlleTabellen          ' Erstellt alle 10 Tabellen + Indizes + Config

3. Konfiguration prüfen/anpassen:
   ? LeseConfig("ExportBasisPfad")
   ' → Standard: C:\Users\{user}\OutlookSync\

4. Erster Sync:
   SyncPosteingang "MeinProjekt", "Phase1"
   ' oder:
   SyncOrdner "Torsten.Kugler@rps.bwl.de\Posteingang\FLIWAS", "FLIWAS", "Prod"
```

---

## 9. Aufruf-Beispiele (Direktbereich `STRG+G`)

```vba
' --- Schema ---
ErstelleAlleTabellen              ' Alle Tabellen erstellen (einmalig)
LoescheAlleTabellen               ' ACHTUNG: Löscht alle Daten!

' --- Sync ---
SyncPosteingang                   ' Posteingang synchronisieren (Standard-Einstellungen)
SyncPosteingang "FLIWAS", "Test", 50   ' Max 50 Mails, Projekt=FLIWAS
SyncPosteingang "FLIWAS", "Test", 50, True  ' Mit Unterordnern
SyncOrdner "Torsten.Kugler@rps.bwl.de\Posteingang\FLIWAS", "FLIWAS", "Prod"
SyncOrdner "...\FLIWAS", "FLIWAS", "Prod", 100, True  ' Mit Unterordnern

' --- Postfach-Sync (v0.2) ---
SyncPostfach "Torsten.Kugler@rps.bwl.de", "Standard", "Standard"  ' Ganzes Postfach

' --- Ordnerstruktur ---
SyncOrdnerStruktur 3              ' Alle Ordner bis Tiefe 3 in DB einlesen

' --- Konfiguration ---
? LeseConfig("ExportBasisPfad")
? LeseConfig("LogLevel")
```

---

## 10. Roadmap

| Phase | Beschreibung | Status |
|---|---|---|
| **v0.1** | Schema + Grundmodule (8 .bas Dateien) | ✅ |
| **v0.2** | Kontakt-Erweiterung, Namens-Parsing, Subfolder-Sync, Domain-Lernen (9 Module) | ← AKTUELL |
| v0.3 | CID/Inline-Bild-Auflösung, HTML-Content-Bereinigung | |
| v0.4 | StatusPanel-Integration (Fortschrittsanzeige) | |
| v0.5 | Queue-basierte Verarbeitung (Pause/Abbruch) | |
| v0.6 | Inkrementeller Sync (nur neue Mails seit LetzterSync) | |
| v0.7 | Formular: Sync-Steuerung + Ordnerauswahl | |
| v1.0 | Produktionsreif | |
