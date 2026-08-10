# Sprecherwechsel in Word (wordLangSwitch)

NVDA-Erweiterung, die in Microsoft Word (und im Outlook-Nachrichtenfenster)
bei ausgezeichneten Sprachwechseln im Dokument automatisch die Stimme
wechseln lässt – z. B. bei den Neural-/Natural-Stimmen (OneCore, SAPI 5).

## Das Problem

Im Web erzeugt NVDA aus dem `lang`-Attribut Sprachwechsel-Befehle, worauf der
Synthesizer die passende Stimme wählt. In Word (mit dem heute standardmäßigen
UIA-Zugriff) wird die pro Textabschnitt ausgezeichnete Sprache oft nicht
geliefert, sodass alles mit der Standardstimme gelesen wird.

## Die Lösung

Die Erweiterung ermittelt die im Dokument ausgezeichnete Sprache über das
Word-Objektmodell (`Range.LanguageID`) und speist sie in die Vorlese-Daten
von NVDA ein. NVDA erzeugt daraus Sprachwechsel-Befehle, und der Synthesizer
wechselt die Stimme – genau wie im Web.

## Voraussetzungen

1. NVDA 2023.1 oder neuer.
2. In den **Sprachausgaben-Einstellungen** von NVDA muss
   **„Automatischer Sprachenwechsel (sofern unterstützt)“** aktiviert sein
   (optional auch „Automatischer Dialektwechsel“).
3. Ein Synthesizer, der Sprachwechsel unterstützt und für die gewünschten
   Sprachen Stimmen installiert hat, z. B.:
   * Windows OneCore-Stimmen (inkl. natürlicher Stimmen unter Windows 11),
   * SAPI 5 (ab NVDA 2025.1 mit automatischem Sprachenwechsel),
   * eSpeak NG.
4. Der Text im Word-Dokument muss korrekt mit der Sprache ausgezeichnet sein
   (Überprüfen → Sprache → Sprache für die Korrekturhilfen festlegen).

## Bedienung

* Die Erweiterung arbeitet automatisch beim Lesen mit Pfeiltasten, beim
  satz-/absatzweisen Lesen und bei „Alles Lesen“.
* **NVDA+Alt+W** schaltet die Funktion ein oder aus.
* **NVDA+Alt+K** zeigt eine Diagnose (erkannte Sprachen, Objektmodell,
   Einstellungen) – hilfreich bei der Fehlersuche.
* Beide Tasten sind im „Eingaben“-Dialog änderbar (Kategorie
  „Sprecherwechsel in Word“).

## Einstellungen (NVDA-Einstellungen → „Sprecherwechsel in Word“)

* **Aktivieren/Deaktivieren** der Erweiterung.
* **Dialekte entfernen:** meldet nur die Grundsprache (z. B. `en` statt
  `en_GB`). Empfohlen, wenn der Synthesizer-Treiber exakte Sprachcodes
  erwartet und Dialekte ignoriert (z. B. manche Neural-Voices-Treiber).
* **Sprachzuordnungen:** gezieltes Umschreiben einzelner Sprachcodes,
  Format `quelle=ziel`, mehrere Einträge mit `;` trennen.
  Beispiele: `en_GB=en` oder `en_GB=en_US;fr_CA=fr_FR`.

## Hinweise / Fehlersuche

* Funktioniert der Wechsel nicht, zuerst prüfen, ob er auf einer
  zweisprachigen Webseite funktioniert. Falls nein, liegt es an
  Synthesizer/Stimmen oder an deaktiviertem automatischem Sprachenwechsel.
* Nach der Installation neuer Stimmen NVDA neu starten, damit der Treiber
  seine Stimmenliste neu einliest.
* Alternativ kann in den **erweiterten Einstellungen** von NVDA
  „UI Automation zum Zugriff auf Microsoft Word-Dokumente verwenden“ auf
  **„Nur wenn nötig“** gestellt werden; das klassische Word-Objektmodell
  unterstützt Sprachwechsel von Haus aus.
* Die Erweiterung greift nur, wenn das Word-Objektmodell verfügbar ist
  (nicht in der geschützten Ansicht).

## Lizenz

Diese Erweiterung steht unter der GNU General Public License, Version 2 oder
später (siehe Datei `LICENSE`).
