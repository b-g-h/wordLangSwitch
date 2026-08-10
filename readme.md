# Word Language Switching (wordLangSwitch)

* Authors: Barbara Henn
* Download [stable version][1]
* NVDA compatibility: 2023.1 and later

NVDA add-on that automatically switches the synthesizer voice for marked
language changes in Microsoft Word (and the Outlook message window) – for
example with neural/natural voices (OneCore, SAPI 5).

## The problem

On the web, NVDA turns `lang` attributes into language change commands, and
the synthesizer picks a matching voice. In Word (with the UIA access that is
the default for modern Office versions), the language marked per text run is
often not delivered, so everything is read with the default voice.

## The solution

The add-on obtains the language marked in the document via the Word object
model (`Range.LanguageID`) and feeds it into NVDA's reading data. NVDA then
generates language change commands and the synthesizer switches voices –
exactly as on the web.

## Requirements

1. NVDA 2023.1 or later.
2. **"Automatic language switching (when supported)"** must be enabled in
   NVDA's **speech settings** (optionally also "Automatic dialect switching").
3. A synthesizer that supports language switching and has voices installed
   for the desired languages, e.g.:
   * Windows OneCore voices (including natural voices on Windows 11),
   * SAPI 5 (automatic language switching since NVDA 2025.1),
   * eSpeak NG.
4. The text in the Word document must be correctly marked with its language
   (Review → Language → Set Proofing Language).

## Usage

* The add-on works automatically when reading with the arrow keys, by
  sentence/paragraph, and during say all.
* **NVDA+alt+w** toggles the add-on on or off.
* **NVDA+alt+k** shows diagnostics (detected languages, object model,
  settings) – helpful for troubleshooting.
* Both gestures can be changed in the Input Gestures dialog (category
  "Word language switching").

## Settings (NVDA settings → "Word language switching")

* **Enable/disable** the add-on.
* **Strip dialects:** report only the base language (e.g. `en` instead of
  `en_GB`). Recommended if the synthesizer driver expects exact language
  codes and ignores dialects (e.g. some neural voice drivers).
* **Language mappings:** rewrite individual language codes, format
  `source=target`, separate multiple entries with `;`.
  Examples: `en_GB=en` or `en_GB=en_US;fr_CA=fr_FR`.

## Notes / troubleshooting

* If switching does not work, first check whether it works on a bilingual
  web page. If not, the cause is the synthesizer/voices or automatic
  language switching being disabled.
* After installing new voices, restart NVDA so the driver reloads its voice
  list.
* Alternatively, "Use UI Automation to access Microsoft Word document
  controls" can be set to **"Only when necessary"** in NVDA's advanced
  settings; the classic Word object model supports language switching
  natively.
* The add-on only takes effect when the Word object model is available
  (not in protected view).

## Building from source

The add-on is built with SCons, using the project setup from the current NVDA
add-on template.

Prerequisites:

* [`uv`](https://docs.astral.sh/uv/)
* gettext tools (`xgettext`, `msgfmt`) on your `PATH`

Build:

    uv sync
    uv run scons

This creates `wordLangSwitch-<version>.nvda-addon` in the root folder.
Run `uv run scons pot` to generate the translation template (`.pot`).

Alternatively, with your own Python environment:

    pip install scons Markdown
    scons

## License

This add-on is licensed under the GNU General Public License, version 2 or
later (see the `LICENSE` file).

[1]: https://github.com/b-g-h/wordLangSwitch/releases/latest
