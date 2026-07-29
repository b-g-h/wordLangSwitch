# -*- coding: utf-8 -*-
# wordLangSwitch: Automatic voice switching for language changes in Microsoft Word
# A global plugin for NVDA.
#
# Copyright (C) 2026 Barbara Henn
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License along
# with this program; if not, see <https://www.gnu.org/licenses/>.
#
# Background:
# On the web, NVDA turns lang attributes into LangChangeCommands, which lets
# synthesizers with neural/natural voices (OneCore, SAPI 5 etc.) switch to a
# matching voice automatically. In Microsoft Word (UIA access, the default
# with modern Office) the per-run language attribute is often missing, so no
# voice switching happens.
#
# This plugin obtains the language marked in the document via the Word object
# model (Range.LanguageID) and injects it as a "language" attribute into the
# format fields NVDA processes while reading. NVDA then generates
# LangChangeCommands and the synthesizer switches voices, provided
# "Automatic language switching" is enabled in NVDA's speech settings.

import config
import globalPluginHandler
import languageHandler
import textInfos
import ui
from comtypes import COMError
from logHandler import log
from scriptHandler import script

try:
	import addonHandler

	addonHandler.initTranslation()
except Exception:
	pass
if "_" not in globals():

	def _(msg):
		return msg


# Word constants
wdCharacter = 1
wdUndefined = 9999999  # mixed formatting/language
wdLanguageNone = 0
wdNoProofing = 1024

# Limits so large documents are not slowed down
MAX_CHUNK_LEN = 8000  # maximum text length of a spoken chunk
MAX_SPLIT_DEPTH = 8  # maximum recursion depth for mixed-language ranges
SEARCH_BACK = 400  # search window before the current selection (characters)
SEARCH_AHEAD = 4000  # search window after the current selection (characters)

confspec = {
	"enabled": "boolean(default=True)",
	"stripDialects": "boolean(default=False)",
	"languageMap": "string(default='')",
}
config.conf.spec["wordLangSwitch"] = confspec


def _mapLocale(locale):
	"""Applies the configured language mapping (e.g. en_GB=en_US)."""
	if not locale:
		return locale
	try:
		mapping = config.conf["wordLangSwitch"]["languageMap"]
		strip = config.conf["wordLangSwitch"]["stripDialects"]
	except KeyError:
		return locale
	if mapping:
		for pair in str(mapping).replace(",", ";").split(";"):
			if "=" not in pair:
				continue
			src, dst = pair.split("=", 1)
			if src.strip().lower().replace("-", "_") == locale.lower().replace("-", "_"):
				# Keep the target untouched so formats like "en-GB" (dash)
				# or plain "en" are possible.
				dst = dst.strip()
				return dst or locale
	if strip:
		return locale.split("_")[0]
	return locale


def _shouldEnrich():
	"""Only active if the plugin is enabled and NVDA should process language changes."""
	try:
		if not config.conf["wordLangSwitch"]["enabled"]:
			return False
	except KeyError:
		pass
	try:
		from speech import languageHandling

		return languageHandling.shouldMakeLangChangeCommand()
	except Exception:
		try:
			return bool(config.conf["speech"]["autoLanguageSwitching"])
		except KeyError:
			return False


def _lcidToLocale(lid):
	"""Converts a Word language ID (LCID) into an NVDA locale name."""
	try:
		lid = int(lid)
	except (TypeError, ValueError):
		return None
	if lid in (wdLanguageNone, wdNoProofing, wdUndefined) or lid <= 0:
		return None
	try:
		return languageHandler.windowsLCIDToLocaleName(lid)
	except Exception:
		return None


def _collectLanguageRuns(rng, out, depth=0):
	"""Recursively splits a Word range into runs of a uniform language.

	Result: list of [length, languageID].
	"""
	try:
		start = rng.start
		end = rng.end
	except COMError:
		return
	length = end - start
	if length <= 0:
		return
	try:
		lid = rng.languageID
	except COMError:
		lid = wdLanguageNone
	if lid == wdUndefined and length > 1 and depth < MAX_SPLIT_DEPTH:
		mid = start + (length // 2)
		left = rng.duplicate
		left.end = mid
		right = rng.duplicate
		right.start = mid
		_collectLanguageRuns(left, out, depth + 1)
		_collectLanguageRuns(right, out, depth + 1)
	else:
		out.append([length, lid])


def _runsFromOmRange(rng, textLen):
	"""Computes language runs for an object model range."""
	rawRuns = []
	_collectLanguageRuns(rng, rawRuns)
	if not rawRuns:
		return None
	runs = []
	prevLocale = None
	for length, lid in rawRuns:
		locale = _mapLocale(_lcidToLocale(lid))
		if locale is None:
			# Runs without a language (e.g. "no proofing") inherit the
			# previous language so the voice does not flip back and forth.
			locale = prevLocale
		if runs and runs[-1][1] == locale:
			runs[-1][0] += length
		else:
			runs.append([length, locale])
		prevLocale = locale
	# If the object model range was shorter than the spoken text,
	# stretch the last run.
	total = sum(r[0] for r in runs)
	if total < textLen and runs:
		runs[-1][0] += textLen - total
	return runs


def _omRangeFromTextInfo(ti):
	"""With the classic object model text info the range is directly available."""
	try:
		from NVDAObjects.window.winword import WordDocumentTextInfo as OMWordTextInfo
	except Exception:
		return None
	if isinstance(ti, OMWordTextInfo):
		try:
			return ti._rangeObj.duplicate
		except (COMError, AttributeError):
			return None
	return None


def _findOmRangeForText(ti, text):
	"""Finds the object model range matching the text about to be spoken.

	Strategy 0: classic object model text info (range directly available).
	Strategy 1: text search in a window around the current selection
	(works for line/sentence/paragraph reading and say all).
	Strategy 2: screen coordinates of the chunk start (RangeFromPoint).
	"""
	rng = _omRangeFromTextInfo(ti)
	if rng is not None:
		return rng
	obj = ti.obj
	try:
		sel = obj.WinwordSelectionObject
	except (COMError, AttributeError):
		sel = None
	keys = [text]
	stripped = text.rstrip("\r\n\x07 ")
	if stripped and stripped != text:
		keys.append(stripped)
	if "\n" in text:
		keys.append(text.replace("\n", "\r"))
	keys = [k for k in keys if k.strip()]
	if not keys:
		return None
	if sel:
		try:
			selRange = sel.range
			window = selRange.duplicate
			window.collapse()
			window.moveStart(wdCharacter, -SEARCH_BACK)
			window.moveEnd(wdCharacter, SEARCH_AHEAD + len(text))
			wtext = window.text or ""
			winStart = window.start
			target = max(0, selRange.start - winStart)
			for key in keys:
				# Prefer a match close to the selection
				idx = wtext.find(key, max(0, target - len(key) - 100))
				if idx < 0:
					idx = wtext.find(key)
				if idx >= 0:
					rng = window.duplicate
					rng.start = winStart + idx
					rng.end = winStart + idx + len(key)
					return rng
		except COMError:
			pass
	# Strategy 2: via screen coordinates
	try:
		point = ti.pointAtStart
		om = obj.WinwordWindowObject
		if om is None:
			return None
		rng = om.rangeFromPoint(point.x, point.y)
		rng = rng.duplicate
		rng.moveEnd(wdCharacter, len(keys[0]))
		omText = rng.text or ""
		probe = keys[0][:10]
		if probe.strip() and not omText.replace("\x07", "").startswith(probe):
			return None
		return rng
	except (COMError, LookupError, AttributeError, NameError):
		return None


def _applyRuns(fields, runs):
	"""Injects the computed language runs into the field list."""
	out = []
	runIdx = 0
	remaining, curLang = runs[0]
	lastFormat = None
	langPending = True
	for item in fields:
		if (
			isinstance(item, textInfos.FieldCommand)
			and item.command == "formatChange"
			and isinstance(item.field, dict)
		):
			if curLang:
				item.field["language"] = curLang
			else:
				item.field.pop("language", None)
			lastFormat = item.field
			langPending = False
			out.append(item)
		elif isinstance(item, str) and item:
			pos = 0
			while pos < len(item):
				if remaining <= 0 and runIdx + 1 < len(runs):
					runIdx += 1
					remaining, newLang = runs[runIdx]
					if newLang != curLang:
						curLang = newLang
						langPending = True
					continue
				if langPending:
					f = textInfos.FormatField()
					if lastFormat:
						f.update(lastFormat)
					if curLang:
						f["language"] = curLang
					else:
						f.pop("language", None)
					out.append(textInfos.FieldCommand("formatChange", f))
					lastFormat = f
					langPending = False
				take = len(item) - pos if remaining <= 0 else min(len(item) - pos, remaining)
				out.append(item[pos : pos + take])
				remaining -= take
				pos += take
		else:
			out.append(item)
	return out


def _mapFieldLanguages(fields):
	"""Applies language mapping/dialect stripping to all existing language fields.

	Important for language info Word itself provides via UIA.
	"""
	changed = False
	for item in fields:
		if (
			isinstance(item, textInfos.FieldCommand)
			and item.command == "formatChange"
			and isinstance(item.field, dict)
		):
			lang = item.field.get("language")
			if lang:
				mapped = _mapLocale(lang)
				if mapped != lang:
					item.field["language"] = mapped
					changed = True
	return changed


def _enrichFields(ti, fields):
	"""Enriches the field list with language information.

	Returns (fields, reason); reason describes the outcome for diagnostics/logging.
	"""
	text = "".join(item for item in fields if isinstance(item, str))
	if not text.strip():
		return fields, "empty text"
	if len(text) > MAX_CHUNK_LEN:
		_mapFieldLanguages(fields)
		return fields, "text too long"
	rng = _findOmRangeForText(ti, text)
	if rng is None:
		mapped = _mapFieldLanguages(fields)
		return fields, "object model range not found%s" % (
			", existing languages mapped" if mapped else ""
		)
	runs = _runsFromOmRange(rng, len(text))
	if not runs:
		_mapFieldLanguages(fields)
		return fields, "no language runs found"
	if all(r[1] is None for r in runs):
		mapped = _mapFieldLanguages(fields)
		return fields, "no language markup in text%s" % (
			", existing languages mapped" if mapped else ""
		)
	fields = _applyRuns(fields, runs)
	_mapFieldLanguages(fields)
	return fields, "enriched: %r" % (runs,)


try:
	import wx
	import gui
	from gui.settingsDialogs import SettingsPanel

	class WordLangSwitchSettingsPanel(SettingsPanel):
		# Translators: Title of the add-on's settings panel.
		title = _("Word language switching")

		def makeSettings(self, settingsSizer):
			from gui import guiHelper

			helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
			self.enabledCheckbox = helper.addItem(
				wx.CheckBox(
					self,
					# Translators: Label of a checkbox in the add-on's settings panel.
					label=_("&Enable voice switching on language changes in Word"),
				),
			)
			self.enabledCheckbox.SetValue(config.conf["wordLangSwitch"]["enabled"])
			self.stripDialectsCheckbox = helper.addItem(
				wx.CheckBox(
					self,
					# Translators: Label of a checkbox in the add-on's settings panel.
					label=_("Strip &dialects (e.g. report en_GB as plain en to the synthesizer)"),
				),
			)
			self.stripDialectsCheckbox.SetValue(config.conf["wordLangSwitch"]["stripDialects"])
			self.languageMapEdit = helper.addLabeledControl(
				# Translators: Label of an edit field in the add-on's settings panel.
				_("Language &mappings (format: en_GB=en_US;fr_CA=fr_FR):"),
				wx.TextCtrl,
			)
			self.languageMapEdit.SetValue(config.conf["wordLangSwitch"]["languageMap"])

		def onSave(self):
			config.conf["wordLangSwitch"]["enabled"] = self.enabledCheckbox.IsChecked()
			config.conf["wordLangSwitch"]["stripDialects"] = self.stripDialectsCheckbox.IsChecked()
			config.conf["wordLangSwitch"]["languageMap"] = self.languageMapEdit.GetValue().strip()
except Exception:
	WordLangSwitchSettingsPanel = None

_patchedClasses = {}
_stats = {"calls": 0}
MAX_INFO_LOGS = 25


def _makePatched(orig):
	def patchedGetTextWithFields(self, formatConfig=None):
		fields = orig(self, formatConfig)
		if not _shouldEnrich():
			return fields
		try:
			fields, reason = _enrichFields(self, fields)
		except Exception:
			log.warning("wordLangSwitch: enrichment failed", exc_info=True)
			return fields
		_stats["calls"] += 1
		if _stats["calls"] <= MAX_INFO_LOGS:
			log.info("wordLangSwitch: call %d: %s" % (_stats["calls"], reason))
		return fields

	return patchedGetTextWithFields


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	# Translators: Name of the add-on's category in the input gestures dialog.
	scriptCategory = _("Word language switching")

	def __init__(self):
		super().__init__()
		self._patch()
		if WordLangSwitchSettingsPanel is not None:
			try:
				gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(
					WordLangSwitchSettingsPanel,
				)
			except Exception:
				log.debugWarning("wordLangSwitch: settings panel not registered", exc_info=True)

	def _patch(self):
		targets = []
		try:
			from NVDAObjects.UIA import wordDocument as uiaWordDocument

			targets.append(("UIA", uiaWordDocument.WordDocumentTextInfo))
		except Exception:
			log.warning("wordLangSwitch: UIA Word class not available", exc_info=True)
		try:
			from NVDAObjects.window import winword

			targets.append(("object model", winword.WordDocumentTextInfo))
		except Exception:
			log.warning("wordLangSwitch: object model Word class not available", exc_info=True)
		for label, cls in targets:
			if cls in _patchedClasses:
				continue
			try:
				_patchedClasses[cls] = cls.getTextWithFields
				cls.getTextWithFields = _makePatched(_patchedClasses[cls])
				log.info("wordLangSwitch: patched %s Word text info (%s)" % (label, cls.__name__))
			except Exception:
				_patchedClasses.pop(cls, None)
				log.error("wordLangSwitch: patching failed (%s)" % label, exc_info=True)

	def terminate(self):
		for cls, orig in list(_patchedClasses.items()):
			try:
				cls.getTextWithFields = orig
			except Exception:
				log.debugWarning("wordLangSwitch: unpatching failed", exc_info=True)
			_patchedClasses.pop(cls, None)
		if WordLangSwitchSettingsPanel is not None:
			try:
				gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(
					WordLangSwitchSettingsPanel,
				)
			except Exception:
				pass
		super().terminate()

	@script(
		# Translators: Description of a script in the input gestures dialog.
		description=_("Shows diagnostic information about language switching in Word"),
		gesture="kb:NVDA+shift+k",
	)
	def script_diagnose(self, gesture):
		import api

		msgs = []
		obj = api.getFocusObject()
		msgs.append("Focus object: %s.%s" % (obj.__class__.__module__, obj.__class__.__qualname__))
		try:
			import synthDriverHandler

			synth = synthDriverHandler.getSynth()
			msgs.append("Synthesizer: %s" % (synth.name if synth else "none"))
		except Exception as e:
			msgs.append("Synthesizer: error %r" % e)
		try:
			msgs.append(
				"Automatic language switching: %s" % config.conf["speech"]["autoLanguageSwitching"],
			)
		except KeyError:
			msgs.append("Automatic language switching: setting not found")
		msgs.append("Plugin enabled: %s" % config.conf["wordLangSwitch"]["enabled"])
		msgs.append("Strip dialects: %s" % config.conf["wordLangSwitch"]["stripDialects"])
		msgs.append("Language mappings: %r" % config.conf["wordLangSwitch"]["languageMap"])
		msgs.append("Test mapping en_GB -> %s" % _mapLocale("en_GB"))
		msgs.append("Enrichment active (_shouldEnrich): %s" % _shouldEnrich())
		msgs.append("Patched classes: %s" % ", ".join(c.__name__ for c in _patchedClasses) or "none")
		ti = None
		try:
			ti = obj.makeTextInfo(textInfos.POSITION_CARET)
			ti.expand(textInfos.UNIT_LINE)
		except Exception as e:
			msgs.append("Could not create text info: %r" % e)
		if ti is not None:
			msgs.append("TextInfo class: %s.%s" % (ti.__class__.__module__, ti.__class__.__qualname__))
			msgs.append(
				"getTextWithFields patched: %s"
				% (getattr(ti.getTextWithFields, "__name__", "?") == "patchedGetTextWithFields"),
			)
			for name in ("WinwordSelectionObject", "WinwordDocumentObject", "WinwordWindowObject"):
				try:
					v = getattr(obj, name, None)
					msgs.append("%s: %s" % (name, "OK" if v is not None else "MISSING"))
				except Exception as e:
					msgs.append("%s: error %r" % (name, e))
			try:
				fields = ti.getTextWithFields()
				text = "".join(i for i in fields if isinstance(i, str))
				msgs.append("Line text (%d characters): %r" % (len(text), text[:100]))
				langs = [
					f.field.get("language")
					for f in fields
					if isinstance(f, textInfos.FieldCommand)
					and f.command == "formatChange"
					and isinstance(f.field, dict)
				]
				msgs.append("Languages in format fields (final): %r" % (langs,))
				orig = _patchedClasses.get(ti.__class__)
				if orig is not None:
					origLangs = [
						f.field.get("language")
						for f in orig(ti)
						if isinstance(f, textInfos.FieldCommand)
						and f.command == "formatChange"
						and isinstance(f.field, dict)
					]
					msgs.append("Languages from Word/UIA (unmodified): %r" % (origLangs,))
				rng = _findOmRangeForText(ti, text)
				if rng is None:
					msgs.append("Object model range: not found")
				else:
					msgs.append("Object model range text: %r" % ((rng.text or "")[:100],))
					msgs.append("Language runs: %r" % (_runsFromOmRange(rng, len(text)),))
			except Exception as e:
				log.error("wordLangSwitch: diagnostics error", exc_info=True)
				msgs.append("Error during analysis: %r" % e)
		message = "\n".join(msgs)
		log.info("wordLangSwitch diagnostics:\n%s" % message)
		try:
			# Translators: Title of the diagnostics browseable message.
			ui.browseableMessage(message, _("Word language switching diagnostics"))
		except Exception:
			# Translators: Spoken if the diagnostics window could not be shown.
			ui.message(_("Diagnostics written to the NVDA log"))

	@script(
		# Translators: Description of a script in the input gestures dialog.
		description=_("Toggles automatic voice switching for language changes in Microsoft Word"),
		gesture="kb:NVDA+shift+l",
	)
	def script_toggleWordLangSwitch(self, gesture):
		newValue = not config.conf["wordLangSwitch"]["enabled"]
		config.conf["wordLangSwitch"]["enabled"] = newValue
		_stats["calls"] = 0
		if newValue:
			# Translators: Spoken when the add-on has been enabled.
			ui.message(_("Word language switching on"))
		else:
			# Translators: Spoken when the add-on has been disabled.
			ui.message(_("Word language switching off"))
