# Dependency-free build script for the wordLangSwitch NVDA add-on.
# Alternative to the SCons build from the NVDA add-on template:
# generates manifests from buildVars.py, compiles .po translations to .mo,
# converts readme.md files to HTML documentation and packs the .nvda-addon.

import gettext
import glob
import html
import os
import re
import struct
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

try:
	import buildVars  # noqa: E402
except ImportError:
	# The template helpers pull in SCons; provide light-weight stand-ins
	# so buildVars.py can be imported without SCons being installed.
	import types

	for modName in ("site_scons", "site_scons.site_tools", "site_scons.site_tools.NVDATool"):
		sys.modules.setdefault(modName, types.ModuleType(modName))
	typings = types.ModuleType("site_scons.site_tools.NVDATool.typings")
	typings.AddonInfo = dict
	typings.BrailleTables = dict
	typings.SymbolDictionaries = dict
	utils = types.ModuleType("site_scons.site_tools.NVDATool.utils")
	utils._ = lambda arg: arg
	sys.modules["site_scons.site_tools.NVDATool.typings"] = typings
	sys.modules["site_scons.site_tools.NVDATool.utils"] = utils

	import buildVars  # noqa: E402

ADDON_DIR = os.path.join(HERE, "addon")
info = dict(buildVars.addon_info)


def compilePo(poPath, moPath):
	"""Minimal msgfmt: compiles a .po file into a .mo file."""
	messages = {}
	state = {"msgid": None, "msgstr": None, "section": None, "fuzzy": False, "pendingFuzzy": False}

	def add():
		if state["msgid"] is not None and state["msgstr"] is not None:
			if not (state["fuzzy"] and state["msgid"]):
				messages[state["msgid"]] = state["msgstr"]

	with open(poPath, encoding="utf-8") as f:
		for line in f:
			line = line.strip()
			if line.startswith("#,") and "fuzzy" in line:
				state["pendingFuzzy"] = True
				continue
			if not line or line.startswith("#"):
				continue
			if line.startswith("msgid "):
				add()
				state["fuzzy"] = state["pendingFuzzy"]
				state["pendingFuzzy"] = False
				state["msgid"] = eval(line[6:], {}, {})
				state["msgstr"] = None
				state["section"] = "id"
				continue
			if line.startswith("msgstr "):
				state["msgstr"] = eval(line[7:], {}, {})
				state["section"] = "str"
				continue
			if line.startswith('"'):
				value = eval(line, {}, {})
				if state["section"] == "id":
					state["msgid"] += value
				elif state["section"] == "str":
					state["msgstr"] += value
	add()
	keys = sorted(messages)
	entries = []
	for key in keys:
		entries.append((key.encode("utf-8"), messages[key].encode("utf-8")))
	keyStart = 7 * 4 + 16 * len(entries)
	valueStart = keyStart + sum(len(k) + 1 for k, v in entries)
	kOffsets, vOffsets = [], []
	kData = vData = b""
	for k, v in entries:
		kOffsets.append((len(k), keyStart + len(kData)))
		kData += k + b"\x00"
		vOffsets.append((len(v), valueStart + len(vData)))
		vData += v + b"\x00"
	output = struct.pack(
		"Iiiiiii",
		0x950412DE,  # magic
		0,  # version
		len(entries),
		7 * 4,  # offset of key table
		7 * 4 + 8 * len(entries),  # offset of value table
		0,
		0,
	)
	for length, offset in kOffsets:
		output += struct.pack("ii", length, offset)
	for length, offset in vOffsets:
		output += struct.pack("ii", length, offset)
	output += kData + vData
	os.makedirs(os.path.dirname(moPath), exist_ok=True)
	with open(moPath, "wb") as f:
		f.write(output)


def mdToHtml(mdPath, htmlPath, title, lang):
	"""Converts markdown to HTML; uses the markdown package when available."""
	with open(mdPath, encoding="utf-8") as f:
		mdText = f.read()
	try:
		import markdown

		body = markdown.markdown(mdText, extensions=buildVars.markdownExtensions)
	except ImportError:
		body = _miniMarkdown(mdText)
	doc = (
		'<!DOCTYPE html>\n<html lang="%s">\n<head>\n<meta charset="utf-8">\n'
		'<link rel="stylesheet" type="text/css" href="../style.css" media="screen">\n'
		"<title>%s</title>\n</head>\n<body>\n%s\n</body>\n</html>\n"
	) % (lang, html.escape(title), body)
	os.makedirs(os.path.dirname(htmlPath), exist_ok=True)
	with open(htmlPath, "w", encoding="utf-8") as f:
		f.write(doc)


def _miniMarkdown(text):
	"""Very small markdown subset: headings, unordered/ordered lists, paragraphs."""
	outLines = []
	listStack = []
	paragraph = []

	def flushParagraph():
		if paragraph:
			outLines.append("<p>%s</p>" % _inline(" ".join(paragraph)))
			paragraph.clear()

	def closeLists(depth=0):
		while len(listStack) > depth:
			outLines.append("</%s>" % listStack.pop())

	def _inline(s):
		s = html.escape(s, quote=False)
		s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
		s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
		s = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', s)
		s = re.sub(r"\[([^\]]+)\]\[\d+\]", r"\1", s)
		return s

	for line in text.splitlines():
		heading = re.match(r"^(#{1,6})\s+(.*)$", line)
		bullet = re.match(r"^\s*[*+-]\s+(.*)$", line)
		numbered = re.match(r"^\s*\d+\.\s+(.*)$", line)
		if heading:
			flushParagraph()
			closeLists()
			level = len(heading.group(1))
			outLines.append("<h%d>%s</h%d>" % (level, _inline(heading.group(2)), level))
		elif bullet or numbered:
			flushParagraph()
			tag = "ul" if bullet else "ol"
			if not listStack or listStack[-1] != tag:
				closeLists()
				listStack.append(tag)
				outLines.append("<%s>" % tag)
			outLines.append("<li>%s</li>" % _inline((bullet or numbered).group(1)))
		elif not line.strip():
			flushParagraph()
			closeLists()
		else:
			if listStack:
				# continuation line of a list item
				outLines[-1] = outLines[-1][:-5] + " " + _inline(line.strip()) + "</li>"
			else:
				paragraph.append(line.strip())
	flushParagraph()
	closeLists()
	return "\n".join(outLines)


def main():
	name = info["addon_name"]
	version = info["addon_version"]
	buildDir = os.path.join(HERE, "build")
	os.makedirs(buildDir, exist_ok=True)

	# 1. Manifest from template
	with open(os.path.join(HERE, "manifest.ini.tpl"), encoding="utf-8") as f:
		manifest = f.read().format(**info)
	manifestPath = os.path.join(buildDir, "manifest.ini")
	with open(manifestPath, "w", encoding="utf-8") as f:
		f.write(manifest)

	# 2. Compile translations and generate translated manifests
	with open(os.path.join(HERE, "manifest-translated.ini.tpl"), encoding="utf-8") as f:
		translatedTpl = f.read()
	translatedManifests = {}
	moFiles = {}
	for poPath in glob.glob(os.path.join(ADDON_DIR, "locale", "*", "LC_MESSAGES", "nvda.po")):
		lang = os.path.basename(os.path.dirname(os.path.dirname(poPath)))
		moPath = os.path.join(buildDir, lang, "nvda.mo")
		compilePo(poPath, moPath)
		moFiles[lang] = moPath
		with open(moPath, "rb") as f:
			translations = gettext.GNUTranslations(f)
		localized = dict(info)
		for key in ("addon_summary", "addon_description", "addon_changelog"):
			localized[key] = translations.gettext(info[key])
		translatedManifests[lang] = translatedTpl.format(**localized)

	# 3. Documentation from markdown
	docs = []  # (zipPath, fsPath)
	enHtml = os.path.join(buildDir, "en_readme.html")
	mdToHtml(os.path.join(HERE, "readme.md"), enHtml, info["addon_summary"], "en")
	docs.append(("doc/en/readme.html", enHtml))
	for mdPath in glob.glob(os.path.join(ADDON_DIR, "doc", "*", "readme.md")):
		lang = os.path.basename(os.path.dirname(mdPath))
		langHtml = os.path.join(buildDir, "%s_readme.html" % lang)
		mdToHtml(mdPath, langHtml, info["addon_summary"], lang)
		docs.append(("doc/%s/readme.html" % lang, langHtml))
	docs.append(("doc/style.css", os.path.join(HERE, "style.css")))

	# 4. Pack the add-on
	outFile = os.path.join(HERE, "%s-%s.nvda-addon" % (name, version))
	with zipfile.ZipFile(outFile, "w", zipfile.ZIP_DEFLATED) as z:
		z.write(manifestPath, "manifest.ini")
		z.write(os.path.join(HERE, "LICENSE"), "LICENSE")
		for root, dirs, files in os.walk(ADDON_DIR):
			for fileName in files:
				if fileName.endswith((".pyc", ".html")) or fileName == "readme.md":
					continue
				fullPath = os.path.join(root, fileName)
				rel = os.path.relpath(fullPath, ADDON_DIR).replace(os.sep, "/")
				z.write(fullPath, rel)
		for lang, moPath in moFiles.items():
			z.write(moPath, "locale/%s/LC_MESSAGES/nvda.mo" % lang)
		for lang, content in translatedManifests.items():
			z.writestr("locale/%s/manifest.ini" % lang, content)
		for zipPath, fsPath in docs:
			z.write(fsPath, zipPath)
	print("Created: %s" % outFile)
	with zipfile.ZipFile(outFile) as z:
		for n in sorted(z.namelist()):
			print("  %s" % n)


if __name__ == "__main__":
	main()
