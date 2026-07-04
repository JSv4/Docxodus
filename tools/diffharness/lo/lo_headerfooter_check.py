#!/usr/bin/env /usr/bin/python3
"""Headless LibreOffice header/footer oracle.

Usage: lo_headerfooter_check.py DOC.docx [EXPECT ...] [--absent TEXT ...] [--port PORT]

Loads DOC in an ISOLATED headless LibreOffice (private UserInstallation + socket)
and verifies, INDEPENDENTLY of our own XML checks, that:

  1. The document LOADS without error.
  2. The header/footer stories LibreOffice renders (across every page style:
     shared, left, and first-page variants) contain every EXPECT substring —
     i.e. a real renderer surfaces the story content our diff produced.
  3. No harvested story contains any --absent substring (e.g. text a redline's
     accept view must no longer show).

Output is one machine-readable summary line plus per-failure lines.
Exit 0 + "RESULT: OK" on clean; exit 1 + per-failure lines otherwise.
"""
import os
import subprocess
import sys
import time
import uno
from com.sun.star.beans import PropertyValue


def p(name, value):
    pv = PropertyValue()
    pv.Name = name
    pv.Value = value
    return pv


def url(path):
    return "file://" + os.path.abspath(path)


def harvest_stories(doc):
    """Concatenated text of every header/footer story of every page style."""
    stories = {}
    styles = doc.getStyleFamilies().getByName("PageStyles")
    for i in range(styles.getCount()):
        style = styles.getByIndex(i)
        try:
            if not style.isInUse():
                continue
        except Exception:  # noqa: BLE001
            pass
        for prop in ("HeaderText", "HeaderTextLeft", "HeaderTextFirst",
                     "FooterText", "FooterTextLeft", "FooterTextFirst"):
            try:
                text_obj = style.getPropertyValue(prop)
                if text_obj is not None:
                    text = text_obj.getString()
                    if text and text.strip():
                        stories["%s.%s" % (style.Name, prop)] = text
            except Exception:  # noqa: BLE001
                pass
    return stories


def main():
    doc_path = sys.argv[1]
    rest = sys.argv[2:]
    expects, absents, port = [], [], "2013"
    mode = "expect"
    i = 0
    while i < len(rest):
        a = rest[i]
        if a == "--absent":
            mode = "absent"
        elif a == "--port":
            i += 1
            port = rest[i]
        elif mode == "expect":
            expects.append(a)
        else:
            absents.append(a)
        i += 1
    profile = "file:///tmp/lo_profile_hf_%s" % port

    soffice = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--norestore", "--nologo",
        "--nofirststartwizard", "--nodefault",
        "--accept=socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port,
        "-env:UserInstallation=" + profile,
    ])
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        last_err = None
        for _ in range(60):
            try:
                ctx = resolver.resolve(
                    "uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port)
                break
            except Exception as e:  # noqa: BLE001
                last_err = e
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("could not connect to soffice: %s" % last_err)

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            url(doc_path), "_blank", 0, (p("Hidden", True),))
        if doc is None:
            print("doc=%s LOAD=FAILED" % os.path.basename(doc_path))
            return 1

        stories = harvest_stories(doc)
        blob = "\n".join(stories.values())
        for key, text in sorted(stories.items()):
            print("story %s: %r" % (key, text[:120]))

        failures = []
        for want in expects:
            if want not in blob:
                failures.append("expected header/footer text %r not rendered" % want)
        for bad in absents:
            if bad in blob:
                failures.append("text %r must NOT appear in any header/footer" % bad)

        print("doc=%s stories=%d" % (os.path.basename(doc_path), len(stories)))
        if failures:
            for fail in failures[:40]:
                print("  FAIL: " + fail)
            print("RESULT: FAIL (%d issue(s))" % len(failures))
            doc.close(False)
            return 1
        print("RESULT: OK")
        doc.close(False)
        return 0
    finally:
        try:
            subprocess.run(["soffice", "--headless",
                            "-env:UserInstallation=" + profile,
                            "--unaccept=socket,host=localhost,port=%s;urp;" % port],
                           timeout=10)
        except Exception:
            pass
        soffice.terminate()


if __name__ == "__main__":
    sys.exit(main())
