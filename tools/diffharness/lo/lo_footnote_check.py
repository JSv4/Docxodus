#!/usr/bin/env /usr/bin/python3
"""Headless LibreOffice footnote-integrity backstop.

Usage: lo_footnote_check.py DOC1.docx [DOC2.docx ...] [--port PORT]

Loads each .docx in an ISOLATED headless soffice (private UserInstallation +
socket, never colliding with the user's open soffice) and reports, per file,
how many footnotes/endnotes LibreOffice actually parsed and the concatenated
note text it loaded. This is the INDEPENDENT validity backstop for the DocxDiff
footnote campaign: a Compare output with duplicate footnote ids or a dangling
footnoteReference makes LibreOffice silently DROP a note (loss) or refuse to
load it (repair). A clean load whose note count + text matches expectation is
the cross-renderer confirmation that the OOXML the engine emitted is sound.

Output is one machine-readable line per file:
    FILE=<path> LOAD=ok FOOTNOTES=<n> ENDNOTES=<m> NOTETEXT=<sha-free joined text>
or, on failure to load:
    FILE=<path> LOAD=FAILED ERR=<message>
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


def note_text(notes):
    parts = []
    for i in range(notes.Count):
        try:
            parts.append(notes.getByIndex(i).getString())
        except Exception:  # noqa: BLE001
            parts.append("?")
    return " | ".join(t.strip() for t in parts if t.strip())


def main():
    port = "2008"
    raw = sys.argv[1:]
    args = []
    skip = False
    for i, a in enumerate(raw):
        if skip:
            skip = False
            continue
        if a == "--port":
            port = raw[i + 1] if i + 1 < len(raw) else port
            skip = True
        elif not a.startswith("--"):
            args.append(a)
    profile = "file:///tmp/lo_profile_fn_%s" % port

    soffice = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--norestore", "--nologo",
        "--nofirststartwizard", "--nodefault",
        "--accept=socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port,
        "-env:UserInstallation=" + profile,
    ])
    rc = 0
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
            except Exception as e:  # noqa: BLE001 - retry until the socket is up
                last_err = e
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("could not connect to soffice: %s" % last_err)

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        for path in args:
            try:
                doc = desktop.loadComponentFromURL(
                    url(path), "_blank", 0, (p("Hidden", True),))
                if doc is None:
                    print("FILE=%s LOAD=FAILED ERR=null-component" % path)
                    rc = 1
                    continue
                fn = doc.Footnotes
                en = doc.Endnotes
                print("FILE=%s LOAD=ok FOOTNOTES=%d ENDNOTES=%d NOTETEXT=%s"
                      % (path, fn.Count, en.Count, note_text(fn) + " || " + note_text(en)))
                doc.close(False)
            except Exception as e:  # noqa: BLE001
                print("FILE=%s LOAD=FAILED ERR=%s" % (path, e))
                rc = 1
    finally:
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(1)
        soffice.terminate()
        try:
            soffice.wait(timeout=10)
        except Exception:  # noqa: BLE001
            soffice.kill()
    sys.exit(rc)


if __name__ == "__main__":
    main()
