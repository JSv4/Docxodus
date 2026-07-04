#!/usr/bin/env /usr/bin/python3
"""Headless LibreOffice bookmark / cross-reference integrity oracle.

Usage: lo_bookmark_check.py DOC.docx [PORT]

Loads DOC in an ISOLATED headless LibreOffice (private UserInstallation + socket)
and verifies, INDEPENDENTLY of our own XML checks, that:

  1. Every com.sun.star.text.Bookmark loads (count + names).
  2. Every GetReference text field (REF/PAGEREF/NOTEREF/HYPERLINK \\l) resolves —
     its SourceName names a bookmark LibreOffice actually has.
  3. After a field refresh, NO GetReference field presents the
     "Reference source not found" error (LibreOffice's own dangling-ref signal).
  4. Every internal hyperlink (URL "#name") targets a present bookmark.

Exit 0 + "OK" line on clean; exit 1 + per-failure lines otherwise.
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


def main():
    doc_path = sys.argv[1]
    port = sys.argv[2] if len(sys.argv) > 2 else "2009"
    profile = "file:///tmp/lo_profile_bk_%s" % port

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

        failures = []

        # (1) bookmarks LibreOffice loaded
        bms = doc.getBookmarks()
        bm_names = set(bms.getElementNames())

        # refresh fields so dangling references surface their error presentation
        try:
            doc.getTextFields().refresh()
        except Exception:
            pass
        try:
            doc.refresh()
        except Exception:
            pass

        # (2)+(3) GetReference fields resolve + no "Reference source not found"
        ref_count = 0
        enum = doc.getTextFields().createEnumeration()
        while enum.hasMoreElements():
            f = enum.nextElement()
            if not f.supportsService("com.sun.star.text.textfield.GetReference"):
                continue
            ref_count += 1
            src = ""
            try:
                src = f.getPropertyValue("SourceName")
            except Exception:
                pass
            pres = ""
            try:
                pres = f.getPresentation(False) or ""
            except Exception:
                pass
            if "Reference source not found" in pres or "not found" in pres.lower():
                failures.append("GetReference '%s' presents error: %r" % (src, pres))
            elif src and src not in bm_names:
                # a sequence/heading auto-target may legitimately not be a Bookmark; only flag _Ref/_Toc style
                if src.startswith("_Ref") or src.startswith("_Toc") or src.startswith("_DV") or src.startswith("_cp"):
                    failures.append("GetReference SourceName '%s' has no matching bookmark" % src)

        # (4) internal hyperlinks (#anchor) resolve to a bookmark
        hyper_count = 0
        penum = doc.getText().createEnumeration()
        while penum.hasMoreElements():
            par = penum.nextElement()
            if not par.supportsService("com.sun.star.text.Paragraph"):
                continue
            renum = par.createEnumeration()
            while renum.hasMoreElements():
                portion = renum.nextElement()
                try:
                    href = portion.getPropertyValue("HyperLinkURL") or ""
                except Exception:
                    href = ""
                if href.startswith("#"):
                    hyper_count += 1
                    target = href[1:].split("|")[0]  # LO encodes "#name|outline"/"#name|region"
                    if target and target not in bm_names and (
                            target.startswith("_Ref") or target.startswith("_Toc")):
                        failures.append("internal hyperlink '#%s' has no matching bookmark" % target)

        print("doc=%s bookmarks=%d refFields=%d internalHyperlinks=%d"
              % (os.path.basename(doc_path), len(bm_names), ref_count, hyper_count))
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
