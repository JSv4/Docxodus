#!/usr/bin/env /usr/bin/python3
"""Headless LibreOffice comment integrity oracle.

Usage: lo_comment_check.py DOC.docx [PORT]

Loads DOC in an ISOLATED headless LibreOffice (private UserInstallation + socket)
and verifies, INDEPENDENTLY of our own XML checks, that:

  1. The document LOADS without error.
  2. Every comment (com.sun.star.text.textfield.Annotation) LibreOffice parsed has
     an Author + is anchored in the text (a dangling w:commentReference whose
     w:comment definition is missing would otherwise drop the annotation or leave
     it unanchored — LibreOffice's own integrity signal).
  3. After a field + document refresh, the comment set is unchanged (no comment
     silently dropped / orphaned on refresh).
  4. Every THREADED reply (an Annotation with a ParentName) names a parent comment
     LibreOffice actually loaded — so commentsExtended reply links survive.

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


def collect_comments(doc):
    """Return list of dicts {name, author, content, parent, anchored} for every Annotation field."""
    out = []
    enum = doc.getTextFields().createEnumeration()
    while enum.hasMoreElements():
        f = enum.nextElement()
        if not f.supportsService("com.sun.star.text.textfield.Annotation"):
            continue
        rec = {"name": "", "author": "", "content": "", "parent": "", "anchored": False}
        for prop, key in (("Name", "name"), ("Author", "author"),
                          ("Content", "content"), ("ParentName", "parent")):
            try:
                rec[key] = f.getPropertyValue(prop) or ""
            except Exception:  # noqa: BLE001
                pass
        try:
            rec["anchored"] = f.getAnchor() is not None
        except Exception:  # noqa: BLE001
            pass
        out.append(rec)
    return out


def main():
    doc_path = sys.argv[1]
    port = sys.argv[2] if len(sys.argv) > 2 else "2011"
    profile = "file:///tmp/lo_profile_cmt_%s" % port

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

        failures = []
        before = collect_comments(doc)
        names = {c["name"] for c in before if c["name"]}

        # (2) every loaded comment has an author + is anchored
        for c in before:
            if not c["author"]:
                failures.append("comment %r has no author" % (c["name"] or c["content"][:20]))
            if not c["anchored"]:
                failures.append("comment %r is not anchored in the text" % (c["name"] or c["content"][:20]))

        # (4) every threaded reply names a parent comment that loaded
        for c in before:
            if c["parent"] and c["parent"] not in names:
                failures.append("reply %r names missing parent %r" % (c["name"], c["parent"]))

        # (3) refresh: the comment set must be unchanged
        try:
            doc.getTextFields().refresh()
        except Exception:
            pass
        try:
            doc.refresh()
        except Exception:
            pass
        after = collect_comments(doc)
        if len(after) != len(before):
            failures.append("comment count changed on refresh: %d -> %d" % (len(before), len(after)))

        threads = sum(1 for c in before if c["parent"])
        print("doc=%s comments=%d threadedReplies=%d authors=%d"
              % (os.path.basename(doc_path), len(before), threads,
                 len({c["author"] for c in before if c["author"]})))
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
