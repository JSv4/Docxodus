#!/usr/bin/env /usr/bin/python3
"""Headless LibreOffice 'Compare Documents' driver.

Usage: lo_compare.py BASE.docx VARIANT.docx OUT.docx [PORT]

Opens VARIANT (the edited/right document), runs Writer's Compare-to against
BASE (the original/left document), and saves the redlined result to OUT.
Redlines therefore represent the base->variant transform (ins = right-only).

Runs an ISOLATED soffice instance (private UserInstallation + socket) so it
never collides with a soffice the user already has open.
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
    base, variant, out = sys.argv[1], sys.argv[2], sys.argv[3]
    port = sys.argv[4] if len(sys.argv) > 4 else "2002"
    profile = "file:///tmp/lo_profile_diff_%s" % port

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
            except Exception as e:  # noqa: BLE001 - retry until the socket is up
                last_err = e
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("could not connect to soffice: %s" % last_err)

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        dispatcher = smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx)

        # Load the VARIANT (right) VISIBLE (the compare slot lives in the Writer
        # view shell and is not registered on a Hidden frame); CompareDocuments
        # brings in the BASE (left). Redlines = base->variant transform.
        doc = desktop.loadComponentFromURL(url(variant), "_blank", 0, ())
        frame = doc.CurrentController.Frame
        dispatcher.executeDispatch(frame, ".uno:CompareDocuments", "", 0, (p("URL", url(base)),))

        doc.storeToURL(url(out), (p("FilterName", "MS Word 2007 XML"),))
        # Report redline count for the spike.
        try:
            redlines = doc.Redlines
            n = redlines.Count if hasattr(redlines, "Count") else len(list(redlines))
            print("REDLINES=%d" % n)
        except Exception as e:  # noqa: BLE001
            print("REDLINES=? (%s)" % e)
        doc.close(False)
        print("OK wrote %s" % out)
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


if __name__ == "__main__":
    main()
