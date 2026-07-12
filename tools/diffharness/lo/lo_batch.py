#!/usr/bin/env /usr/bin/python3
"""Batch LibreOffice driver over a diffharness corpus, in ONE isolated soffice instance.

Usage: lo_batch.py CORPUS_DIR [PORT]

For each scenario folder (with left.docx / right.docx / ours.docx) it:
  1. Runs Writer 'Compare Documents' (open right, compare-to left) -> <id>/lo.docx,
     recording the redline count LibreOffice's own algorithm produces.
  2. Loads our diff output <id>/ours.docx and records whether it opens and how many
     redlines LibreOffice RECOGNIZES in our markup (the 'renders correctly' proxy:
     LO parsing our w:ins/w:del and agreeing they are tracked changes).

Writes <id>/lo_meta.json per scenario and a corpus-level lo_summary.json.
Isolated profile + socket so it never collides with a user's open soffice.
"""
import json
import os
import subprocess
import sys
import time
from typing import Any

import uno
from com.sun.star.beans import PropertyValue


def p(n: str, v: Any) -> PropertyValue:
    pv = PropertyValue(); pv.Name = n; pv.Value = v; return pv


def url(path: str) -> str:
    return "file://" + os.path.abspath(path)


def redline_summary(doc: Any) -> tuple[int, list[dict[str, Any]]]:
    """Return (count, [{type, author, text}]) for a document's tracked changes.

    Each redline element is an XPropertySet (also XText); read its RedlineType/Author
    via properties and the affected text via getString().
    """
    out: list[dict[str, Any]] = []
    try:
        rl = doc.Redlines
        n = rl.Count
        for i in range(n):
            r = rl.getByIndex(i)
            try:
                txt = r.getString()
            except Exception:  # noqa: BLE001
                txt = None
            out.append({"type": r.RedlineType, "author": r.RedlineAuthor, "text": txt})
        return n, out
    except Exception as e:  # noqa: BLE001
        return -1, [{"error": str(e)}]


def main() -> None:
    corpus = sys.argv[1]
    port = sys.argv[2] if len(sys.argv) > 2 else "2010"
    profile = "file:///tmp/lo_profile_batch_%s" % port
    manifest = json.load(open(os.path.join(corpus, "manifest.json")))

    sof = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--norestore", "--nologo",
        "--nofirststartwizard", "--nodefault",
        "--accept=socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port,
        "-env:UserInstallation=" + profile,
    ])
    summary = []
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(80):
            try:
                ctx = resolver.resolve(
                    "uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port)
                break
            except Exception:  # noqa: BLE001
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("could not connect to soffice")
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)

        for s in manifest:
            sid = s["id"]
            d = os.path.join(corpus, sid)
            left, right, ours = (os.path.join(d, f) for f in ("left.docx", "right.docx", "ours.docx"))
            rec = {"id": sid}

            # 1) LibreOffice's own compare.
            try:
                doc = desktop.loadComponentFromURL(url(right), "_blank", 0, ())
                frame = doc.CurrentController.Frame
                dispatcher.executeDispatch(frame, ".uno:CompareDocuments", "", 0, (p("URL", url(left)),))
                lo_count, lo_list = redline_summary(doc)
                doc.storeToURL(url(os.path.join(d, "lo.docx")), (p("FilterName", "MS Word 2007 XML"),))
                doc.close(False)
                rec["loCompareRedlines"] = lo_count
                rec["loTypes"] = _types(lo_list)
                rec["loRedlines"] = lo_list
            except Exception as e:  # noqa: BLE001
                rec["loCompareRedlines"] = -1
                rec["loError"] = str(e)

            # 2) Does OUR output open, and does LO recognize our redlines?
            try:
                odoc = desktop.loadComponentFromURL(url(ours), "_blank", 0, ())
                ours_count, ours_list = redline_summary(odoc)
                rec["oursLoadOk"] = True
                rec["oursRedlinesSeenByLo"] = ours_count
                rec["oursTypes"] = _types(ours_list)
                rec["oursRedlines"] = ours_list
                odoc.close(False)
            except Exception as e:  # noqa: BLE001
                rec["oursLoadOk"] = False
                rec["oursError"] = str(e)

            json.dump(rec, open(os.path.join(d, "lo_meta.json"), "w"), indent=2)
            summary.append(rec)
            print("%-26s loCompare=%s  ours: load=%s seenByLO=%s" % (
                sid, rec.get("loCompareRedlines"), rec.get("oursLoadOk"),
                rec.get("oursRedlinesSeenByLo")))

        json.dump(summary, open(os.path.join(corpus, "lo_summary.json"), "w"), indent=2)
    finally:
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(1)
        sof.terminate()
        try:
            sof.wait(timeout=10)
        except Exception:  # noqa: BLE001
            sof.kill()


def _types(lst: list[dict[str, Any]]) -> dict[Any, int]:
    out: dict[Any, int] = {}
    for x in lst:
        t = x.get("type", "?")
        out[t] = out.get(t, 0) + 1
    return out


if __name__ == "__main__":
    main()
