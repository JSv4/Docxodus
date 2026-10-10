import hashlib
import itertools
import json
from pathlib import Path
import tempfile
import unittest

from tracked_drawings_matrix import fixtures


class MatrixContractTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        (self.folder / "generated").mkdir()
        # These bytes stand in for externally validated DOCX files. This unit test exercises
        # the manifest contract; the .NET generator supplies real packages for integration runs.
        def report(name, drawings=0, pictures=0, generated=False):
            file = self.folder / ("generated" if generated else "") / name
            payload = name.encode()
            file.write_bytes(payload)
            return dict(File=name, Sha256=hashlib.sha256(payload).hexdigest(), Errors=[],
                        Drawings=drawings, Pictures=pictures)

        matrix = [dict(name=f"control-{number}.docx", grouped=grouped, textbox=textbox,
                       anchored=anchored, header=header, revision=revision)
                  for number, (grouped, textbox, anchored, header, revision) in enumerate(
                      itertools.product([False, True], [False, True], [False, True],
                                        [False, True], ["plain", "ins", "del"]))]
        inputs = [report(row["name"]) for row in matrix]
        absent = report("header-absent.docx")
        inputs.append(absent)
        pairs = []
        self.results = []
        for kind in ["group", "vml"]:
            drawings, pictures = (1, 0) if kind == "group" else (0, 1)
            present = report(f"header-{kind}-plain.docx", drawings, pictures)
            inputs.append(present)
            for revision in ["ins", "del"]:
                inputs.append(report(f"header-{kind}-{revision}.docx", drawings, pictures))
            for revision in ["insert", "delete"]:
                name = f"header-{revision}-{kind}"
                left, right = (absent, present) if revision == "insert" else (present, absent)
                pairs.append(dict(name=name, left=left["File"], right=right["File"]))
                self.results.append(dict(Name=name, Left=left, Right=right,
                    Compared=report(name + ".docx", drawings, pictures, True),
                    Accepted=report(name + "-accepted.docx", right["Drawings"], right["Pictures"], True),
                    Rejected=report(name + "-rejected.docx", left["Drawings"], left["Pictures"], True)))
        for revision in ["ins", "del"]:
            name = "header-group-paragraph-mark-" + revision
            source = report(name + ".docx", 1)
            inputs.append(source)
            self.results.append(dict(Name=name, Mode="paragraph-mark-diagnostic", Source=source,
                Accepted=report(name + "-accepted.docx", 1, generated=True),
                Rejected=report(name + "-rejected.docx", 1, generated=True)))
        self.write("matrix.json", matrix)
        self.write("pairs.json", pairs)
        self.write("input-validation.json", inputs)
        self.write("results.json", self.results)

    def write(self, name, value):
        (self.folder / name).write_text(json.dumps(value))

    def test_comparison_controls_cannot_be_omitted(self):
        self.write("pairs.json", [])
        with self.assertRaisesRegex(ValueError, "four insertion/deletion comparison controls"):
            fixtures(self.folder)

    def test_accepted_drawing_presence_must_match_right_input(self):
        next(row for row in self.results if row["Name"] == "header-insert-group")["Accepted"]["Drawings"] = 0
        self.write("results.json", self.results)
        with self.assertRaisesRegex(ValueError, "accept/reject content differs"):
            fixtures(self.folder)

    def test_rejected_vml_presence_must_match_left_input(self):
        next(row for row in self.results if row["Name"] == "header-delete-vml")["Rejected"]["Pictures"] = 0
        self.write("results.json", self.results)
        with self.assertRaisesRegex(ValueError, "accept/reject content differs"):
            fixtures(self.folder)


if __name__ == "__main__":
    unittest.main()
