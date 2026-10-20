"""Run the independent cases against a caller-selected Docxodus checkout."""
import json
from pathlib import Path
import subprocess
import sys

ROOT=Path(__file__).resolve().parent
REPO=Path(sys.argv[1]).resolve()
OUT=Path(sys.argv[2]).resolve()
OUT.mkdir(parents=True,exist_ok=True)
jobs=[]
for src in sorted((ROOT/"inputs").glob("*.docx")):
    jobs.append({"operation":"validate","input":str(src),"output":str(OUT/(src.stem+".validation.json"))})
    jobs.append({"operation":"html","input":str(src),"output":str(OUT/(src.stem+".html"))})
    if src.stem.startswith("revisions-"):
        for op in ("accept","reject"):
            jobs.append({"operation":op,"input":str(src),"output":str(OUT/(src.stem+"."+op+".docx"))})
jobfile=OUT/"jobs.jsonl"
jobfile.write_text("".join(json.dumps(j)+"\n" for j in jobs))
subprocess.run(["dotnet","build",str(ROOT/"native/Probe.csproj"),"-c","Release","-t:Rebuild",
                "-p:NuGetAudit=false",f"-p:DocxodusRepo={REPO}","-o",str(OUT/"runner"),"--nologo"],check=True)
with (OUT/"results.jsonl").open("w") as log:
    subprocess.run(["dotnet",str(OUT/"runner/Probe.dll"),str(jobfile)],stdout=log,check=True)
print(OUT/"results.jsonl")
