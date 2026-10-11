Synthetic formatting comparison reproductions. The generator creates the inputs from literal XML and invented text, then checks accepted and rejected output with the public Docxodus APIs.

Run from the repository root:
dotnet run --project repros/synthetic-formatting/Repro.csproj -c Release -- /tmp/docx-formatting-repro

Recorded outputs in fixtures/ were generated on e5512006ca50e09c4df1ceb008d08587086baf6c using .NET SDK 10.0.301 and DocumentFormat.OpenXml 3.5.1.
