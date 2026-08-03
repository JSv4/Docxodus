# README redline screenshot

This fixture regenerates `docs/images/redline.png` from the public
[NVCA Model Voting Agreement (October 2025)](https://nvca.org/wp-content/uploads/2024/10/NVCA-Model-VA-10-1-2025.docx).
It deterministically applies the edit round described in the README, compares the source and
modified copies with `DocxDiff`, renders tracked changes to HTML, and asserts the definition-list
markers before a browser is allowed to capture the image.

From the repository root:

```bash
curl -L --fail --output /tmp/NVCA-Model-VA-10-1-2025.docx \
  https://nvca.org/wp-content/uploads/2024/10/NVCA-Model-VA-10-1-2025.docx

dotnet run --project tools/screenshots/redline/redline-screenshot.csproj -- \
  /tmp/NVCA-Model-VA-10-1-2025.docx /tmp/docxodus-redline-screenshot

# Requires npm install in npm/ and Chrome at /usr/bin/google-chrome.
# Set CHROME_PATH when Chrome is installed elsewhere.
node tools/screenshots/redline/capture.mjs \
  /tmp/docxodus-redline-screenshot/redline.html docs/images/redline.png
```

The generator fails unless the rendered definition markers are exactly:

```text
(a), (b), (c), (d), (e), (f), (g), (g), (h), (i), (j), (k), (l), (m), (n), (o), (p)
```

The struck/live `(g)` duplicate proves that the deleted paragraph does not consume a final-document
number. The inserted `(k)` and following `(l)`–`(o)` prove that inserted paragraphs do consume a
number and cascade the remaining live items. The final struck `(p)` is the source of the clause moved
to live `(a)`.
