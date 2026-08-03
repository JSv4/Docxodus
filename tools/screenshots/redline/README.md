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

The generator fails unless the rendered definition markers and their revision states are exactly:

```text
(a)+, (a)-/(b)+, (b)-/(c)+, (c)-/(d)+, (d)-/(e)+, (e)-/(f)+,
(f)-, (g), (h), (i), (j), (k)+, (k)-/(l)+, (l)-/(m)+,
(m)-/(n)+, (n)-/(o)+, (o)-
```

`+` means inserted and `-` means deleted. The paired markers prove the cascade itself is visible:
moving the former `(o)` definition to new `(a)` shifts `(a)`–`(e)` to `(b)`–`(f)`, deleting former
`(f)` restores `(g)` onward, and inserting new `(k)` shifts former `(k)`–`(n)` to `(l)`–`(o)`.
