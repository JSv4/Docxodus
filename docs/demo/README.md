# Social demo setup

These pages are static and require no application server. They load the pinned
`docxodus@9.1.1` embed bundle from jsDelivr and the branded product guide from
the same GitHub Pages directory. Regenerate that document with
`python tools/generate-demo-guide.py` from the repository root.

## Publish

1. Publish `docxodus@9.1.1` and confirm
   `https://cdn.jsdelivr.net/npm/docxodus@9.1.1/dist/embed.bundle.js` returns JavaScript.
2. In GitHub **Settings → Pages**, choose **GitHub Actions** as the publishing source.
   The `Deploy static demo to GitHub Pages` workflow uploads `/docs` without
   running Jekyll.
3. Open `https://jsv4.github.io/Docxodus/demo/` and verify the status becomes `live`.
4. If the site is hosted anywhere else, update the absolute `og:url`, `og:image`,
   and `twitter:image` values in `index.html` before sharing.

Both pages accept query overrides for local or preview testing:

```text
?engine=<module URL of embed.bundle.js>&doc=<CORS-readable DOCX URL>
```

## Social-platform expectations

- LinkedIn uses the landing page's Open Graph title, description, and image. It
  does not run the editor inside the feed; the user clicks through to the live page.
- X receives a `summary_large_image` link card. Its current developer docs no
  longer document historical Player Cards, so the page does not promise an
  interactive editor inside a post. `player.html` is for iframe-capable sites.
- The Playwright specs validate the static metadata, boot-on-tap behavior, and
  live editor locally. They cannot prove how an external social client will render a post.
