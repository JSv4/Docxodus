# Social demo setup

These pages are static and require no application server. They load the pinned
`docxodus@9.1.0` embed bundle from jsDelivr and the sample DOCX from
`raw.githubusercontent.com`.

## Publish

1. Publish `docxodus@9.1.0` and confirm
   `https://cdn.jsdelivr.net/npm/docxodus@9.1.0/dist/embed.bundle.js` returns JavaScript.
2. In GitHub **Settings → Pages**, choose **GitHub Actions** as the publishing source.
   The `Deploy static demo to GitHub Pages` workflow uploads `/docs` without
   running Jekyll.
3. Open `https://jsv4.github.io/Docxodus/demo/` and verify the status becomes `live`.
4. If the site is hosted anywhere else, update the absolute `og:url`, `og:image`,
   `twitter:player`, and `twitter:image` values in `index.html` before sharing.

Both pages accept query overrides for local or preview testing:

```text
?engine=<module URL of embed.bundle.js>&doc=<CORS-readable DOCX URL>
```

## Social-platform expectations

- LinkedIn uses the landing page's Open Graph title, description, and image. It
  does not run the editor inside the feed; the user clicks through to the live page.
- `twitter:player` is experimental. X no longer documents Player Cards and may
  show only a normal link/card or ignore the player metadata. `player.html` is a
  progressive enhancement, never the only route to the demo.
- The Playwright specs validate the static metadata, boot-on-tap behavior, and
  live editor locally. They cannot prove how an external social client will render a post.
