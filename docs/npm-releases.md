# npm releases

The `v*` release workflow builds and tests `docxodus` and `@docxodus/export`, then
submits both with `npm stage publish`. A successful workflow means the npm
packages are staged for review; a maintainer separately approves each package
with their own two-factor authentication before it becomes public. NuGet and
the independent `docx-scalpel` PyPI workflow retain their existing publishing flow.

The workflow pins npm 11.19.1. Local staging and approval commands require npm
11.15.0 or newer and Node 22.14.0 or newer. Both npm packages need a trusted
publisher matching `JSv4/Docxodus`, workflow `publish.yml`, environment `Publisher`,
with staging allowed. Direct-publishing permission is unnecessary for this flow.

## Review and approve

1. Find each stage ID in the workflow's **npm packages awaiting 2FA approval**
   summary. A partial failure still reports any package that was successfully staged.
2. Review the packages in npm's **Staged Packages** tab, or use the CLI:

   ```sh
   npm stage list docxodus
   npm stage list @docxodus/export
   npm stage view <stage-id>
   ```

3. Approve `docxodus` first, then the matching `@docxodus/export` stage, since the
   companion requires that exact browser-package version:

   ```sh
   npm stage approve <stage-id>
   ```

   Complete the 2FA prompt yourself. Approval is an explicit maintainer action;
   the workflow does not perform it. A local npm older than 11.15.0 can run these
   commands through `npx npm@11.19.1 stage ...`.

4. Confirm both packages are publicly available at the intended version. Then
   confirm the browser bundle is served by jsDelivr before updating consumer-facing
   CDN snippets. The demo site builds its own matching runtime;
   see [the demo publishing instructions](demo/README.md#publish).

## Partial releases

Staged and published versions share the same version namespace. If one package
stages successfully and the other fails, keep the successful stage for review
and retry only the missing package from the release tag. Re-running the entire
npm job will encounter the existing version. An already validated tarball can
be submitted with `npm stage publish ./package.tgz --access public --tag latest`.

See npm's [staged publishing guide](https://docs.npmjs.com/staged-publishing/)
and [trusted publisher configuration](https://docs.npmjs.com/trusted-publishers/).
