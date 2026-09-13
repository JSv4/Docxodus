import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const xml = value => value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;');

// Run the real MSBuild targets around a tiny stand-in for the expensive compiler.
// Real profile-switch publishes are also measured in benchmarks/aot-coverage/.
test('AOT caching follows profile contents and successful compilation', async t => {
  const dir = mkdtempSync(join(tmpdir(), 'docxodus-aot-cache-'));
  try {
    const profile = join(dir, 'recording.aotprofile');
    const alternate = join(dir, 'alternate.aotprofile');
    const stamp = join(dir, 'docxodus-aot-profile.sha256');
    const project = join(dir, 'CacheTest.proj');
    writeFileSync(profile, 'original profile');
    writeFileSync(join(dir, 'runtime.o'), 'unrelated native runtime object');
    writeFileSync(project, `<Project DefaultTargets="_WasmAotCompileApp">
      <PropertyGroup>
        <_WasmShouldAOT>true</_WasmShouldAOT>
        <_WasmIntermediateOutputPath>${xml(dir)}/</_WasmIntermediateOutputPath>
      </PropertyGroup>
      <Import Project="${xml(join(repo, 'wasm/DocxodusWasm/AotProfile.targets'))}" />
      <Target Name="_WasmAotCompileApp">
        <ItemGroup>
          <AotOutput Include="${xml(dir)}/Example.dll.bc;${xml(dir)}/Example.dll.o;${xml(dir)}/aot_compiler_cache.json;${xml(dir)}/monoAotPropertyValues.txt;${xml(dir)}/tokens/Example.dll.bin" />
        </ItemGroup>
        <Error Condition="'$(ExpectInvalidated)' == 'true' and Exists('%(AotOutput.Identity)')"
               Text="Stale AOT output survived: %(AotOutput.Identity)" />
        <Error Condition="'$(ExpectInvalidated)' == 'false' and !Exists('%(AotOutput.Identity)')"
               Text="Unchanged AOT output was deleted: %(AotOutput.Identity)" />
        <MakeDir Directories="${xml(dir)}/tokens" />
        <WriteLinesToFile File="%(AotOutput.Identity)" Lines="compiled output" Overwrite="true" />
        <Error Condition="'$(FailCompilation)' == 'true'" Text="Simulated AOT failure" />
      </Target>
    </Project>`);
    const publish = (selectedProfile = profile, invalidated = true, extra = []) => {
      const result = spawnSync('dotnet', ['msbuild', project, '-nologo', '-verbosity:minimal',
        `-p:WasmAotProfilePath=${selectedProfile}`, `-p:ExpectInvalidated=${invalidated}`, ...extra],
      { encoding: 'utf8' });
      assert.ifError(result.error);
      assert.equal(readFileSync(join(dir, 'runtime.o'), 'utf8'), 'unrelated native runtime object');
      return result;
    };
    const succeeds = (...args) => {
      const result = publish(...args);
      assert.equal(result.status, 0, result.stdout + result.stderr);
    };

    await t.test('first build, unchanged profile, and legacy cache without a stamp', () => {
      succeeds();
      succeeds(profile, false);
      rmSync(stamp);
      succeeds();
    });
    await t.test('re-recording at the same path, even with an older timestamp', () => {
      writeFileSync(profile, 'expanded profile');
      utimesSync(profile, new Date(0), new Date(0));
      succeeds();
    });
    await t.test('another path with the same contents reuses AOT', () => {
      writeFileSync(alternate, readFileSync(profile));
      succeeds(alternate, false);
    });
    await t.test('switching profiles and switching back both invalidate AOT', () => {
      writeFileSync(alternate, 'different profile');
      succeeds(alternate);
      succeeds(profile);
    });
    await t.test('missing profile fails without changing existing outputs', () => {
      const before = readFileSync(stamp, 'utf8');
      assert.notEqual(publish(join(dir, 'missing.aotprofile')).status, 0);
      assert.equal(readFileSync(stamp, 'utf8'), before);
      succeeds(profile, false);
    });
    await t.test('failed compilation does not record a successful profile', () => {
      const before = readFileSync(stamp, 'utf8');
      assert.notEqual(publish(alternate, true, ['-p:FailCompilation=true']).status, 0);
      assert.equal(readFileSync(stamp, 'utf8'), before);
      succeeds(alternate);
    });
    await t.test('full AOT and profile-guided AOT have distinct cache entries', () => {
      succeeds('');
      assert.equal(readFileSync(stamp, 'utf8').trim(), 'full-aot');
      succeeds('', false);
      succeeds(profile);
    });
    await t.test('interpreter builds leave AOT outputs and profile stamp alone', () => {
      const before = readFileSync(stamp, 'utf8');
      succeeds(alternate, false, ['-p:_WasmShouldAOT=false']);
      assert.equal(readFileSync(stamp, 'utf8'), before);
    });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
