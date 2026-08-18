import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

type SchemaNode = Record<string, unknown>;

const testDirectory = dirname(fileURLToPath(import.meta.url));
const schemaPath = resolve(testDirectory, '../../docs/schemas/render-report-v3.schema.json');
const schema = JSON.parse(readFileSync(schemaPath, 'utf8')) as SchemaNode;
const definitions = schema.$defs as Record<string, SchemaNode>;

function objectProperties(node: SchemaNode): Record<string, SchemaNode> {
  return node.properties as Record<string, SchemaNode>;
}

function strings(node: SchemaNode, key: string): string[] {
  return node[key] as string[];
}

test.describe('render-report v3 font schema', () => {
  test('keeps the complete and failed report branches wired to the same font definitions', () => {
    const complete = definitions.complete;
    const failed = definitions.failed;
    const baseProperties = definitions.baseProperties;
    const fonts = baseProperties.fonts as SchemaNode;

    expect(strings(complete, 'required')).toContain('fontIdentity');
    expect(strings(complete, 'required')).toContain('fontReadiness');
    expect(strings(failed, 'required')).not.toContain('fontIdentity');
    expect(strings(failed, 'required')).toContain('fontReadiness');
    expect(objectProperties(complete).fontIdentity).toEqual({
      $ref: '#/$defs/fontIdentity',
    });
    expect(objectProperties(failed).fontIdentity).toEqual({
      $ref: '#/$defs/fontIdentity',
    });
    expect(baseProperties.fontIdentity).toEqual({ $ref: '#/$defs/fontIdentity' });
    expect(fonts.items).toEqual({
      $ref: '#/$defs/fontResolution',
    });
  });

  test('keeps exact final font readiness separate from semantic resolution evidence', () => {
    const readiness = definitions.fontReadinessProbe;
    const properties = objectProperties(readiness);

    expect(readiness.additionalProperties).toBe(false);
    expect(strings(readiness, 'required')).toEqual([
      'requestKey', 'requestedFamily', 'available',
    ]);
    expect(Object.keys(properties).sort()).toEqual([
      'available', 'requestKey', 'requestedFamily',
    ]);
    expect(properties.requestKey.pattern).toBe('^[0-9a-f]{64}$');
    expect(properties.available.type).toBe('boolean');
    expect((definitions.baseProperties.fontReadiness as SchemaNode).items).toEqual({
      $ref: '#/$defs/fontReadinessProbe',
    });
  });

  test('exactly mirrors the path-free production FontResolution vocabulary', () => {
    const resolution = definitions.fontResolution;
    const properties = objectProperties(resolution);

    expect(resolution.additionalProperties).toBe(false);
    expect(strings(resolution, 'required')).toEqual([
      'requestId', 'requestedFamily', 'requestedFamilies', 'requestedFamilyKinds', 'requestedStyle',
      'requestedWeight', 'requestedStretch', 'sampleCodePointCount', 'sampleDigest',
      'status', 'source',
    ]);
    expect(Object.keys(properties).sort()).toEqual([
      'browserFallbackAvailable', 'faceMatch', 'fileSha256', 'format', 'glyphCoverage', 'licenseEvidence',
      'metricCompatible', 'missingCodePointCount', 'requestId', 'requestedFamilies',
      'requestedFamily', 'requestedFamilyKinds', 'requestedStretch', 'requestedStyle', 'requestedWeight',
      'resolvedFace', 'resolvedFamily', 'sampleCodePointCount', 'sampleDigest', 'source',
      'status', 'version',
    ].sort());
    expect(properties.status.enum).toEqual([
      'resolved', 'substituted', 'missing', 'load_failed', 'unverified',
    ]);
    expect(properties.source.enum).toEqual(['browser', 'configured', 'attested']);
    expect(properties.requestedStyle.enum).toEqual(['normal', 'italic', 'oblique']);
    expect((properties.requestedFamilyKinds.items as SchemaNode).enum).toEqual(['named', 'generic']);
    const equalLengthBranches = resolution.oneOf as SchemaNode[];
    expect(equalLengthBranches).toHaveLength(64);
    equalLengthBranches.forEach((branch, index) => {
      const branchProperties = objectProperties(branch);
      expect(branchProperties.requestedFamilies).toMatchObject({
        minItems: index + 1,
        maxItems: index + 1,
      });
      expect(branchProperties.requestedFamilyKinds).toMatchObject({
        minItems: index + 1,
        maxItems: index + 1,
      });
    });
    expect(properties.format.enum).toEqual(['ttf', 'otf', 'woff', 'woff2']);
    expect(properties.faceMatch.enum).toEqual(['exact', 'synthesized']);
    expect(properties.glyphCoverage.enum).toEqual(['complete', 'partial', 'unverified']);
    expect(properties.browserFallbackAvailable.type).toBe('boolean');
    expect(properties.licenseEvidence).toEqual({ $ref: '#/$defs/fontLicenseEvidence' });
    expect(properties).not.toHaveProperty('path');
    expect(properties).not.toHaveProperty('bytesBase64');
  });

  test('requires the complete PDF digest and volatility contract together', () => {
    const requiredSets: string[][] = [];
    const visit = (value: unknown): void => {
      if (!value || typeof value !== 'object') return;
      if (Array.isArray(value)) {
        value.forEach(visit);
        return;
      }
      const record = value as SchemaNode;
      if (Array.isArray(record.required)) requiredSets.push(record.required as string[]);
      Object.values(record).forEach(visit);
    };
    visit(definitions.complete);
    expect(requiredSets).toContainEqual([
      'pdfDigest', 'pdfByteDeterministic', 'volatilePdfMetadata',
    ]);
  });

  test('freezes the resolver and license evidence identities', () => {
    const identity = definitions.fontIdentity;
    const identityProperties = objectProperties(identity);
    const license = definitions.fontLicenseEvidence;
    const licenseProperties = objectProperties(license);

    expect(identity.additionalProperties).toBe(false);
    expect(strings(identity, 'required')).toEqual([
      'resolverContract', 'substitutionContractVersion', 'substitutionContractDigest',
      'resolutionDigest',
    ]);
    expect(identityProperties.resolverContract.const)
      .toBe('https://docxodus.dev/contracts/font-resolver/v1');
    expect(identityProperties.substitutionContractVersion.const).toBe(1);
    expect(identityProperties.substitutionContractDigest.pattern).toBe('^[0-9a-f]{64}$');
    expect(identityProperties.resolutionDigest.pattern).toBe('^[0-9a-f]{64}$');

    expect(license.additionalProperties).toBe(false);
    expect(strings(license, 'required')).toEqual(['kind', 'identity', 'noSubsetting']);
    expect(licenseProperties.kind.enum)
      .toEqual(['installable', 'previewPrint', 'editable', 'attested']);
    expect(licenseProperties.identity.pattern).toBe('^[0-9a-f]{64}$');
  });
});
