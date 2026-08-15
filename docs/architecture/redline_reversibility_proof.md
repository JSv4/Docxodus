# Redline reversibility proof

Issue #464 turns the informal `Accept(Compare(left, right))` / `Reject(...)`
round-trip claim into a versioned verification result. The proof is deliberately
stronger than a body-text assertion: it resolves only revisions attributed to the
generated comparison, preserves pre-existing review state, and compares the complete
package on both paths.

## Inputs and baseline policy

The operation takes three immutable byte arrays:

- `baseline`: the exact document version that rejecting generated revisions must
  recover;
- `intendedFinal`: the exact document version that accepting generated revisions must
  recover;
- `redline`: the native tracked-change document being proved.

Optional proof settings carry the shared package-inspection limits and exact-byte
policy.

The caller chooses the baseline policy before invoking the proof. For example, a
workflow that deliberately accepted all prior revisions supplies that accepted
document as `baseline`; a workflow that preserves prior review state supplies the
review-bearing document. The proof never silently accepts or rejects pre-existing
revisions to manufacture a match.

## Generated-revision ownership

Revision ownership is established from the live, part-qualified revision registry,
not from author names alone. A stable identity contains the owning part, revision
family, and native constituent IDs. The proof inventories baseline, final, and
redline revisions, then classifies the redline inventory as:

- `preExisting`: an unchanged identity already present in the selected baseline;
- `generated`: an identity present only in the redline;
- `conflicted`: an identity that overlaps a baseline identity but changes its family,
  constituents, resolution status, or package location.

Every generated revision must be selectively resolvable. A malformed, ambiguous,
unsupported, or ownership-conflicted revision fails closed. Pre-existing identities
must remain live and unchanged after both proof paths. This avoids the unsafe shortcut
of calling accept-all/reject-all on a redline that contains older review markup.

## Two resolution paths

The operation clones `redline` twice. It accepts each generated revision in the first
clone and rejects each generated revision in the second clone through the same
selective resolver exposed by `DocxSession`. Resolution is deterministic and rebuilds
the live registry after every operation, because resolving an outer revision can
expose or detach nested markup.

Resolving one native revision can atomically consume a linked generated sibling (for
example, paragraph-mark and paragraph-property markup that Word treats as one
closure). The proof records that sibling separately as implicitly resolved only when
none of its native constituents remain live. A changed identity or surviving
constituent still fails closed, and target equivalence remains mandatory.

Each path records:

- the generated revision IDs requested, explicitly resolved, and implicitly consumed;
- the surviving pre-existing revision identities;
- the actual output package identity;
- the expected package identity;
- modeled semantic equivalence;
- normalized whole-package equivalence;
- ordered OPC-content equivalence;
- exact package-byte equivalence;
- every divergent part and the first divergent part/anchor;
- structured findings and the revision IDs relevant to a failure.

The proof succeeds only when accept-to-final and reject-to-baseline are both
normalized-whole-package equivalent and all generated revisions resolve without
altering pre-existing review identities. Exact package-byte equivalence is reported
separately and is not required unless the caller explicitly requests it.

## Equivalence layers

The proof reports, and never conflates, four layers:

1. **Modeled semantic equivalence** uses the versioned semantic change set. An empty
   modeled change projection proves equality only for that schema's understood surface.
   `OpaquePackagePart` changes remain visible in the underlying change set but are
   excluded from this layer and reported as unknown/unmodeled package divergences.
2. **Normalized whole-package equivalence** uses the package manifest's normalized
   XML/binary identity. This covers document text and formatting, lists, tables,
   sections, headers/footers, notes, comments, bookmarks, fields, content controls,
   relationships, media, and opaque vendor parts while ignoring only documented XML
   serialization choices.
3. **Ordered OPC-content equivalence** compares exact uncompressed entry bytes while
   ignoring ZIP entry order, timestamp, and compression choices.
4. **Raw package-byte equivalence** compares the supplied ZIP byte arrays exactly.

If modeled semantics match but normalized package identity differs, the result is not
called equivalent. The divergent manifest entries are emitted as unknown/unmodeled
package differences. A divergence also states whether the semantic change set has a
modeled change in that part, but this never claims the modeled projection exhaustively
explains every XML node in the part. This conservative residual rule is the guard
against silently losing unsupported content mixed into an otherwise modeled part.

## Result and receipt embedding

The canonical JSON schema is
`https://docxodus.dev/schemas/verification/redline-reversibility-proof/v1`.
Serialization is deterministic. It includes algorithm-labelled input/output digests,
revision classifications, both path results, part divergences, and findings. The
delivery receipt embeds this object as a value and includes its canonical JSON digest;
it does not reinterpret or flatten proof fields.

## Dependency boundary

The implementation consumes the package manifest and shared inspection limits from
#456 and the semantic change set from #457. It does not define another ZIP reader,
XML normalizer, digest profile, location vocabulary, or safety policy. The delivery
receipt from #458 consumes the proof but is not a dependency of proof generation.
