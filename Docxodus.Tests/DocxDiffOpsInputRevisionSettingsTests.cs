#nullable enable

using Docxodus;
using Docxodus.Internal;
using Xunit;
using static Docxodus.Tests.Ir.Diff.RevisionsInInputFixtures;

namespace Docxodus.Tests;

/// <summary>
/// Wire-level coverage for the two mutually exclusive policies for tracked
/// revisions already present in DocxDiff inputs. WASM and the Python stdio host
/// share <see cref="DocxDiffOps.ParseSettings"/>, so these tests pin the owner
/// rather than duplicating transport-specific behavior.
/// </summary>
public class DocxDiffOpsInputRevisionSettingsTests
{
    [Fact]
    public void ParseSettings_reads_input_revision_policy_flags_for_compare_and_consolidate()
    {
        const string json = "{\"preAcceptInputRevisions\":true,\"preserveInputRevisions\":true}";

        var diff = DocxDiffOps.ParseSettings(json);
        var consolidate = DocxDiffOps.ParseConsolidateSettings(json);

        Assert.True(diff.PreAcceptInputRevisions);
        Assert.True(diff.PreserveInputRevisions);
        Assert.True(consolidate.Diff.PreAcceptInputRevisions);
        Assert.True(consolidate.Diff.PreserveInputRevisions);
    }

    [Fact]
    public void Compare_honors_input_revision_policy_flags_from_wire()
    {
        var left = MultiScopeRevisionDoc(
            "Body alpha", "PriorAlice", "ins-a", "del-a", "PriorBob", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "PriorAlice", "ins-g", "del-g", "PriorBob", "hdr-prior", "fn-prior");

        var preAccepted = new WmlDocument("pre-accepted.docx", DocxDiffOps.Compare(
            left.DocumentByteArray,
            right.DocumentByteArray,
            "{\"preAcceptInputRevisions\":true,\"authorForRevisions\":\"NewDiff\"}"));
        Assert.DoesNotContain("PriorBob", RevisionAuthorsAllScopes(preAccepted));

        var preserved = new WmlDocument("preserved.docx", DocxDiffOps.Compare(
            left.DocumentByteArray,
            right.DocumentByteArray,
            "{\"preAcceptInputRevisions\":true,\"preserveInputRevisions\":true,\"authorForRevisions\":\"NewDiff\"}"));

        // Preserve wins over pre-accept, exactly as the public setting documents.
        Assert.Contains("PriorBob", RevisionAuthorsAllScopes(preserved));
    }
}
