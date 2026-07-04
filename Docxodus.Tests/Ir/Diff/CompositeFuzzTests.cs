#nullable enable
using System.Linq;
using Docxodus;
using Xunit;
namespace Docxodus.Tests.Ir.Diff;
public class CompositeFuzzTests
{
    [Theory]
    [InlineData(3)] [InlineData(4)] [InlineData(5)]
    public void Composite_round_trips_reject_equals_base(int reviewerCount)
    {
        for (int seed = 0; seed < 50; seed++)
        {
            var fc = DiffFuzzer.GenerateComposite(seed, reviewerCount);
            var baseDoc = new WmlDocument("b.docx", fc.Base);
            var reviewers = fc.Reviewers
                .Select(r => new DocxDiffReviewer { Document = new WmlDocument("r.docx", r.Doc), Author = r.Author })
                .ToList();
            var merged = DocxDiff.Consolidate(baseDoc, reviewers);
            Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        }
    }

    [Theory]
    [InlineData(3)]
    public void Composite_round_trips_structurally_with_tables_and_footnotes(int reviewerCount)
    {
        for (int seed = 0; seed < 50; seed++)
        {
            var fc = DiffFuzzer.GenerateCompositeWithStructure(seed, reviewerCount);
            var baseDoc = new WmlDocument("b.docx", fc.Base);
            var reviewers = fc.Reviewers
                .Select(r => new DocxDiffReviewer { Document = new WmlDocument("r.docx", r.Doc), Author = r.Author })
                .ToList();
            var merged = DocxDiff.Consolidate(baseDoc, reviewers);
            var rejected = RevisionProcessor.RejectRevisions(merged);
            // Structural: rejecting all revisions must restore the base body structure (incl. tables),
            // not just the body paragraph text. Docs.StructuralBody walks body w:p AND w:tbl (descending
            // into rows/cells), so a consolidate that corrupts or drops a table on the reject path differs.
            Assert.Equal(Docs.StructuralBody(baseDoc), Docs.StructuralBody(rejected));
        }
    }

    [Theory]
    [InlineData(3)] [InlineData(4)]
    public void Composite_apply_verifier_holds(int reviewerCount)
    {
        for (int seed = 0; seed < 30; seed++)
        {
            var fc = DiffFuzzer.GenerateComposite(seed, reviewerCount);
            var baseDoc = new WmlDocument("b.docx", fc.Base);
            var revs = fc.Reviewers.Select(r => (r.Author, (WmlDocument)new WmlDocument("r.docx", r.Doc))).ToList();
            var dd = revs.Select(r => new DocxDiffReviewer { Document = r.Item2, Author = r.Author }).ToList();
            var merged = DocxDiff.Consolidate(baseDoc, dd);
            var acceptedText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
            IrCompositeVerifier.Verify(baseDoc, revs, ConflictResolution.BaseWins, acceptedText);
        }
    }

    /// <summary>
    /// STRUCTURAL-OP composite fuzz coverage. The reviewer pool here ADDITIONALLY includes the structural
    /// mutations (Relocate / Split / Merge), so each consolidate exercises the merger's
    /// <c>LowerStructuralOps</c> lowering AND its contested-relocation branch — paths the paragraph-only
    /// composite fuzz never reaches. For every (seed, reviewerCount) the case must satisfy, under the
    /// BaseWins policy the verifier reconstructs against:
    /// <list type="bullet">
    /// <item>(a) reject ≡ base — <c>RejectRevisions</c> plain text equals the base body text; AND</item>
    /// <item>(b) no content loss on accept — the rendered accepted body text equals the independently
    /// reconstructed policy-resolved body (the <see cref="IrCompositeVerifier"/> apply-oracle).</item>
    /// </list>
    /// A future refactor that silently breaks the lowering or contested branch would fail one of these on
    /// some seed even with the rest of the suite green.
    /// </summary>
    [Theory]
    [InlineData(3)] [InlineData(4)] [InlineData(5)]
    public void Composite_with_structural_ops_round_trips(int reviewerCount)
    {
        for (int seed = 0; seed < 60; seed++)
        {
            var fc = DiffFuzzer.GenerateCompositeWithStructuralOps(seed, reviewerCount);
            var baseDoc = new WmlDocument("b.docx", fc.Base);
            var revs = fc.Reviewers
                .Select(r => (r.Author, (WmlDocument)new WmlDocument("r.docx", r.Doc)))
                .ToList();
            var dd = revs.Select(r => new DocxDiffReviewer { Document = r.Item2, Author = r.Author }).ToList();

            // The verifier reconstructs the BaseWins-resolved body, so consolidate under BaseWins too.
            var merged = DocxDiff.Consolidate(
                baseDoc, dd, new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.BaseWins });

            // (a) reject ≡ base.
            Assert.Equal(
                Docs.PlainText(baseDoc),
                Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));

            // (b) no content loss on accept: the apply-oracle reconstructs the policy-resolved accepted body
            // purely from the script ops and asserts it equals the rendered accept.
            var acceptedText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
            IrCompositeVerifier.Verify(baseDoc, revs, ConflictResolution.BaseWins, acceptedText);
        }
    }

    /// <summary>
    /// FOLLOW-ON B per-cell table composition fuzz: DIFFERENT reviewers edit DIFFERENT table cells/rows (the
    /// disjoint composed-table happy path). For every (seed, reviewerCount) the consolidated table must satisfy
    /// (a) reject ≡ base STRUCTURALLY (table-aware, incl. row/cell content) — the composed table restores the
    /// base table exactly — and (b) the table-aware apply-verifier (cell ops applied to the base table) matches
    /// the rendered accepted body. A future refactor that breaks the per-cell compose, the recursion, or the
    /// composed-table renderer would fail one of these on some seed even with the rest of the suite green.
    /// </summary>
    [Theory]
    [InlineData(2)] [InlineData(3)] [InlineData(4)]
    public void Composite_disjoint_table_cells_round_trip(int reviewerCount)
    {
        for (int seed = 0; seed < 60; seed++)
        {
            var fc = DiffFuzzer.GenerateCompositeDisjointTableCells(seed, reviewerCount);
            var baseDoc = new WmlDocument("b.docx", fc.Base);
            var revs = fc.Reviewers
                .Select(r => (r.Author, (WmlDocument)new WmlDocument("r.docx", r.Doc)))
                .ToList();
            var dd = revs.Select(r => new DocxDiffReviewer { Document = r.Item2, Author = r.Author }).ToList();

            var merged = DocxDiff.Consolidate(
                baseDoc, dd, new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.BaseWins });

            // (a) reject ≡ base, table-aware (StructuralBody descends rows/cells).
            Assert.Equal(
                Docs.StructuralBody(baseDoc),
                Docs.StructuralBody(RevisionProcessor.RejectRevisions(merged)));

            // (b) table-aware apply-verifier: the composed table's cell ops reconstruct the rendered accept.
            IrCompositeVerifier.Verify(baseDoc, revs, ConflictResolution.BaseWins, Docs.AcceptStructuralBody(merged));
        }
    }
}
