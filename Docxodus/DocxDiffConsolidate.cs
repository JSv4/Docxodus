#nullable enable
namespace Docxodus;

/// <summary>How a consolidate resolves a span edited with DIFFERING edits by two+ reviewers.</summary>
public enum ConflictResolution
{
    /// <summary>Leave the base text at the conflicted span; record every competitor. The default.</summary>
    BaseWins,
    /// <summary>Apply the first reviewer (list order) inline; record the others.</summary>
    FirstReviewerWins,
    /// <summary>Emit each competing edit in sequence at the site; record all.</summary>
    StackAll,
}
