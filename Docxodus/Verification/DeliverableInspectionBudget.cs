// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

/// <summary>Shared fail-closed work budget for all semantic detectors inspecting one package.</summary>
internal sealed class DeliverableInspectionBudget
{
    private readonly DeliverableVerificationOptions _options;
    private long _nodes;
    private long _relationships;
    private long _textCharacters;
    private long _regexMatches;
    private long _steps;

    internal DeliverableInspectionBudget(DeliverableVerificationOptions options) => _options = options;

    internal bool Exhausted { get; private set; }
    internal string? ExhaustedResource { get; private set; }

    internal bool Node(long count = 1) => Consume(ref _nodes, count, _options.MaxDetectorNodes, "nodes");
    internal bool Relationship(long count = 1) => Consume(
        ref _relationships, count, _options.MaxDetectorRelationships, "relationships");
    internal bool Text(long count) => Consume(
        ref _textCharacters, count, _options.MaxDetectorTextCharacters, "text_characters");
    internal bool RegexMatch(long count = 1) => Consume(
        ref _regexMatches, count, _options.MaxDetectorRegexMatches, "regex_matches");
    internal bool Step(long count = 1) => Consume(ref _steps, count, _options.MaxDetectorSteps, "steps");

    private bool Consume(ref long current, long count, long maximum, string resource)
    {
        if (Exhausted) return false;
        if (count < 0 || current > maximum - count)
        {
            Exhausted = true;
            ExhaustedResource = resource;
            return false;
        }
        current += count;
        return true;
    }
}
