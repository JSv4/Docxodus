// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using LegalEval;

try
{
    var options = Options.Parse(args);
    var outcome = new LegalEvaluationRunner().Run(
        new EvaluationRunOptions(
            options.CorpusPath,
            options.Subset,
            options.ScenarioId,
            options.CandidateDirectory,
            options.ArtifactRoot,
            options.ReportPath,
            options.Render ? ArtifactRenderMode.TrustedDocuments : ArtifactRenderMode.Disabled),
        Console.Out,
        Console.Error);
    return outcome.ExitCode;
}
catch (Exception exception) when (Docxodus.Verification.DeliverableExceptionBoundary.IsRecoverable(exception))
{
    Console.Error.WriteLine(exception.Message);
    return 2;
}

internal sealed record Options(
    string CorpusPath,
    string Subset,
    string? ScenarioId,
    string? CandidateDirectory,
    string ArtifactRoot,
    string? ReportPath,
    bool Render)
{
    private static readonly HashSet<string> ValueOptions = new(StringComparer.Ordinal)
    {
        "--corpus", "--subset", "--scenario", "--candidate-dir", "--artifacts", "--report",
    };

    public static Options Parse(string[] args)
    {
        var values = new Dictionary<string, string?>(StringComparer.Ordinal);
        var render = false;
        for (var index = 0; index < args.Length; index++)
        {
            var argument = args[index];
            if (argument is "--render" or "--keep-passes")
            {
                if (argument == "--render") render = true;
                continue;
            }
            if (!ValueOptions.Contains(argument) || index + 1 >= args.Length)
                throw new ArgumentException($"Unknown or incomplete argument '{argument}'.");
            if (!values.TryAdd(argument, args[++index]))
                throw new ArgumentException($"Argument '{argument}' was supplied more than once.");
        }
        var subset = values.GetValueOrDefault("--subset") ?? "fast";
        if (subset is not ("fast" or "full"))
            throw new ArgumentException("--subset must be 'fast' or 'full'.");
        return new Options(
            values.GetValueOrDefault("--corpus") ?? Path.Combine("eval", "legal", "corpus.json"),
            subset,
            values.GetValueOrDefault("--scenario"),
            values.GetValueOrDefault("--candidate-dir"),
            values.GetValueOrDefault("--artifacts") ?? Path.Combine("artifacts", "legal-eval"),
            values.GetValueOrDefault("--report"),
            render);
    }
}
