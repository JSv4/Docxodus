// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;

namespace LegalEval;

public sealed class EvaluationScorer
{
    private const string WNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly IEvaluationPackageValidator _packageValidator;
    private readonly IEvaluationArtifactRenderer? _artifactRenderer;

    public EvaluationScorer(
        IEvaluationPackageValidator? packageValidator = null,
        IEvaluationArtifactRenderer? artifactRenderer = null)
    {
        _packageValidator = packageValidator ?? new InterimEvaluationPackageValidator();
        _artifactRenderer = artifactRenderer;
    }

    public EvaluationScore Score(
        LegalScenario scenario,
        BaselineExecution baseline,
        byte[] candidate,
        ScoreKind kind,
        string artifactRoot,
        ArtifactRenderMode renderMode = ArtifactRenderMode.Disabled)
    {
        var metrics = new List<MetricResult>();
        string? packageSafetyError = null;
        try
        {
            _packageValidator.Validate(baseline.Input, "evaluation input");
            _packageValidator.Validate(baseline.Expected, "pinned expected document");
            _packageValidator.Validate(candidate, "candidate document");
            metrics.Add(new MetricResult("document-validity.package-safety", "document_validity",
                "passed", "candidate passed interim bounded OPC validation", 1));
        }
        catch (Exception exception)
        {
            packageSafetyError = exception.Message;
            metrics.Add(new MetricResult("document-validity.package-safety", "document_validity",
                "failed", exception.Message, 0));
        }
        (MetricResult Metric, byte[]? RedlineBytes) redline;
        (MetricResult Metric, string? CandidateHtml) render;
        if (packageSafetyError is null)
        {
            metrics.AddRange(EvaluateInvariants(scenario, candidate));
            metrics.Add(SafeMetric("target-precision.reference-equivalence", "target_precision",
                () => EvaluateTargetPrecision(baseline.Expected, candidate)));
            metrics.AddRange(SafeMetrics(
                new[]
                {
                    ("unintended-change.parts", "unintended_change"),
                    ("unintended-change.anchors", "unintended_change"),
                },
                () => EvaluateChangeBudget(scenario, baseline.Input, candidate)));
            metrics.Add(SafeMetric("document-validity.openxml", "document_validity",
                () => EvaluateValidity(candidate)));
            redline = EvaluateRedlineReversibility(baseline.Input, candidate);
            metrics.Add(redline.Metric);
            render = EvaluateRendering(baseline.Expected, candidate);
            metrics.Add(render.Metric);
        }
        else
        {
            metrics.AddRange(BlockedMetrics(scenario, packageSafetyError));
            redline = (new MetricResult("redline-reversibility.interim-text-projection",
                "redline_reversibility", "failed",
                $"blocked by package safety validation: {packageSafetyError}", 0), null);
            metrics.Add(redline.Metric);
            render = (new MetricResult("rendering-regression.html-projection",
                "rendering_regression", "failed",
                $"blocked by package safety validation: {packageSafetyError}", 0), null);
            metrics.Add(render.Metric);
        }

        var status = metrics.Any(value => value.Status != "passed") ? "failed" : "passed";
        var artifactDirectory = ArtifactWriter.ResolveScoreDirectory(artifactRoot, scenario.Id, kind);
        var semanticDiff = packageSafetyError is null
            ? TrySemanticDiff(baseline.Input, candidate)
            : (Json: (string?)null, Error: packageSafetyError);
        var targetSemanticDiff = TrySemanticDiff(baseline.Input, baseline.Expected);
        var publication = ArtifactWriter.Write(
            artifactDirectory,
            scenario.Id,
            kind,
            status,
            metrics,
            scenario.ExpectedOutputs,
            baseline.OperationLog,
            baseline.Input,
            candidate,
            baseline.Expected,
            semanticDiff.Json ?? ErrorEnvelope("semantic-diff", "failed",
                semanticDiff.Error ?? packageSafetyError ?? "semantic diff generation failed"),
            semanticDiff.Json is not null,
            targetSemanticDiff.Json ?? ErrorEnvelope("semantic-diff", "failed",
                targetSemanticDiff.Error ?? "target semantic diff generation failed"),
            targetSemanticDiff.Json is not null,
            render.CandidateHtml,
            redline.RedlineBytes,
            renderMode,
            allowExternalRenderer: kind == ScoreKind.EngineBaseline,
            packageSafetyError,
            _artifactRenderer);

        return new EvaluationScore(scenario.Id, kind, publication.Status,
            publication.Metrics, publication.Artifacts, artifactDirectory);
    }

    private static IReadOnlyList<MetricResult> BlockedMetrics(
        LegalScenario scenario, string packageSafetyError)
    {
        var detail = $"blocked by package safety validation: {packageSafetyError}";
        var results = scenario.Invariants.Select(value =>
                new MetricResult(value.Id, value.Metric, "failed", detail, 0))
            .ToList();
        results.Add(new MetricResult("target-precision.reference-equivalence", "target_precision",
            "failed", detail, 0));
        results.Add(new MetricResult("unintended-change.parts", "unintended_change",
            "failed", detail, 0));
        results.Add(new MetricResult("unintended-change.anchors", "unintended_change",
            "failed", detail, 0));
        results.Add(new MetricResult("document-validity.openxml", "document_validity",
            "failed", detail, 0));
        return results;
    }

    private static MetricResult SafeMetric(
        string id, string category, Func<MetricResult> evaluate)
    {
        try
        {
            return evaluate();
        }
        catch (Exception exception)
        {
            return new MetricResult(id, category, "failed",
                $"evaluation error: {exception.Message}", 0);
        }
    }

    private static IReadOnlyList<MetricResult> SafeMetrics(
        IReadOnlyList<(string Id, string Category)> expected,
        Func<IReadOnlyList<MetricResult>> evaluate)
    {
        try
        {
            return evaluate();
        }
        catch (Exception exception)
        {
            return expected.Select(value => new MetricResult(
                value.Id, value.Category, "failed",
                $"evaluation error: {exception.Message}", 0)).ToList();
        }
    }

    private static IReadOnlyList<MetricResult> EvaluateInvariants(
        LegalScenario scenario, byte[] candidate)
    {
        var results = new List<MetricResult>(scenario.Invariants.Count);
        foreach (var invariant in scenario.Invariants)
        {
            try
            {
                var observed = Probe(candidate, invariant.Probe);
                var passed = Compare(observed, invariant.Operator, invariant.Expected);
                results.Add(new MetricResult(
                    invariant.Id,
                    invariant.Metric,
                    passed ? "passed" : "failed",
                    $"operator={invariant.Operator}; expected={Json(invariant.Expected)}; observed={Json(observed)}",
                    passed ? 1 : 0));
            }
            catch (Exception exception)
            {
                results.Add(new MetricResult(invariant.Id, invariant.Metric, "failed",
                    $"probe error: {exception.Message}", 0));
            }
        }
        return results;
    }

    private static JsonNode Probe(byte[] candidate, JsonObject probe)
    {
        var kind = RequiredString(probe, "kind");
        return kind switch
        {
            "xmlCount" => JsonValue.Create(XmlCount(candidate,
                RequiredString(probe, "part"), RequiredString(probe, "xpath")))!,
            "partExists" => JsonValue.Create(PartBytes(candidate,
                RequiredString(probe, "part"), required: false) is not null)!,
            "textCount" => JsonValue.Create(TextCount(candidate,
                RequiredString(probe, "text")))!,
            _ => throw new ScenarioValidationException($"unknown invariant probe kind '{kind}'"),
        };
    }

    private static int XmlCount(byte[] bytes, string part, string xpath)
    {
        var payload = PartBytes(bytes, part, required: true)!;
        using var reader = XmlReader.Create(new MemoryStream(payload), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        });
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        var namespaces = new XmlNamespaceManager(new NameTable());
        namespaces.AddNamespace("w", WNamespace);
        namespaces.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        namespaces.AddNamespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
        namespaces.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        namespaces.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");
        namespaces.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships");
        var value = document.XPathEvaluate(xpath, namespaces);
        if (value is IEnumerable<object> objects) return objects.Count();
        if (value is IEnumerable<XElement> elements) return elements.Count();
        if (value is IEnumerable<XAttribute> attributes) return attributes.Count();
        if (value is double number) return checked((int)number);
        throw new InvalidOperationException($"XPath must return a node-set or number, got {value?.GetType().Name}");
    }

    private static int TextCount(byte[] bytes, string needle)
    {
        var projection = WmlToMarkdownConverter.Convert(
            new WmlDocument("candidate.docx", bytes),
            new WmlToMarkdownConverterSettings
            {
                AnchorMode = AnchorRenderMode.None,
                Scopes = ProjectionScopes.All,
                TrackedChanges = TrackedChangeMode.RenderInline,
            });
        var count = 0;
        var cursor = 0;
        while ((cursor = projection.Markdown.IndexOf(needle, cursor, StringComparison.Ordinal)) >= 0)
        {
            count++;
            cursor += needle.Length;
        }
        return count;
    }

    private static bool Compare(JsonNode observed, string operation, JsonNode? expected) =>
        operation switch
        {
            "equals" => Json(observed) == Json(expected),
            "atLeast" => Number(observed) >= Number(expected),
            "contains" => observed.GetValue<string>().Contains(
                expected?.GetValue<string>() ?? string.Empty, StringComparison.Ordinal),
            "setEquals" => Set(observed).SetEquals(Set(expected)),
            _ => throw new ScenarioValidationException($"unknown invariant operator '{operation}'"),
        };

    private static MetricResult EvaluateTargetPrecision(byte[] expected, byte[] candidate)
    {
        var expectedDigest = PackageDigest(expected);
        var candidateDigest = PackageDigest(candidate);
        var differences = ChangedParts(expectedDigest, candidateDigest);
        var passed = differences.Count == 0;
        return new MetricResult("target-precision.reference-equivalence", "target_precision",
            passed ? "passed" : "failed",
            passed
                ? "candidate is normalized-package equivalent to the scripted expected output"
                : $"candidate differs from scripted expected output in: {string.Join(", ", differences)}",
            passed ? 1 : 0);
    }

    private static IReadOnlyList<MetricResult> EvaluateChangeBudget(
        LegalScenario scenario, byte[] input, byte[] candidate)
    {
        var changedParts = ChangedParts(PackageDigest(input), PackageDigest(candidate));
        var unexpectedParts = changedParts
            .Where(value => !scenario.ChangeBudget.AllowedChangedParts.Contains(value))
            .OrderBy(value => value, StringComparer.Ordinal).ToList();
        var partPass = unexpectedParts.Count == 0;

        var inputAccepted = RevisionProcessor.AcceptRevisions(
            new WmlDocument("input.docx", input));
        var candidateAccepted = RevisionProcessor.AcceptRevisions(
            new WmlDocument("candidate.docx", candidate));
        var revisions = DocxDiff.GetRevisions(inputAccepted, candidateAccepted,
            DiffSettings("Legal Evaluation Anchor Metric"));
        var anchors = revisions.SelectMany(value => new[] { value.LeftAnchor, value.RightAnchor })
            .Where(value => value is not null).Cast<string>().Distinct(StringComparer.Ordinal).ToList();
        var anchorPass = anchors.Count <= scenario.ChangeBudget.MaximumChangedAnchors;

        return new[]
        {
            new MetricResult("unintended-change.parts", "unintended_change",
                partPass ? "passed" : "failed",
                partPass
                    ? $"changed parts are within budget: {string.Join(", ", changedParts)}"
                    : $"unexpected changed parts: {string.Join(", ", unexpectedParts)}",
                partPass ? 1 : 0),
            new MetricResult("unintended-change.anchors", "unintended_change",
                anchorPass ? "passed" : "failed",
                $"distinct DocxDiff anchors={anchors.Count}; maximum={scenario.ChangeBudget.MaximumChangedAnchors}",
                anchorPass ? 1 : 0),
        };
    }

    private static MetricResult EvaluateValidity(byte[] candidate)
    {
        using var stream = new MemoryStream(candidate);
        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(document)
            .Where(value => !(value.Description ?? string.Empty).Contains(
                "powertools.codeplex.com", StringComparison.Ordinal)).ToList();
        return new MetricResult("document-validity.openxml", "document_validity",
            errors.Count == 0 ? "passed" : "failed",
            errors.Count == 0
                ? "OpenXmlValidator reported zero material errors"
                : string.Join(" | ", errors.Take(5).Select(value => value.Description)),
            errors.Count == 0 ? 1 : 0);
    }

    private static (MetricResult Metric, byte[]? RedlineBytes) EvaluateRedlineReversibility(
        byte[] input, byte[] candidate)
    {
        try
        {
            var flatInput = RevisionProcessor.AcceptRevisions(new WmlDocument("input.docx", input));
            var flatCandidate = RevisionProcessor.AcceptRevisions(new WmlDocument("candidate.docx", candidate));
            var redline = DocxDiff.Compare(flatInput, flatCandidate,
                DiffSettings("Legal Evaluation Redline"));
            var accepted = RevisionProcessor.AcceptRevisions(redline);
            var rejected = RevisionProcessor.RejectRevisions(redline);
            var acceptPass = RevisionModeledSignature(accepted.DocumentByteArray)
                == RevisionModeledSignature(flatCandidate.DocumentByteArray);
            var rejectPass = RevisionModeledSignature(rejected.DocumentByteArray)
                == RevisionModeledSignature(flatInput.DocumentByteArray);
            var passed = acceptPass && rejectPass;
            return (new MetricResult("redline-reversibility.interim-text-projection",
                    "redline_reversibility", passed ? "passed" : "failed",
                    $"acceptEqualsCandidate={acceptPass}; rejectEqualsBaseline={rejectPass}; "
                    + "interimScope=body/header/footer text projection; not a full-surface proof; "
                    + "pre-existing revisions flattened",
                    passed ? 1 : 0),
                GeneratedPackageNormalizer.Normalize(redline.DocumentByteArray));
        }
        catch (Exception exception)
        {
            return (new MetricResult("redline-reversibility.interim-text-projection",
                "redline_reversibility", "failed", exception.Message, 0), null);
        }
    }

    private static string RevisionModeledSignature(byte[] bytes)
    {
        var projection = WmlToMarkdownConverter.Convert(new WmlDocument("signature.docx", bytes),
            new WmlToMarkdownConverterSettings
            {
                AnchorMode = AnchorRenderMode.None,
                Scopes = ProjectionScopes.Body | ProjectionScopes.Headers | ProjectionScopes.Footers,
                TrackedChanges = TrackedChangeMode.Accept,
                ResolveNumbering = true,
            });
        return projection.Markdown.Replace("\r\n", "\n", StringComparison.Ordinal);
    }

    private static (MetricResult Metric, string? CandidateHtml) EvaluateRendering(
        byte[] expected, byte[] candidate)
    {
        try
        {
            var expectedHtml = RenderHtml(expected);
            var candidateHtml = RenderHtml(candidate);
            var passed = NormalizeHtml(expectedHtml) == NormalizeHtml(candidateHtml);
            return (new MetricResult("rendering-regression.html-projection", "rendering_regression",
                    passed ? "passed" : "failed",
                    passed
                        ? "sanitized Docxodus HTML projection matches the pinned expected output; this is not a visual-layout proof"
                        : "sanitized Docxodus HTML projection differs from the pinned expected output; this is not a visual-layout proof",
                    passed ? 1 : 0),
                candidateHtml);
        }
        catch (Exception exception)
        {
            return (new MetricResult("rendering-regression.html-projection", "rendering_regression",
                "failed", exception.Message, 0), null);
        }
    }

    internal static string RenderHtml(byte[] bytes)
    {
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = true,
            CssClassPrefix = "legal-eval-",
            PageTitle = "Docxodus legal evaluation",
        };
        var html = WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument("candidate.docx", bytes), settings);
        HardenPreview(html);
        return html.ToString(SaveOptions.DisableFormatting);
    }

    private static void HardenPreview(XElement html)
    {
        var head = html.Elements().FirstOrDefault(value => value.Name.LocalName == "head");
        head?.AddFirst(new XElement(html.Name.Namespace + "meta",
            new XAttribute("http-equiv", "Content-Security-Policy"),
            new XAttribute("content",
                "default-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'; "
                + "base-uri 'none'; form-action 'none'; frame-ancestors 'none'")));

        foreach (var element in html.DescendantsAndSelf())
        {
            foreach (var attribute in element.Attributes()
                .Where(value => value.Name.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                .ToList())
                attribute.Remove();

            var href = element.Attribute("href");
            if (href is not null && !href.Value.StartsWith('#'))
            {
                element.SetAttributeValue("title", "External link suppressed in evaluation preview");
                href.Value = "#";
            }
            var source = element.Attribute("src");
            if (source is not null && !source.Value.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                source.Remove();
        }
    }

    private static string NormalizeHtml(string html) =>
        html.Replace("\r\n", "\n", StringComparison.Ordinal).Trim();

    private static (string? Json, string? Error) TrySemanticDiff(byte[] input, byte[] candidate)
    {
        try
        {
            return (DocxDiff.GetEditScriptJson(
                new WmlDocument("before.docx", input),
                new WmlDocument("after.docx", candidate),
                DiffSettings("Legal Evaluation Artifact")), null);
        }
        catch (Exception exception)
        {
            return (null, exception.Message);
        }
    }

    internal static (string Json, bool Succeeded) SemanticDiffArtifact(
        byte[] input, byte[] candidate, string artifact)
    {
        var result = TrySemanticDiff(input, candidate);
        return (result.Json ?? ErrorEnvelope(artifact, "failed",
            result.Error ?? "semantic diff generation failed"), result.Json is not null);
    }

    internal static string ErrorEnvelope(string artifact, string status, string detail) =>
        JsonSerializer.Serialize(new
        {
            schemaVersion = "docxodus.evaluation-error/1.0",
            artifact,
            status,
            detail,
        }, new JsonSerializerOptions { WriteIndented = true });

    private static DocxDiffSettings DiffSettings(string author) => new()
    {
        AuthorForRevisions = author,
        Deterministic = true,
        DateTimeForRevisions = "2026-01-15T12:00:00Z",
        CompareHeadersFooters = true,
        TrackBlockFormatChanges = true,
        PreAcceptInputRevisions = true,
    };

    internal static IReadOnlyDictionary<string, string> PackageDigest(byte[] bytes)
    {
        var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read, leaveOpen: false);
        foreach (var entry in archive.Entries.OrderBy(value => value.FullName, StringComparer.Ordinal))
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            var payload = copy.ToArray();
            if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                payload = CanonicalXml(payload);
            result["/" + entry.FullName] = Convert.ToHexString(SHA256.HashData(payload)).ToLowerInvariant();
        }
        return result;
    }

    private static byte[] CanonicalXml(byte[] payload)
    {
        try
        {
            using var reader = XmlReader.Create(new MemoryStream(payload), new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
            });
            var document = XDocument.Load(reader, LoadOptions.None);
            if (document.Root is null) return payload;
            var canonical = CanonicalElement(document.Root);
            return Encoding.UTF8.GetBytes(canonical.ToString(SaveOptions.DisableFormatting));
        }
        catch (XmlException)
        {
            return payload;
        }
    }

    private static XElement CanonicalElement(XElement element) =>
        new(element.Name,
            element.Attributes().Where(value => !value.IsNamespaceDeclaration)
                .OrderBy(value => value.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(value => value.Name.LocalName, StringComparer.Ordinal)
                .Select(value => new XAttribute(value.Name, value.Value)),
            element.Nodes().Where(value => value is not XText text || !string.IsNullOrWhiteSpace(text.Value))
                .Select(value => value is XElement child
                    ? CanonicalElement(child)
                    : value is XText text ? new XText(text.Value) : value));

    private static IReadOnlyList<string> ChangedParts(
        IReadOnlyDictionary<string, string> left,
        IReadOnlyDictionary<string, string> right) =>
        left.Keys.Union(right.Keys, StringComparer.Ordinal)
            .Where(key => !left.TryGetValue(key, out var leftHash)
                || !right.TryGetValue(key, out var rightHash)
                || leftHash != rightHash)
            .OrderBy(value => value, StringComparer.Ordinal).ToList();

    private static byte[]? PartBytes(byte[] bytes, string part, bool required)
    {
        var normalized = part.TrimStart('/');
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read, leaveOpen: false);
        var entry = archive.GetEntry(normalized);
        if (entry is null)
        {
            if (required) throw new InvalidOperationException($"package part '{part}' is absent");
            return null;
        }
        using var stream = entry.Open();
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private static string RequiredString(JsonObject parent, string name) =>
        parent[name]?.GetValue<string>()
            ?? throw new ScenarioValidationException($"probe property '{name}' must be a string");

    private static string Json(JsonNode? node) =>
        node?.ToJsonString(new JsonSerializerOptions { WriteIndented = false }) ?? "null";

    private static double Number(JsonNode? node)
    {
        if (node is not JsonValue value)
            throw new ScenarioValidationException("numeric invariant operand is null or non-scalar");
        if (value.TryGetValue<int>(out var intValue)) return intValue;
        if (value.TryGetValue<long>(out var longValue)) return longValue;
        if (value.TryGetValue<double>(out var doubleValue)) return doubleValue;
        if (value.TryGetValue<decimal>(out var decimalValue)) return (double)decimalValue;
        throw new ScenarioValidationException("numeric invariant operand is not a number");
    }

    private static HashSet<string> Set(JsonNode? node) =>
        node is JsonArray array
            ? array.Select(value => value?.GetValue<string>()
                ?? throw new ScenarioValidationException("set members must be strings"))
                .ToHashSet(StringComparer.Ordinal)
            : throw new ScenarioValidationException("set operand must be an array");
}
