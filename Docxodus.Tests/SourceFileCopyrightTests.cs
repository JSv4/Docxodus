// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Diagnostics;
using System.Text;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Enforces the two source-hygiene rules stated in CLAUDE.md. Both were prose-only rules that a
/// single copy-paste had already broken across ~195 files before anyone noticed, so they are
/// asserted here rather than left to review: StyleCop cannot express either one. Its
/// <c>SA1636</c> compares every header against a single configured company name, which a
/// repository with two legitimate copyright holders can never satisfy — that rule is switched
/// off in <c>rules.ruleset</c> and this class is what replaced it.
/// </summary>
public class SourceFileCopyrightTests
{
    /// <summary>
    /// Files whose content descends from Microsoft's original OpenXmlPowerTools source and must
    /// keep its copyright notice. Every name here exists in
    /// <see href="https://github.com/OfficeDev/Open-Xml-PowerTools">OfficeDev/Open-Xml-PowerTools</see>;
    /// nothing written after the fork belongs in this list.
    /// </summary>
    private static readonly HashSet<string> InheritedFromMicrosoft = new(StringComparer.Ordinal)
    {
        "Docxodus/ColorParser.cs",
        "Docxodus/DocumentBuilder.cs",
        "Docxodus/FieldRetriever.cs",
        "Docxodus/FormattingAssembler.cs",
        "Docxodus/GetListItemText_Default.cs",
        "Docxodus/GetListItemText_fr_FR.cs",
        "Docxodus/GetListItemText_ru_RU.cs",
        "Docxodus/GetListItemText_sv_SE.cs",
        "Docxodus/GetListItemText_tr_TR.cs",
        "Docxodus/GetListItemText_zh_CN.cs",
        "Docxodus/HtmlToWmlConverter.cs",
        "Docxodus/HtmlToWmlConverterCore.cs",
        "Docxodus/HtmlToWmlCssApplier.cs",
        "Docxodus/HtmlToWmlCssParser.cs",
        "Docxodus/ListItemRetriever.cs",
        "Docxodus/MarkupSimplifier.cs",
        "Docxodus/MetricsGetter.cs",
        "Docxodus/OpenXmlRegex.cs",
        "Docxodus/Properties/AssemblyInfo.cs",
        "Docxodus/PtOpenXmlDocument.cs",
        "Docxodus/PtOpenXmlUtil.cs",
        "Docxodus/PtUtil.cs",
        "Docxodus/RevisionAccepter.cs",
        "Docxodus/RevisionProcessor.cs",
        "Docxodus/ScalarTypes.cs",
        "Docxodus/TestUtil.cs",
        "Docxodus/UnicodeMapper.cs",
        "Docxodus/WmlDocument.cs",
        "Docxodus/WmlToHtmlConverter.cs",
        "Docxodus.Tests/DocumentBuilderTests.cs",
        "Docxodus.Tests/HtmlConverterTests.cs",
        "Docxodus.Tests/HtmlToWmlConverterTests.cs",
        "Docxodus.Tests/HtmlToWmlReadAsXElement.cs",
        "Docxodus.Tests/MarkupSimplifierTests.cs",
        "Docxodus.Tests/MetricsGetterTests.cs",
        "Docxodus.Tests/OpenXmlRegexTests.cs",
        "Docxodus.Tests/RevisionAccepterTests.cs",
        "Docxodus.Tests/RevisionProcessorTests.cs",
        "Docxodus.Tests/TestsBase.cs",
        "Docxodus.Tests/UnicodeMapperTests.cs",
    };

    private const string MicrosoftNotice = "// Copyright (c) Microsoft. All rights reserved.";
    private const string ProjectNotice = "// Copyright (c) John Scrudato IV. All rights reserved.";
    private const string LicenseNotice =
        "// Licensed under the MIT license. See LICENSE file in the project root for full license information.";

    [Fact]
    public void OnlyFilesInheritedFromOpenXmlPowerToolsClaimMicrosoftsCopyright()
    {
        var claiming = SourceFiles()
            .Where(file => Header(file.Path).Any(line => line.StartsWith("// Copyright (c) Microsoft", StringComparison.Ordinal)))
            .Select(file => file.Relative)
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();

        // Set equality in both directions: a new post-fork file that copy-pasted the Microsoft
        // header fails, and so does stripping the notice off genuinely inherited source.
        Assert.Equal(InheritedFromMicrosoft.OrderBy(path => path, StringComparer.Ordinal).ToArray(), claiming);
    }

    [Fact]
    public void EveryFileHeaderUsesOneOfTheTwoApprovedNotices()
    {
        foreach (var file in SourceFiles())
        {
            var header = Header(file.Path);
            var copyright = header.FirstOrDefault(line => line.StartsWith("// Copyright", StringComparison.Ordinal));
            if (copyright is null) continue; // Most of the library carries no header at all; SA1633 covers that.

            var expected = InheritedFromMicrosoft.Contains(file.Relative) ? MicrosoftNotice : ProjectNotice;
            Assert.True(copyright == expected,
                $"{file.Relative} header reads \"{copyright}\" but must read \"{expected}\".");
            Assert.Contains(LicenseNotice, header);
        }
    }

    /// <summary>
    /// CLAUDE.md forbids per-file nullable directives because Docxodus.csproj already sets
    /// <c>&lt;Nullable&gt;enable&lt;/Nullable&gt;</c> project-wide. <c>#nullable disable</c> is the
    /// half that is an outright regression, and issue #645 finished removing the last one.
    /// </summary>
    [Fact]
    public void NoSourceFileOptsOutOfNullableReferenceTypes()
    {
        var disabled = SourceFiles()
            .Where(file => File.ReadLines(file.Path)
                .Any(line => line.TrimStart().StartsWith("#nullable disable", StringComparison.Ordinal)))
            .Select(file => file.Relative)
            .ToArray();

        Assert.Empty(disabled);
    }

    private static string[] Header(string path) => File.ReadLines(path).Take(5).ToArray();

    /// <summary>
    /// Every C# file git lists — tracked, or new and not ignored — so a gitignored scratch
    /// checkout under the root (issue #822) is not mistaken for source. Without a git
    /// checkout (a source tarball has no ignored content) it walks the tree minus build output.
    /// </summary>
    private static IEnumerable<(string Path, string Relative)> SourceFiles()
    {
        var root = RepositoryRoot();
        var relative = Path.Exists(Path.Combine(root, ".git"))
            ? GitListedSourceFiles(root)
            : Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories)
                .Select(path => Path.GetRelativePath(root, path).Replace('\\', '/'))
                .Where(path => !path.Contains("/bin/", StringComparison.Ordinal)
                    && !path.Contains("/obj/", StringComparison.Ordinal));
        return relative
            .Select(path => (Path: Path.Combine(root, path), Relative: path))
            .Where(file => File.Exists(file.Path)) // listed in the index but deleted from the worktree
            .OrderBy(file => file.Relative, StringComparer.Ordinal);
    }

    private static string[] GitListedSourceFiles(string root)
    {
        var start = new ProcessStartInfo("git", new[] { "ls-files", "-z", "--cached", "--others", "--exclude-standard", "--", "*.cs" })
        {
            WorkingDirectory = root,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
        };
        using var git = Process.Start(start)!;
        var error = git.StandardError.ReadToEndAsync(); // drain concurrently so neither pipe blocks
        var output = git.StandardOutput.ReadToEnd();
        git.WaitForExit();

        // No silent fallback to the directory walk: that would bring back the bug this avoids.
        Assert.True(git.ExitCode == 0, $"git ls-files exited {git.ExitCode} in {root}: {error.Result}");

        // An unmerged path is listed once per conflict stage.
        return output.Split('\0', StringSplitOptions.RemoveEmptyEntries).Distinct().ToArray();
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Docxodus.sln")))
            directory = directory.Parent;
        Assert.True(directory is not null, "Could not locate the repository root from " + AppContext.BaseDirectory);
        return directory!.FullName;
    }
}
