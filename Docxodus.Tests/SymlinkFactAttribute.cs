// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A fact that creates symbolic links. Windows lets a process create one only with Developer Mode
/// on or the SeCreateSymbolicLinkPrivilege ("A required privilege is not held by the client"); there
/// such a test is skipped with that reason instead of failing. Where links can be made, as on Linux
/// and macOS, it always runs. The OS is asked once per test run.
/// </summary>
public sealed class SymlinkFactAttribute : FactAttribute
{
    private static readonly Lazy<string?> Refusal = new(Probe);

    public SymlinkFactAttribute()
    {
        if (Refusal.Value is { } reason)
            Skip = reason;
    }

    /// <summary>Why this process cannot create symbolic links, or null when it can.</summary>
    internal static string? RefusalReason => Refusal.Value;

    private static string? Probe()
    {
        var directory = Directory.CreateTempSubdirectory("docxodus-symlink-probe-");
        try
        {
            File.CreateSymbolicLink(Path.Combine(directory.FullName, "link"), Path.Combine(directory.FullName, "target"));
            return null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
        {
            // The message's first clause; Windows appends the probe's path after " : ".
            return $"This OS does not let the test process create symbolic links ({ex.Message.Split(" : ")[0].Trim()}); "
                + "on Windows, turn on Developer Mode to run it.";
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }
}

public class SymlinkFactAttributeTests
{
    /// <summary>The skip is for Windows' privilege rule only: on Linux and macOS, where CI runs these
    /// tests, a probe that failed would silently skip every symlink test.</summary>
    [Fact]
    public void LinksCanBeCreatedOffWindows()
    {
        if (!OperatingSystem.IsWindows())
            Assert.Null(SymlinkFactAttribute.RefusalReason);
    }
}
