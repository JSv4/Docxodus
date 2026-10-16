using System;
using System.Diagnostics;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="TestProcess"/> returns on time even when a process it started leaves a descendant holding
/// the output pipe (issue #1158). The commands use <c>sh</c>, so these run on Linux and macOS only.
/// </summary>
public class TestProcessTests
{
    private static ProcessStartInfo Shell(string script) => new("sh", new[] { "-c", script });

    [Fact]
    public void Run_ReturnsWhenADescendantKeepsTheOutputPipeOpen()
    {
        if (OperatingSystem.IsWindows()) return;

        // The shell exits at once; the background sleep it started keeps stdout's write end for 60 s.
        var watch = Stopwatch.StartNew();
        var result = TestProcess.Run(Shell("sleep 60 & echo $!"), TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(3));
        watch.Stop();
        KillQuietly(result.Output);

        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(15), $"Run took {watch.Elapsed}.");
        Assert.False(result.TimedOut);
        Assert.True(result.OutputIncomplete);
        Assert.Matches(@"^\d+\s*$", result.Output);
    }

    [Fact]
    public void Run_KillsAProcessThatOutlivesTheTimeoutAndKeepsItsOutput()
    {
        if (OperatingSystem.IsWindows()) return;

        var watch = Stopwatch.StartNew();
        var result = TestProcess.Run(Shell("echo started; sleep 60"), TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(2));
        watch.Stop();

        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(15), $"Run took {watch.Elapsed}.");
        Assert.True(result.TimedOut);
        // The kill takes the shell's child too, so nothing is left holding the pipe.
        Assert.False(result.OutputIncomplete);
        Assert.Equal("started", result.Output.Trim());
    }

    [Fact]
    public void Run_CollectsStandardOutputThenStandardError()
    {
        if (OperatingSystem.IsWindows()) return;

        var result = TestProcess.Run(Shell("echo to-stderr >&2; echo to-stdout"), TimeSpan.FromSeconds(30));

        Assert.False(result.TimedOut);
        Assert.False(result.OutputIncomplete);
        Assert.Equal("to-stdout\nto-stderr\n", result.Output);
    }

    private static void KillQuietly(string pidText)
    {
        if (!int.TryParse(pidText.Trim(), out var pid)) return;
        try { Process.GetProcessById(pid).Kill(); }
        catch (ArgumentException) { } // already gone
        catch (InvalidOperationException) { }
    }
}
