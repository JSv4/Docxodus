#nullable enable

using System.Diagnostics;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

public sealed class DocxHistoryProcessRecoveryTests(ITestOutputHelper output)
{
    [Theory]
    [InlineData("version", "before")]
    [InlineData("version", "after")]
    [InlineData("restore", "before")]
    [InlineData("restore", "after")]
    [InlineData("text", "before")]
    [InlineData("text", "after")]
    [InlineData("conflict", "before")]
    [InlineData("conflict", "after")]
    public async Task KilledProducerRecoversExactRequestInFreshProcess(string kind, string boundary)
    {
        var root = Directory.CreateTempSubdirectory("history-process-proof-").FullName;
        try
        {
            var seed = await Run("seed"); Assert.Equal(0, seed.Code);
            var crash = await Run("crash");
            Assert.NotEqual(0, crash.Code);
            if (OperatingSystem.IsLinux()) Assert.Equal(137, crash.Code); // Actual SIGKILL, not a caught exception/recreated service.
            Assert.Contains("terminating-at-" + boundary + "-cas", crash.Text);
            var recovery = await Run("recover"); Assert.Equal(0, recovery.Code);
            Assert.Contains("recovered " + kind + " " + boundary, recovery.Text);
            Assert.Contains("reflection disabled", recovery.Text);
        }
        finally { Directory.Delete(root, recursive: true); }

        async Task<(int Code, string Text)> Run(string phase)
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            var configuration = directory.Parent!.Name;
            while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Docxodus.Tests", "Docxodus.Tests.csproj")))
                directory = directory.Parent;
            Assert.NotNull(directory);
            var probe = Path.Combine(directory!.FullName, "tools", "history-recovery-probe", "bin", configuration, "net10.0", "HistoryRecoveryProbe.dll");
            Assert.True(File.Exists(probe), "Project-reference-built recovery probe is missing: " + probe);
            var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet")
            { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
            foreach (var value in new[] { probe, phase, root, kind, boundary }) start.ArgumentList.Add(value);
            using var process = Process.Start(start)!;
            var stdout = process.StandardOutput.ReadToEndAsync(); var stderr = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            try { await process.WaitForExitAsync(timeout.Token); }
            catch { if (!process.HasExited) process.Kill(entireProcessTree: true); await process.WaitForExitAsync(); throw; }
            var text = await stdout + await stderr;
            output.WriteLine(phase + ": " + text);
            return (process.ExitCode, text);
        }
    }
}
