using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Docxodus.Tests;

/// <summary>What <see cref="TestProcess.Run"/> observed.</summary>
/// <param name="Output">Everything read from standard output, then standard error.</param>
/// <param name="TimedOut">The process was still running at the deadline and was killed.</param>
/// <param name="OutputIncomplete">A pipe was still open after the process ended: some process it
/// started kept the write end, so <paramref name="Output"/> holds only what had arrived by then.</param>
public sealed record TestProcessResult(string Output, bool TimedOut, bool OutputIncomplete);

/// <summary>
/// Runs a child process from a test without letting it hang the run (issue #1158).
/// </summary>
/// <remarks>
/// A test that reads a redirected stream to its end before waiting for the process blocks until every
/// holder of the pipe's write end closes it. A grandchild that outlives the process (LibreOffice's
/// <c>soffice.bin</c>, for one) keeps the pipe open, so the read never returns and the exit timeout
/// that follows it never runs. This reads both streams on background tasks, waits for the process with
/// a deadline, kills its whole process tree when the deadline passes, and then gives the streams a short
/// grace period: a pipe still held open after that is reported, not waited on.
/// </remarks>
public static class TestProcess
{
    /// <summary>How long to keep reading after the process has ended or been killed.</summary>
    public static readonly TimeSpan DefaultDrainGrace = TimeSpan.FromSeconds(5);

    public static TestProcessResult Run(ProcessStartInfo start, TimeSpan timeout, TimeSpan? drainGrace = null)
    {
        start.RedirectStandardOutput = true;
        start.RedirectStandardError = true;
        start.UseShellExecute = false;
        // Python scripts run from the source tree; keep their bytecode caches out of it.
        start.Environment["PYTHONDONTWRITEBYTECODE"] = "1";

        var stdout = new StringBuilder();
        var stderr = new StringBuilder();
        var process = Process.Start(start)
            ?? throw new InvalidOperationException($"Could not start {start.FileName}.");
        var pumps = new[] { Pump(process.StandardOutput, stdout), Pump(process.StandardError, stderr) };

        // With no async event handlers registered, the timed wait waits only for the process to exit.
        var timedOut = !process.WaitForExit(timeout);
        if (timedOut)
        {
            try { process.Kill(entireProcessTree: true); }
            catch (InvalidOperationException) { } // it exited between the wait and the kill
            process.WaitForExit(drainGrace ?? DefaultDrainGrace);
        }

        var drained = Task.WaitAll(pumps, drainGrace ?? DefaultDrainGrace);
        string output;
        lock (stdout) lock (stderr) output = stdout.ToString() + stderr.ToString();
        // A pump still reading owns its stream until whoever holds the pipe lets go; disposing the
        // process under it would race that read, so the handles are left to the finalizer then.
        if (drained) process.Dispose();
        return new TestProcessResult(output, timedOut, !drained);
    }

    // Reads in chunks rather than to the end, so whatever arrived is available when the drain gives up.
    // Each pump is its own thread, not a pool task: a busy test host can starve the thread pool for
    // longer than the drain grace, and then nothing would be read at all.
    private static Task Pump(StreamReader reader, StringBuilder sink)
    {
        var done = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            var buffer = new char[4096];
            try
            {
                int read;
                while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
                    lock (sink) sink.Append(buffer, 0, read);
            }
            catch (Exception e) when (e is IOException or ObjectDisposedException)
            {
                // The pipe broke or the reader was closed: what arrived so far is all there is.
            }
            finally
            {
                done.TrySetResult();
            }
        })
        { IsBackground = true, Name = "TestProcess output pump" };
        thread.Start();
        return done.Task;
    }
}
