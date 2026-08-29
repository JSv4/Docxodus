#nullable enable

using System;
using System.Threading.Tasks;

namespace Docxodus.Internal;

/// <summary>
/// Fan independent, pure work out across threads — but only where the runtime actually has them.
/// </summary>
/// <remarks>
/// <para>The diff engine reads its input documents concurrently: the reads share no state and are pure
/// functions of their arguments, so running them together is a wall-clock win and nothing else. That
/// holds on a server. It does not hold in the browser.</para>
///
/// <para><b>Why the guard exists.</b> <c>wasm/DocxodusWasm/DocxodusWasm.csproj</c> does not set
/// <c>WasmEnableThreads</c>, so the browser runtime is single-threaded. There <c>Task.Run</c> does not
/// start a second thread: it queues the delegate for the ONE thread — which is the very thread about to
/// block on the result. A blocking join would then wait forever for work that cannot start, hanging the
/// page rather than merely failing to be faster. The fan-out is therefore compiled out of the WASM
/// assembly entirely, and gated at runtime besides, since a single-core host gains nothing from it.</para>
///
/// <para>Both paths compute the same values in the same order. This is a scheduling decision, never a
/// semantic one.</para>
/// </remarks>
internal static class ParallelWork
{
    /// <summary>
    /// Whether independent work should be fanned out. False in the browser (single-threaded runtime)
    /// and on a single-core host, where the sequential path is both correct and no slower.
    /// </summary>
    internal static bool CanFanOut =>
#if WASM_BUILD
        false;
#else
        Environment.ProcessorCount > 1;
#endif

    /// <summary>
    /// Evaluate <paramref name="first"/> and <paramref name="second"/>, concurrently where the runtime
    /// allows it, and return their results in argument order.
    /// </summary>
    public static (T First, T Second) Pair<T>(Func<T> first, Func<T> second)
    {
        if (!CanFanOut)
            return (first(), second());

        // The caller's thread runs the second half rather than idling on two queued tasks.
        var firstTask = Task.Run(first);
        T secondResult;
        try
        {
            secondResult = second();
        }
        catch
        {
            // Sequentially, first runs first, so ITS failure is the one a caller would have seen.
            // Joining here both preserves that ordering and observes the task, so a concurrent
            // failure never becomes an unobserved task exception.
            firstTask.GetAwaiter().GetResult();
            throw;
        }

        return (firstTask.GetAwaiter().GetResult(), secondResult);
    }

    /// <summary>
    /// Evaluate <paramref name="head"/> and every element of <paramref name="rest"/>, concurrently where
    /// the runtime allows it. Results keep their argument order, which callers rely on (reviewer order is
    /// significant to consolidate conflict reporting).
    /// </summary>
    public static (T Head, T[] Others) Fan<T>(Func<T> head, Func<T>[] rest)
    {
        if (!CanFanOut)
        {
            var sequential = new T[rest.Length];
            var headOnly = head();
            for (var i = 0; i < rest.Length; i++)
                sequential[i] = rest[i]();
            return (headOnly, sequential);
        }

        var tasks = new Task<T>[rest.Length];
        for (var i = 0; i < rest.Length; i++)
        {
            var work = rest[i];
            tasks[i] = Task.Run(work);
        }

        T headResult;
        try
        {
            headResult = head();
        }
        catch
        {
            // Head's failure is the one the sequential order would have surfaced; drain the rest so
            // a concurrent failure is observed rather than left dangling, then rethrow it.
            foreach (var task in tasks)
            {
                try
                {
                    _ = task.GetAwaiter().GetResult();
                }
                catch
                {
                    // Deliberately swallowed: head's exception is the one being propagated.
                }
            }

            throw;
        }

        var results = new T[rest.Length];
        for (var i = 0; i < rest.Length; i++)
            results[i] = tasks[i].GetAwaiter().GetResult();
        return (headResult, results);
    }
}
