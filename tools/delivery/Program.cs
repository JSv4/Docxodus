#nullable enable

namespace Docxodus.DeliveryCli;

internal static class Program
{
    private static Task<int> Main(string[] args) =>
        DeliveryCli.RunAsync(args, Console.Out, Console.Error);
}
