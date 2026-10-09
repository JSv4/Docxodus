// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using Docxodus.OpCodegen;

// Regenerate the session plumbing from tools/op-descriptions (issue #1027).
//   dotnet run --project tools/op-codegen            rewrite the generated files
//   dotnet run --project tools/op-codegen -- --check exit 1 if any generated file is stale
var check = args.Contains("--check");
var rootIndex = Array.IndexOf(args, "--root");
var root = rootIndex >= 0 && rootIndex + 1 < args.Length ? Path.GetFullPath(args[rootIndex + 1]) : FindRoot();

string? Read(string path)
{
    var full = Path.Combine(root, path);
    return File.Exists(full) ? File.ReadAllText(full) : null;
}

IReadOnlyDictionary<string, string> expected;
try
{
    expected = PlumbingGenerator.Render(PlumbingGenerator.Generate(root), Read);
}
catch (CodegenException e)
{
    Console.Error.WriteLine($"op-codegen: {e.Message}");
    return 2;
}

var stale = 0;
foreach (var (path, text) in expected)
{
    var current = Read(path);
    if (current is not null && PlumbingGenerator.Normalize(current) == text) continue;
    stale++;
    if (check)
    {
        Console.Error.WriteLine($"stale: {path}");
        continue;
    }

    var full = Path.Combine(root, path);
    Directory.CreateDirectory(Path.GetDirectoryName(full)!);
    var bom = current is not null && current.StartsWith('﻿');
    File.WriteAllText(full, text, new UTF8Encoding(bom));
    Console.WriteLine($"wrote: {path}");
}

if (check && stale > 0)
{
    Console.Error.WriteLine("op-codegen: generated files are stale; run `dotnet run --project tools/op-codegen`");
    return 1;
}

Console.WriteLine(stale == 0 ? "op-codegen: generated files are up to date" : $"op-codegen: wrote {stale} file(s)");
return 0;

static string FindRoot()
{
    for (var dir = new DirectoryInfo(Directory.GetCurrentDirectory()); dir is not null; dir = dir.Parent)
    {
        if (File.Exists(Path.Combine(dir.FullName, "Docxodus.sln"))) return dir.FullName;
    }

    throw new InvalidOperationException("run from inside the Docxodus repository, or pass --root <path>");
}
