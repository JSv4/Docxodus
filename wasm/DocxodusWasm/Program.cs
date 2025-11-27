// Docxodus WASM Entry Point
// This file initializes the WASM runtime and keeps it alive

Console.WriteLine("Docxodus WASM Library Initialized");
Console.WriteLine($".NET Runtime: {Environment.Version}");

// Keep the runtime alive for JavaScript interop
await Task.Delay(-1);
