// Docxodus WASM Bootstrap
// This file initializes the .NET WASM runtime and exposes exports to JavaScript

import { dotnet } from './_framework/dotnet.js';

const { getAssemblyExports, getConfig } = await dotnet
    .withDiagnosticTracing(false)
    .create();

const config = getConfig();
const exports = await getAssemblyExports(config.mainAssemblyName);

// Expose Docxodus APIs to window for testing
window.Docxodus = {
    DocumentConverter: exports.DocxodusWasm.DocumentConverter,
    DocumentComparer: exports.DocxodusWasm.DocumentComparer
};

console.log("Docxodus WASM loaded successfully");
console.log("Version:", window.Docxodus.DocumentConverter.GetVersion());

// Signal that Docxodus is ready
window.DocxodusReady = true;
if (window.onDocxodusReady) {
    window.onDocxodusReady();
}
