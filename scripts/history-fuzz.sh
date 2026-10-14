#!/usr/bin/env bash
# Deterministic extended version-history corpus. Every failing case prints its seed and trace.
set -euo pipefail
cd "$(dirname "$0")/.."
export DOCXODUS_HISTORY_FUZZ_SEEDS="${DOCXODUS_HISTORY_FUZZ_SEEDS:-64}"
export DOCXODUS_HISTORY_FUZZ_STEPS="${DOCXODUS_HISTORY_FUZZ_STEPS:-128}"
dotnet test Docxodus.Tests/Docxodus.Tests.csproj -c Release -m:1 /p:UseSharedCompilation=false \
  --filter 'FullyQualifiedName~DocxVersionModelFuzz|FullyQualifiedName~DocxVersionRequestFault|FullyQualifiedName~HistoryRequestJournal' \
  --logger 'console;verbosity=normal' --logger 'trx;LogFileName=history-fuzz.trx' "$@"
