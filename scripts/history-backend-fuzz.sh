#!/usr/bin/env bash
# Extended backend stream campaign plus real subprocess-crash recovery and fault matrices.
set -euo pipefail
cd "$(dirname "$0")/.."
export DOCXODUS_BACKEND_FUZZ_SEEDS="${DOCXODUS_BACKEND_FUZZ_SEEDS:-32}"
export DOCXODUS_BACKEND_FUZZ_ROUNDS="${DOCXODUS_BACKEND_FUZZ_ROUNDS:-16}"
dotnet test Docxodus.Tests/Docxodus.Tests.csproj -c Release -m:1 /p:UseSharedCompilation=false \
  --filter 'FullyQualifiedName~DocxBackend|FullyQualifiedName~DocxHistoryProcessRecovery' \
  --logger 'console;verbosity=normal' --logger 'trx;LogFileName=history-backend-fuzz.trx' "$@"
