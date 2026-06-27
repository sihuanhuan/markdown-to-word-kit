#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [[ $# -eq 0 ]]; then
  cat <<'USAGE'
Usage:
  ./convert.sh input.md [output.docx] [options]

Examples:
  ./convert.sh thesis.md
  ./convert.sh thesis.md thesis.docx
  ./convert.sh thesis.md thesis.docx --auto-diagrams

Options:
  --no-toc          Do not generate a table of contents
  --auto-diagrams   Render Mermaid/PlantUML only when tools are available
  --diagrams        Require Mermaid/PlantUML rendering tools
USAGE
  exit 2
fi

input="$1"
shift

output=""
if [[ $# -gt 0 && "$1" == *.docx ]]; then
  output="$1"
  shift
fi

if [[ -n "$output" ]]; then
  python3 "$SCRIPT_DIR/_md2word/md2docx.py" "$input" -o "$output" "$@"
else
  python3 "$SCRIPT_DIR/_md2word/md2docx.py" "$input" "$@"
fi
