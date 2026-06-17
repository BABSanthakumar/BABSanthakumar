#!/bin/bash
set -euo pipefail

LOG_DIR="${CLAUDE_PROJECT_DIR:-$(pwd)}/logs"
LOG_FILE="$LOG_DIR/audit.jsonl"
mkdir -p "$LOG_DIR"

INPUT=$(cat)

TIMESTAMP=$(date -u +"%Y-%m-%dT%H:%M:%SZ")
USER="${USER:-$(whoami 2>/dev/null || echo 'unknown')}"

python3 -c "
import json, sys

raw = '''$INPUT'''
try:
    d = json.loads(raw)
except:
    d = {}

usage = d.get('usage', {})
entry = {
    'timestamp': '$TIMESTAMP',
    'session_id': d.get('session_id', 'unknown'),
    'user': '$USER',
    'event': 'session_end',
    'input_tokens': usage.get('input_tokens', 0),
    'output_tokens': usage.get('output_tokens', 0),
    'total_tokens': usage.get('input_tokens', 0) + usage.get('output_tokens', 0),
}
print(json.dumps(entry))
" >> "$LOG_FILE" 2>/dev/null || true
