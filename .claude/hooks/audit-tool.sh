#!/bin/bash
set -euo pipefail

LOG_DIR="${CLAUDE_PROJECT_DIR:-$(pwd)}/logs"
LOG_FILE="$LOG_DIR/audit.jsonl"
mkdir -p "$LOG_DIR"

INPUT=$(cat)

TIMESTAMP=$(date -u +"%Y-%m-%dT%H:%M:%SZ")
SESSION_ID=$(echo "$INPUT" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('session_id','unknown'))" 2>/dev/null || echo "unknown")
TOOL=$(echo "$INPUT" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('tool_name','unknown'))" 2>/dev/null || echo "unknown")
USER="${USER:-$(whoami 2>/dev/null || echo 'unknown')}"

# Extract file path for file-modifying tools
FILE_PATH=$(echo "$INPUT" | python3 -c "
import sys, json
d = json.load(sys.stdin)
inp = d.get('tool_input', {})
print(inp.get('file_path', inp.get('path', inp.get('notebook_path', ''))))
" 2>/dev/null || echo "")

# Extract command for Bash tool
COMMAND=$(echo "$INPUT" | python3 -c "
import sys, json
d = json.load(sys.stdin)
inp = d.get('tool_input', {})
cmd = inp.get('command', '')
print(cmd[:200] if cmd else '')
" 2>/dev/null || echo "")

python3 -c "
import json, sys
entry = {
    'timestamp': '$TIMESTAMP',
    'session_id': '$SESSION_ID',
    'user': '$USER',
    'tool': '$TOOL',
    'file': '$FILE_PATH',
    'command': $(echo "$COMMAND" | python3 -c "import sys,json; print(json.dumps(sys.stdin.read().strip()))"),
}
print(json.dumps(entry))
" >> "$LOG_FILE" 2>/dev/null || true
