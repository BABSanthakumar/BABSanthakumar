#!/usr/bin/env python3
"""View audit logs in a readable table format."""
import json
import sys
from pathlib import Path
from collections import defaultdict

LOG_FILE = Path(__file__).parent / "audit.jsonl"

if not LOG_FILE.exists():
    print("No audit log found yet.")
    sys.exit(0)

entries = []
with open(LOG_FILE) as f:
    for line in f:
        line = line.strip()
        if line:
            try:
                entries.append(json.loads(line))
            except json.JSONDecodeError:
                pass

if not entries:
    print("Log file is empty.")
    sys.exit(0)

# Summary
print(f"\n{'='*70}")
print(f"  CLAUDE CODE AUDIT LOG  ({len(entries)} events)")
print(f"{'='*70}\n")

# Token summary
sessions = defaultdict(lambda: {"input": 0, "output": 0, "user": "unknown"})
for e in entries:
    if e.get("event") == "session_end":
        sid = e["session_id"]
        sessions[sid]["input"] += e.get("input_tokens", 0)
        sessions[sid]["output"] += e.get("output_tokens", 0)
        sessions[sid]["user"] = e.get("user", "unknown")

if sessions:
    print("TOKEN USAGE BY SESSION:")
    print(f"  {'Session ID':<20} {'User':<15} {'Input':>10} {'Output':>10} {'Total':>10}")
    print(f"  {'-'*67}")
    for sid, data in sessions.items():
        total = data["input"] + data["output"]
        print(f"  {sid[:20]:<20} {data['user']:<15} {data['input']:>10,} {data['output']:>10,} {total:>10,}")
    print()

# Files modified
modified = [e for e in entries if e.get("file") and e.get("tool") in ("Edit", "Write", "NotebookEdit")]
if modified:
    print("FILES MODIFIED:")
    seen = set()
    for e in modified:
        key = (e["session_id"][:8], e["file"])
        if key not in seen:
            seen.add(key)
            print(f"  [{e['timestamp']}] {e['user']} | {e['tool']} | {e['file']}")
    print()

# All tool events
print("ALL EVENTS (most recent 20):")
print(f"  {'Timestamp':<22} {'User':<12} {'Tool':<18} {'Detail'}")
print(f"  {'-'*70}")
for e in entries[-20:]:
    if e.get("event") == "session_end":
        detail = f"tokens: {e.get('total_tokens', 0):,}"
        tool = "SESSION_END"
    else:
        tool = e.get("tool", "")
        detail = e.get("file") or (e.get("command", "")[:40])
    print(f"  {e['timestamp']:<22} {e.get('user','?'):<12} {tool:<18} {detail}")
