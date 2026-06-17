# Claude Code Audit Tracker

Tracks all Claude Code activity: requests, file modifications, tokens used, and user identity.

## What Gets Logged

| Field | Description |
|---|---|
| timestamp | UTC time of the event |
| session_id | Unique Claude Code session ID |
| user | OS username running the session |
| tool | Tool used (Edit, Write, Bash, Read, etc.) |
| file | File path (for file-modifying tools) |
| command | Command run (for Bash tool) |
| input_tokens | Tokens sent to Claude (logged at session end) |
| output_tokens | Tokens received from Claude (logged at session end) |

## Log Location

All logs are written to `logs/audit.jsonl` — one JSON object per line.

## View Logs

```bash
python3 logs/view-audit.py
```

Shows:
- Token usage by session and user
- All files modified
- Recent activity table

## Architecture

```
Claude Code Session
      ↓ (every tool call)
PostToolUse hook → logs/audit.jsonl
      ↓ (session ends)
Stop hook → logs session token totals
```

## How It Works

- `.claude/settings.json` registers two hooks
- `PostToolUse` fires after every tool call — logs tool name, file, command, user
- `Stop` fires when the session ends — logs total input/output tokens
- Logs are append-only JSONL — never overwritten
