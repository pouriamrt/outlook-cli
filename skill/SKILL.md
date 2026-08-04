---
name: outlook-cli
description: Use this skill when the user asks about their Outlook mail or calendar — reading, sending, replying, scheduling, searching, summarizing threads, drafting replies, finding meeting times, or any task involving their Microsoft 365 inbox or calendar. Wraps the local 'outlook' CLI (must be authenticated via 'outlook login'). Triggers on phrases like "check my email", "what's on my calendar", "reply to that message from Alice", "find time with Bob", "summarize this thread".
---

# outlook-cli skill

You are wrapping the local `outlook` command-line tool. The CLI talks to the
Outlook REST API v2.0 on behalf of the signed-in user. It is the **only** way
you should access their mailbox — never import `outlook_cli` Python modules,
never make raw Graph/Outlook HTTP calls, never browse the web for their mail.

## Iron rules

1. **Always shell out.** Use `Bash` to run `outlook ...`. Never `import outlook_cli`.
2. **Always pass `--json`** when reading data (it goes before the subcommand:
   `outlook --json mail list ...`). Parse the JSON to extract what the user asked
   for. Render a concise summary to the user — do not dump raw JSON unless asked.
3. **Read-only is automatic, writes require confirmation.** See tables below.
   For ANY state-changing command, draft the exact command + body, show it to
   the user, call `AskUserQuestion`, run only on explicit approval.
4. **Indices are session-scoped.** The CLI assigns short integers `1, 2, 3, ...`
   to the items in the last list/search. They reset per family (`mail` vs `cal`).
   If you have not just listed, you must list before resolving an index.
5. **Never run `outlook login` yourself.** It requires an interactive browser +
   bookmarklet. If `whoami` returns 77, instruct the user to log in and STOP.
6. **Never invent flags or commands.** Stick strictly to what's documented here.
   If you need a flag you don't see, run `outlook <cmd> --help` first.

## Pre-flight (run before any other command)

```bash
outlook whoami --json
```

| Exit | What it means | What you do |
|---|---|---|
| `0`  | Logged in. JSON has `username`, `tenant_id`, `home_account_id`, `id_token_claims`. | Proceed. |
| `77` | Session expired or never logged in. | Stop. Tell the user: *"Your Outlook session has expired. Please run `outlook login` in your terminal and then ask me again."* |
| other | Unexpected. | Surface stderr to the user. Do not retry blindly. |

## Read-only commands (use freely, no confirmation)

`--json` flag goes **before** the subcommand. All read commands populate the
per-family index cache (mail commands populate the `mail` cache, cal commands
the `cal` cache).

### Mail (read)

| Intent | Command |
|---|---|
| List inbox (default = inbox, 25 items) | `outlook --json mail list` |
| List with filters | `outlook --json mail list --folder NAME --unread --flagged --from ADDR --subject TEXT --since "2d" --top N --skip N --all` |
| Read one message (markdown body) | `outlook --json mail read <index-or-id>` |
| Read one message (raw HTML body) | `outlook --json mail read <index-or-id> --raw` |
| Save attachments | `outlook --json mail read <index> --save-attachments ./path/` |
| Read full conversation | `outlook --json mail thread <index-or-id>` |
| KQL search | `outlook --json mail search "from:bob subject:Q3 hasattachment:true" --folder NAME --top N` |

**`--since` accepts**: `"7d"`, `"2h"`, `"yesterday"`, ISO datetime, or natural language ("monday").

**`mail list` filters compose with AND** — `--unread --from alice --since 3d` returns unread messages from Alice in the last 3 days.

**KQL operators for `mail search`**: `from:` `to:` `cc:` `subject:` `body:`
`hasattachment:true` `received>=2026-05-01` `kind:meetings` `importance:high`.
Quoted phrases preserved. Operators can be combined.

### Calendar (read)

| Intent | Command |
|---|---|
| Today | `outlook --json cal today` |
| Tomorrow | `outlook --json cal tomorrow` |
| This week (7 days) | `outlook --json cal week` |
| Next week | `outlook --json cal week --next` |
| Arbitrary range | `outlook --json cal list --start "monday" --end "friday" --calendar "Vacation Calendar"` |
| Show one event | `outlook --json cal show <index-or-id>` |
| Show with attendees | `outlook --json cal show <index> --attendees` |
| Export as .ics | `outlook cal show <index> --ics` (no `--json`; pipes RFC 5545 to stdout) |
| Find meeting times | `outlook --json cal find-time --with a@x,b@y --duration 30m --window "this week" --per-day 3` |

**`--window` for find-time**: `"today"`, `"this week"`, `"next week"`, or any date. Week windows are calendar-week aligned (Mon–Fri, working hours); `this week` drops days already past. By default **every** free slot per day is shown. Pass `--per-day N` to trim each day to N slots sampled evenly across it (morning→afternoon), e.g. `--per-day 3` for a short list. Times display in your local timezone.

**`--duration` format**: `30m`, `1h`, `1h30m`, `90m`. No spaces.

### Meta

| Intent | Command |
|---|---|
| Current version + API target | `outlook version` |
| JSON Schema for any command | `outlook --json-schema mail.list` (or `mail.read`, `cal.today`, etc.) |
| Config list | `outlook config list` |
| Config get | `outlook config get default_folder` |

## State-changing commands (CONFIRM with `AskUserQuestion` first)

For each of these, your workflow is:

1. **Draft.** Compose the exact command, including the body if applicable.
2. **Show.** Print the full command and the body to the user in fenced blocks.
3. **Confirm.** Call `AskUserQuestion` with the draft. Recommended options:
   `"Send it"`, `"Edit first"`, `"Cancel"`.
4. **Execute.** Only on explicit "Send it" approval, run the command.
   Pipe multi-line bodies via stdin using `--body -`.
5. **Report.** Show the CLI's confirmation (exit 0 + message) to the user.

### Mail (write)

| Intent | Command |
|---|---|
| Send mail | `outlook mail send --to a@x [--to b@y] [--cc c@z] [--bcc d@w] --subject "S" --body -` (stdin) |
| Send HTML mail | add `--html` |
| Send with attachments | `--attach ./file1.pdf --attach ./file2.png` (repeatable) |
| Save as draft (don't send) | add `--draft` |
| Set importance | `--importance low|normal|high` (default: `normal`) |
| Reply | `outlook mail reply <index> --body -` |
| Reply-all | `outlook mail reply <index> --all --body -` |
| Reply CC/BCC | `--cc a@x --bcc b@y` |
| Forward | `outlook mail forward <index> --to a@x --body -` |
| Move to folder | `outlook mail move <index> <folder-name>` |
| Soft-delete (to Deleted Items) | `outlook mail delete <index>` |
| **Hard-delete (irreversible)** | `outlook mail delete <index> --purge` ⚠️ |
| Flag for follow-up | `outlook mail flag <index> [--due "fri"]` |
| Unflag | `outlook mail unflag <index>` |
| Mark read | `outlook mail mark <index> --read` |
| Mark unread | `outlook mail mark <index> --unread` |

**`--body -`** reads the body from stdin — that's how you pipe a multi-line
draft. If you omit `--body` entirely on `mail send`, the CLI opens `$EDITOR`,
which will hang in a non-interactive session — always pass `--body -` with stdin.

**`--purge` is one-way.** Always extra-confirm hard deletes ("This permanently
deletes the message and cannot be undone. Confirm?").

**`mark` is low-stakes and reversible.** `--read`/`--unread` only flips the
`isRead` state in the user's own mailbox — never visible to senders or other
recipients, and instantly reversible. Don't heavy-confirm it like `--purge`.
But still never mark messages read as a silent side effect of reading or
summarizing a thread (see the prohibitions section).

### Calendar (write)

| Intent | Command |
|---|---|
| Create event (duration) | `outlook cal create --title T --start "tomorrow 3pm" --duration 30m --location "Room 2"` |
| Create event (end time) | `outlook cal create --title T --start "..." --end "..."` |
| Create with attendees | `--invitees a@x,b@y` (comma-separated or repeated) |
| Create as Teams meeting | add `--online` |
| Create all-day | add `--all-day` |
| With pre-written body | `--body "agenda text"` (omit to open $EDITOR — don't omit in non-interactive) |
| Respond to invite | `outlook cal respond <index> --accept|--decline|--tentative [--comment "..."]` |
| Cancel (organizer only) | `outlook cal cancel <index> --comment "Apologies, rescheduling"` |

**`--start`/`--end` accept**: ISO datetime (`2026-05-23T14:00`) or natural
language (`"tomorrow 3pm"`, `"monday 9am"`, `"in 2 hours"`).

**You must provide either `--end` OR `--duration`** to `cal create`. Not both.

## Indices in depth

Lists and searches assign integers `1..N` to results. The cache is
**per-family**:

- The mail family covers `mail list`, `mail search`. Index `3` from a search
  refers to whatever message was the 3rd result of *that* search until you
  run another list/search.
- The cal family covers `cal today`, `cal tomorrow`, `cal week`, `cal list`.
  Same per-family semantics.

**Mail and cal indices never collide.** `mail read 1` and `cal show 1` refer to
unrelated items.

**If the user says "the third one" but the family cache is stale or empty:**
re-run the appropriate list first, *then* resolve.

**You can always use the long Graph ID** instead of an index. The JSON output
of any list/show includes `id` — pass that in place of the integer for an
explicit, unambiguous reference (handy when the user has been chatting for a
while and you're not sure the cache is fresh).

## JSON contract

The `--json` output has a stable shape across versions. Top-level keys are
frozen; new fields may be added but never renamed or removed.

```jsonc
// outlook --json mail list
{
  "items": [
    {
      "id": "AAMkAD...",
      "index": 1,
      "from": { "name": "Alice", "address": "alice@example.com" },
      "to":   [{ "name": "...", "address": "..." }],
      "subject": "Q3 plans",
      "received": "2026-05-22T14:33:00Z",
      "is_read": false,
      "is_flagged": false,
      "has_attachments": false,
      "preview": "First 255 chars of the body..."
    }
  ]
}

// outlook --json mail read <index>
{
  "id": "...", "index": 1,
  "from": {...}, "to": [...], "cc": [...], "bcc": [...],
  "subject": "...", "received": "...", "sent": "...",
  "is_read": true, "is_flagged": false,
  "body_text": "plaintext-or-converted-markdown",
  "body_html": "<html>...</html>",
  "attachments": [{ "name": "file.pdf", "size": 12345, "content_type": "application/pdf" }],
  "conversation_id": "..."
}
```

For any other command, get the exact schema with
`outlook --json-schema <command>` (e.g. `mail.list`, `mail.read`, `cal.today`).

## Exit codes

| Code | Meaning | Your reaction |
|---|---|---|
| `0` | Success | Proceed. |
| `1` | Generic error (stderr has detail) | Surface the stderr line to the user. |
| `2` | Usage error (bad flags) | You made a flag mistake — re-check `outlook <cmd> --help`. |
| `64` | Not found (bad index or ID) | The index is stale; re-run the appropriate list, then retry. |
| `77` | Session expired | Stop. Tell user to `outlook login`. |

## Worked examples

### 1. "Summarize my unread"

```bash
outlook --json mail list --unread --top 10
```

For each item with `is_read: false`, optionally:
```bash
outlook --json mail read <index>
```

Then produce a per-sender summary, mentioning subject lines and dates. **Do
not** mark anything read or send anything.

### 2. "Reply to Bob's email about Q3"

```bash
outlook --json mail search "from:bob subject:Q3" --top 5
```

Identify the right message; if ambiguous, ask the user which index.
Read the body:

```bash
outlook --json mail read <index>
```

Draft the reply. Show the user the draft and the exact command:
```bash
outlook mail reply <index> --body -
```
with the draft body that would be piped via stdin shown in a fenced block.

Use `AskUserQuestion` with options like `"Send it"`, `"Edit first"`,
`"Cancel"`. On `"Send it"`, run:

```bash
printf '%s\n' "$DRAFT_BODY" | outlook mail reply <index> --body -
```

### 3. "Find a 30-minute slot with Alice and Bob this week, then book it"

```bash
outlook --json cal find-time --with alice@example.com,bob@example.com --duration 30m --window "this week"
```

Parse `.suggestions`. Show the top 3 with start times and confidence.
If the user picks one, draft:

```bash
outlook cal create \
  --title "Quick sync" \
  --start "wed 14:00" \
  --duration 30m \
  --invitees alice@example.com,bob@example.com \
  --online \
  --body "agenda goes here"
```

Confirm with `AskUserQuestion`. On approval, run it. Tell the user the new
event's index (parsed from the CLI's success line).

### 4. "Forward the email from finance to the team with a note"

```bash
outlook --json mail search "from:finance" --top 5
```

Pick the right one (ask if ambiguous).

Draft:
```bash
outlook mail forward <index> --to team@example.com --body -
```

with the note body. Confirm. Send.

### 5. "Decline the 3pm meeting"

```bash
outlook --json cal today
```

Find the 3pm event. Draft:
```bash
outlook cal respond <index> --decline --comment "Sorry, schedule conflict"
```

Confirm. Send.

### 6. "Archive everything older than 7 days from newsletters@"

This is **multiple state changes** — confirm the batch plan first, not each
individual move. Draft:

```bash
# Step 1: identify
outlook --json mail list --from newsletters@example.com --since 7d --all
# Step 2: for each item.index, move to "Archive"
```

Show the user the count + the move command pattern. After approval, loop:
```bash
for idx in $INDICES; do outlook mail move $idx Archive; done
```

### 7. "What's the JSON shape of mail read?"

```bash
outlook --json-schema mail.read
```

Show the schema. Don't fetch real mail to demo.

## Error recovery

| Symptom | Fix |
|---|---|
| Exit `77` | Stop. *"Your Outlook session has expired. Run `outlook login` in your terminal."* |
| Exit `64` after `mail read 3` | Re-run `mail list`, then re-resolve `3`. |
| Exit `64` after `cal show 3` | Re-run `cal today` / `cal week`, then re-resolve. |
| Exit `2` | You used a flag that doesn't exist on that subcommand. Check `outlook <cmd> --help`. |
| Exit `1` with stderr "Search returned no results" | The query was valid, just empty. Tell the user. |
| Exit `1` with stderr about HTTP / network | Likely transient; show the user; offer to retry. Do not retry blindly. |
| `mail send` hangs | You forgot `--body -` and the CLI is waiting on `$EDITOR`. Cancel, retry with `--body -` and stdin. |

## What I will NOT do

- **Send mail** without explicit user approval via `AskUserQuestion`.
- **Delete** (especially `--purge`) without explicit user approval.
- **Cancel meetings** without explicit user approval.
- **Mark messages read** as a side effect of summarizing (this changes server state).
- **Move messages** to other folders without explicit approval.
- **Forward** without explicit approval (data exfiltration risk).
- **Import `outlook_cli` Python modules** — always shell out via `Bash`.
- **Run `outlook login`** myself — requires interactive browser + bookmarklet.
- **Invent flags** — if a flag isn't in this skill or in `--help`, it doesn't exist.
- **Promise behavior I can't verify** — if uncertain, run `--help` or check
  `--json-schema`.

## Configuration

The user has a config file at `~/.config/outlook-cli/config.toml`. You can
inspect it (read-only) with `outlook config list`. Don't write to it without
explicit instruction — it changes CLI defaults across all sessions.

## Environment variable overrides (for diagnostics)

If the user reports unexpected behavior, these env vars exist:

| Var | Effect |
|---|---|
| `OUTLOOK_CLI_API_BASE` | Override API base (default `https://outlook.office.com/api/v2.0`). |
| `OUTLOOK_CLI_API_SCOPE` | Override scope (default `https://outlook.office.com/.default`). |
| `OUTLOOK_CLI_CA_BUNDLE` | Path to corporate root CA bundle (for MITM proxies). |
| `OUTLOOK_CLI_INSECURE=1` | Disable TLS verification entirely. **Diagnostic only — never suggest as a fix.** |

Do not set these for the user — only mention them if they ask about TLS or
endpoint issues.
