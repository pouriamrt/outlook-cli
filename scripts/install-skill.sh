#!/usr/bin/env bash
# Installs the outlook-cli Claude Code skill to ~/.claude/skills/outlook-cli/
set -euo pipefail

SKILL_DIR="${HOME}/.claude/skills/outlook-cli"
SOURCE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)/skill"

if [[ ! -d "$SOURCE_DIR" ]]; then
  echo "Error: skill/ directory not found at $SOURCE_DIR" >&2
  exit 1
fi

mkdir -p "$SKILL_DIR"
cp "$SOURCE_DIR/SKILL.md" "$SKILL_DIR/SKILL.md"
echo "Installed skill to: $SKILL_DIR/SKILL.md"
echo
echo "Verify with:"
echo "  ls -la $SKILL_DIR"
echo
echo "Then restart Claude Code (or start a new conversation) so the skill is picked up."
