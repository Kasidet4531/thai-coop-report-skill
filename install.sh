#!/usr/bin/env bash
# Install the coop-report-docx skill for Claude Code, Codex and/or Antigravity.
#
#   bash install.sh [claude|codex|antigravity|antigravity-cli|project|all] ...
#   curl -fsSL https://raw.githubusercontent.com/Kasidet4531/thai-coop-report-skill/main/install.sh | bash -s -- codex antigravity
#
# Run from a clone it copies the local skill; piped from curl it downloads the repo zip.
# "project" installs into ./.agents/skills of the current folder (read by Codex and Antigravity).
set -euo pipefail

REPO="Kasidet4531/thai-coop-report-skill"
SKILL="coop-report-docx"
targets=("$@")
[ ${#targets[@]} -eq 0 ] && targets=(all)

dest_for() {
  case "$1" in
    claude)          echo "$HOME/.claude/skills" ;;
    codex)           echo "$HOME/.agents/skills" ;;
    antigravity)     echo "$HOME/.gemini/config/skills" ;;
    antigravity-cli) echo "$HOME/.gemini/antigravity-cli/skills" ;;
    project)         echo "$PWD/.agents/skills" ;;
    *) echo "unknown target: $1 (use claude, codex, antigravity, antigravity-cli, project, all)" >&2; exit 1 ;;
  esac
}

expanded=()
for t in "${targets[@]}"; do
  if [ "$t" = all ]; then expanded+=(claude codex antigravity); else expanded+=("$t"); fi
done

# Locate the skill source: local clone, or download the repo zip into a temp dir.
here="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" 2>/dev/null && pwd || true)"
if [ -n "$here" ] && [ -f "$here/skills/$SKILL/SKILL.md" ]; then
  src="$here/skills/$SKILL"
else
  tmp="$(mktemp -d)"
  trap 'rm -rf "$tmp"' EXIT
  echo "Downloading $REPO ..."
  curl -fsSL "https://github.com/$REPO/archive/refs/heads/main.zip" -o "$tmp/repo.zip"
  (cd "$tmp" && unzip -q repo.zip)
  src="$(find "$tmp" -type d -path "*/skills/$SKILL" | head -1)"
fi

for t in "${expanded[@]}"; do
  d="$(dest_for "$t")"
  mkdir -p "$d"
  rm -rf "${d:?}/$SKILL"
  cp -R "$src" "$d/$SKILL"
  rm -rf "$d/$SKILL/evals"   # test cases are for developing the skill, not using it
  echo "installed for $t -> $d/$SKILL"
done

# Python packages the scripts need
py=""
for c in python3 python; do
  # skip stubs such as the Windows Store "python3" alias, which exist but don't run
  if command -v "$c" >/dev/null 2>&1 && "$c" -c "import sys" >/dev/null 2>&1; then py="$c"; break; fi
done
if [ -n "$py" ]; then
  "$py" -m pip install --quiet --user python-docx pythainlp \
    && echo "python packages ready: python-docx, pythainlp" \
    || echo "warning: run 'pip install python-docx pythainlp' yourself"
else
  echo "warning: Python not found - install Python 3, then: pip install python-docx pythainlp"
fi
echo "Done. Restart the app (or start a new session) so it picks up the skill."
