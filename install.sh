#!/usr/bin/env bash
# ppt-skills installer — symlinks pptx & ppt-direct into ~/.claude/skills/
# Two modes:
#   1. Run inside a cloned repo:  ./install.sh
#   2. Piped via curl:            curl -fsSL https://raw.githubusercontent.com/xu-jin-cs/ppt-skills/main/install.sh | bash
set -euo pipefail

REPO_TARBALL="https://codeload.github.com/xu-jin-cs/ppt-skills/tar.gz/refs/heads/main"
SKILLS=(pptx ppt-direct)

# --- locate repo root -------------------------------------------------------
if [ -n "${BASH_SOURCE[0]:-}" ] && [ -f "${BASH_SOURCE[0]}" ]; then
  REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
  CLEANUP_DIR=""
else
  # curl-pipe mode: download and extract the repo to a temp dir
  TMP="$(mktemp -d /tmp/ppt-skills.XXXXXX)"
  echo "Downloading ppt-skills to $TMP ..."
  curl -fsSL "$REPO_TARBALL" | tar -xz -C "$TMP"
  REPO_ROOT="$(echo "$TMP"/ppt-skills-*)"
  CLEANUP_DIR="$TMP"
fi

for s in "${SKILLS[@]}"; do
  [ -d "$REPO_ROOT/$s" ] || { echo "error: $REPO_ROOT/$s not found (bad repo layout?)" >&2; exit 1; }
done

# --- symlink into ~/.claude/skills ------------------------------------------
TARGET_DIR="$HOME/.claude/skills"
mkdir -p "$TARGET_DIR"

for s in "${SKILLS[@]}"; do
  dest="$TARGET_DIR/$s"
  if [ -e "$dest" ] || [ -L "$dest" ]; then
    echo "skip: $dest already exists (not overwriting)"
  else
    ln -s "$REPO_ROOT/$s" "$dest"
    echo "linked: $dest -> $REPO_ROOT/$s"
  fi
done

if [ -n "$CLEANUP_DIR" ]; then
  echo
  echo "note: curl-pipe mode installed from a temp dir ($REPO_ROOT)."
  echo "      The symlinks point there — do NOT delete it, or clone the repo"
  echo "      and re-run ./install.sh from the clone for a permanent location."
fi

cat <<'EOF'

Done. The skills are now available in Claude Code as /pptx and /ppt-direct.

Using another agent? Symlink the two skill dirs into its skills directory:
  Codex:      ln -s <repo>/pptx <repo>/ppt-direct  into  ~/.codex/skills/
  Kimi Code:  ln -s <repo>/pptx <repo>/ppt-direct  into  ~/.agents/skills/

Dependencies: pip install python-pptx numpy Pillow PyMuPDF
              + LibreOffice (soffice) for template replication / render QC.
EOF
