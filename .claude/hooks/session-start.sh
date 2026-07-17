#!/bin/bash
# SessionStart hook: keep CLAUDE.md in sync with the repo's default branch.
#
# Older / resumed chats run in a container that was cloned before newer rule
# changes were merged. This hook refreshes CLAUDE.md from the default branch at
# session start so every session follows the latest standing instructions
# without anyone having to fetch it by hand.
set -euo pipefail

# Only run in Claude Code on the web (remote sessions).
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-.}"

# Nothing to do if this isn't a git repo with an 'origin' remote.
git rev-parse --is-inside-work-tree >/dev/null 2>&1 || exit 0
git remote get-url origin >/dev/null 2>&1 || exit 0

# Determine origin's default branch; fall back to the known default.
DEFAULT_BRANCH="$(git symbolic-ref --quiet --short refs/remotes/origin/HEAD 2>/dev/null | sed 's#^origin/##' || true)"
DEFAULT_BRANCH="${DEFAULT_BRANCH:-claude/optimistic-euler-5lx59e}"

# Fetch just the default branch. Never fail the session on network hiccups.
git fetch --quiet origin "$DEFAULT_BRANCH" 2>/dev/null || exit 0

# Refresh CLAUDE.md from the default branch only when it actually differs.
if git cat-file -e "origin/${DEFAULT_BRANCH}:CLAUDE.md" 2>/dev/null; then
  remote_content="$(git show "origin/${DEFAULT_BRANCH}:CLAUDE.md" 2>/dev/null)"
  if [ ! -f CLAUDE.md ] || [ "$remote_content" != "$(cat CLAUDE.md 2>/dev/null)" ]; then
    printf '%s\n' "$remote_content" > CLAUDE.md
    echo "Synced CLAUDE.md from origin/${DEFAULT_BRANCH}"
  fi
fi

exit 0
