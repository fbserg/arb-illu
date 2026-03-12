#!/usr/bin/env bash
set -euo pipefail

INPUT=$(cat)
CMD=$(echo "$INPUT" | python3 -c \
  "import sys,json; print(json.load(sys.stdin).get('tool_input',{}).get('command',''))" \
  2>/dev/null || echo "")

# Only intercept git push
if ! echo "$CMD" | grep -q 'git push'; then
    exit 0
fi

HEAD=$(git rev-parse HEAD 2>/dev/null || echo "unknown")
FLAG="/tmp/.pre_push_simplify_${HEAD}"

if [ -f "$FLAG" ]; then
    rm -f "$FLAG"
    exit 0
fi

touch "$FLAG"
echo "Run /simplify before pushing." >&2
exit 2
