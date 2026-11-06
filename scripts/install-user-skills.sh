#!/usr/bin/env bash
# 내 PC의 모든 Claude Code 세션(새 세션, --resume 한 세션)에 영업 스킬과 공통 규칙을 설치합니다.
set -euo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"
dst="$HOME/.claude"
mkdir -p "$dst/skills"
cp -r "$root/.claude/skills/." "$dst/skills/"
cp "$root/claude-config/sales-rules.md" "$dst/sales-rules.md"
touch "$dst/CLAUDE.md"
grep -qs '@sales-rules.md' "$dst/CLAUDE.md" || printf '\n@sales-rules.md\n' >> "$dst/CLAUDE.md"
echo "설치 완료: 스킬 $(ls "$root/.claude/skills" | tr '\n' ' ')+ 공통 규칙 sales-rules.md"
