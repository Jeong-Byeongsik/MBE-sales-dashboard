#!/usr/bin/env bash
# 이 레포의 .claude/skills 를 내 PC 전체 세션에서 쓰도록 ~/.claude/skills 에 설치합니다.
set -euo pipefail
src="$(cd "$(dirname "$0")/.." && pwd)/.claude/skills"
mkdir -p "$HOME/.claude/skills"
cp -r "$src"/. "$HOME/.claude/skills/"
echo "설치 완료: $(ls "$src" | tr '\n' ' ')"
