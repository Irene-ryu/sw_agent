#!/bin/bash

# 커밋 메시지에 현재 날짜와 시간을 포함
commit_msg="auto: $(date '+%Y-%m-%d %H:%M:%S') update"

# 1. 변경사항 스테이징
git add .

# 2. 커밋 (변경 사항이 있을 때만)
if ! git diff --cached --quiet; then
  git commit -m "$commit_msg"
else
  echo "No changes to commit."
fi

# 3. 원격 저장소에 push
git push origin main

