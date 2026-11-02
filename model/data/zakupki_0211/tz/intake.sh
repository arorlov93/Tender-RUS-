#!/bin/bash
# копирует новые загрузки (по префиксам) в raw/ и распаковывает в x/<префикс>/
B=/home/user/Tender-RUS-/model/data/zakupki_0211/tz; U=/root/.claude/uploads/e981d7a6-b89f-5df6-824c-c1907918e938
for p in "$@"; do for f in $U/$p-*; do b=$(basename "$f"); [ -e "$B/raw/$b" ] || cp "$f" $B/raw/; d="$B/x/$p"; mkdir -p "$d"
  case "$f" in *.zip|*.ZIP|*.rar|*.RAR|*.7z) (cd "$d" && unar -q -f "$B/raw/$b" >/dev/null 2>&1 || 7z x -y "$B/raw/$b" >/dev/null);; *) cp "$B/raw/$b" "$d/";; esac; done; done
